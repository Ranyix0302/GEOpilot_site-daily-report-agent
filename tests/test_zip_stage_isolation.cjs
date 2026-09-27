const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

class Element {
  constructor() {
    this.textContent = ''; this.style = {}; this.disabled = false;
    this.children = new Map(); this.listeners = {};
    const classes = new Set();
    this.classList = {
      add: (...names) => names.forEach(n => classes.add(n)),
      remove: (...names) => names.forEach(n => classes.delete(n)),
      contains: name => classes.has(name),
      toggle: (name, force) => {
        const enabled = force === undefined ? !classes.has(name) : force;
        if (enabled) classes.add(name); else classes.delete(name);
      },
    };
  }
  querySelector(selector) {
    if (!this.children.has(selector)) this.children.set(selector, new Element());
    return this.children.get(selector);
  }
  addEventListener(type, callback) {
    (this.listeners[type] ||= []).push(callback);
  }
}

const nodes = new Map();
const node = selector => {
  if (!nodes.has(selector)) nodes.set(selector, new Element());
  return nodes.get(selector);
};
const context = vm.createContext({
  document: {querySelector: node, querySelectorAll: () => [], addEventListener: () => {}, dispatchEvent: () => {}, documentElement: {lang: 'en'}},
  setTimeout: callback => callback(), console, FormData: class {append() {}}, confirm: () => true,
  localStorage: {getItem: () => null, setItem: () => {}},
  CustomEvent: class {constructor(type, options) {this.type=type;this.detail=options?.detail}},
});
const i18nSource = fs.readFileSync(path.join(__dirname, '../static/i18n.js'), 'utf8');
vm.runInContext(i18nSource, context);
const source = fs.readFileSync(path.join(__dirname, '../static/app.js'), 'utf8');
vm.runInContext(source.slice(0, source.lastIndexOf('init().catch')), context);
const run = text => vm.runInContext(text, context);
run('renderStageStatus=()=>{};toast=()=>{};updateInterface=()=>{};askConfirmation=async()=>true;');

async function main() {
  assert.equal(run("t('dropZip')"), 'Drop a ZIP file here, or click to select');
  run("pendingZips.operation={name:'kept.zip',size:99};setLanguage('zh')");
  assert.equal(run("t('dropZip')"), '把 ZIP 文件拖到这里，或点击选择文件');
  assert.equal(run('pendingZips.operation.name'), 'kept.zip');
  run("setLanguage('en');pendingZips.operation=null");

  run("current='operation';pendingZips.operation={name:'op.zip',size:1000};zipStates.operation={status:'running',completed:3,total:10,message:'op progress'};renderZipSelection()");
  assert.equal(node('#zipDropTitle').textContent, 'op.zip');
  assert.equal(node('#zipProgressBar').style.width, '30%');
  run("current='silo';renderZipSelection()");
  assert.equal(node('#selectedZipName').textContent, '');
  assert.equal(node('#zipDropTitle').textContent, 'Drop a ZIP file here, or click to select');
  assert.equal(node('#zipFileReview').classList.contains('hidden-panel'), true);
  assert.equal(node('#zipProgressWrap').classList.contains('hidden-panel'), true);
  assert.equal(node('#zipProgressBar').style.width, '0%');

  run("pendingZips.silo={name:'silo.zip',size:2000};zipStates.silo={status:'running',completed:1,total:5,message:'silo progress'};renderZipSelection()");
  assert.equal(node('#zipDropTitle').textContent, 'silo.zip');
  assert.equal(node('#zipProgressBar').style.width, '20%');
  const before = node('#selectedZipSize').textContent;
  run("jsonFetch=async()=>({status:'completed',completed:10,total:10,message:'op finished',result:{count:10}})");
  await run("waitForMediaJob('op-job','operation',zipStates.operation)");
  assert.equal(node('#selectedZipSize').textContent, before);
  assert.equal(node('#zipDropTitle').textContent, 'silo.zip');
  assert.equal(node('#zipProgressBar').style.width, '20%');
  assert.equal(run('zipStates.operation.completed'), 10);
  assert.equal(run('zipStates.silo.completed'), 1);

  run("jsonFetch=async()=>({status:'failed',completed:2,total:10,error:'op failed',message:'failure'})");
  await assert.rejects(run("waitForMediaJob('op-job','operation',zipStates.operation)"), /op failed/);
  assert.equal(node('#selectedZipSize').textContent, before);
  assert.equal(run('zipStates.silo.status'), 'running');

  const stale = run('zipStates.operation');
  run('zipStates.operation=newZipState()');
  context.stale = stale;
  assert.equal(await run("waitForMediaJob('old-job','operation',stale)"), null);

  // A completion arriving while viewing the other stage must not enable its
  // download button or replace its upload controls.
  run("current='operation';zipStates.operation=newZipState();operationReady=false;runId='run';jsonFetch=async url=>url.endsWith('/operation')?{job_id:'job'}:{rows:[]};waitForMediaJob=()=>new Promise(resolve=>{finishJob=resolve})");
  const task = node('#confirmZip').listeners.click[0]();
  await new Promise(resolve => setImmediate(resolve));
  run("current='silo';renderZipSelection()");
  node('#downloadStage').disabled = true;
  run("finishJob({count:10,message:'done'})");
  await task;
  assert.equal(node('#zipDropTitle').textContent, 'silo.zip');
  assert.equal(node('#downloadStage').disabled, true);
  assert.equal(run('operationReady'), true);
  assert.equal(run('siloReady'), false);
  console.log('PASS: empty stage reset, independent progress, hidden completion/failure, stale run protection.');
}
main().catch(error => { console.error(error); process.exitCode = 1; });
