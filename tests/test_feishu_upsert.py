import unittest
from unittest.mock import patch

from agent.app import RUNS, app, new_run


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


class FakeRequests:
    def __init__(self, records):
        self.records = records
        self.posts = []

    def get(self, url, **kwargs):
        return FakeResponse({"code": 0, "data": {"items": self.records, "has_more": False}})

    def post(self, url, **kwargs):
        self.posts.append((url, kwargs.get("json") or {}))
        records = (kwargs.get("json") or {}).get("records") or []
        return FakeResponse({"code": 0, "data": {"records": [{} for _ in records]}})


def construction_row(pile, equipment):
    return {
        "Date": "2026-09-14", "Pile NO.": pile, "Equipment NO.": equipment,
        "Drilling Start Time": "08:00", "Drilling Complete Time": "08:40",
        "Point Complete Time": "09:00", "Operator": "", "Cement Content": "",
        "_start_dt": "2026-09-14T08:00", "_drill_dt": "2026-09-14T08:40",
        "_complete_dt": "2026-09-14T09:00",
    }


class FeishuUpsertTests(unittest.TestCase):
    def setUp(self):
        self.client = app.test_client()
        self.run = new_run("2026-09-14")
        self.run["construction"] = {
            "DCM-1": [construction_row("E-E-0101", "DCM-1")],
            "DCM-2": [construction_row("E-E-0102", "DCM-2")],
        }

    def tearDown(self):
        RUNS.pop(self.run["run_id"], None)

    @patch("agent.app._operator_field", return_value=None)
    def test_existing_pile_updates_and_new_pile_creates(self, _operator):
        fake = FakeRequests([{"record_id": "rec-existing", "fields": {"Pile NO.": "E-E-0101"}}])
        with patch("agent.app._feishu_context", return_value=(fake, "https://example.test", "token", "app", "table")):
            response = self.client.post(f"/api/runs/{self.run['run_id']}/write-feishu")
        self.assertEqual(response.status_code, 200)
        body = response.get_json()
        self.assertEqual((body["created"], body["updated"]), (1, 1))
        create = next(payload for url, payload in fake.posts if url.endswith("batch_create"))
        update = next(payload for url, payload in fake.posts if url.endswith("batch_update"))
        self.assertEqual(create["records"][0]["fields"]["Pile NO."], "E-E-0102")
        self.assertEqual(update["records"][0]["record_id"], "rec-existing")

    @patch("agent.app._operator_field", return_value=None)
    def test_duplicate_existing_pile_stops_without_writing(self, _operator):
        fake = FakeRequests([
            {"record_id": "rec-a", "fields": {"Pile NO.": "E-E-0101"}},
            {"record_id": "rec-b", "fields": {"Pile NO.": "E-E-0101"}},
        ])
        with patch("agent.app._feishu_context", return_value=(fake, "https://example.test", "token", "app", "table")):
            response = self.client.post(f"/api/runs/{self.run['run_id']}/write-feishu")
        self.assertEqual(response.status_code, 409)
        self.assertEqual(fake.posts, [])

    @patch("agent.app._operator_field", return_value=None)
    def test_operation_only_can_create_partial_record(self, _operator):
        self.run["construction"] = {}
        self.run["operation"] = [{
            "Date": "2026-09-14", "Pile NO.": "E-E-0201", "Equipment NO.": "",
            "Drilling Start Time": "", "Drilling Complete Time": "", "Point Complete Time": "",
            "Operator": "Raj Pillai", "Cement Content": "",
        }]
        fake = FakeRequests([])
        with patch("agent.app._feishu_context", return_value=(fake, "https://example.test", "token", "app", "table")):
            response = self.client.post(f"/api/runs/{self.run['run_id']}/write-feishu")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["created"], 1)
        create = next(payload for url, payload in fake.posts if url.endswith("batch_create"))
        fields = create["records"][0]["fields"]
        self.assertEqual(fields["Pile NO."], "E-E-0201")
        self.assertEqual(fields["Operator"], "Raj Pillai")
        self.assertNotIn("Equipment NO.", fields)

    @patch("agent.app._operator_field", return_value=None)
    def test_two_partial_groups_merge_into_one_pile(self, _operator):
        self.run["construction"] = {}
        self.run["operation"] = [{
            "Date": "2026-09-14", "Pile NO.": "E-E-0301", "Equipment NO.": "",
            "Drilling Start Time": "", "Drilling Complete Time": "", "Point Complete Time": "",
            "Operator": "Dev Sharma", "Cement Content": "",
        }]
        self.run["silo"] = [{
            "Date": "2026-09-14", "Pile NO.": "E-E-0301", "Equipment NO.": "",
            "Drilling Start Time": "", "Drilling Complete Time": "", "Point Complete Time": "",
            "Operator": "", "Cement Content": 14,
        }]
        fake = FakeRequests([])
        with patch("agent.app._feishu_context", return_value=(fake, "https://example.test", "token", "app", "table")):
            response = self.client.post(f"/api/runs/{self.run['run_id']}/write-feishu")
        self.assertEqual(response.status_code, 200)
        create = next(payload for url, payload in fake.posts if url.endswith("batch_create"))
        self.assertEqual(len(create["records"]), 1)
        fields = create["records"][0]["fields"]
        self.assertEqual(fields["Operator"], "Dev Sharma")
        self.assertEqual(fields["Cement Content"], 14.0)


if __name__ == "__main__":
    unittest.main()
