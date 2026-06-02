"""
Google Sheets helper utilizing Google Apps Script API.
Reads/writes data to the spreadsheet completely via HTTP requests to an Apps Script Web App,
eliminating the need for Google Cloud project service accounts, credentials.json, and complicated setup.
Also handles automatic header matching and bulk multi-tab data splitting.
"""

import requests
from typing import Dict, List, Optional

# human-readable label mapping
TABLES = {
    "Demographics": "Demographics",
    "DMP": "Detection Method & Presentation",
    "RMD": "Reproductive & Menstrual Data",
    "FamilyPersonalHistory": "Family & Personal History",
    "Pathology": "Pathology",
    "Followup": "Treatment & Follow-up",
    "Samples": "Blood & Tissue Samples",
    "Analysis": "Analysis",
}

FILTER_COLUMNS = {
    "FamilyPersonalHistory": ["Family_history_of_Ca_Breast", "Multiple_neoplasia"],
    "Pathology": ["Type_of_tumor", "Grade", "Lympho_infiltrate", "Desmoplasia", "Adj_breast", "DCIS"],
    "RMD": ["Menopausal_status"],
    "DMP": ["Clinical_staging_simplified"],
}


class SheetsDB:
    """Read/write interface communicating with Google Sheets through Apps Script Web App."""

    def __init__(self, script_url: str):
        self.script_url = script_url

    def _get(self, params: dict) -> dict:
        """Helper to send GET requests to Google Apps Script."""
        try:
            resp = requests.get(self.script_url, params=params, timeout=30)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            return {"success": False, "message": f"HTTP GET error: {str(e)}"}

    def _post(self, payload: dict) -> dict:
        """Helper to send POST requests to Google Apps Script."""
        try:
            resp = requests.post(self.script_url, json=payload, timeout=90)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            return {"success": False, "message": f"HTTP POST error: {str(e)}"}

    # ── Read methods ──────────────────────────────────────────────────

    def get_sheet_headers(self) -> Dict[str, List[str]]:
        """Fetch headers from all worksheets to map uploads automatically."""
        res = self._get({"action": "getHeaders"})
        if res.get("success"):
            return res.get("headers", {})
        return {tab: [] for tab in TABLES}

    def joined_data(self, filters: Optional[dict] = None) -> List[dict]:
        """Fetch consolidated patients joined by Patient_ID with optional filters."""
        params = {"action": "getJoinedData"}
        if filters:
            for k, v in filters.items():
                if v:
                    params[k] = ",".join(v)
        res = self._get(params)
        if res.get("success"):
            return res.get("data", [])
        raise Exception(res.get("message", "Failed to retrieve joined data"))

    def filter_values(self) -> Dict[str, List[str]]:
        """Fetch sorted unique values for all criteria filter sidebars."""
        res = self._get({"action": "getFilters"})
        if res.get("success"):
            return res.get("filters", {})
        return {}

    def total_patients(self) -> int:
        """Return count of distinct Patient_IDs in demographics."""
        try:
            data = self.joined_data()
            return len(data)
        except Exception:
            return 0

    def get_authorized_users(self) -> List[dict]:
        """Fetch list of username-password pairs from AuthorizedUsers sheet tab."""
        res = self._get({"action": "getAuthorizedUsers"})
        if res.get("success"):
            return res.get("users", [])
        return []

    def all_columns(self) -> List[str]:
        """Return all merged columns across all sheets (used for view table headers)."""
        cols = []
        seen = set()
        headers = self.get_sheet_headers()
        for tab in TABLES:
            for col in headers.get(tab, []):
                if col and col not in seen:
                    cols.append(col)
                    seen.add(col)
        # Ensure Patient_ID is the first column
        if "Patient_ID" in cols:
            cols.remove("Patient_ID")
            cols.insert(0, "Patient_ID")
        return cols

    # ── Upload auto-mapping & writing ─────────────────────────────────

    def detect_best_tab(self, excel_headers: List[str]) -> str:
        """
        Compares Excel headers against existing Google Sheet tabs' headers
        and automatically maps the upload to the sheet tab with the highest match.
        Defaults to 'Demographics' if no overlap is found.
        """
        # Standardize matching set
        excel_set = set()
        for h in excel_headers:
            if not h:
                continue
            h_std = h.replace(" ", "_").replace("-", "_")
            if h_std.lower() not in ("patient_id", "patientid", "id_patient", "pat_id"):
                excel_set.add(h_std.lower())

        if not excel_set:
            return "Demographics"

        all_headers = self.get_sheet_headers()
        best_tab = "Demographics"
        max_overlap = -1

        for tab in TABLES:
            tab_headers = all_headers.get(tab, [])
            tab_set = set()
            for h in tab_headers:
                if not h:
                    continue
                h_std = h.replace(" ", "_").replace("-", "_")
                if h_std.lower() not in ("patient_id", "patientid", "id_patient", "pat_id"):
                    tab_set.add(h_std.lower())

            overlap = len(excel_set.intersection(tab_set))
            if overlap > max_overlap:
                max_overlap = overlap
                best_tab = tab

        # If zero overlap exists across all tables, guess based on specific keywords
        if max_overlap <= 0:
            lower_headers = {h.lower() for h in excel_headers}
            if any("tumor" in h or "grade" in h or "dcis" in h for h in lower_headers):
                return "Pathology"
            if any("menopaus" in h or "pregnan" in h or "abort" in h for h in lower_headers):
                return "RMD"
            if any("stage" in h or "size" in h or "nod" in h for h in lower_headers):
                return "DMP"
            if any("fam" in h or "breast" in h for h in lower_headers):
                return "FamilyPersonalHistory"
            if any("treat" in h or "therap" in h or "recur" in h for h in lower_headers):
                return "Followup"
            if any("sample" in h or "blood" in h or "tissu" in h for h in lower_headers):
                return "Samples"
            if any("analy" in h or "gene" in h or "mutat" in h for h in lower_headers):
                return "Analysis"

        return best_tab

    def bulk_upload(self, tables_payload: dict) -> dict:
        """Sends all tables' split records to Apps Script Web App for bulk atomic updates."""
        payload = {
            "action": "bulk_upload",
            "tables": tables_payload
        }
        res = self._post(payload)
        if res.get("success"):
            return {
                "results": res.get("results", {}),
                "new_columns": res.get("new_columns", {}),
                "warnings": res.get("warnings", [])
            }
        raise Exception(res.get("message", "Bulk upload API request failed."))

    def upsert_rows(self, tab_name: str, rows_data: List[dict], excel_headers: List[str]) -> dict:
        """Post rows to Google Apps Script for automated sheets upsert/column expansion."""
        payload = {
            "action": "upsert",
            "tabName": tab_name,
            "headers": excel_headers,
            "rows": rows_data
        }
        res = self._post(payload)
        if res.get("success"):
            return {
                "updated": res.get("stats", {}).get("updated", 0),
                "inserted": res.get("stats", {}).get("inserted", 0),
                "skipped": res.get("stats", {}).get("skipped", 0),
                "warnings": res.get("warnings", []),
                "new_columns": res.get("new_columns", [])
            }
        raise Exception(res.get("message", "API Request failed"))
