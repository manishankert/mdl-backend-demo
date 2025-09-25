import os, requests

KEY = "4LlIXwxP7dHQa5fMTPj8i6Gi0RSNNoJ3L9zaC4KV"
BASE = "https://api.fac.gov"
H = {"X-Api-Key": KEY}

def first_with_zero(year=2024):
    r = requests.get(f"{BASE}/general", headers=H, params={
        "audit_year": f"eq.{year}",
        "select": "report_id,auditee_name,auditee_ein",
        "order": "fac_accepted_date.desc",
        "limit": "300"
    }); r.raise_for_status()
    for row in r.json():
        rid = row["report_id"]
        f = requests.get(f"{BASE}/findings", headers=H, params={
            "report_id": f"eq.{rid}", "select": "reference_number", "limit": "1"
        }); f.raise_for_status()
        if len(f.json()) == 0:
            return {"auditee_name": row["auditee_name"], "ein": row["auditee_ein"], "audit_year": year, "report_id": rid}

def first_with_small(year=2024, max_findings=3):
    r = requests.get(f"{BASE}/general", headers=H, params={
        "audit_year": f"eq.{year}",
        "select": "report_id,auditee_name,auditee_ein",
        "order": "fac_accepted_date.desc",
        "limit": "300"
    }); r.raise_for_status()
    for row in r.json():
        rid = row["report_id"]
        # Probe with limit=4
        probe = requests.get(f"{BASE}/findings", headers=H, params={
            "report_id": f"eq.{rid}", "select": "reference_number", "order":"reference_number.asc", "limit": "4"
        }); probe.raise_for_status()
        if len(probe.json()) < 4:
            # Confirm exact count for this single report
            full = requests.get(f"{BASE}/findings", headers=H, params={
                "report_id": f"eq.{rid}", "select": "reference_number", "limit": "200"
            }); full.raise_for_status()
            cnt = len(full.json())
            if 1 <= cnt <= max_findings:
                return {"auditee_name": row["auditee_name"], "ein": row["auditee_ein"], "audit_year": year, "report_id": rid, "count": cnt}

print("ZERO:", first_with_zero())
print("SMALL:", first_with_small())

