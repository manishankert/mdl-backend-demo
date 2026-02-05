#!/usr/bin/env python3
"""
Test script to verify FAC API data for City of Carson (EIN: 952513547)
and test the auditee_name retrieval logic.

Run with: python test_carson.py
"""

import os
import sys
import requests

# Add the project root to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Load .env file if python-dotenv is available
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # Manually load .env if dotenv not installed
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if os.path.exists(env_path):
        with open(env_path) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key.strip()] = value.strip()

from config import FAC_BASE, FAC_KEY


def fac_headers():
    key = os.getenv("FAC_API_KEY") or FAC_KEY
    if not key:
        print("ERROR: FAC_API_KEY not configured")
        print("Please set FAC_API_KEY environment variable")
        sys.exit(1)
    return {"X-Api-Key": key}


def test_carson_ein():
    """Test FAC API data for City of Carson"""

    ein = "952513547"  # City of Carson
    test_year = 2022   # Input year for test

    print(f"\n{'='*60}")
    print(f"Testing FAC API for City of Carson (EIN: {ein})")
    print(f"Input audit year: {test_year}")
    print(f"{'='*60}\n")

    # Step 1: Fetch ALL recent audit records
    print("Step 1: Fetching all recent audit records...")
    params = {
        "auditee_ein": f"eq.{ein}",
        "select": "report_id, audit_year, fac_accepted_date, auditee_name, auditee_contact_name, auditee_contact_title",
        "order": "audit_year.desc,fac_accepted_date.desc",
        "limit": 10
    }

    try:
        r = requests.get(f"{FAC_BASE}/general", headers=fac_headers(), params=params, timeout=20)
        r.raise_for_status()
        all_audits = r.json()
    except Exception as e:
        print(f"ERROR: Failed to fetch from FAC API: {e}")
        sys.exit(1)

    if not all_audits:
        print(f"ERROR: No FAC records found for EIN {ein}")
        sys.exit(1)

    # Display all available audits
    print(f"\nFound {len(all_audits)} audit records:\n")
    print(f"{'Year':<8} {'Accepted Date':<15} {'Auditee Name':<25} {'POC Name':<25} {'POC Title':<20}")
    print("-" * 95)

    for audit in all_audits:
        year = audit.get("audit_year", "N/A")
        date = audit.get("fac_accepted_date", "N/A")[:10] if audit.get("fac_accepted_date") else "N/A"
        name = audit.get("auditee_name", "(empty)")
        poc_name = audit.get("auditee_contact_name", "(empty)")
        poc_title = audit.get("auditee_contact_title", "(empty)")
        if not name or not name.strip():
            name = "(empty)"
        if not poc_name or not poc_name.strip():
            poc_name = "(empty)"
        if not poc_title or not poc_title.strip():
            poc_title = "(empty)"
        print(f"{year:<8} {date:<15} {name:<25} {poc_name:<25} {poc_title:<20}")

    # Step 2: Find latest record with valid auditee_name (NEW LOGIC)
    print(f"\n{'='*60}")
    print("Step 2: Finding latest record with valid auditee_name...")
    print(f"{'='*60}\n")

    gen_latest = None
    auditee_name_from_latest = ""
    poc_name_from_latest = ""
    poc_title_from_latest = ""

    for audit_record in all_audits:
        candidate_name = (audit_record.get("auditee_name") or "").strip()
        if candidate_name:
            gen_latest = audit_record
            auditee_name_from_latest = candidate_name
            poc_name_from_latest = (audit_record.get("auditee_contact_name") or "").strip()
            poc_title_from_latest = (audit_record.get("auditee_contact_title") or "").strip()
            break

    if gen_latest:
        latest_year = gen_latest.get("audit_year")
        print(f"✓ Found latest record with auditee_name:")
        print(f"  Year: {latest_year}")
        print(f"  Auditee Name: {auditee_name_from_latest}")
        print(f"  POC Name: {poc_name_from_latest or '(empty)'}")
        print(f"  POC Title: {poc_title_from_latest or '(empty)'}")
    else:
        print("✗ No record found with valid auditee_name!")

    # Step 3: Fetch INPUT year data
    print(f"\n{'='*60}")
    print(f"Step 3: Fetching input year ({test_year}) data...")
    print(f"{'='*60}\n")

    params_input = {
        "audit_year": f"eq.{test_year}",
        "auditee_ein": f"eq.{ein}",
        "select": "report_id, fac_accepted_date, auditee_name, auditee_contact_name, auditee_contact_title",
        "order": "fac_accepted_date.desc",
        "limit": 1
    }

    r = requests.get(f"{FAC_BASE}/general", headers=fac_headers(), params=params_input, timeout=20)
    r.raise_for_status()
    gen_input = r.json()

    if gen_input:
        input_year_name = gen_input[0].get("auditee_name", "(empty)")
        input_year_poc = gen_input[0].get("auditee_contact_name", "(empty)")
        input_year_poc_title = gen_input[0].get("auditee_contact_title", "(empty)")
        print(f"Input year ({test_year}) auditee_name: {input_year_name}")
        print(f"Input year ({test_year}) POC: {input_year_poc} ({input_year_poc_title})")
    else:
        print(f"No record found for input year {test_year}")

    # Step 4: Summary - which name and POC will be used?
    print(f"\n{'='*60}")
    print("SUMMARY: AUDITEE NAME & POC SOURCE VERIFICATION")
    print(f"{'='*60}\n")

    if gen_latest and gen_input:
        print(f"  AUDITEE NAME:")
        print(f"    Input year ({test_year}): {gen_input[0].get('auditee_name', '(empty)')}")
        print(f"    Latest year ({latest_year}): {auditee_name_from_latest}")
        print(f"    ➜ USING (from latest year {latest_year}): {auditee_name_from_latest}")

        if gen_input[0].get('auditee_name') != auditee_name_from_latest:
            print(f"    ⚠️  Names are DIFFERENT between years!")
        else:
            print(f"    ✓ Names are the same between years")

        print(f"\n  POINT OF CONTACT:")
        print(f"    Input year ({test_year}): {gen_input[0].get('auditee_contact_name', '(empty)')} ({gen_input[0].get('auditee_contact_title', '(empty)')})")
        print(f"    Latest year ({latest_year}): {poc_name_from_latest or '(empty)'} ({poc_title_from_latest or '(empty)'})")
        print(f"    ➜ USING (from latest year {latest_year}): {poc_name_from_latest or '(empty)'} ({poc_title_from_latest or '(empty)'})")

        input_poc = gen_input[0].get('auditee_contact_name', '')
        if input_poc != poc_name_from_latest:
            print(f"    ⚠️  POC names are DIFFERENT between years!")
        else:
            print(f"    ✓ POC names are the same between years")

    print(f"\n{'='*60}")
    print("Test completed successfully!")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    test_carson_ein()
