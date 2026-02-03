# models/schemas.py
from pydantic import BaseModel
from typing import List, Dict, Any, Optional


class BuildDocx(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    body_html: Optional[str] = None
    body_html_b64: Optional[str] = None
    dest_path: Optional[str] = None
    filename: Optional[str] = None


class BuildFromFAC(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    fac_general: List[Dict[str, Any]] = []
    fac_findings: List[Dict[str, Any]] = []
    fac_findings_text: List[Dict[str, Any]] = []
    fac_caps: List[Dict[str, Any]] = []
    federal_awards: List[Dict[str, Any]] = []
    dest_path: Optional[str] = None
    filename: Optional[str] = None


class BuildByReport(BaseModel):
    auditee_name: str
    ein: str
    audit_year: int
    report_id: str
    dest_path: Optional[str] = None
    only_flagged: bool = False
    max_refs: int = 15
    include_awards: bool = True


class BuildByReportTemplated(BuildByReport):
    auditor_name: Optional[str] = None
    fy_end_text: Optional[str] = None
    recipient_name: Optional[str] = None
    street_address: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    poc_name: Optional[str] = None
    poc_title: Optional[str] = None
    template_path: Optional[str] = None
    treasury_listings: Optional[List[str]] = None


class BuildAuto(BaseModel):
    # required
    auditee_name: str
    ein: str
    audit_year: int

    # options (all optional)
    dest_path: Optional[str] = None
    max_refs: int = 15
    only_flagged: bool = False
    include_awards: bool = True
    treasury_listings: Optional[List[str]] = None

    # header overrides (all optional)
    recipient_name: Optional[str] = None
    fy_end_text: Optional[str] = None
    auditor_name: Optional[str] = None
    street_address: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    poc_name: Optional[str] = None
    poc_title: Optional[str] = None

    # template & mappings (optional)
    template_path: Optional[str] = None
    aln_reference_xlsx: Optional[str] = None
    treasury_contact_email: Optional[str] = None

    # optional flags
    include_no_qc_line: bool = True
    include_no_cap_line: bool = False
