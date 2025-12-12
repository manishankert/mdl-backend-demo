"""
MDL Generator - Complete Implementation Based on Template Comments

Requirements extracted from 40 comments in the template document:

FIELD REQUIREMENTS:
1. Current Date - Format: "Month DD, YYYY" (e.g., December 12, 2025)
2. Recipient Name - Standard case (not ALL CAPS), from [auditee_name]
   - Address block: WITHOUT "The" prefix
   - Narrative: WITH "The" prefix
3. EIN - Format: XX-XXXXXXX (with dash), from [auditee_ein]
4. Street Address - Standard case, from [auditee_address_line_1]
5. City, State, Zip - Standard case, from [auditee_city], [auditee_state], [auditee_zip]
6. Point of Contact - Standard case, from [auditee_contact_name], [auditee_contact_title]
7. Fiscal Year End Date - Format: "Month Day, Year", from [fy_end_date]
8. Auditor Name - Add "the" prefix in narrative, standard case

PLURALIZATION (based on total finding count):
- is/are, finding/findings, issue/issues, violates/violate, CAP/CAPs, corrective action/actions

TABLE REQUIREMENTS:
- One table per program (ALN)
- Tables in ALN order (21.023, then 21.026, then 21.027, etc.)
- Format: [ALN]/ [Program Name] ([Acronym])

FINDING REQUIREMENTS:
- Finding Number: Reference number from SF-SAC
- Repeat Finding: If [is_repeat_finding]=Y, add "Repeat of [prior_finding_ref_numbers]"
- Compliance Type: Mapped from [type_requirement] letter code
- Finding Summary: Matched from standardized list
- Audit Finding Determination: "Sustained"
- Questioned Cost: "Questioned Cost:\nNone\nDisallowed Cost:\nNone"
- CAP Determination: "Accepted" if CAP exists, else "Not Applicable"
"""

import os
import re
import logging
from io import BytesIO
from datetime import datetime, timedelta
from typing import Optional, List, Dict, Any, Tuple
from dataclasses import dataclass, field

import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


# ============================================================
# CONFIGURATION
# ============================================================

class Config:
    FAC_API_BASE = os.getenv("FAC_API_BASE", "https://api.fac.gov")
    FAC_API_KEY = os.getenv("FAC_API_KEY", "")
    OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
    AZURE_CONN_STR = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "")
    AZURE_CONTAINER = os.getenv("AZURE_BLOB_CONTAINER", "mdl-output")
    LOCAL_SAVE_DIR = os.getenv("LOCAL_SAVE_DIR", "./_out")
    PUBLIC_BASE_URL = os.getenv("PUBLIC_BASE_URL", "http://localhost:8000")
    MDL_TEMPLATE_PATH = os.getenv("MDL_TEMPLATE_PATH", "templates/MDL_Template.docx")
    TREASURY_EMAIL = os.getenv("TREASURY_CONTACT_EMAIL", "ORP_SingleAudits@treasury.gov")


# ============================================================
# STATIC DATA - From Comments & mdl_helpers.py
# ============================================================

# Treasury Programs (Comment #26, #27): ALN -> (Full Name, Acronym)
TREASURY_PROGRAMS = {
    "21.019": ("Coronavirus Relief Fund", "CRF"),
    "21.023": ("Emergency Rental Assistance Program", "ERA"),
    "21.026": ("Homeowner Assistance Fund", "HAF"),
    "21.027": ("Coronavirus State and Local Fiscal Recovery Funds", "SLFRF"),
    "21.029": ("Capital Projects Fund", "CPF"),
    "21.031": ("State Small Business Credit Initiative", "SSBCI"),
    "21.032": ("Local Assistance and Tribal Consistency Fund", "LATCF"),
}

# Compliance Types (Comment #32, #34): Letter -> Full Description
COMPLIANCE_TYPES = {
    "A": "Activities allowed or unallowed",
    "B": "Allowable costs/cost principles",
    "C": "Cash management",
    "E": "Eligibility",
    "F": "Equipment and real property management",
    "G": "Matching, level of effort, earmarking",
    "H": "Period of performance (or availability) of Federal funds",
    "I": "Procurement and suspension and debarment",
    "J": "Program income",
    "L": "Reporting",
    "M": "Subrecipient monitoring",
    "N": "Special tests and provisions",
    "P": "Other",
}

# Standard Finding Summaries (Comment #37, #38)
FINDING_SUMMARIES = [
    "Deficient Subrecipient Monitoring or Deficient Subaward",
    "Failure to file FFATA report for subawards",
    "Lack of evidence of competitive procurement",
    "Lack of evidence of suspension and debarment verification",
    "Lack of time and effort documentation",
    "Failure to retain adequate supporting documentation",
    "Inaccurate Treasury Reporting",
    "Lack of Eligibility Support",
    "Unallowable expenditures due to being incurred outside of period of performance",
    "Lack of Written Policies and/or Procedures - Management of Federal Funds",
    "Lack of Segregation of Duties",
    "Lack of Internal Controls - Grants Management",
]

# Keywords for classification fallback
FINDING_KEYWORDS = {
    "subrecipient": "Deficient Subrecipient Monitoring or Deficient Subaward",
    "subaward": "Deficient Subrecipient Monitoring or Deficient Subaward",
    "sub-recipient": "Deficient Subrecipient Monitoring or Deficient Subaward",
    "ffata": "Failure to file FFATA report for subawards",
    "procurement": "Lack of evidence of competitive procurement",
    "competitive bid": "Lack of evidence of competitive procurement",
    "sole source": "Lack of evidence of competitive procurement",
    "bid": "Lack of evidence of competitive procurement",
    "suspension and debarment": "Lack of evidence of suspension and debarment verification",
    "sam.gov": "Lack of evidence of suspension and debarment verification",
    "debarment": "Lack of evidence of suspension and debarment verification",
    "suspended": "Lack of evidence of suspension and debarment verification",
    "debarred": "Lack of evidence of suspension and debarment verification",
    "time and effort": "Lack of time and effort documentation",
    "timesheet": "Lack of time and effort documentation",
    "labor cost": "Lack of time and effort documentation",
    "personnel": "Lack of time and effort documentation",
    "documentation": "Failure to retain adequate supporting documentation",
    "supporting documentation": "Failure to retain adequate supporting documentation",
    "records": "Failure to retain adequate supporting documentation",
    "treasury report": "Inaccurate Treasury Reporting",
    "quarterly report": "Inaccurate Treasury Reporting",
    "project and expenditure": "Inaccurate Treasury Reporting",
    "p&e report": "Inaccurate Treasury Reporting",
    "period of performance": "Unallowable expenditures due to being incurred outside of period of performance",
    "outside the period": "Unallowable expenditures due to being incurred outside of period of performance",
    "eligibility": "Lack of Eligibility Support",
    "eligible": "Lack of Eligibility Support",
    "internal control": "Lack of Internal Controls - Grants Management",
    "policies and procedures": "Lack of Written Policies and/or Procedures - Management of Federal Funds",
    "written policies": "Lack of Written Policies and/or Procedures - Management of Federal Funds",
    "segregation of duties": "Lack of Segregation of Duties",
    "segregation": "Lack of Segregation of Duties",
}


# ============================================================
# FORMATTING UTILITIES (Based on Comments #1, #5, #14, #16)
# ============================================================

def to_standard_case(name: str) -> str:
    """
    Convert ALL CAPS to Standard Case (Comment #1: "Everything must be in standard case")
    Preserves known acronyms like LLC, LLP, etc.
    """
    if not name:
        return ""
    name = name.strip()
    
    if not name.isupper():
        return name  # Already mixed case, preserve
    
    # Known acronyms to preserve
    acronyms = {"LLC", "LLP", "PC", "PA", "CPA", "USA", "US", "II", "III", "IV"}
    # Words to keep lowercase (except at start)
    lowercase_words = {"and", "or", "the", "of", "for", "to", "in", "on", "by", "with", "a", "an"}
    
    words = []
    for i, word in enumerate(name.split()):
        word_upper = word.upper()
        word_lower = word.lower()
        
        if word_upper in acronyms:
            words.append(word_upper)
        elif word_lower in lowercase_words and i > 0:
            words.append(word_lower)
        else:
            words.append(word.capitalize())
    
    return " ".join(words)


def format_ein(ein: str) -> str:
    """
    Format EIN as XX-XXXXXXX (Comment #5: "Must be XX-XXXXXXX format")
    """
    ein = (ein or "").replace("-", "").replace(" ", "").strip()
    if len(ein) == 9 and ein.isdigit():
        return f"{ein[:2]}-{ein[2:]}"
    return ein


def format_date(date_str: Optional[str]) -> str:
    """
    Format date as "Month Day, Year" (Comment #14: "Must be in [Month] [Day], [Year] format")
    Example: "June 30, 2024"
    """
    if not date_str:
        return ""
    try:
        # Try ISO format (2024-06-30)
        if "-" in date_str and len(date_str) >= 10:
            dt = datetime.fromisoformat(date_str.replace("Z", "+00:00").split("T")[0])
            return dt.strftime("%B %d, %Y")
        return date_str
    except:
        return date_str


def add_the_prefix(name: str) -> str:
    """
    Add "The" prefix if not present (Comment #16: "Must add 'The' before recipient name")
    """
    if not name:
        return ""
    name = name.strip()
    if name.lower().startswith("the "):
        return name
    return f"The {name}"


def add_the_lowercase_prefix(name: str) -> str:
    """
    Add "the" prefix (lowercase) for auditor name in narrative.
    """
    if not name:
        return ""
    name = name.strip()
    if name.lower().startswith("the "):
        return name
    return f"the {name}"


def get_current_date() -> str:
    """
    Get current date in required format (Comment #0: "Current Date")
    """
    return datetime.now().strftime("%B %d, %Y")


# ============================================================
# PLURALIZATION (Comments #19, #20, #40)
# ============================================================

def get_pluralization(count: int) -> Dict[str, str]:
    """
    Get singular/plural forms based on finding count.
    Comment #19: "Add S if plural, remove if singular finding"
    """
    singular = (count == 1)
    
    return {
        # For "[is/are]" placeholder
        "is/are": "is" if singular else "are",
        
        # For "[The]" placeholder before recipient name in body
        # Comment #16: Must add "The" before recipient name
        # But for singular, we may not want "The" (based on template context)
        "The": "" if singular else "The",
        
        # For "[the]" placeholder before auditor name  
        "the": "the",  # Always lowercase "the" before auditor
        
        # Additional pluralization helpers (if template uses these)
        "finding_s": "finding" if singular else "findings",
        "issue_s": "issue" if singular else "issues",
        "violate_s": "violates" if singular else "violate",
        "CAP_s": "CAP" if singular else "CAPs",
    }


# ============================================================
# FAC CLIENT
# ============================================================

class FACClient:
    """Federal Audit Clearinghouse API client."""
    
    def __init__(self, api_key: str = ""):
        self.base_url = Config.FAC_API_BASE
        self.session = requests.Session()
        self.session.headers["Accept"] = "application/json"
        if api_key:
            self.session.headers["X-Api-Key"] = api_key
    
    def _get(self, endpoint: str, params: dict) -> list:
        url = f"{self.base_url}/{endpoint}"
        try:
            resp = self.session.get(url, params=params, timeout=30)
            resp.raise_for_status()
            return resp.json()
        except requests.HTTPError as e:
            logger.error(f"FAC API error: {e}")
            return []
        except Exception as e:
            logger.error(f"FAC API error: {e}")
            return []
    
    def _or_param(self, field: str, values: List[str]) -> str:
        inner = ",".join([f"{field}.eq.{v}" for v in values])
        return f"({inner})"
    
    def find_report(self, ein: str, year: int) -> Optional[dict]:
        """Find most recent report for EIN/year."""
        logger.info(f"Searching FAC: EIN={ein}, year={year}")
        data = self._get("general", {
            "audit_year": f"eq.{year}",
            "auditee_ein": f"eq.{ein}",
            "select": "report_id,fac_accepted_date,auditee_name,auditee_address_line_1,"
                     "auditee_city,auditee_state,auditee_zip,auditor_firm_name,"
                     "fy_end_date,auditee_contact_name,auditee_contact_title",
            "order": "fac_accepted_date.desc",
            "limit": "1",
        })
        if data:
            logger.info(f"Found report: {data[0].get('report_id')}")
            return data[0]
        logger.warning("No report found")
        return None
    
    def get_findings(self, report_id: str, max_refs: int = 15, only_flagged: bool = False) -> List[dict]:
        """Get findings for a report."""
        params = {
            "report_id": f"eq.{report_id}",
            "select": "reference_number,award_reference,type_requirement,"
                     "is_material_weakness,is_significant_deficiency,is_questioned_costs,"
                     "is_modified_opinion,is_other_findings,is_other_matters,"
                     "is_repeat_finding,prior_finding_ref_numbers",
            "order": "reference_number.asc",
            "limit": str(max_refs),
        }
        if only_flagged:
            flagged = ["is_material_weakness", "is_significant_deficiency", "is_questioned_costs",
                      "is_modified_opinion", "is_other_findings", "is_other_matters", "is_repeat_finding"]
            params["or"] = "(" + ",".join(f"{f}.eq.true" for f in flagged) + ")"
        return self._get("findings", params) or []
    
    def get_findings_text(self, report_id: str, refs: List[str]) -> Dict[str, str]:
        """Get finding text by reference."""
        if not refs:
            return {}
        data = self._get("findings_text", {
            "report_id": f"eq.{report_id}",
            "select": "finding_ref_number,finding_text",
            "order": "finding_ref_number.asc",
            "limit": str(len(refs) + 5),
            "or": self._or_param("finding_ref_number", refs),
        }) or []
        return {d.get("finding_ref_number", ""): d.get("finding_text", "") for d in data}
    
    def get_caps(self, report_id: str, refs: List[str]) -> Dict[str, str]:
        """Get corrective action plans by reference."""
        if not refs:
            return {}
        data = self._get("corrective_action_plans", {
            "report_id": f"eq.{report_id}",
            "select": "finding_ref_number,planned_action",
            "order": "finding_ref_number.asc",
            "limit": str(len(refs) + 5),
            "or": self._or_param("finding_ref_number", refs),
        }) or []
        return {d.get("finding_ref_number", ""): d.get("planned_action", "") for d in data}
    
    def get_awards(self, report_id: str) -> List[dict]:
        """Get federal awards for a report."""
        return self._get("federal_awards", {
            "report_id": f"eq.{report_id}",
            "select": "award_reference,federal_program_name,assistance_listing",
            "order": "award_reference.asc",
            "limit": "200",
        }) or []


# ============================================================
# CLASSIFIER (Comments #37, #38)
# ============================================================

def classify_finding_openai(text: str, labels: List[str] = FINDING_SUMMARIES) -> Optional[str]:
    """Classify finding using OpenAI."""
    if not Config.OPENAI_API_KEY or not text:
        return None
    
    prompt = f"""Classify this audit finding into exactly one category.

Categories:
{chr(10).join(f"- {label}" for label in labels)}

Finding text (first 2000 chars):
{text[:2000]}

Reply with ONLY the category name from the list above, nothing else."""

    try:
        resp = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {Config.OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "gpt-4o-mini",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0,
            },
            timeout=15,
        )
        result = resp.json()
        answer = (result.get("choices", [{}])[0].get("message", {}).get("content") or "").strip()
        if answer in labels:
            return answer
        # Try partial match
        for label in labels:
            if label.lower() in answer.lower() or answer.lower() in label.lower():
                return label
    except Exception as e:
        logger.warning(f"OpenAI classification failed: {e}")
    return None


def classify_finding_keywords(text: str) -> Optional[str]:
    """Classify finding using keyword matching."""
    if not text:
        return None
    text_lower = text.lower()
    for keyword, summary in FINDING_KEYWORDS.items():
        if keyword in text_lower:
            return summary
    return None


def classify_finding(text: str) -> str:
    """
    Classify finding text (Comment #37: "Finding Summary to be selected from list")
    Tries OpenAI first, then keywords, then extracts first sentence.
    """
    if not text:
        return "Other"
    
    # Try OpenAI first
    result = classify_finding_openai(text)
    if result:
        logger.info(f"OpenAI classified: {result}")
        return result
    
    # Fallback to keywords
    result = classify_finding_keywords(text)
    if result:
        logger.info(f"Keyword classified: {result}")
        return result
    
    # Last resort: extract first meaningful sentence
    clean = re.sub(r"\s+", " ", text.strip())
    # Try to get first sentence
    match = re.match(r'^([^.!?]+[.!?])', clean)
    if match and len(match.group(1)) < 150:
        return match.group(1)
    
    return clean[:100] + "..." if len(clean) > 100 else clean


# ============================================================
# DOCUMENT HELPERS
# ============================================================

def add_table_borders(table):
    """Add borders to table."""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def add_shading(cell, hex_color: str = "D9D9D9"):
    """Add background shading to cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def normalize_aln(aln: str) -> str:
    """Normalize ALN to XX.XXX format."""
    aln = (aln or "").strip()
    if aln and "." not in aln and len(aln) >= 5:
        return f"{aln[:2]}.{aln[2:]}"
    return aln


# ============================================================
# MDL GENERATOR
# ============================================================

@dataclass
class Finding:
    """Single finding data."""
    finding_id: str
    compliance_type: str  # Full text like "Procurement and suspension and debarment"
    summary: str  # Matched summary like "Lack of evidence of suspension and debarment verification"
    audit_determination: str = "Sustained"
    questioned_cost: str = "Questioned Cost:\nNone\nDisallowed Cost:\nNone"
    cap_determination: str = "Accepted"
    is_repeat: bool = False
    prior_finding_ref: str = ""


@dataclass  
class Program:
    """Program with findings."""
    aln: str  # e.g., "21.027"
    name: str  # e.g., "Coronavirus State and Local Fiscal Recovery Funds"
    acronym: str  # e.g., "SLFRF"
    findings: List[Finding] = field(default_factory=list)
    
    @property
    def header(self) -> str:
        """Format: ALN/ Program Name (Acronym)"""
        return f"{self.aln}/ {self.name} ({self.acronym})"


class MDLGenerator:
    """
    MDL Generator using simple template replacement.
    Implements all requirements from template comments.
    """
    
    def __init__(self, template_path: str = ""):
        self.template_path = template_path or Config.MDL_TEMPLATE_PATH
        self.fac = FACClient(Config.FAC_API_KEY)
    
    def generate_from_fac(
        self,
        auditee_name: str,
        ein: str,
        audit_year: int,
        treasury_listings: Optional[List[str]] = None,
        max_refs: int = 15,
        only_flagged: bool = False,
        # Optional overrides
        recipient_name: Optional[str] = None,
        street_address: Optional[str] = None,
        city: Optional[str] = None,
        state: Optional[str] = None,
        zip_code: Optional[str] = None,
        poc_name: Optional[str] = None,
        poc_title: Optional[str] = None,
        auditor_name: Optional[str] = None,
        fiscal_year_end: Optional[str] = None,
        treasury_email: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Generate MDL from FAC data."""
        
        if not treasury_listings:
            treasury_listings = ["21.027", "21.023", "21.026"]
        
        try:
            # 1. Find report
            report = self.fac.find_report(ein, audit_year)
            if not report:
                return {"ok": False, "message": f"No FAC report found for EIN {ein} in {audit_year}"}
            
            report_id = report["report_id"]
            
            # 2. Fetch all data
            findings_raw = self.fac.get_findings(report_id, max_refs, only_flagged)
            refs = [f["reference_number"] for f in findings_raw if f.get("reference_number")]
            findings_text = self.fac.get_findings_text(report_id, refs)
            caps = self.fac.get_caps(report_id, refs)
            awards = self.fac.get_awards(report_id)
            
            # 3. Build award lookup (award_reference -> ALN)
            award_to_aln = {}
            for a in awards:
                ref = a.get("award_reference")
                if not ref:
                    continue
                aln = normalize_aln(a.get("assistance_listing") or "")
                if aln:
                    award_to_aln[ref] = aln
            
            # 4. Group findings by program (Comment #21, #22: separate table per program)
            programs_map: Dict[str, Program] = {}
            
            for f in findings_raw:
                ref = f.get("reference_number")
                if not ref:
                    continue
                
                # Get ALN from award
                award_ref = f.get("award_reference") or ""
                aln = award_to_aln.get(award_ref, "")
                
                # Skip non-Treasury programs
                if aln not in treasury_listings:
                    continue
                
                # Get program info (Comment #26, #27)
                if aln in TREASURY_PROGRAMS:
                    prog_name, prog_acronym = TREASURY_PROGRAMS[aln]
                else:
                    prog_name = "Unknown Program"
                    prog_acronym = "UNK"
                
                # Get or create program
                if aln not in programs_map:
                    programs_map[aln] = Program(
                        aln=aln,
                        name=prog_name,
                        acronym=prog_acronym,
                    )
                
                # Classify finding (Comment #37)
                text = findings_text.get(ref, "")
                summary = classify_finding(text)
                
                # Get compliance type (Comment #32, #34)
                ctype_code = (f.get("type_requirement") or "")[:1].upper()
                ctype_label = COMPLIANCE_TYPES.get(ctype_code, "Other")
                
                # If type is "P" (Other), use heading from finding text (Comment #33)
                if ctype_code == "P" and text:
                    # Try to extract a heading from the finding
                    first_line = text.split('\n')[0].strip()
                    if first_line and len(first_line) < 100:
                        ctype_label = first_line
                
                # CAP determination
                cap_text = caps.get(ref)
                cap_det = "Accepted" if cap_text else "Not Applicable"
                
                # Check for repeat finding (Comment #29, #30)
                is_repeat = f.get("is_repeat_finding") in [True, "Y", "Yes", "true", "yes"]
                prior_ref = f.get("prior_finding_ref_numbers") or ""
                
                programs_map[aln].findings.append(Finding(
                    finding_id=ref,
                    compliance_type=ctype_label,
                    summary=summary,
                    cap_determination=cap_det,
                    is_repeat=is_repeat,
                    prior_finding_ref=prior_ref,
                ))
            
            if not programs_map:
                return {"ok": False, "message": f"No Treasury findings found for {treasury_listings}"}
            
            # Sort programs by ALN (Comment #21: "Put tables in ALN order")
            programs = [programs_map[aln] for aln in sorted(programs_map.keys())]
            
            # 5. Build template replacement data
            # Count total findings for pluralization (Comment #19, #20, #40)
            total_findings = sum(len(p.findings) for p in programs)
            plural = get_pluralization(total_findings)
            
            # Get recipient name in standard case (Comment #1, #3)
            raw_recipient = recipient_name or report.get("auditee_name") or auditee_name
            recipient_standard = to_standard_case(raw_recipient)
            
            # Get auditor name in standard case
            raw_auditor = auditor_name or report.get("auditor_firm_name") or ""
            auditor_standard = to_standard_case(raw_auditor)
            
            # Build data dict
            data = {
                # Date (Comment #0)
                "Date XX, 2025": get_current_date(),
                
                # Recipient (Comments #1, #3) - standard case, NO "The" in address
                "Recipient Name": recipient_standard,
                
                # EIN (Comments #4, #5, #6) - XX-XXXXXXX format
                "EIN": format_ein(ein),
                
                # Address (Comments #7, #8, #9, #10)
                "Street Address": to_standard_case(street_address or report.get("auditee_address_line_1") or ""),
                "City": to_standard_case(city or report.get("auditee_city") or ""),
                "State": (state or report.get("auditee_state") or "").upper(),
                "Zip Code": zip_code or report.get("auditee_zip") or "",
                
                # Point of Contact (Comments #11, #12)
                "Point of Contact": to_standard_case(poc_name or report.get("auditee_contact_name") or ""),
                "Point of Contact Title": to_standard_case(poc_title or report.get("auditee_contact_title") or ""),
                
                # Auditor (with "the" prefix for narrative)
                "Auditor Name": auditor_standard,
                
                # Fiscal Year End (Comments #13, #14, #15)
                "Fiscal Year End Date": fiscal_year_end or format_date(report.get("fy_end_date")),
                
                # Email
                "treasury_contact_email": treasury_email or Config.TREASURY_EMAIL,
                
                # Pluralization (Comments #19, #20, #40)
                **plural,
            }
            
            # 6. Generate document
            docx_bytes = self._generate_document(data, programs, auditor_standard)
            
            # 7. Save
            filename = f"MDL-{self._sanitize(auditee_name)}-{ein.replace('-', '')}-{audit_year}.docx"
            url = self._save_document(docx_bytes, filename)
            
            logger.info(f"Generated: {filename} ({total_findings} findings in {len(programs)} programs)")
            
            return {
                "ok": True,
                "url": url,
                "blob_path": f"{Config.AZURE_CONTAINER}/{filename}" if Config.AZURE_CONN_STR else filename,
                "report_id": report_id,
                "findings_count": total_findings,
                "programs_count": len(programs),
            }
            
        except Exception as e:
            logger.exception("Generation failed")
            return {"ok": False, "message": str(e)}
    
    def _generate_document(
        self,
        data: Dict[str, str],
        programs: List[Program],
        auditor_name: str,
    ) -> bytes:
        """Generate document from template."""
        
        # Load template
        if os.path.exists(self.template_path):
            doc = Document(self.template_path)
            logger.info(f"Loaded template: {self.template_path}")
        else:
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        
        # Build replacement map
        replacements = {}
        for key, value in data.items():
            # Special handling for "Date XX, 2025" (literal text in template)
            if key.startswith("Date "):
                replacements[key] = value or ""
            else:
                # All other keys are bracketed placeholders
                replacements[f"[{key}]"] = value or ""
        
        # Special: "[The] [Recipient Name]" -> "The City of Ann Arbor" (Comment #16)
        # But only in narrative, not in address block
        # The address block just has "[Recipient Name]" without "[The]"
        
        # Special: "[the] [Auditor Name]" -> "the Rehmann Robson LLC"
        replacements["[the]"] = "the"
        
        # Replace in all paragraphs
        for para in doc.paragraphs:
            self._replace_in_paragraph(para, replacements)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        self._replace_in_paragraph(para, replacements)
        
        # Replace in headers/footers
        for section in doc.sections:
            for para in section.header.paragraphs:
                self._replace_in_paragraph(para, replacements)
            for para in section.footer.paragraphs:
                self._replace_in_paragraph(para, replacements)
        
        # Find "[ALN]/ [Program Name] [(Program Acronym)]" and replace with actual program headers
        # Then insert findings tables
        self._insert_program_sections(doc, programs)
        
        # Save to bytes
        buffer = BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    
    def _replace_in_paragraph(self, para, replacements: Dict[str, str]):
        """
        Replace placeholders in paragraph - handles placeholders split across runs.
        Key insight: Always check the full paragraph text, not just individual runs.
        """
        # Get full text first
        full_text = para.text
        
        # Check if any replacement needed
        needs_replacement = False
        for old in replacements:
            if old in full_text:
                needs_replacement = True
                break
        
        if not needs_replacement:
            return
        
        # Apply replacements to full text
        new_text = full_text
        for old, new in replacements.items():
            new_text = new_text.replace(old, new)
        
        if new_text == full_text:
            return  # Nothing changed
        
        # Rebuild paragraph: put all text in first run, clear others
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""
    
    def _insert_program_sections(self, doc: Document, programs: List[Program]):
        """
        Find the program header line and table template, then replicate for each program.
        Comment #21, #22: Create new table for each program, in ALN order.
        """
        # Find the template elements
        program_line_para = None
        template_table = None
        template_row_idx = None
        
        for para in doc.paragraphs:
            text = para.text
            if "[ALN]" in text:
                program_line_para = para
                break
        
        # Find the findings table (has [Finding Number])
        for table in doc.tables:
            for ri, row in enumerate(table.rows):
                for cell in row.cells:
                    if "[Finding Number]" in cell.text:
                        template_table = table
                        template_row_idx = ri
                        break
                if template_table:
                    break
            if template_table:
                break
        
        if not template_table or not program_line_para:
            logger.warning("Could not find template table or program line")
            return
        
        first_program = programs[0]
        
        # Replace program line placeholders
        for run in program_line_para.runs:
            text = run.text
            text = text.replace("[ALN]", first_program.aln)
            text = text.replace("[Program Name]", first_program.name)
            text = text.replace("[(Program Acronym)]", f"({first_program.acronym})")
            run.text = text
        
        # Also check if split across runs
        full_text = program_line_para.text
        if "[ALN]" in full_text or "[Program Name]" in full_text:
            new_text = full_text
            new_text = new_text.replace("[ALN]", first_program.aln)
            new_text = new_text.replace("[Program Name]", first_program.name)
            new_text = new_text.replace("[(Program Acronym)]", f"({first_program.acronym})")
            if program_line_para.runs:
                program_line_para.runs[0].text = new_text
                for run in program_line_para.runs[1:]:
                    run.text = ""
        
        # Fill the template table with first program's first finding
        if first_program.findings:
            self._fill_template_row(template_table.rows[template_row_idx], first_program.findings[0])
            
            # Add additional rows for more findings
            for finding in first_program.findings[1:]:
                new_row = template_table.add_row()
                self._fill_finding_row(new_row, finding)
        
        # For additional programs, insert after first table
        if len(programs) > 1:
            insert_after = template_table._element
            
            for program in programs[1:]:
                # Add spacing paragraph
                blank_p = doc.add_paragraph()
                insert_after.addnext(blank_p._element)
                insert_after = blank_p._element
                
                # Add program header
                header_p = doc.add_paragraph()
                run = header_p.add_run("Assistance Listing Number/Program Name:")
                run.bold = True
                insert_after.addnext(header_p._element)
                insert_after = header_p._element
                
                # Add program name line
                name_p = doc.add_paragraph(f"{program.aln}/ {program.name} ({program.acronym})")
                insert_after.addnext(name_p._element)
                insert_after = name_p._element
                
                # Add blank
                blank_p2 = doc.add_paragraph()
                insert_after.addnext(blank_p2._element)
                insert_after = blank_p2._element
                
                # Add table
                table = self._create_findings_table(doc, program.findings)
                insert_after.addnext(table._element)
                insert_after = table._element
    
    def _fill_template_row(self, row, finding: Finding):
        """Fill the template row by replacing placeholders in all paragraphs."""
        replacements = {
            "[Finding Number]": finding.finding_id,
            "[Compliance Type]": finding.compliance_type,
            "[Audit Finding Summary]": finding.summary,
        }
        
        for cell in row.cells:
            # Process ALL paragraphs in the cell (template has 2 paragraphs in column 1)
            for para in cell.paragraphs:
                full_text = para.text
                
                # Check if any replacement needed
                needs_replacement = False
                for old in replacements:
                    if old in full_text:
                        needs_replacement = True
                        break
                
                if needs_replacement:
                    new_text = full_text
                    for old, new in replacements.items():
                        new_text = new_text.replace(old, new)
                    
                    # Rebuild paragraph
                    if para.runs:
                        para.runs[0].text = new_text
                        for r in para.runs[1:]:
                            r.text = ""
    
    def _fill_findings_table(self, table, findings: List[Finding]):
        """Fill existing template table with findings."""
        # The template has a header row and one data row template
        # We need to fill the first finding, then add rows for rest
        
        if len(table.rows) < 2:
            return
        
        # Fill first finding in existing row
        if findings:
            self._fill_finding_row(table.rows[1], findings[0])
        
        # Add additional rows
        for finding in findings[1:]:
            # Add a new row by copying structure
            new_row = table.add_row()
            self._fill_finding_row(new_row, finding)
    
    def _fill_finding_row(self, row, finding: Finding):
        """Fill a table row with finding data."""
        cells = row.cells
        
        # Column 0: Finding Number (with repeat indicator if applicable)
        # Comment #29, #30: If repeat, add "Repeat of [prior_ref]"
        finding_text = finding.finding_id
        if finding.is_repeat and finding.prior_finding_ref:
            finding_text += f"\nRepeat of {finding.prior_finding_ref}"
        
        cells[0].text = ""
        p = cells[0].paragraphs[0]
        run = p.add_run(finding_text)
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Column 1: Compliance Type - Summary
        cells[1].text = ""
        p = cells[1].paragraphs[0]
        run = p.add_run(finding.compliance_type)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        p.add_run("  ̶\n")  # Em dash
        run = p.add_run(finding.summary)
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        
        # Column 2: Audit Finding Determination
        cells[2].text = ""
        p = cells[2].paragraphs[0]
        run = p.add_run(finding.audit_determination)
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Column 3: Questioned Cost Determination
        cells[3].text = ""
        p = cells[3].paragraphs[0]
        run = p.add_run(finding.questioned_cost)
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Column 4: CAP Determination
        cells[4].text = ""
        p = cells[4].paragraphs[0]
        run = p.add_run(finding.cap_determination)
        run.font.size = Pt(10)
        run.font.name = "Calibri"
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    def _create_findings_table(self, doc: Document, findings: List[Finding]):
        """Create a new findings table."""
        table = doc.add_table(rows=len(findings) + 1, cols=5)
        add_table_borders(table)
        
        # Column widths
        widths = [Inches(0.8), Inches(2.8), Inches(1.0), Inches(1.2), Inches(0.8)]
        for row in table.rows:
            for idx, cell in enumerate(row.cells):
                cell.width = widths[idx]
        
        # Header row
        headers = ["Audit\nFinding #", "Compliance Type - Audit Finding Summary",
                   "Audit Finding\nDetermination", "Questioned Cost\nDetermination", "CAP\nDetermination"]
        
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(h)
            run.bold = True
            run.font.size = Pt(10)
            run.font.name = "Calibri"
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            add_shading(cell)
        
        # Data rows
        for row_idx, finding in enumerate(findings, start=1):
            self._fill_finding_row(table.rows[row_idx], finding)
        
        return table
    
    def _save_document(self, data: bytes, filename: str) -> str:
        """Save document and return URL."""
        if Config.AZURE_CONN_STR:
            return self._save_to_azure(data, filename)
        return self._save_local(data, filename)
    
    def _save_to_azure(self, data: bytes, filename: str) -> str:
        """Save to Azure Blob Storage."""
        from azure.storage.blob import BlobServiceClient, generate_blob_sas, BlobSasPermissions
        
        parts = dict(p.split("=", 1) for p in Config.AZURE_CONN_STR.split(";") if "=" in p)
        account_name = parts.get("AccountName", "")
        account_key = parts.get("AccountKey", "")
        
        client = BlobServiceClient.from_connection_string(Config.AZURE_CONN_STR)
        container = client.get_container_client(Config.AZURE_CONTAINER)
        try:
            container.create_container()
        except:
            pass
        
        blob = container.get_blob_client(filename)
        blob.upload_blob(data, overwrite=True)
        
        sas = generate_blob_sas(
            account_name=account_name,
            container_name=Config.AZURE_CONTAINER,
            blob_name=filename,
            account_key=account_key,
            permission=BlobSasPermissions(read=True),
            expiry=datetime.utcnow() + timedelta(hours=2),
        )
        
        return f"https://{account_name}.blob.core.windows.net/{Config.AZURE_CONTAINER}/{filename}?{sas}"
    
    def _save_local(self, data: bytes, filename: str) -> str:
        """Save locally."""
        os.makedirs(Config.LOCAL_SAVE_DIR, exist_ok=True)
        path = os.path.join(Config.LOCAL_SAVE_DIR, filename)
        with open(path, "wb") as f:
            f.write(data)
        return f"{Config.PUBLIC_BASE_URL}/local/{filename}"
    
    def _sanitize(self, s: str) -> str:
        """Sanitize for filename."""
        return re.sub(r"[^A-Za-z0-9._-]+", "_", s or "").strip("_")[:50]


# ============================================================
# FASTAPI APP
# ============================================================

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

app = FastAPI(title="MDL DOCX Builder", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


class BuildRequest(BaseModel):
    """Request model."""
    auditee_name: str
    ein: str
    audit_year: int
    treasury_listings: Optional[List[str]] = None
    max_refs: int = Field(default=15, ge=1, le=100)
    only_flagged: bool = False
    include_awards: bool = True
    recipient_name: Optional[str] = None
    street_address: Optional[str] = None
    city: Optional[str] = None
    state: Optional[str] = None
    zip_code: Optional[str] = None
    poc_name: Optional[str] = None
    poc_title: Optional[str] = None
    auditor_name: Optional[str] = None
    fy_end_text: Optional[str] = None
    template_path: Optional[str] = None
    aln_reference_xlsx: Optional[str] = None
    dest_path: Optional[str] = None


@app.get("/healthz")
def healthz():
    return {"ok": True, "service": "mdl-generator", "version": "2.0.0"}


@app.post("/build-mdl-docx-auto")
def build_mdl_docx_auto(req: BuildRequest):
    """Main endpoint."""
    try:
        template = req.template_path or Config.MDL_TEMPLATE_PATH
        gen = MDLGenerator(template_path=template)
        
        result = gen.generate_from_fac(
            auditee_name=req.auditee_name,
            ein=req.ein,
            audit_year=req.audit_year,
            treasury_listings=req.treasury_listings,
            max_refs=req.max_refs,
            only_flagged=req.only_flagged,
            recipient_name=req.recipient_name,
            street_address=req.street_address,
            city=req.city,
            state=req.state,
            zip_code=req.zip_code,
            poc_name=req.poc_name,
            poc_title=req.poc_title,
            auditor_name=req.auditor_name,
            fiscal_year_end=req.fy_end_text,
        )
        
        return result
        
    except Exception as e:
        logger.exception("Unhandled error")
        return JSONResponse(status_code=200, content={"ok": False, "message": str(e)})


@app.post("/build-mdl")
def build_mdl(req: BuildRequest):
    return build_mdl_docx_auto(req)


@app.get("/local/{path:path}")
def get_local_file(path: str):
    full = os.path.join(Config.LOCAL_SAVE_DIR, path)
    if not os.path.isfile(full):
        raise HTTPException(404, "Not found")
    return FileResponse(
        full,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# ============================================================
# CLI
# ============================================================

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) >= 4:
        gen = MDLGenerator()
        result = gen.generate_from_fac(
            auditee_name=sys.argv[1],
            ein=sys.argv[2],
            audit_year=int(sys.argv[3]),
        )
        if result.get("ok"):
            print(f"✓ Generated: {result.get('url')}")
            print(f"  Report: {result.get('report_id')}")
            print(f"  Findings: {result.get('findings_count')}")
        else:
            print(f"✗ Error: {result.get('message')}")
            sys.exit(1)
    else:
        import uvicorn
        print("Starting MDL Generator API on port 8000...")
        uvicorn.run(app, host="0.0.0.0", port=8000)
