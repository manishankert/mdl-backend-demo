from typing import Dict, List, Optional, Tuple
from decimal import Decimal, ROUND_DOWN

# --- Generated artifacts ---
finding_summaries_list: List[str] = [
"Deficient Record Management",
"Deficient Subrecipient Monitoring or Deficient Subaward",
"Failure to file FFATA report for subawards",
"Failure to include required provisions in subawards supported with federal funds",
"Failure to monitor contracts/third party service providers",
"Failure to monitor subrecipient cash draws",
"Failure to retain adequate supporting documentation",
"Failure to submit Treasury report(s)",
"Failure to submit Single Audit report timely",
"Fraud",
"Improper internal recording of expenditures",
"Inaccurate calculation of revenue loss",
"Inaccurate Treasury Reporting",
"Inadequate Limitations related to Administrative Costs in Contractual Agreements",
"Lack of Adequate System Controls - Change Management",
"Lack of Adequate System Controls - Security and Access",
"Lack of an Independent Review Prior to Disbursement of Funds",
"Lack of an Independent Review for Reporting",
"Lack of Cash Management Controls - Interest Earned",
"Lack of Eligibility Support",
"Lack of evidence of competitive procurement",
"Lack of evidence of suspension and debarment verification",
"Lack of evidence to support that costs are necessary and reasonable",
"Lack of Federal Government Approval",
"Lack of Internal Controls- Equipment and Real Property Management",
"Lack of Internal Controls - Financial Statement Preparation",
"Lack of Internal Controls- Grants Management",
"Lack of Internal Controls- IT Risk Management",
"Lack of Internal Controls - Matching, Level of Effort, and Earmarking",
"Lack of Internal Controls- Payroll",
"Lack of Internal Controls- SEFA Preparation",
"Lack of Segregation of Duties",
"Lack of time and effort documentation",
"Lack of Written Policies and/or Procedures - Procurement, Suspension, and Debarment",
"Lack of Written Policies and/or Procedures - Management of Federal Funds",
"Lack of Written Policies and/or Procedures - Subrecipient Monitoring",
"Late submission of Treasury report(s)",
"Multiple Issues",
"Unallowable expenditures due to being incurred outside of period of performance",
"Unallowable expenditures due to duplicative benefits with another"
]
finding_types: Dict[str, str] = {
"A": "Activities allowed or unallowed",
"B": "Allowable costs/cost principles",
"C": "Cash management",
"D": "Reserved",
"E": "Eligibility",
"F": "Equipment and real property management",
"G": "Matching, level of effort, earmarking",
"H": "Period of performance (or availability) of Federal funds",
"I": "Procurement and suspension and debarment",
"J": "Program income",
"K": "Reserved",
"L": "Reporting",
"M": "Subrecipient monitoring",
"N": "Special tests and provisions",
"P": "Other"
}
aln_program_acronym: List[tuple] = [
[
"21.029",
"Capital Projects Fund",
"CPF"
],
[
"21.019",
"Coronavirus Relief Fund",
"CRF"
],
[
"21.023",
"Emergency Rental Assistance Program",
"ERA"
],
[
"21.026",
"Homeowner Assistance Fund",
"HAF"
],
[
"21.032",
"Local Assistance and Tribal Consistency Fund",
"LATCF"
],
[
"21.031",
"State Small Business Credit Initiative",
"SSBCI"
],
[
"21.027",
"State and Local Fiscal Recovery Funds",
"SLFRF"
]
]

# ---------- ALN helpers ----------

def _normalize_aln(aln: object) -> str:
    """
    Normalize ALN to '##.###' (string) with zero padding, tolerant to floats/strings.
    Examples: 21.27 -> '21.270', '21.027' -> '21.027', 21 -> '21.000'
    """
    if aln is None:
        return ""
    s = str(aln).strip()
    if s.replace(".", "", 1).isdigit():
        try:
            # Use Decimal for stable zero padding (no scientific notation issues)
            d = Decimal(s)
            return f"{d.quantize(Decimal('0.000'), rounding=ROUND_DOWN):f}"
        except Exception:
            pass
    # Non-numeric inputs: try to coerce if it looks like '21-027' or '21_027'
    s2 = s.replace("-", ".").replace("_", ".").replace(" ", "")
    if s2.count(".") == 1:
        left, right = s2.split(".")
        if left.isdigit() and right.isdigit():
            right = (right + "000")[:3]  # pad/crop to 3
            return f"{int(left)}.{right}"
    # Last resort: return as-is
    return s

# Build index: '##.###' -> (program, acronym)
_acronym_index: Dict[str, Tuple[str, str]] = {}
for a, program, acr in aln_program_acronym:
    key = _normalize_aln(a)
    if key:
        _acronym_index[key] = (program, acr)

def get_program_acronym(aln: object, program_hint: Optional[str] = None) -> Optional[Dict[str, str]]:
    """Return a dict with 'aln', 'program', 'acronym' for the given ALN.
    If not found by ALN, and program_hint given, try a loose program-name match.
    """
    key = _normalize_aln(aln)
    tup = _acronym_index.get(key)
    if tup:
        program, acr = tup
        return {"aln": key, "program": program, "acronym": acr}

    # fallback: loose program name search if provided
    if program_hint:
        ph = program_hint.strip().lower()
        for k, (p, acr) in _acronym_index.items():
            if ph == p.lower() or ph in p.lower():
                return {"aln": k, "program": p, "acronym": acr}
    return None

def format_program_header_line(aln: object, program_hint: Optional[str] = None) -> str:
    """Format: '<ALN> <Program Name> (<Acronym>)' with good defaults when missing."""
    info = get_program_acronym(aln, program_hint)
    if info:
        return f"{info['aln']}/{info['program']} ({info['acronym']})"
    # Unknown path
    aln_str = _normalize_aln(aln) or "Unknown"
    prog_str = program_hint or "Unknown Program"
    return f"{aln_str} {prog_str}"

# ---------- Table text helpers ----------

def _combine_comp_summary(f: Dict[str, str]) -> str:
    """Build 'Compliance Type – Audit Finding Summary' with mapped finding type text."""
    raw_comp = (f.get("compliance_type") or "").strip()
    comp_text = (finding_types.get(raw_comp)
                    or finding_types.get(raw_comp.upper())
                    or finding_types.get(raw_comp.title())
                    or raw_comp)

    summ = (f.get("summary") or "").strip()
    return f"{comp_text} – {summ}" if comp_text and summ else (comp_text or summ)

def map_compliance_type(code: str) -> str:
    return (finding_types.get(code)
            or finding_types.get((code or '').upper())
            or finding_types.get((code or '').title())
            or (code or ''))

# --- OpenAI classification helper (user must set OPENAI_API_KEY) ---
def classify_top_category(summary: str, candidates: List[str], model: str = "gpt-4o-mini") -> Optional[str]:
    """Return the single best-matching category from `candidates` for `summary`.
    Requires `openai` python package >= 1.0 and network access.
    """
    try:
        from openai import OpenAI
        client = OpenAI()
        system = "You are a strict classifier. Pick the single best category label from the provided list."
        user = "Summary: " + summary + "\n\nCategories:\n- " + "\n- ".join(candidates) + "\n\nRespond with exactly one category label from the list."
        resp = client.chat.completions.create(
            model=model,
            messages=[
                { "role": "system", "content": system },
                { "role": "user", "content": user }
            ],
            temperature=0
        )
        choice = resp.choices[0].message.content.strip()
        for c in candidates:
            if choice.lower() == c.lower():
                return c
        for c in candidates:
            if c.lower() in choice.lower():
                return c
        return choice
    except Exception:
        return None
