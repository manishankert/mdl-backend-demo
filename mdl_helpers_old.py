
from typing import Dict, List, Optional

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

def _combine_comp_summary(f: Dict[str, str]) -> str:
    """
    Build the 'Compliance Type – Audit Finding Summary' cell.
    Replaces the finding type code (e.g., 'I') with the full description from `finding_types`.
    Expected keys in `f`: 'compliance_type', 'summary'
    """
    raw_comp = (f.get("compliance_type") or "").strip()
    # exact, uppercase, and title-case lookups
    comp_text = (finding_types.get(raw_comp)
                 or finding_types.get(raw_comp.upper())
                 or finding_types.get(raw_comp.title())
                 or raw_comp)

    summ = (f.get("summary") or "").strip()
    if comp_text and summ:
        return f"{comp_text} – {summ}"  # en dash
    return comp_text or summ

def map_compliance_type(code: str) -> str:
    return (finding_types.get(code)
            or finding_types.get((code or '').upper())
            or finding_types.get((code or '').title())
            or code)

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
        # return exact match if present
        for c in candidates:
            if choice.lower() == c.lower():
                return c
        for c in candidates:
            if c.lower() in choice.lower():
                return c
        return choice
    except Exception:
        return None
