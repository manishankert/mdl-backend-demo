#!/usr/bin/env bash
set -euo pipefail

YEAR=2024
LIST=(
  "City of Boulder|846000566"
  "City of Tempe|866000262"
  "City of Raleigh|566000236"
  "City of Madison|396005507"
  "City of Phoenix|866000256"
  "County of Marin|946000519"
  "City of Ann Arbor|386004534"
  "City of Cambridge|046001383"
  "City of Austin|746000085"
  "Oconto County|396005722"
)

good=0
for entry in "${LIST[@]}"; do
  name="${entry%%|*}"; ein="${entry##*|}"
  rid=$(curl -s "https://api.fac.gov/general?audit_year=eq.${YEAR}&auditee_ein=eq.${ein}&select=report_id&order=fac_accepted_date.desc&limit=1" \
          -H "X-Api-Key: $FAC_API_KEY" | jq -r '.[0].report_id')
  [[ "$rid" == "null" || -z "$rid" ]] && { echo "skip: $name ($ein) — no report"; continue; }

  # try findings with a “flagged” OR filter; if none, check for any findings_text
  flagged=$(curl -s "https://api.fac.gov/findings?report_id=eq.${rid}&select=reference_number&limit=1&or=(is_material_weakness.eq.true,is_significant_deficiency.eq.true,is_questioned_costs.eq.true,is_modified_opinion.eq.true,is_other_findings.eq.true,is_other_matters.eq.true,is_repeat_finding.eq.true)" \
              -H "X-Api-Key: $FAC_API_KEY" | jq 'length')
  if [[ "$flagged" -gt 0 ]]; then
    echo "GOOD: $name|$ein|$YEAR  (flagged findings present)"
    ((good++))
    continue
  fi

  has_text=$(curl -s "https://api.fac.gov/findings_text?report_id=eq.${rid}&select=finding_ref_number&limit=1" \
               -H "X-Api-Key: $FAC_API_KEY" | jq 'length')
  if [[ "$has_text" -gt 0 ]]; then
    echo "OK:   $name|$ein|$YEAR  (narrative findings present; flags may be missing)"
    ((good++))
  else
    echo "skip: $name ($ein) — no findings"
  fi
done

echo "Found $good good candidates."
