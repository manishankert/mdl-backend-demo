KEY="4LlIXwxP7dHQa5fMTPj8i6Gi0RSNNoJ3L9zaC4KV"
BASE="https://api.fac.gov"

curl -s -H "X-Api-Key: $KEY" \
  "$BASE/general?audit_year=eq.2024&select=report_id,auditee_name,auditee_ein&order=fac_accepted_date.desc&limit=300" \
| jq -r '.[] | @tsv' \
| while IFS=$'\t' read -r REPORT_ID NAME EIN; do
    # Quick probe: if <4 returned with limit=4, total is likely <=3
    ARR=$(curl -s -H "X-Api-Key: $KEY" \
      "$BASE/findings?report_id=eq.$REPORT_ID&select=reference_number&order=reference_number.asc&limit=4")
    CNT=$(echo "$ARR" | jq 'length')
    if [ "$CNT" -lt 4 ]; then
      # (Optional) confirm exact count by fetching more for this single candidate
      TRUECNT=$(curl -s -H "X-Api-Key: $KEY" \
        "$BASE/findings?report_id=eq.$REPORT_ID&select=reference_number&limit=200" | jq 'length')
      if [ "$TRUECNT" -ge 1 ] && [ "$TRUECNT" -le 3 ]; then
        echo "SMALL-FINDINGS($TRUECNT): auditee_name=\"$NAME\" | ein=\"$EIN\" | audit_year=2024 | report_id=$REPORT_ID"
        break
      fi
    fi
  done

