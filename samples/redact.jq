# Applied to the raw Actor output before committing:
#   jq -f samples/redact.jq raw.json > samples/toronto-listings.json
#
# The sample is real output of yugenox/realtor-ca-property-scraper with the listing agents
# pseudonymized. Listing and brokerage fields are left as published.
# - Each listing agent's name becomes "LISTING AGENT 01", "LISTING AGENT 02" ... (in order of
#   first appearance). Their individualId and phone numbers are removed (null / []).
# - The brokerage name and organizationId stay: they identify the brokerage, not a person.
# - publicRemarks (free text that can name people) is replaced when present.

([.[] | .agents[]?.name] | reduce .[] as $n ([]; if index([$n]) then . else . + [$n] end)) as $names
| def pad($n): ($n | tostring | if length < 2 then "0" + . else . end);
  def alias($n): "LISTING AGENT " + pad(($names | index([$n])) + 1);
  [ .[]
    | .agents |= (if type == "array" then map(
        .name = (if .name then alias(.name) else null end)
        | .individualId = null
        | .phone = null
        | .phones = []) else . end)
    | .publicRemarks |= (if . == null then null else "(removed from this sample)" end) ]
