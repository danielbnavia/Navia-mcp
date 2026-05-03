#!/usr/bin/env bash
# Smoke-test the Cloudflare-fronted navia connectors.
#
# Required env:
#   NAVIA_CF_ACCESS_CLIENT_ID
#   NAVIA_CF_ACCESS_CLIENT_SECRET
#   NAVIA_MCP_TOKEN          (AAD bearer with API.Access role)
# Optional env:
#   NAVIA_MCP_HOST           (default navia-mcp.naviafreight.com)
#   NAVIA_AGENTIC_HOST       (default navia-agentic-api.naviafreight.com,
#                             falls back to .danielb-ca5.workers.dev if no
#                             custom domain yet)
#
# Exit code 0 = all checks passed.

set -u

MCP_HOST="${NAVIA_MCP_HOST:-navia-mcp.naviafreight.com}"
AGENT_HOST="${NAVIA_AGENTIC_HOST:-navia-agentic-api.naviafreight.com}"

fail=0

run() {
  local name="$1"
  local cmd="$2"
  local expected="$3"
  printf '%-55s ' "$name"
  local body status
  body=$(eval "$cmd" 2>&1) || true
  status=$(printf '%s' "$body" | tail -n1)
  if [[ "$status" == "$expected" ]]; then
    echo "OK ($status)"
  else
    echo "FAIL (got $status, expected $expected)"
    echo "  body: $(printf '%s' "$body" | head -n-1)"
    fail=1
  fi
}

echo "== Direct origin (no Cloudflare) =="
run "  Azure ACA navia-mcp /health" \
  "curl -s -o /dev/null -w '%{http_code}' https://navia-mcp.livelyground-86f971d8.australiasoutheast.azurecontainerapps.io/health" \
  "200"
run "  Workers.dev navia-agentic-api /api/health" \
  "curl -s -o /dev/null -w '%{http_code}' https://navia-agentic-api.danielb-ca5.workers.dev/api/health" \
  "200"

echo
echo "== Public MCP servers (used by Claude Code/Desktop) =="
run "  Microsoft Learn MCP (expect 405 on GET)" \
  "curl -s -o /dev/null -w '%{http_code}' https://learn.microsoft.com/api/mcp" \
  "405"
run "  Asana MCP (expect 401 without OAuth)" \
  "curl -s -o /dev/null -w '%{http_code}' https://mcp.asana.com/v2/mcp" \
  "401"

echo
echo "== Cloudflare-fronted navia-mcp =="
if [[ -z "${NAVIA_CF_ACCESS_CLIENT_ID:-}" || -z "${NAVIA_CF_ACCESS_CLIENT_SECRET:-}" ]]; then
  echo "  SKIP: NAVIA_CF_ACCESS_CLIENT_ID / _SECRET not set"
else
  run "  Without CF Access creds (expect 302 to login)" \
    "curl -s -o /dev/null -w '%{http_code}' https://$MCP_HOST/health" \
    "302"
  run "  With CF Access service token (expect 200)" \
    "curl -s -o /dev/null -w '%{http_code} ' -H 'CF-Access-Client-Id: $NAVIA_CF_ACCESS_CLIENT_ID' -H 'CF-Access-Client-Secret: $NAVIA_CF_ACCESS_CLIENT_SECRET' https://$MCP_HOST/health" \
    "200"
fi

echo
echo "== /mcp gate (CF Access + AAD) =="
if [[ -z "${NAVIA_CF_ACCESS_CLIENT_ID:-}" || -z "${NAVIA_MCP_TOKEN:-}" ]]; then
  echo "  SKIP: NAVIA_CF_ACCESS_CLIENT_ID and/or NAVIA_MCP_TOKEN not set"
else
  run "  /mcp tools/list (expect 200 + JSON-RPC envelope)" \
    "curl -s -o /dev/null -w '%{http_code}' -X POST -H 'Content-Type: application/json' -H 'CF-Access-Client-Id: $NAVIA_CF_ACCESS_CLIENT_ID' -H 'CF-Access-Client-Secret: $NAVIA_CF_ACCESS_CLIENT_SECRET' -H 'Authorization: Bearer $NAVIA_MCP_TOKEN' --data '{\"jsonrpc\":\"2.0\",\"id\":1,\"method\":\"tools/list\"}' https://$MCP_HOST/mcp" \
    "200"
fi

echo
if (( fail == 0 )); then
  echo "All checks passed."
else
  echo "One or more checks failed -- see output above."
fi
exit $fail
