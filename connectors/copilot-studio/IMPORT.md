# Re-importing the connectors into Copilot Studio

For each cleaned-up swagger in this directory:

1. **Copilot Studio -> Tools -> Add a tool -> Custom connector -> Import an
   OpenAPI file** -> upload the YAML / JSON.
2. Set the connector name to match the file basename (e.g. `navia-mcp`,
   `navia-orchestrator`, etc.) so existing rebind instructions in the
   agent docs line up.
3. Authentication, per connector:

   | Connector | Auth in Copilot Studio |
   |---|---|
   | `navia-mcp.swagger.json` | OAuth 2.0 (AAD) -- App Registration that issues `API.Access` role tokens. CF Access SSO covers the outer gate. |
   | `navia-orchestrator.swagger.json` | API key (`x-api-key` header) -- supply the orchestrator key from Google Secret Manager. |
   | `navia-flare-triage.swagger.json` | OAuth 2.0 / Bearer -- supply the AAD bearer that `navia-agentic-api` accepts. |
   | `asana-nf-mcp.swagger.json` | OAuth 2.0 -- existing Asana app's client id + secret. |
   | `nf-collector.swagger.json` | API key (`x-api-key`). **Skip if not deployed.** |
   | `outlook-triage-a2a.swagger.json` | API key (`x-api-key`). |
   | `zapier-bridge.swagger.json` | API key. Use the bridge's shared secret. |

4. Rebind any agent action that previously pointed at the legacy connector
   (the duplicates with `navia-mcp.ngrok.io` / `livelyground-...azure...` /
   `unparasitic-denisse-palaestric.ngrok-free.dev` hosts) to the newly
   imported one.
5. Republish the bot. For the Outlook Triage agent, use
   `scripts/publish-bot-v3.ps1` from `outlook-triage-knowledge`.
6. After 24h healthy in production, delete the old connectors so each
   capability has exactly one route from Copilot Studio to the backend.

## What was changed in the swaggers vs. the originals

- Hosts swapped to the canonical `*.naviafreight.com` names listed in
  `connectors/README.md`. Until DNS / Cloudflare Tunnel ingress is in
  place, Copilot Studio will fail to call them; that's intentional --
  prevents agents drifting back to the legacy URLs during the migration
  window.
- Standardised on `x-api-key` as the header name across REST connectors
  (one swagger had it set to `Authorization`, several had no auth
  declared at all).
- Removed `x-ms-openai-data` blocks from connectors that aren't currently
  surfaced as Copilot Studio AI plugins (keeps the manifest minimal).
- The duplicate of #7 with the leaked Anthropic key in the
  `securityDefinitions.api_key.name` field has been **deleted**, not
  re-published. Rotate the leaked key and use the cleaned `outlook-triage-a2a.swagger.json` instead.
