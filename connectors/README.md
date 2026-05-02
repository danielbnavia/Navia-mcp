# Navia connector registry

Source of truth for every connector that needs to be reachable from
Microsoft Copilot Studio **and** Claude Code / Claude Desktop.

Derived from the 9 swagger exports the team had floating around in Power
Platform (Default Solution). This registry consolidates them onto a single
Cloudflare-fronted topology so the same set of tools is available to every
agent surface.

## Inventory

| # | Connector | Canonical URL (target) | Currently deployed at | Protocol | Auth | Status |
|---|---|---|---|---|---|---|
| 1 | `navia-mcp` | `https://navia-mcp.naviafreight.com/mcp` | `navia-mcp.livelyground-86f971d8.australiasoutheast.azurecontainerapps.io` + `navia-mcp.ngrok.io` | MCP Streamable HTTP | CF Access service token + AAD `requireRole("API.Access")` | Migrating |
| 2 | `navia-mcp-intel` (alias of #1) | same | same | REST -> MCP tools | `x-api-key` | **Merge into #1** |
| 3 | `navia-orchestrator` | `https://navia-orchestrator.naviafreight.com` | `navia-orchestrator-115739894314.us-west1.run.app` (GCP Cloud Run) | REST | `x-api-key` | Move behind CF (proxy or migrate) |
| 4 | `navia-flare-triage` | `https://navia-agentic-api.naviafreight.com/api/triage` | `navia-agentic-api.danielb-ca5.workers.dev` | REST (CF Worker) | `Authorization: Bearer` | Add custom domain on the Worker |
| 5 | `asana-nf-mcp` | `https://mcp.asana.com/v2/mcp` | (Asana SaaS) | MCP Streamable HTTP | OAuth2 | Use as-is |
| 6 | `nf-collector` | `https://nf-collector.naviafreight.com/collector` | `REPLACE_ME_WITH_YOUR_HOST` | REST | `x-api-key` | **Decide: deploy or delete** |
| 7 | `outlook-triage-a2a` | `https://outlook-triage.naviafreight.com/a2a/rest` | `unparasitic-denisse-palaestric.ngrok-free.dev` | A2A 1.0 | `x-api-key` | Move behind CF Tunnel |
| 8 | _Researchescustomerserviceti.swagger.json_ | -- | same ngrok host | A2A 1.0 | leaked Anthropic key | **DELETED, key rotated** |
| 9 | `zapier-bridge` | `https://zapier-bridge.naviafreight.com` | same ngrok host | REST -> Zapier MCP | shared secret | Move behind CF Tunnel |

## Why a registry

Before this PR there was no single place that listed which connectors exist,
where they live, or who consumes them. The result: duplicate connectors in
Copilot Studio (#1 and #2 are the same backend), an Anthropic API key
leaked into the security definition of #8, and ngrok URLs that change on
restart. This file is the canonical answer to "how does my agent reach X".

## Per-surface wiring

### Microsoft Copilot Studio

Each connector is imported as a **custom connector** from its swagger in
`connectors/copilot-studio/`. The swaggers in this directory are the
_corrected_ versions:

- canonical Cloudflare hostnames substituted in;
- leaked secret stripped from the duplicate of #7;
- `x-api-key` standardised as the header name across REST connectors;
- auth section trimmed to a single security scheme per connector.

### Claude Code / Claude Desktop

Client configs in `connectors/clients/`:

- `claude-code-mcp.json` -- drop into a project's `.mcp.json`.
- `claude-desktop.json` -- drop into `claude_desktop_config.json`.

**Direct MCP** servers (#1, #5) are referenced natively. REST / A2A
connectors (#3, #4, #6, #7, #9) are reached via tool wrappers exposed by
the `navia-mcp` server itself, so the Claude clients only configure one
endpoint and still get the full toolbelt. See
[`docs/cloudflare-tunnel-setup.md`](../docs/cloudflare-tunnel-setup.md)
for the auth model.

## Action items

1. **Rotate the Anthropic API key** that was hardcoded into swagger #8.
   The key value is `sk-ant-api03-R9GAM4...` -- treat as fully compromised.
2. Decide on connector #6 (`nf-collector` -- host placeholder, never
   deployed?). If it's intended, deploy and update the registry; if dead,
   delete the Copilot Studio custom connector to stop confusing the agent.
3. Add Cloudflare Tunnel ingress entries for #3, #7, #9 alongside the
   existing `navia-mcp` entry in `cloudflared/config.example.yml`.
4. Add a `navia-agentic-api.naviafreight.com` custom domain on the
   existing Cloudflare Worker for #4 (no tunnel needed -- Workers can
   bind a hostname directly).
5. After all hostnames resolve, re-import the corrected swaggers into
   Copilot Studio and rebind every agent action to the new connector
   names. See `connectors/copilot-studio/IMPORT.md`.
6. Once Copilot Studio is on the new connectors, retire the ngrok
   URLs (`navia-mcp.ngrok.io`, `unparasitic-denisse-palaestric.ngrok-free.dev`).
