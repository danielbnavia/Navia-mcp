# Cloudflare front for navia-mcp (live deployment topology)

Reality check from probes on 2026-05-03:

| URL | Status |
|---|---|
| `https://navia-mcp.livelyground-86f971d8.australiasoutheast.azurecontainerapps.io/health` | **200 OK** -- MCP server is live on Azure Container Apps |
| `https://navia-mcp.ngrok.io/health` | **404** -- ngrok tunnel dead |
| `https://navia-agentic-api.danielb-ca5.workers.dev/api/health` | **200 OK** -- Worker live |

So the unification work isn't "deploy navia-mcp to Cloudflare" -- it's **"put Cloudflare in front of the existing Azure Container Apps deploy"**. Two paths; pick one.

## Path A (recommended): CF DNS proxy -> Azure Container Apps custom domain

No Docker, no tunnel, no sidecars. Just DNS + a custom domain on ACA.

```
[Copilot Studio / Claude Desktop / Claude Code]
              |
              v
  navia-mcp.naviafreight.com (Cloudflare proxied, orange-cloud)
              |  CF Access policy (SSO + service token)
              v
  navia-mcp.livelyground-86f971d8.australiasoutheast.azurecontainerapps.io
              |
              v
  Existing AAD requireRole("API.Access") gate on /mcp
```

### Steps

1. **Add the custom hostname on Azure Container Apps**

   In the Azure Portal -> Container Apps -> `navia-mcp` -> **Custom domains** ->
   *Add*. Hostname: `navia-mcp.naviafreight.com`. ACA prints the validation
   TXT and CNAME records.

2. **Create the records in Cloudflare DNS**

   - `_dnsauth.navia-mcp` TXT -> the value ACA gave you (validation only).
   - `navia-mcp` CNAME -> `navia-mcp.livelyground-86f971d8.australiasoutheast.azurecontainerapps.io` -- proxy = **DNS only (grey cloud)** for the validation step.

3. **Validate + bind cert in Azure**

   Click *Validate* on the ACA custom-domain blade. ACA issues a managed cert.

4. **Switch the CNAME to proxied (orange cloud)** in Cloudflare. SSL/TLS
   mode for the zone must be **Full (strict)** so the CF -> ACA hop verifies
   against the new ACA cert.

5. **Add Cloudflare Access in front** (Zero Trust dashboard):
   - Application type: Self-hosted
   - Domain: `navia-mcp.naviafreight.com`
   - Policy 1 (Service Auth): create service token `navia-mcp-claude` -- the
     Client ID + Secret become `NAVIA_CF_ACCESS_CLIENT_ID` /
     `NAVIA_CF_ACCESS_CLIENT_SECRET` env vars in the Claude clients.
   - Policy 2 (Allow): emails ending `@naviafreight.com` for Copilot Studio
     SSO + interactive browser.

6. **Smoke test** -- see `scripts/verify-cloudflare-front.sh`.

## Path B: Cloudflare Tunnel (use when ACA isn't an option)

Keep the legacy `cloudflared/config.example.yml` + `docker-compose.cloudflare.yml`
artifacts in this PR for the case where you want to run navia-mcp on a Docker
host outside ACA. They're still valid -- just unnecessary for the current
Azure deployment.

## Other connectors

| Connector | Action to put Cloudflare in front |
|---|---|
| `navia-agentic-api` (CF Worker) | Add custom domain `navia-agentic-api.naviafreight.com` in the Worker dashboard. Already on Cloudflare -- 1-click. |
| `navia-orchestrator` (GCP Cloud Run) | CF DNS CNAME -> the run.app host, proxied; or move behind a CF Worker reverse-proxy. Auth stays as `x-api-key`. |
| `outlook-triage-a2a` (currently ngrok) | Replace ngrok with Cloudflare Tunnel (use the existing `cloudflared/config.example.yml` template). |
| `zapier-bridge` (currently ngrok) | Same -- replace ngrok with CF Tunnel. |
| `asana-nf-mcp` | None -- already on Asana SaaS, leave as-is. |

## Smoke test

After setup, run from any machine with the env vars set:

```bash
bash scripts/verify-cloudflare-front.sh
```
