# Cloudflare Tunnel + Access setup for navia-mcp

Replaces the ngrok-based exposure (`npm run tunnel`) with a Cloudflare-fronted
endpoint so Copilot Studio, Claude Desktop, and Claude Code can all reach the
same MCP server through a stable HTTPS URL.

## Topology

```
Copilot Studio  --+
Claude Desktop  --+--> navia-mcp.naviafreight.com (Cloudflare edge)
Claude Code     --+              |
                                 v
                         Cloudflare Access (policy gate)
                                 |
                                 v
                         Cloudflare Tunnel
                                 |
                                 v
                    Docker host: navia-mcp:88 (Express + MCP /mcp)
                                 |
                                 v
                    Existing AAD requireRole("API.Access") gate on /mcp
```

The existing AAD bearer / role check on `/mcp` is preserved. Cloudflare Access
sits in front and adds a second gate so Claude Code / Desktop can authenticate
via a Cloudflare Access **service token** rather than minting AAD tokens by hand.

## One-time setup

### 1. Create the tunnel

```bash
cloudflared tunnel login
cloudflared tunnel create navia-mcp
# Note the tunnel UUID it prints and the credentials file at ~/.cloudflared/<UUID>.json
```

Move the credentials file onto the host that will run the tunnel (the Docker
host) at `/etc/cloudflared/<UUID>.json`, then copy `cloudflared/config.example.yml`
to `cloudflared/config.yml` and replace `<TUNNEL_UUID>` with the UUID.

If you prefer the token-based mode (used by `docker-compose.cloudflare.yml`),
grab the tunnel token from **Zero Trust > Networks > Tunnels > Configure >
Install** instead and skip the credentials-file step.

### 2. Route DNS

```bash
cloudflared tunnel route dns navia-mcp navia-mcp.naviafreight.com
```

If your zone is something other than `naviafreight.com`, swap the hostname here
and update the matching line in:
- `cloudflared/config.example.yml`
- `docker-compose.cloudflare.yml`
- every sample in `client-configs/`
- every `.mcp.json` checked in to the agent repos

### 3. Add Cloudflare Access policies

In the Cloudflare Zero Trust dashboard:

1. **Access > Applications > Add an application > Self-hosted**.
2. Application domain: `navia-mcp.naviafreight.com`.
3. Add two policies on the application:
   - **Service token (Claude Code / Claude Desktop)**
     - Action: Service Auth
     - Include: Service Token > create `navia-mcp-claude`. Save the Client ID
       and Client Secret it generates -- these become `NAVIA_CF_ACCESS_CLIENT_ID`
       and `NAVIA_CF_ACCESS_CLIENT_SECRET` env vars in the client repos.
   - **SSO (Copilot Studio + interactive browser)**
     - Action: Allow
     - Include: emails ending in `@naviafreight.com` (or your IdP rule).

### 4. Run

```bash
docker compose -f docker-compose.cloudflare.yml up -d
docker compose -f docker-compose.cloudflare.yml logs -f cloudflared
```

Verify with:

```bash
curl -H "CF-Access-Client-Id: $NAVIA_CF_ACCESS_CLIENT_ID" \
     -H "CF-Access-Client-Secret: $NAVIA_CF_ACCESS_CLIENT_SECRET" \
     https://navia-mcp.naviafreight.com/health
```

Should return `{"status":"healthy","server":"navia-mcp-server",...}`.

## Client-side configuration

See `client-configs/`:
- `claude-desktop.json` -- drop-in for `claude_desktop_config.json`.
- `claude-code.mcp.json` -- drop-in for any project's `.mcp.json`.
- `copilot-studio-openapi.yaml` -- OpenAPI manifest for the Copilot Studio MCP
  custom connector.

All three reference the same env vars so secrets stay out of git:

| Env var | Purpose |
|---|---|
| `NAVIA_CF_ACCESS_CLIENT_ID` | Cloudflare Access service token Client ID |
| `NAVIA_CF_ACCESS_CLIENT_SECRET` | Cloudflare Access service token Client Secret |
| `NAVIA_MCP_TOKEN` | AAD bearer with `API.Access` role for the existing `requireRole` gate. Copilot Studio supplies this via the connector's OAuth connection; Claude clients need a fresh token in shell env until we add a long-lived service principal flow. (Same env var name dans-admin already uses in `lib/mcp/client.ts`.) |

## Migrating off ngrok (follow-up PR)

Once the tunnel is verified:

1. Replace the hardcoded `navia-mcp.ngrok.io` references in `src/http-server.ts`
   (in the `/config`, `/tab`, `/dashboard` HTML and the `/api/config` JSON
   response) with `process.env.MCP_BASE_URL`.
2. Remove the `tunnel` script from `package.json`.
3. Decommission the ngrok agent on the host.

These source changes are intentionally **not** in this PR so it stays purely
additive and reversible.
