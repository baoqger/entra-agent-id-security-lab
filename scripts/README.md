# Scripts — Entra Agent ID Bootstrap

Shared PowerShell and bash tooling that provisions the Entra objects every sample in this repo depends on:

- a **Blueprint application** with a client credential
- an **Agent Identity** with Microsoft Graph `User.Read.All` granted
- a **Client SPA** for the On-Behalf-Of (OBO) flow
- OAuth2 permission grants wiring the three together

Run these once per tenant. After that, the IDs go into each sample's `.env` (local dev) or into federated credentials (Azure production).

## Prerequisites

- **PowerShell 7.4+** (`pwsh`) on macOS, Linux, or Windows. The workflow uses `Microsoft.Graph.*` modules 2.35+.
- **Entra role** on the user signing in — one of:
  - `Global Administrator`
  - `Agent ID Administrator` (template `db506228-d27e-4b7d-95e5-295956d6615f`) — recommended
  - `Agent ID Developer` (template `adb2368d-a9be-41b5-8667-d96778e081b0`) — least privilege
- **Graph modules:** `Install-Module Microsoft.Graph -Scope CurrentUser`

## Files in this folder

| File | Purpose |
|---|---|
| [`EntraAgentID-Functions.ps1`](EntraAgentID-Functions.ps1) | The main module. Exposes `Start-EntraAgentIDWorkflow` plus a set of lower-level helpers. |
| [`setup-obo-blueprint.ps1`](setup-obo-blueprint.ps1) | Adds the `access_as_user` scope to a Blueprint and grants a Client SPA permission against it. PowerShell edition. |
| [`setup-obo-blueprint.sh`](setup-obo-blueprint.sh) | Same workflow in bash, for CI systems without PowerShell. |
| [`setup-obo-client-app.ps1`](setup-obo-client-app.ps1) | Creates the Client SPA if one doesn't already exist in the tenant. PowerShell edition. |
| [`setup-obo-client-app.sh`](setup-obo-client-app.sh) | Same workflow in bash, for CI systems without PowerShell. |

## One-shot bootstrap

```powershell
# From the repo root
cd scripts
. ./EntraAgentID-Functions.ps1
Start-EntraAgentIDWorkflow
```

This interactively creates a Blueprint, mints an Agent Identity, and grants the Agent `User.Read.All`. You'll see output like:

```
Blueprint App ID:     <your-blueprint-app-id>
Blueprint Secret:     <generated>
Agent App ID:         <your-agent-app-id>
Agent Permissions:    User.Read.All (granted)
```

Save these values — you'll paste them into each sample's `.env` next.

## Walkthrough: end-to-end verification

This section proves the bootstrap worked by running the sidecar against a freshly created Agent Identity and calling Microsoft Graph with the minted token. It assumes you completed the one-shot bootstrap above.

### You'll learn

- How the Blueprint credential is consumed by the sidecar (not your agent code).
- What a two-token exchange (T1 → T2) looks like on the wire.
- How to inspect the Agent token's claims and confirm `User.Read.All` is present.
- The difference between the `/AuthorizationHeader` (token-only) and `/DownstreamApi` (token + proxied call) sidecar endpoints.

### Step 1 — Prepare environment configuration

```powershell
# In the repo root
cd sidecar/dev          # or sidecar/aws
Copy-Item .env.example .env
```

### Step 2 — Paste the bootstrap values into `.env`

```
TENANT_ID=<your-tenant-id>
BLUEPRINT_APP_ID=<blueprint-app-id from Step 1>
BLUEPRINT_CLIENT_SECRET=<secret from Start-EntraAgentIDWorkflow>
AGENT_CLIENT_ID=<agent-app-id from Step 1>
```

> When deployed on Azure App Service / Container Apps, **do not set `BLUEPRINT_CLIENT_SECRET`**. Use `SourceType=SignedAssertionFromManagedIdentity` instead. See [`../deploy/azure/container-apps/`](../deploy/azure/container-apps/).

### Step 3 — Start the sidecar

```powershell
docker-compose up -d sidecar
# For dev edition, also bring up Ollama and the agent; for aws, also the agent + token-refresher
```

Verify it's healthy:

```powershell
Invoke-RestMethod http://localhost:5000/health
```

### Step 4 — Get an Agent Identity token

```powershell
$agentAppId = "<your-agent-app-id>"

$response = Invoke-RestMethod `
  -Uri "http://localhost:5000/AuthorizationHeaderUnauthenticated/graph?AgentIdentity=$agentAppId" `
  -Method GET

$token = $response.authorizationHeader   # "Bearer eyJ..."
Write-Host "Got agent token: $($token.Substring(0,50))..."
```

What just happened:

1. You asked the sidecar for a Graph authorization header scoped to your Agent Identity.
2. The sidecar used the Blueprint credential to get a Blueprint token (T1) from Entra.
3. The sidecar exchanged T1 for an Agent token (T2) scoped to your Agent's permissions.
4. You received `Bearer <T2>`.

### Step 5 — Call Microsoft Graph with the Agent token

```powershell
$users = Invoke-RestMethod `
  -Uri "https://graph.microsoft.com/v1.0/users" `
  -Headers @{ Authorization = $token }

$users.value | Select-Object displayName, userPrincipalName | Format-Table
```

If you get a list of users back, the Agent's `User.Read.All` grant is working end-to-end.

### Step 6 — Verify the token claims (optional)

```powershell
$jwt = $token -replace "^Bearer ", ""
$payload = $jwt.Split(".")[1]
# Base64url decode and parse
$padded = $payload.PadRight($payload.Length + ((4 - $payload.Length % 4) % 4), "=")
$json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($padded.Replace("-", "+").Replace("_", "/")))
$claims = $json | ConvertFrom-Json

"App ID (appid):           $($claims.appid)"
"Audience (aud):           $($claims.aud)"
"Issuer (iss):             $($claims.iss)"
"Roles:                    $($claims.roles -join ', ')"
"Agent Identity marker:    xms_par_app_azp=$($claims.xms_par_app_azp)"
```

The key claim is `xms_par_app_azp` — it identifies the Blueprint that minted this token. That's how downstream APIs recognize the token as an Agent Identity token (and not a plain app-only token).

### Step 7 — Simplified pattern: sidecar does the call for you

Instead of two round-trips (get token, then call Graph), the sidecar can proxy the call:

```powershell
$result = Invoke-RestMethod `
  -Uri "http://localhost:5000/DownstreamApiUnauthenticated/graph?optionsOverride.RelativePath=users&AgentIdentity=$agentAppId" `
  -Method POST

$users = ($result.content | ConvertFrom-Json).value
$users | Select-Object displayName | Format-Table
```

Response shape: `{"statusCode": 200, "content": "<json string>"}`. The sidecar calls Graph for you and returns the response wrapped.

## Endpoints reference

| Endpoint | Authenticated? | Returns |
|---|---|---|
| `GET /AuthorizationHeader/graph` | Yes — requires incoming Bearer (OBO) | `Bearer <TF1>` (agent-on-behalf-of-user) |
| `GET /AuthorizationHeaderUnauthenticated/graph` | No (app-only) | `Bearer <TR>` (autonomous agent) |
| `POST /DownstreamApi/graph` | Yes (OBO) | `{statusCode, content}` from Graph |
| `POST /DownstreamApiUnauthenticated/graph` | No (app-only) | `{statusCode, content}` from Graph |
| `GET /health` | No | `Healthy` |

Always use `-Method POST` when hitting the `/DownstreamApi*` endpoints even if the downstream verb is GET. This is a [known requirement](https://learn.microsoft.com/en-us/entra/msidweb/agent-id-sdk/endpoints) of the SDK.

## Troubleshooting

### `Access token is empty`
The Blueprint credential in `.env` is blank or placeholder. Re-run `Start-EntraAgentIDWorkflow` and paste the real `BLUEPRINT_APP_ID` + `BLUEPRINT_CLIENT_SECRET`.

### `405 Method Not Allowed`
You're calling `/DownstreamApi*` with `GET`. Use `-Method POST`.

### Token has no `roles` claim
Graph permissions weren't granted-and-consented. Run `Start-EntraAgentIDWorkflow` to completion — it calls `Add-MgServicePrincipalAppRoleAssignment` to grant the role.

### `403 Authorization_RequestDenied` creating the Blueprint
Your signing-in user lacks an Agent ID role. `Application Administrator` is **not** sufficient. Add `Agent ID Developer` or `Agent ID Administrator` to your user.

### Container name conflict on restart
```powershell
docker-compose down
docker ps -a | Select-String entra-sidecar | ForEach-Object { docker rm -f ($_ -split '\s+')[0] }
docker-compose up -d
```

### Endpoint returns 404
Your sidecar is older than `mcr.microsoft.com/entra-sdk/auth-sidecar:latest`. Pull the latest:
```powershell
docker pull mcr.microsoft.com/entra-sdk/auth-sidecar:latest
docker-compose up -d --force-recreate sidecar
```

## Cleanup

```powershell
docker-compose down -v
```

To remove the Entra objects from your tenant:
```powershell
. ./EntraAgentID-Functions.ps1
Remove-EntraAgentIDWorkflow -BlueprintAppId <blueprint-app-id> -AgentAppId <agent-app-id>
```

## Summary

After completing this walkthrough you have:

- a Blueprint and Agent Identity in your tenant, with `User.Read.All` granted to the Agent
- a running sidecar that mints Agent tokens on demand, zero secrets in your own code
- verified end-to-end: sidecar → Entra → Graph → real user data
- the same `.env` values that the [`sidecar/dev`](../sidecar/dev/README.md) and [`sidecar/aws`](../sidecar/aws/README.md) samples consume

Proceed to one of the full samples to see an LLM drive this flow against the `weather-api`:

- [`sidecar/dev/README.md`](../sidecar/dev/README.md) — local-LLM (Ollama) + LangChain
- [`sidecar/aws/README.md`](../sidecar/aws/README.md) — AWS Bedrock (Claude) + LangChain
