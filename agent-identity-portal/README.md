# Agent Identity Portal

A self-service governance portal for **Microsoft Entra Agent ID** — request
and approve **Agent Blueprints**, **Agent Identities**, and **Permissions**
with built-in approval workflows, email + Teams notifications, and a full
audit log.

> **Status:** Sample code. Intended to demonstrate governance patterns for
> Entra Agent ID. **Not an officially supported Microsoft product.**

## Architecture

```
React SPA (Fluent UI + MSAL)            ── REST ──►      FastAPI backend
  · Blueprint catalog                                    · MSAL token validation (PyJWT)
  · Identity / permission request forms                  · Role-based authorization (Entra directory roles)
  · Admin approvals / audit log                          · Approval workflows + audit log
                                                         · Notifications (SMTP, Teams webhook)
                                                         · SQLite persistence
                                                                        │
                                                                        ▼
                                                   Microsoft Graph (beta) — live mode only
                                                   · agentIdentityBlueprint
                                                   · agentIdentity
                                                   · servicePrincipal / requiredResourceAccess
```

## Features

- **Agent Blueprint Catalog** — pre-approved templates with recommended permission sets
- **Identity Requests** — request a new Entra app registration from a blueprint or custom
- **Permission Requests** — add permissions to an existing agent identity
- **Approval Dashboard** — review, approve, or reject requests
- **Blocked-permission validation** — rejects Graph permissions that are not
  allowed for agent identities (admin-tier writes, high-risk resource writes, etc.)
- **Notifications** — HTML email + Teams Adaptive Card on new requests
- **Audit log** — every state transition stored with actor, action, and timestamp
- **Mock mode** — full UX without touching Entra ID, for demos and local dev

## Quick start (mock mode, no Azure setup)

```powershell
# backend
cd backend
python -m venv .venv
. .venv/Scripts/Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
# Leave AZURE_* blank and set DEV_AUTH_BYPASS=true for fully local dev.
$env:DEV_AUTH_BYPASS = "true"
$env:SEED_DEMO_DATA = "true"
uvicorn main:app --reload --port 8000
```

```powershell
# frontend (separate terminal)
cd frontend
npm install
Copy-Item .env.example .env
npm start
```

Open <http://localhost:3000>.

## Live mode (against your own Entra tenant)

### 1. Entra app registrations (automated)

From the repo root:

```powershell
cd agent-identity-portal/scripts
Copy-Item ..\backend\.env.example ..\backend\.env
Copy-Item ..\frontend\.env.example ..\frontend\.env
pwsh ./setup-portal-entra-apps.ps1 -TenantId '<your-tenant-id>' -UpdateEnv
```

This script creates or updates:

1. **Backend (confidential client)** — exposes `api://{backend-client-id}/access_as_user`, grants Microsoft Graph **application** permissions (see [scripts/README.md](scripts/README.md)), and optionally mints a dev client secret.
2. **SPA (public client)** — redirect URIs, delegated `access_as_user` on the backend API, admin consent.

Add a deployed SPA URL:

```powershell
pwsh ./setup-portal-entra-apps.ps1 -TenantId '<id>' `
  -RedirectUris @('https://your-spa.azurestaticapps.net') `
  -FrontendOrigin 'http://localhost:3000,https://your-spa.azurestaticapps.net' `
  -UpdateEnv
```

Use `-SkipClientSecret` when the backend will use managed identity or certificate only.

> The portal's *approvers* are identified by holding the
> **Global Administrator** or **Agent ID Administrator** Entra directory role.

Manual portal steps are still documented in [scripts/README.md](scripts/README.md) (permission table and troubleshooting).

### 2. Backend `.env`

```ini
PROVISIONING_MODE=live
AZURE_TENANT_ID=<your-tenant-id>
AZURE_CLIENT_ID=<backend-app-client-id>
FRONTEND_CLIENT_ID=<spa-app-client-id>
FRONTEND_ORIGIN=https://<your-spa-host>
```

Pick **one** credential option for the backend → Graph calls, in order of
preference:

| Option | When to use | Config |
|--------|-------------|--------|
| **Managed Identity** (recommended in Azure) | Running on App Service / Container Apps / AKS / VM | `USE_MANAGED_IDENTITY=true` (+ optional `MANAGED_IDENTITY_CLIENT_ID` for user-assigned) |
| **Certificate** (recommended off-Azure) | Kubernetes off-cloud, on-prem | `AZURE_CLIENT_CERTIFICATE_PATH=/secure/path/key.pem` and upload the public cert to the app registration |
| **Secret from Key Vault** | Azure-hosted but app registration requires a secret | `AZURE_CLIENT_SECRET_KEYVAULT_URI=https://<vault>.vault.azure.net/secrets/<name>` (combined with Managed Identity to fetch it) |
| **Inline secret** (local dev only) | Demos, laptops | `AZURE_CLIENT_SECRET=<secret>` — never commit |

The backend resolves them in that order; the first configured option wins.
All four fields are blank in `.env.example` by default.

Optional notification settings:

```ini
SMTP_HOST=smtp.office365.com
SMTP_USERNAME=noreply@example.com
SMTP_PASSWORD=<smtp-password>
TEAMS_WEBHOOK_URL=https://<tenant>.webhook.office.com/...
```

### 3. Frontend `.env`

```ini
REACT_APP_API_BASE_URL=https://<your-backend-host>
REACT_APP_AZURE_CLIENT_ID=<spa-app-client-id>
REACT_APP_API_CLIENT_ID=<backend-api-client-id>
REACT_APP_AZURE_TENANT_ID=<your-tenant-id>
REACT_APP_AZURE_REDIRECT_URI=https://<your-spa-host>
```

`REACT_APP_AZURE_CLIENT_ID` is the SPA (MSAL sign-in). `REACT_APP_API_CLIENT_ID` is the backend API app id used for `api://{id}/access_as_user` tokens.

## API reference

| Method | Path | Auth | Description |
|--------|------|------|-------------|
| GET | `/api/health` | none | Health + mode |
| GET | `/api/blueprints` | user | List local blueprints |
| GET | `/api/blueprints/tenant` | user | List Entra-tenant blueprints via Graph |
| POST | `/api/blueprints` | user | Submit a blueprint (approvers create as `active`) |
| PATCH | `/api/blueprints/{id}/deprecate` | approver | Deprecate a blueprint |
| GET/POST | `/api/identity-requests` | user | List/submit identity requests |
| GET/POST | `/api/permission-requests` | user | List/submit permission requests |
| GET | `/api/permission-requests/blocked-permissions` | user | Scopes blocked for agent identities |
| GET | `/api/approvals/pending` | approver | All pending requests |
| POST | `/api/approvals/blueprint/{id}` | approver | Approve/reject blueprint |
| POST | `/api/approvals/identity/{id}` | approver | Approve/reject identity |
| POST | `/api/approvals/permission/{id}` | approver | Approve/reject permission |
| GET | `/api/admin/audit-log` | approver | Audit log (max 500 rows) |
| GET | `/api/admin/stats` | user | Dashboard counts |

OpenAPI spec is served at `/openapi.json` **only when** `ENABLE_DOCS=true`.

## Security notes

- All endpoints (except `/api/health`) require a valid MSAL access token whose
  audience is the **backend** API app (`api://{AZURE_CLIENT_ID}`); unauthenticated
  requests return 401.
- Approver-only routes additionally require the caller to hold the
  Global Administrator or Agent ID Administrator Entra directory role.
- Token validation uses PyJWT with Entra's JWKS endpoint; audience and
  issuer are enforced.
- CORS is locked to the configured `FRONTEND_ORIGIN` (comma-separated list
  supported).
- Standard security headers (HSTS, XFO, XCTO, Referrer-Policy,
  Permissions-Policy) are set on every response.
- Blocked-permission list in `backend/services/provisioning.py` rejects
  admin-tier and high-risk Graph scopes from both blueprints and permission
  requests before anything is written to Entra.
- Secrets **must never** be committed. The bundled `.gitignore` blocks
  `.env`, `*.sqlite`, and similar patterns.

Report vulnerabilities via the [repository SECURITY.md](../SECURITY.md)
process — do not file public GitHub issues.

## License

Licensed under the repository-level [MIT LICENSE](../LICENSE).

## Trademarks

This project may contain trademarks or logos for projects, products, or
services. Authorized use of Microsoft trademarks or logos is subject to and
must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project
must not cause confusion or imply Microsoft sponsorship. Any use of
third-party trademarks or logos is subject to those third parties' policies.
