# Agent Identity Portal — Entra bootstrap

PowerShell tooling to register the portal **backend API** and **SPA** in Microsoft Entra ID, grant Microsoft Graph and Agent ID **application** permissions to the backend, wire delegated `access_as_user` from the SPA to the API, and optionally update local `.env` files.

## Prerequisites

- **PowerShell 7.4+** (`pwsh`)
- **Microsoft.Graph.Authentication** — `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`
- **Entra directory role** on the account that signs in — **Global Administrator** or **Agent ID Administrator** (portal approvers use these same roles at runtime)

The script uses delegated Graph scopes for **setup only**:

`Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`, `DelegatedPermissionGrant.ReadWrite.All`, `Directory.Read.All`, `User.Read`

## Quick start

```powershell
cd agent-identity-portal/scripts

# Copy env templates if you have not already
Copy-Item ..\backend\.env.example ..\backend\.env
Copy-Item ..\frontend\.env.example ..\frontend\.env

pwsh ./setup-portal-entra-apps.ps1 -TenantId '<your-tenant-guid>' -UpdateEnv
```

Sign in when the browser opens. The script is idempotent: re-running updates redirect URIs, API permissions, and consent without creating duplicate apps (matched by display name).

## Parameters

| Parameter | Purpose |
|-----------|---------|
| `-TenantId` | Target tenant. If omitted, reads `AZURE_TENANT_ID` from `backend/.env`. |
| `-RedirectUri` | Primary SPA redirect (default `http://localhost:3000`). |
| `-RedirectUris` | Additional redirects (e.g. Azure Static Web Apps URL). Merged with existing URIs on the app. |
| `-FrontendOrigin` | CORS value for `FRONTEND_ORIGIN`. Defaults to comma-separated redirect URIs. |
| `-UpdateEnv` | Patch `backend/.env` and `frontend/.env` (creates from `.env.example` if missing). |
| `-SkipClientSecret` | Do not call `addPassword`; use certificate, managed identity, or Key Vault instead. |
| `-BackendDisplayName` / `-SpaDisplayName` | App registration display names (defaults in script header). |

## What gets configured

### Backend API (confidential client)

- Application ID URI: `api://{backend-app-id}`
- Exposed delegated scope: `access_as_user`
- Microsoft Graph **application** permissions (admin consent on the backend service principal):

| Permission | Purpose |
|------------|---------|
| `Application.ReadWrite.OwnedBy` | Create/update apps the portal owns |
| `Directory.Read.All` | Resolve sign-in user directory roles |
| `AgentIdentityBlueprint.CreateAsManager` or `AgentIdentityBlueprint.Create` | Live blueprint provisioning |
| `AgentIdentity.CreateAsManager` | Live agent identity provisioning |
| `AgentIdentityBlueprintPrincipal.Create` | Blueprint principal registration |
| `Application.Read.All` | Read applications / blueprints |
| `AppRoleAssignment.ReadWrite.All` | Grant permissions to agent identities |

Unless `-SkipClientSecret` is set, a client secret named `portal-local-dev` is created and written only to `backend/.env` when `-UpdateEnv` is used.

### SPA (public client)

- Redirect URIs (merged, not replaced)
- Delegated permission: `access_as_user` on the backend API
- Tenant-wide admin consent for SPA → API

### `.env` keys when `-UpdateEnv` is used

**Backend:** `PROVISIONING_MODE=live`, `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `FRONTEND_CLIENT_ID`, `FRONTEND_ORIGIN`, optional `AZURE_CLIENT_SECRET`

**Frontend:** `REACT_APP_AZURE_CLIENT_ID` (SPA), `REACT_APP_API_CLIENT_ID` (backend, for token scope), tenant, redirect URI, `REACT_APP_API_BASE_URL`

## Production notes

- Prefer **managed identity** or **certificate** for the backend in Azure; use `-SkipClientSecret` for hosted environments.
- Add deployed SPA and API hosts with `-RedirectUris` and `-FrontendOrigin` (comma-separated origins match backend CORS).
- Never commit `.env` files or secrets.

## Troubleshooting

| Symptom | Check |
|---------|--------|
| 401 from API after sign-in | `REACT_APP_API_CLIENT_ID` must be the **backend** app id; token `aud` is `api://{backend-id}`. |
| 403 on approver routes | User needs Global Administrator or Agent ID Administrator (not only Application Administrator). |
| Live provision 403 | Backend app Graph permissions — re-run this script to refresh admin consent. |
| Graph role not found | Tenant may not expose an Agent ID app role yet; upgrade Graph or check Agent ID preview availability. |
