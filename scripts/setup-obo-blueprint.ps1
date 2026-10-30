# =============================================================================
# Setup Blueprint for OBO Flow (PowerShell)
# =============================================================================
# Following: https://blog.christianposta.com/entra-agent-id-agw/PART-2.html
#
# Uses Connect-MgGraph with specific scopes (avoids Directory.AccessAsUser.All).
# Uses Invoke-MgGraphRequest for all calls (avoids module version conflicts).
# =============================================================================

param(
    [Parameter(Mandatory=$true)][string]$BlueprintAppId,
    [Parameter(Mandatory=$true)][string]$AgentAppId,
    [Parameter(Mandatory=$true)][string]$ClientSpaAppId,
    [Parameter(Mandatory=$true)][string]$TenantId
)

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  Setup Blueprint for OBO Flow"
Write-Host "  Blueprint:  $BlueprintAppId"
Write-Host "  Agent:      $AgentAppId"  
Write-Host "  Client SPA: $ClientSpaAppId"
Write-Host "  Tenant:     $TenantId"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host ""

# ── Import and connect ──────────────────────────────────────────────────────
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Error "Install Microsoft Graph PowerShell: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser"
    exit 1
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "🔐 Connecting to Microsoft Graph..."
Write-Host "   (A browser window will open for sign-in)"
Connect-MgGraph -Scopes @(
    "AgentIdentityBlueprint.ReadWrite.All",
    "Application.ReadWrite.All"
) -TenantId $TenantId -NoWelcome

Write-Host "   ✅ Connected"
Write-Host ""

# ── Step 1: Get Blueprint ──────────────────────────────────────────────────
Write-Host "📋 Step 1: Looking up Blueprint..."
$result = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/beta/applications?`$filter=appId eq '$BlueprintAppId'&`$select=id,displayName,identifierUris,api"

$blueprintApp = $result.value[0]
$blueprintObjectId = $blueprintApp.id

Write-Host "   ✅ Found: $($blueprintApp.displayName)"
Write-Host "   Object ID: $blueprintObjectId"
Write-Host "   Current URIs: $($blueprintApp.identifierUris -join ', ')"
$existingScopes = $blueprintApp.api.oauth2PermissionScopes
Write-Host "   Current scopes: $(($existingScopes | ForEach-Object { $_.value }) -join ', ')"
Write-Host ""

# ── Step 2: Set Application ID URI ─────────────────────────────────────────
$appIdUri = "api://$BlueprintAppId"
Write-Host "🌐 Step 2: Setting Application ID URI → $appIdUri"

if ($blueprintApp.identifierUris -contains $appIdUri) {
    Write-Host "   ✅ Already set"
} else {
    $body = @{ identifierUris = @($appIdUri) } | ConvertTo-Json
    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/beta/applications/$blueprintObjectId" `
        -Body $body -ContentType "application/json"
    Write-Host "   ✅ Set App ID URI to: $appIdUri"
}
Write-Host ""

# ── Step 3: Add access_as_user scope ───────────────────────────────────────
Write-Host "🔑 Step 3: Adding 'access_as_user' delegated scope..."

$hasAccessAsUser = $existingScopes | Where-Object { $_.value -eq "access_as_user" }

if ($hasAccessAsUser) {
    Write-Host "   ✅ Scope already exists (ID: $($hasAccessAsUser.id))"
} else {
    $newScopeId = (New-Guid).ToString()

    $body = @{
        api = @{
            oauth2PermissionScopes = @(
                @{
                    id = $newScopeId
                    adminConsentDescription = "Allow the application to access the agent on behalf of the signed-in user"
                    adminConsentDisplayName = "Access agent as user"
                    isEnabled = $true
                    type = "User"
                    userConsentDescription = "Allow the application to access the agent on your behalf"
                    userConsentDisplayName = "Access agent as user"
                    value = "access_as_user"
                }
            )
        }
    } | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/beta/applications/$blueprintObjectId" `
        -Body $body -ContentType "application/json"
    Write-Host "   ✅ Added 'access_as_user' scope (ID: $newScopeId)"
    Write-Host "   Full scope: $appIdUri/access_as_user"
}
Write-Host ""

# ── Step 4: Add API permission on Client SPA ───────────────────────────────
Write-Host "📝 Step 4: Adding API permission on Client SPA → Blueprint scope..."

# Re-query to get updated scope ID
$bpUpdated = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/beta/applications?`$filter=appId eq '$BlueprintAppId'&`$select=id,api"
$scopeId = ($bpUpdated.value[0].api.oauth2PermissionScopes | Where-Object { $_.value -eq "access_as_user" }).id

# Get Client SPA object ID
$spaResult = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$ClientSpaAppId'&`$select=id"
$clientSpaObjectId = $spaResult.value[0].id

if ($scopeId -and $clientSpaObjectId) {
    $body = @{
        requiredResourceAccess = @(
            @{
                resourceAppId = $BlueprintAppId
                resourceAccess = @(
                    @{
                        id = $scopeId
                        type = "Scope"
                    }
                )
            }
        )
    } | ConvertTo-Json -Depth 5

    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/applications/$clientSpaObjectId" `
        -Body $body -ContentType "application/json"
    Write-Host "   ✅ Added API permission on Client SPA → Blueprint access_as_user"
} else {
    Write-Host "   ⚠️ Missing scopeId=$scopeId or clientSpaObjectId=$clientSpaObjectId"
}
Write-Host ""

# ── Step 5: Grant admin consent for Agent Identity → Graph ──────────────────
Write-Host "🛡️  Step 5: Granting admin consent for Agent Identity → Graph..."

$graphSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'&`$select=id"
$graphSpId = $graphSp.value[0].id

$agentSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AgentAppId'&`$select=id"
$agentSpId = $agentSp.value[0].id

if ($agentSpId -and $graphSpId) {
    try {
        $existingConsent = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$agentSpId' and resourceId eq '$graphSpId'"
        
        if ($existingConsent.value.Count -gt 0) {
            Write-Host "   ✅ Admin consent already exists"
            Write-Host "   Scopes: $($existingConsent.value[0].scope)"
        } else {
            $body = @{
                clientId = $agentSpId
                consentType = "AllPrincipals"
                resourceId = $graphSpId
                scope = "User.Read openid profile offline_access"
            } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
                -Body $body -ContentType "application/json"
            Write-Host "   ✅ Admin consent granted for Agent Identity → Graph"
        }
    } catch {
        Write-Host "   ⚠️ Error: $($_.Exception.Message)"
    }
} else {
    Write-Host "   ⚠️ Could not find service principals (Agent: $agentSpId, Graph: $graphSpId)"
}
Write-Host ""

# ── Step 6: Grant admin consent for Client SPA → Blueprint ──────────────────
Write-Host "🛡️  Step 6: Granting admin consent for Client SPA → Blueprint..."

$blueprintSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$BlueprintAppId'&`$select=id"
$blueprintSpId = $blueprintSp.value[0].id

$clientSpaSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$ClientSpaAppId'&`$select=id"
$clientSpaSpId = $clientSpaSp.value[0].id

if ($clientSpaSpId -and $blueprintSpId) {
    try {
        $existingConsent2 = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$clientSpaSpId' and resourceId eq '$blueprintSpId'"

        if ($existingConsent2.value.Count -gt 0) {
            Write-Host "   ✅ Admin consent already exists"
        } else {
            $body = @{
                clientId = $clientSpaSpId
                consentType = "AllPrincipals"
                resourceId = $blueprintSpId
                scope = "access_as_user"
            } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
                -Body $body -ContentType "application/json"
            Write-Host "   ✅ Admin consent granted for Client SPA → Blueprint"
        }
    } catch {
        Write-Host "   ⚠️ Error: $($_.Exception.Message)"
    }
} else {
    Write-Host "   ⚠️ Could not find service principals (Client SPA: $clientSpaSpId, Blueprint: $blueprintSpId)"
}
Write-Host ""

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  ✅ OBO Blueprint Setup Complete"
Write-Host ""
Write-Host "  Blueprint API:  $appIdUri"
Write-Host "  OBO Scope:      $appIdUri/access_as_user"
Write-Host ""
Write-Host "  MSAL.js requests: scopes=['$appIdUri/access_as_user']"
Write-Host "  User token (Tc) aud: $appIdUri"
Write-Host "  Sidecar AzureAd__ClientId: $BlueprintAppId → aud matches ✅"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
