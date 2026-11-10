# =============================================================================
# Register Weather API resource (custom audience) for Agent Identity (autonomous)
# =============================================================================
# Creates a dedicated Entra app registration that represents the sample
# weather-api downstream. Exposes an application permission (app role)
# Weather.Read, assigns it to your Agent Identity service principal, and
# prints the sidecar scope + compose snippet.
#
# Prerequisites:
#   - Agent Identity already created (AGENT_CLIENT_ID in .env)
#   - Microsoft.Graph.Authentication (Connect-MgGraph)
#   - Entra role: Application Administrator or Global Administrator (or
#     Application.ReadWrite.All + AppRoleAssignment.ReadWrite.All via Graph)
#
# Usage:
#   cd scripts
#   pwsh ./setup-weather-api-resource.ps1 -UpdateEnv
#
#   pwsh ./setup-weather-api-resource.ps1 `
#     -TenantId '<tenant-guid>' `
#     -AgentAppId '<agent-client-id>' `
#     -EnvFile '..\sidecar\dev\.env' `
#     -UpdateEnv
#
# Also adds delegated scope + admin consent for OBO (fixes AADSTS65001).
# =============================================================================

param(
    [string]$TenantId,
    [string]$AgentAppId,
    [string]$EnvFile = (Join-Path $PSScriptRoot "..\sidecar\dev\.env"),
    [string]$WeatherApiAppId,
    [string]$AppDisplayName = "Entra Agent ID Sample - Weather API",
    [string]$AppRoleValue = "Weather.Read",
    [string]$DelegatedScopeValue = "weather.read",
    [switch]$UpdateEnv
)

function Set-DotEnvValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Value,
        [string]$Comment
    )

    $lines = @()
    if (Test-Path -LiteralPath $Path) {
        $lines = @(Get-Content -LiteralPath $Path)
    }

    $found = $false
    $updated = foreach ($line in $lines) {
        if ($line -match "^\s*$([regex]::Escape($Key))\s*=") {
            $found = $true
            "$Key=$Value"
        } else {
            $line
        }
    }

    if (-not $found) {
        if ($updated.Count -gt 0 -and $updated[-1] -ne "") {
            $updated += ""
        }
        if ($Comment) {
            $updated += $Comment
        }
        $updated += "$Key=$Value"
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    Set-Content -LiteralPath $Path -Value $updated -Encoding utf8
}

function Read-DotEnvKey {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Key
    )
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ($line -match "^\s*$([regex]::Escape($Key))\s*=\s*(.+)\s*$") {
            return $Matches[1].Trim().Trim('"').Trim("'")
        }
    }
    return $null
}

# Avoid "$id?`$select=..." in double-quoted strings — PowerShell 7 parses ?$ as
# optional chaining and breaks the object ID segment in Graph URIs.
function Format-GraphResourceUri {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Query
    )
    $uri = "https://graph.microsoft.com/v1.0/$Path"
    if ($Query) {
        $uri += "?$Query"
    }
    return $uri
}

function ConvertTo-GraphOAuth2Scope {
    param([Parameter(Mandatory = $true)]$Scope)
    return [ordered]@{
        adminConsentDescription = [string]$Scope.adminConsentDescription
        adminConsentDisplayName = [string]$Scope.adminConsentDisplayName
        id                      = [string]$Scope.id
        isEnabled               = [bool]$Scope.isEnabled
        type                    = [string]$Scope.type
        userConsentDescription  = [string]$Scope.userConsentDescription
        userConsentDisplayName  = [string]$Scope.userConsentDisplayName
        value                   = [string]$Scope.value
    }
}

function ConvertTo-GraphAppRoleDefinition {
    param(
        [Parameter(Mandatory = $true)]$Role
    )
    $allowed = @($Role.allowedMemberTypes | ForEach-Object { [string]$_ })
    if ($allowed.Count -eq 0) {
        $allowed = @("Application")
    }
    return [ordered]@{
        allowedMemberTypes = $allowed
        description        = [string]$Role.description
        displayName        = [string]$Role.displayName
        id                 = [string]$Role.id
        isEnabled          = [bool]$Role.isEnabled
        value              = [string]$Role.value
    }
}

if (-not $TenantId) {
    $TenantId = Read-DotEnvKey -Path $EnvFile -Key "TENANT_ID"
    if (-not $TenantId) {
        Write-Error "TENANT_ID not found in $EnvFile and -TenantId was not provided"
        exit 1
    }
}

if (-not $AgentAppId) {
    $AgentAppId = Read-DotEnvKey -Path $EnvFile -Key "AGENT_CLIENT_ID"
    if (-not $AgentAppId) {
        Write-Error "AGENT_CLIENT_ID not found in $EnvFile and -AgentAppId was not provided"
        exit 1
    }
}

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  Weather API resource (custom audience) — autonomous + OBO"
Write-Host "  Tenant:     $TenantId"
Write-Host "  Agent:      $AgentAppId"
Write-Host "  .env:       $EnvFile"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host ""

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Error @"
Microsoft Graph PowerShell SDK is not installed.

  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery
"@
    exit 1
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "🔐 Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes @(
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "DelegatedPermissionGrant.ReadWrite.All"
) -TenantId $TenantId -NoWelcome
Write-Host "   ✅ Connected"
Write-Host ""

# ── Step 1: Resolve or create Weather API application ─────────────────────
Write-Host "📋 Step 1: Weather API app registration..."

$weatherAppObjectId = $null
$weatherApiAppIdResolved = $WeatherApiAppId

if ($weatherApiAppIdResolved) {
    $appResult = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$weatherApiAppIdResolved'&`$select=id,appId,displayName,identifierUris,appRoles"
    if (-not $appResult.value -or $appResult.value.Count -eq 0) {
        Write-Error "No application found for -WeatherApiAppId $weatherApiAppIdResolved"
        exit 1
    }
    $weatherApp = $appResult.value[0]
    Write-Host "   ✅ Using existing app: $($weatherApp.displayName) ($weatherApiAppIdResolved)"
} else {
    $filterName = $AppDisplayName.Replace("'", "''")
    $existing = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$filterName'&`$select=id,appId,displayName,identifierUris,appRoles"

    if ($existing.value -and $existing.value.Count -gt 0) {
        $weatherApp = $existing.value[0]
        $weatherApiAppIdResolved = $weatherApp.appId
        Write-Host "   ✅ Found by display name: $weatherApiAppIdResolved"
    } else {
        Write-Host "   Creating: $AppDisplayName"
        $createBody = @{
            displayName    = $AppDisplayName
            signInAudience = "AzureADMyOrg"
        } | ConvertTo-Json
        $created = Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/applications" `
            -Body $createBody -ContentType "application/json"
        $weatherAppObjectId = $created.id
        $weatherApiAppIdResolved = $created.appId
        $weatherApp = @{
            id = $weatherAppObjectId
            appId = $weatherApiAppIdResolved
            identifierUris = @()
            appRoles = @()
        }
        Write-Host "   ✅ Created app ID: $weatherApiAppIdResolved"
        Start-Sleep -Seconds 5
    }
}

if (-not $weatherAppObjectId) {
    $weatherAppObjectId = $weatherApp.id
}
if (-not $weatherApiAppIdResolved) {
    $weatherApiAppIdResolved = $weatherApp.appId
}

$appIdUri = "api://$weatherApiAppIdResolved"
$scopeForSidecar = "$appIdUri/.default"

# ── Step 2: Application ID URI ───────────────────────────────────────────────
Write-Host "🌐 Step 2: Application ID URI → $appIdUri"
$currentUris = @($weatherApp.identifierUris)
if ($currentUris -notcontains $appIdUri) {
  $newUris = @($currentUris + $appIdUri | Select-Object -Unique)
  $body = @{ identifierUris = $newUris } | ConvertTo-Json
  Invoke-MgGraphRequest -Method PATCH `
      -Uri "https://graph.microsoft.com/v1.0/applications/$weatherAppObjectId" `
      -Body $body -ContentType "application/json"
  Write-Host "   ✅ Set identifier URI"
} else {
  Write-Host "   ✅ Already set"
}
Write-Host ""

# ── Step 3: App role Weather.Read (application permission) ─────────────────
Write-Host "🔑 Step 3: App role '$AppRoleValue' (Applications)..."

$appRolesUri = Format-GraphResourceUri -Path "applications/$weatherAppObjectId" -Query '$select=appRoles'
$appRefresh = Invoke-MgGraphRequest -Method GET -Uri $appRolesUri
$existingRoles = @($appRefresh.appRoles)
$role = $existingRoles | Where-Object { $_.value -eq $AppRoleValue }

if ($role) {
    $appRoleId = $role.id
    Write-Host "   ✅ Role already exists (ID: $appRoleId)"
} else {
    $appRoleId = (New-Guid).ToString()
    $newRole = [ordered]@{
        allowedMemberTypes = @("Application")
        description        = "Allows autonomous agents to read weather data from the sample API"
        displayName        = "Read weather"
        id                 = $appRoleId
        isEnabled          = $true
        value              = $AppRoleValue
    }
    $mergedRoles = @(
        foreach ($r in $existingRoles) {
            if ($null -ne $r -and $r.value) {
                ConvertTo-GraphAppRoleDefinition -Role $r
            }
        }
    ) + @((ConvertTo-GraphAppRoleDefinition -Role $newRole))

    $body = @{ appRoles = $mergedRoles } | ConvertTo-Json -Depth 6
    $patchUri = Format-GraphResourceUri -Path "applications/$weatherAppObjectId"
    try {
        Invoke-MgGraphRequest -Method PATCH `
            -Uri $patchUri `
            -Body $body -ContentType "application/json"
    } catch {
        Write-Error "Failed to add app role $AppRoleValue : $($_.Exception.Message)"
        exit 1
    }
    Write-Host "   ✅ Added app role $AppRoleValue (ID: $appRoleId)"
    Start-Sleep -Seconds 3
}
Write-Host ""

# ── Step 3b: Delegated scope (OBO) ──────────────────────────────────────────
Write-Host "🔑 Step 3b: Delegated scope '$DelegatedScopeValue' (OBO)..."

# Do not use "$uri?`$select=..." — PowerShell 7 treats ?$ as optional chaining and breaks the URL.
$apiGetUri = Format-GraphResourceUri -Path "applications/$weatherAppObjectId" -Query '$select=api'
$patchAppUri = Format-GraphResourceUri -Path "applications/$weatherAppObjectId"
$apiSnap = Invoke-MgGraphRequest -Method GET -Uri $apiGetUri
$existingScopes = @()
if ($apiSnap.api -and $apiSnap.api.oauth2PermissionScopes) {
    foreach ($s in @($apiSnap.api.oauth2PermissionScopes)) {
        if ($null -ne $s -and $s.value) { $existingScopes += $s }
    }
}

$scopeAlreadyExists = @($existingScopes | Where-Object { $_.value -eq $DelegatedScopeValue }).Count -gt 0

if ($scopeAlreadyExists) {
    Write-Host "   ✅ Delegated scope already exists"
} else {
    $newScope = [ordered]@{
        id                      = (New-Guid).ToString()
        adminConsentDescription = "Allow the agent to read weather on behalf of the signed-in user"
        adminConsentDisplayName = "Read weather as user"
        isEnabled               = $true
        type                    = "User"
        userConsentDescription  = "Allow the agent to read weather on your behalf"
        userConsentDisplayName  = "Read weather as user"
        value                   = $DelegatedScopeValue
    }
    $scopeByValue = [ordered]@{}
    foreach ($s in $existingScopes) {
        $scopeByValue[[string]$s.value] = ConvertTo-GraphOAuth2Scope -Scope $s
    }
    $scopeByValue[$DelegatedScopeValue] = ConvertTo-GraphOAuth2Scope -Scope $newScope
    $allScopes = @($scopeByValue.Values)

    $tokenVersion = 2
    if ($apiSnap.api -and $apiSnap.api.requestedAccessTokenVersion) {
        $tokenVersion = [int]$apiSnap.api.requestedAccessTokenVersion
    }
    $scopeBody = @{
        api = @{
            requestedAccessTokenVersion = $tokenVersion
            oauth2PermissionScopes      = $allScopes
        }
    } | ConvertTo-Json -Depth 12

    try {
        Invoke-MgGraphRequest -Method PATCH -Uri $patchAppUri -Body $scopeBody -ContentType "application/json"
        Write-Host "   ✅ Added scope $appIdUri/$DelegatedScopeValue"
        Start-Sleep -Seconds 5
    } catch {
        $detail = [string]$_.ErrorDetails.Message
        if ($detail -match 'DuplicateValue') {
            Write-Host "   ✅ Delegated scope already present in tenant (DuplicateValue)"
        } else {
            if ($detail) { Write-Host "   Graph: $detail" }
            Write-Error "Failed to add delegated scope: $($_.Exception.Message)"
            exit 1
        }
    }
}
Write-Host ""

# ── Step 4: Service principal for Weather API ─────────────────────────────
Write-Host "👤 Step 4: Weather API service principal..."
$spResult = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$weatherApiAppIdResolved'&`$select=id,appId"

if ($spResult.value -and $spResult.value.Count -gt 0) {
    $weatherSpId = $spResult.value[0].id
    Write-Host "   ✅ Already exists ($weatherSpId)"
} else {
    $spBody = @{ appId = $weatherApiAppIdResolved } | ConvertTo-Json
    $spCreated = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
        -Body $spBody -ContentType "application/json"
    $weatherSpId = $spCreated.id
    Write-Host "   ✅ Created ($weatherSpId)"
    Start-Sleep -Seconds 5
}

# Re-read app role ID from enterprise SP if app was just patched (SP appRoles sync)
$weatherSpRolesUri = Format-GraphResourceUri -Path "servicePrincipals/$weatherSpId" -Query '$select=appRoles'
$weatherSpFull = Invoke-MgGraphRequest -Method GET -Uri $weatherSpRolesUri
$spRole = $weatherSpFull.appRoles | Where-Object { $_.value -eq $AppRoleValue }
if ($spRole) {
    $appRoleId = $spRole.id
}
Write-Host ""

# ── Step 5: Assign app role to Agent Identity SP ─────────────────────────
Write-Host "🛡️  Step 5: Grant $AppRoleValue → Agent Identity..."

$agentSpResult = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AgentAppId'&`$select=id,displayName"

if (-not $agentSpResult.value -or $agentSpResult.value.Count -eq 0) {
    Write-Error "Agent service principal not found for appId $AgentAppId. Run Start-EntraAgentIDWorkflow first."
    exit 1
}

$agentSpId = $agentSpResult.value[0].id
Write-Host "   Agent SP: $($agentSpResult.value[0].displayName) ($agentSpId)"

$assignments = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$agentSpId/appRoleAssignments"

$alreadyAssigned = $assignments.value | Where-Object {
    $_.resourceId -eq $weatherSpId -and $_.appRoleId -eq $appRoleId
}

if ($alreadyAssigned) {
    Write-Host "   ✅ App role assignment already exists"
} else {
    $assignBody = @{
        principalId = $agentSpId
        resourceId  = $weatherSpId
        appRoleId   = $appRoleId
    } | ConvertTo-Json
    try {
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$agentSpId/appRoleAssignments" `
            -Body $assignBody -ContentType "application/json" | Out-Null
        Write-Host "   ✅ Assigned $AppRoleValue to agent"
    } catch {
        Write-Host "   ⚠️  Assignment failed: $($_.Exception.Message)"
        Write-Host "   Admin consent URL (open signed in as Global Admin):"
        $consentRole = [uri]::EscapeDataString("$appIdUri/$AppRoleValue")
        Write-Host "   https://login.microsoftonline.com/$TenantId/v2.0/adminconsent?client_id=$AgentAppId&role=$consentRole&redirect_uri=https%3A%2F%2Fentra.microsoft.com%2FTokenAuthorize"
    }
}
Write-Host ""

# ── Step 6: Admin consent — Agent → Weather API (OBO) ─────────────────────
Write-Host "🛡️  Step 6: Admin consent Agent → Weather API (OBO)..."

try {
    $grantFilter = "clientId eq '$agentSpId' and resourceId eq '$weatherSpId'"
    $existingGrant = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=$grantFilter"
    $grant = $null
    if ($existingGrant.value -and $existingGrant.value.Count -gt 0) {
        $grant = $existingGrant.value[0]
    }
    if ($grant -and ($grant.scope -split '\s+') -contains $DelegatedScopeValue) {
        Write-Host "   ✅ Already consented (scope: $($grant.scope))"
    } elseif ($grant) {
        $mergedScope = ($grant.scope.Trim() + " $DelegatedScopeValue").Trim() -replace '\s+', ' '
        Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" `
            -Body (@{ scope = $mergedScope } | ConvertTo-Json) `
            -ContentType "application/json"
        Write-Host "   ✅ Updated grant: $mergedScope"
    } else {
        $grantBody = @{
            clientId    = $agentSpId
            consentType = "AllPrincipals"
            resourceId  = $weatherSpId
            scope       = $DelegatedScopeValue
        } | ConvertTo-Json
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
            -Body $grantBody -ContentType "application/json" | Out-Null
        Write-Host "   ✅ Granted Agent → Weather API ($DelegatedScopeValue)"
    }
} catch {
    Write-Host "   ⚠️  OBO consent failed: $($_.Exception.Message)"
}
Write-Host ""

# ── Step 7: Optional .env ──────────────────────────────────────────────────
if ($UpdateEnv) {
    Write-Host "📝 Step 7: Updating $EnvFile ..."
    Set-DotEnvValue -Path $EnvFile -Key "WEATHER_API_APP_ID" -Value $weatherApiAppIdResolved `
        -Comment "# Custom API audience for weather-api (see setup-weather-api-resource.ps1)"
    Write-Host "   ✅ WEATHER_API_APP_ID"
    Write-Host ""
}

# ── Summary ───────────────────────────────────────────────────────────────
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  ✅ Weather API resource ready (autonomous + OBO)"
Write-Host ""
Write-Host "  Weather API App ID:     $weatherApiAppIdResolved"
Write-Host "  Application ID URI:   $appIdUri"
Write-Host "  App role:               $AppRoleValue"
Write-Host "  Delegated scope (OBO):  $DelegatedScopeValue"
Write-Host "  Sidecar scope (both):   $scopeForSidecar"
Write-Host ""
Write-Host "  Verify: pwsh ./verify-weather-obo-entra.ps1"
Write-Host ""
Write-Host "  Optional weather-api audience check (follow-up in Flask app):"
Write-Host "    EXPECTED_AUDIENCE=$appIdUri"
Write-Host ""
Write-Host "  Recreate sidecar after compose change:"
Write-Host "    docker compose --env-file .env up -d --force-recreate sidecar"
Write-Host ""
Write-Host "  Verify autonomous:"
Write-Host "    curl -s -X POST http://localhost:3003/api/chat -H 'Content-Type: application/json' \\"
Write-Host "      -d '{\"message\":\"weather in seattle\",\"use_langchain\":false,\"token_flow\":\"autonomous\"}'"
Write-Host "    TR aud should be $appIdUri ; roles should include $AppRoleValue"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
