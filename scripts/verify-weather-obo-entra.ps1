# =============================================================================
# Verify Entra state for OBO → custom Weather API (AADSTS65001 diagnosis)
# =============================================================================
# Exit 0 = OBO prerequisites look present; non-zero = missing scope and/or grant.
#
#   cd scripts
#   pwsh ./verify-weather-obo-entra.ps1
# =============================================================================

param(
    [string]$EnvFile = (Join-Path $PSScriptRoot "..\sidecar\dev\.env"),
    [string]$DelegatedScopeValue = "weather.read"
)

function Read-DotEnvKey {
    param([string]$Path, [string]$Key)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ($line -match "^\s*$([regex]::Escape($Key))\s*=\s*(.+)\s*$") {
            return $Matches[1].Trim().Trim('"').Trim("'")
        }
    }
    return $null
}

$TenantId = Read-DotEnvKey -Path $EnvFile -Key "TENANT_ID"
$AgentAppId = Read-DotEnvKey -Path $EnvFile -Key "AGENT_CLIENT_ID"
$WeatherAppId = Read-DotEnvKey -Path $EnvFile -Key "WEATHER_API_APP_ID"

if (-not $TenantId -or -not $AgentAppId -or -not $WeatherAppId) {
    Write-Error "Need TENANT_ID, AGENT_CLIENT_ID, WEATHER_API_APP_ID in $EnvFile"
    exit 2
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
$ctx = Get-MgContext -ErrorAction SilentlyContinue
if (-not $ctx) {
    Connect-MgGraph -Scopes @(
        "Application.Read.All",
        "DelegatedPermissionGrant.Read.All"
    ) -TenantId $TenantId -NoWelcome
}

$exitCode = 0
Write-Host ""
Write-Host "OBO Entra checklist (custom Weather API)"
Write-Host "  Tenant:       $TenantId"
Write-Host "  Agent appId:  $AgentAppId"
Write-Host "  Weather appId: $WeatherAppId"
Write-Host ""

$app = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications(appId='$WeatherAppId')?`$select=id,api,identifierUris"

$scopes = @()
if ($app.api.oauth2PermissionScopes) {
    $scopes = @($app.api.oauth2PermissionScopes | ForEach-Object { $_.value })
}
$hasScope = $scopes -contains $DelegatedScopeValue

if ($hasScope) {
    Write-Host "[PASS] Weather API exposes delegated scope '$DelegatedScopeValue'"
} else {
    Write-Host "[FAIL] Weather API missing delegated scope '$DelegatedScopeValue' (found: $($scopes -join ', '))"
    $exitCode = 1
}

$agentSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AgentAppId'&`$select=id,displayName"
$weatherSp = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$WeatherAppId'&`$select=id"

if (-not $agentSp.value -or -not $weatherSp.value) {
    Write-Host "[FAIL] Agent or Weather service principal not found"
    exit 3
}

$agentSpId = $agentSp.value[0].id
$weatherSpId = $weatherSp.value[0].id

$grants = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$agentSpId' and resourceId eq '$weatherSpId'"

$scopeOnGrant = $false
if ($grants.value -and $grants.value.Count -gt 0) {
    foreach ($g in $grants.value) {
        Write-Host "       Grant scope: $($g.scope) (consent: $($g.consentType))"
        if (($g.scope -split '\s+') -contains $DelegatedScopeValue) {
            $scopeOnGrant = $true
        }
    }
}

if ($scopeOnGrant) {
    Write-Host "[PASS] Agent has admin consent (oauth2PermissionGrant) for Weather API"
} else {
    Write-Host "[FAIL] No oauth2PermissionGrant for Agent -> Weather with '$DelegatedScopeValue'"
    Write-Host "       (AADSTS65001 at sidecar OBO is expected until this is fixed)"
    $exitCode = 1
}

$appRoleCount = (Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$agentSpId/appRoleAssignments?`$filter=resourceId eq $weatherSpId").value.Count
Write-Host "[INFO] Agent appRoleAssignments to Weather API: $appRoleCount (autonomous only; does not fix OBO)"
Write-Host ""

if ($exitCode -eq 0) {
    Write-Host "Entra looks ready for OBO. Retry UI OBO after: docker compose --env-file .env up -d --force-recreate sidecar"
} else {
    Write-Host "Fix: pwsh ./setup-weather-api-resource.ps1 -UpdateEnv  (Steps 3b + 6)"
}
Write-Host ""
exit $exitCode
