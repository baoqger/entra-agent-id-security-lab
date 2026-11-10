# =============================================================================
# Agent Identity Portal — Entra app registrations (backend API + SPA)
# =============================================================================
# Creates (or updates) the confidential portal API app and public SPA, exposes
# access_as_user, grants Microsoft Graph + Agent ID application permissions to
# the backend, admin-consents SPA → API, optionally mints a dev client secret,
# and can patch backend/frontend .env files.
#
# Operator prerequisites:
#   - pwsh 7.4+, Microsoft.Graph.Authentication
#   - Entra directory role: Global Administrator or Agent ID Administrator
#   - Run from repo:  cd agent-identity-portal/scripts
#
# Usage:
#   pwsh ./setup-portal-entra-apps.ps1 -UpdateEnv
#
#   pwsh ./setup-portal-entra-apps.ps1 `
#     -TenantId '<tenant-guid>' `
#     -RedirectUri 'http://localhost:3000' `
#     -RedirectUris @('https://my-spa.azurestaticapps.net') `
#     -UpdateEnv
#
#   pwsh ./setup-portal-entra-apps.ps1 -SkipClientSecret -UpdateEnv
# =============================================================================

param(
    [string]$TenantId,
    [string]$BackendDisplayName = "Agent Identity Portal API",
    [string]$SpaDisplayName = "Agent Identity Portal SPA",
    [string]$RedirectUri = "http://localhost:3000",
    [string[]]$RedirectUris = @(),
    [string]$BackendEnvFile = (Join-Path $PSScriptRoot "..\backend\.env"),
    [string]$FrontendEnvFile = (Join-Path $PSScriptRoot "..\frontend\.env"),
    [string]$FrontendOrigin,
    [switch]$UpdateEnv,
    [switch]$SkipClientSecret
)

$ErrorActionPreference = "Stop"

# Bumped when JSON / Graph request helpers change — confirm this string in console output when debugging.
$SetupPortalScriptBuild = "2026-11-10-json-v2"

$GraphResourceAppId = "00000003-0000-0000-c000-000000000000"

$BackendGraphAppRoles = @(
    "Application.ReadWrite.OwnedBy",
    "Directory.Read.All",
    "AgentIdentityBlueprint.CreateAsManager",
    "AgentIdentityBlueprint.Create",
    "AgentIdentity.CreateAsManager",
    "AgentIdentityBlueprintPrincipal.Create",
    "Application.Read.All",
    "AppRoleAssignment.ReadWrite.All"
)

# Resolved at runtime (CreateAsManager preferred over Create for blueprint)
$script:BackendGraphRolesToGrant = $null

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

function Format-GraphResourceUri {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Query
    )
  # Do not build "applications/{id}?$select=..." in one double-quoted string — PowerShell 7
  # treats "?$" as optional chaining and drops the object id from the URI.
    $uri = "https://graph.microsoft.com/v1.0/$Path"
    if ($Query) {
        $uri += '?' + $Query
    }
    return $uri
}

function Escape-ODataString {
    param([string]$Value)
    return ($Value -replace "'", "''")
}

function Get-ApplicationByDisplayName {
    param([string]$DisplayName)
    $escaped = Escape-ODataString -Value $DisplayName
    $uri = "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$escaped'&`$select=id,appId,displayName"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0]
    }
    return $null
}

function Get-ApplicationByAppId {
    param([string]$AppId, [string]$Select = "id,appId,displayName,identifierUris,api,requiredResourceAccess")
    $uri = "https://graph.microsoft.com/v1.0/applications?`$filter=appId eq '$AppId'&`$select=$Select"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0]
    }
    return $null
}

function Ensure-ServicePrincipal {
    param([string]$AppId)
    $spResult = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$AppId'&`$select=id,appId"
    if ($spResult.value -and $spResult.value.Count -gt 0) {
        return $spResult.value[0].id
    }
    $created = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
        -Body (@{ appId = $AppId } | ConvertTo-Json) `
        -ContentType "application/json"
    Start-Sleep -Seconds 3
    return $created.id
}

function Get-GraphServicePrincipal {
    $uri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$GraphResourceAppId'&`$select=id,appRoles"
    $result = Invoke-MgGraphRequest -Method GET -Uri $uri
    if (-not $result.value -or $result.value.Count -eq 0) {
        throw "Microsoft Graph service principal not found in tenant"
    }
    return $result.value[0]
}

function Resolve-BackendGraphRoleValues {
  param($GraphSpAppRoles)
  $resolved = [System.Collections.Generic.List[string]]::new()
  $wantBlueprint = $false
  foreach ($name in $BackendGraphAppRoles) {
    if ($name -eq "AgentIdentityBlueprint.CreateAsManager") {
      $wantBlueprint = $true
      continue
    }
    if ($name -eq "AgentIdentityBlueprint.Create") {
      continue
    }
    if ($GraphSpAppRoles | Where-Object { $_.value -eq $name }) {
      $resolved.Add($name)
    }
  }
  if ($wantBlueprint) {
    if ($GraphSpAppRoles | Where-Object { $_.value -eq "AgentIdentityBlueprint.CreateAsManager" }) {
      $resolved.Add("AgentIdentityBlueprint.CreateAsManager")
    } elseif ($GraphSpAppRoles | Where-Object { $_.value -eq "AgentIdentityBlueprint.Create" }) {
      $resolved.Add("AgentIdentityBlueprint.Create")
    }
  }
  return @($resolved | Select-Object -Unique)
}

function ConvertTo-GraphResourceAccessItem {
    param($Entry)
    if (-not $Entry -or -not $Entry.id -or -not $Entry.type) {
        return $null
    }
    $type = [string]$Entry.type
    if ($type -ne 'Role' -and $type -ne 'Scope') {
        return $null
    }
    return [ordered]@{
        id   = [string]$Entry.id
        type = $type
    }
}

function Normalize-RequiredResourceAccess {
    param([AllowNull()]$RequiredResourceAccess)
    if (-not $RequiredResourceAccess) {
        return , @()
    }
    $byResource = [ordered]@{}
    foreach ($block in @($RequiredResourceAccess)) {
        $resourceAppId = [string]$block.resourceAppId
        if (-not $resourceAppId) { continue }
        if ($byResource.Keys -notcontains $resourceAppId) {
            $byResource[$resourceAppId] = [ordered]@{}
        }
        $accessList = $block.resourceAccess
        if (-not $accessList) { continue }
        foreach ($entry in @($accessList)) {
            $item = ConvertTo-GraphResourceAccessItem -Entry $entry
            if ($item) {
                $byResource[$resourceAppId][$item.id] = $item
            }
        }
    }
    $normalized = [System.Collections.Generic.List[object]]::new()
    foreach ($resourceAppId in $byResource.Keys) {
        $items = @($byResource[$resourceAppId].Values)
        if ($items.Count -eq 0) { continue }
        $normalized.Add([ordered]@{
            resourceAppId  = [string]$resourceAppId
            resourceAccess = $items
        })
    }
    return , $normalized.ToArray()
}

function Merge-RequiredResourceAccess {
    param(
        $Existing,
        [string]$ResourceAppId,
        [array]$NewAccess
    )
    $normalized = Normalize-RequiredResourceAccess -RequiredResourceAccess $Existing
    $map = [ordered]@{}
    foreach ($block in $normalized) {
        $map[[string]$block.resourceAppId] = $block
    }
    $accessById = @{}
    if ($map.Keys -contains $ResourceAppId) {
        foreach ($entry in $map[$ResourceAppId].resourceAccess) {
            $accessById[$entry.id] = $entry
        }
    }
    foreach ($entry in $NewAccess) {
        $item = ConvertTo-GraphResourceAccessItem -Entry $entry
        if ($item) {
            $accessById[$item.id] = $item
        }
    }
    if ($accessById.Count -gt 0) {
        $map[$ResourceAppId] = [ordered]@{
            resourceAppId  = [string]$ResourceAppId
            resourceAccess = @($accessById.Values)
        }
    }
    return , @($map.Values)
}

function Escape-GraphJsonString {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) {
        return "null"
    }
    $escaped = $Value -replace '\\', '\\\\' -replace '"', '\"' `
        -replace "`r", '\r' -replace "`n", '\n' -replace "`t", '\t'
    return "`"$escaped`""
}

function New-GraphRedirectUrisJsonArray {
    param([Parameter(Mandatory = $true)][string[]]$RedirectUris)
    $uris = @($RedirectUris | Where-Object { $_ } | ForEach-Object { [string]$_ } | Select-Object -Unique)
    if ($uris.Count -eq 0) {
        throw "At least one SPA redirect URI is required"
    }
    $encoded = foreach ($u in $uris) {
        Escape-GraphJsonString -Value $u
    }
    return "[$($encoded -join ',')]"
}

function New-GraphSpaRedirectUrisPatchBody {
    param([Parameter(Mandatory = $true)][string[]]$RedirectUris)
    $array = New-GraphRedirectUrisJsonArray -RedirectUris $RedirectUris
    return "{`"spa`":{`"redirectUris`":$array}}"
}

function New-GraphSpaApplicationCreateBody {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string[]]$RedirectUris
    )
    $nameJson = Escape-GraphJsonString -Value $DisplayName
    $array = New-GraphRedirectUrisJsonArray -RedirectUris $RedirectUris
    return "{`"displayName`":$nameJson,`"signInAudience`":`"AzureADMyOrg`",`"spa`":{`"redirectUris`":$array}}"
}

function New-RequiredResourceAccessPatchBody {
    param([Parameter(Mandatory = $true)]$RequiredResourceAccessBlocks)
    $blocks = @($RequiredResourceAccessBlocks)
    if ($blocks.Count -eq 0) {
        return '{ "requiredResourceAccess": [] }'
    }
    $blockJson = foreach ($block in $blocks) {
        $resourceAppId = Escape-GraphJsonString -Value ([string]$block.resourceAppId)
        $accessJson = foreach ($entry in @($block.resourceAccess)) {
            $id = Escape-GraphJsonString -Value ([string]$entry.id)
            $type = Escape-GraphJsonString -Value ([string]$entry.type)
            "{`"id`":$id,`"type`":$type}"
        }
        "{`"resourceAppId`":$resourceAppId,`"resourceAccess`":[$($accessJson -join ',')]}"
    }
    return "{`"requiredResourceAccess`":[$($blockJson -join ',')]}"
}

function Invoke-ApplicationSpaRedirectUrisPatch {
    param(
        [Parameter(Mandatory = $true)][string]$ApplicationObjectId,
        [Parameter(Mandatory = $true)][string[]]$RedirectUris
    )
    $body = New-GraphSpaRedirectUrisPatchBody -RedirectUris $RedirectUris
    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId" `
        -Body $body -ContentType "application/json"
}

function Invoke-ApplicationRequiredResourceAccessPatch {
    param(
        [Parameter(Mandatory = $true)][string]$ApplicationObjectId,
        [Parameter(Mandatory = $true)]$RequiredResourceAccess
    )
    $normalized = @(Normalize-RequiredResourceAccess -RequiredResourceAccess $RequiredResourceAccess)
    $body = New-RequiredResourceAccessPatchBody -RequiredResourceAccessBlocks $normalized
    try {
        Invoke-MgGraphRequest -Method PATCH `
            -Uri "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId" `
            -Body $body -ContentType "application/json"
    } catch {
        Write-Host "   ⚠️  requiredResourceAccess PATCH failed. Payload:"
        Write-Host $body
        throw
    }
}

function Grant-ApplicationPermission {
    param(
        [string]$ClientSpId,
        [string]$ResourceSpId,
        [string]$AppRoleId,
        [string]$RoleValue
    )
    $assignments = Invoke-MgGraphRequest -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ClientSpId/appRoleAssignments"
    $exists = $assignments.value | Where-Object {
        $_.resourceId -eq $ResourceSpId -and $_.appRoleId -eq $AppRoleId
    }
    if ($exists) {
        Write-Host "   ✅ Already consented: $RoleValue"
        return
    }
    $body = @{
        principalId = $ClientSpId
        resourceId  = $ResourceSpId
        appRoleId   = $AppRoleId
    } | ConvertTo-Json
    try {
        Invoke-MgGraphRequest -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ClientSpId/appRoleAssignments" `
            -Body $body -ContentType "application/json" | Out-Null
        Write-Host "   ✅ Granted application permission: $RoleValue"
    } catch {
        Write-Host "   ⚠️  Failed to grant $RoleValue : $($_.Exception.Message)"
    }
}

function Grant-SpaDelegatedConsent {
    param(
        [string]$SpaSpId,
        [string]$ApiSpId,
        [string]$ScopeValue = "access_as_user"
    )
    try {
        $filter = "clientId eq '$SpaSpId' and resourceId eq '$ApiSpId'"
        $existing = Invoke-MgGraphRequest -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=$filter"
        if ($existing.value -and $existing.value.Count -gt 0) {
            $grant = $existing.value[0]
            $scopes = ($grant.scope -split '\s+' | Where-Object { $_ })
            if ($scopes -contains $ScopeValue) {
                Write-Host "   ✅ SPA admin consent already includes $ScopeValue"
                return
            }
            $merged = ($grant.scope.Trim() + " $ScopeValue").Trim() -replace '\s+', ' '
            Invoke-MgGraphRequest -Method PATCH `
                -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($grant.id)" `
                -Body (@{ scope = $merged } | ConvertTo-Json) `
                -ContentType "application/json"
            Write-Host "   ✅ Updated SPA grant: $merged"
        } else {
            $body = @{
                clientId    = $SpaSpId
                consentType = "AllPrincipals"
                resourceId  = $ApiSpId
                scope       = $ScopeValue
            } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
                -Body $body -ContentType "application/json" | Out-Null
            Write-Host "   ✅ Admin consent granted SPA → API ($ScopeValue)"
        }
    } catch {
        Write-Host "   ⚠️  SPA consent failed: $($_.Exception.Message)"
    }
}

# ── Resolve tenant ─────────────────────────────────────────────────────────
$backendExample = Join-Path $PSScriptRoot "..\backend\.env.example"
if (-not $TenantId) {
    $TenantId = Read-DotEnvKey -Path $BackendEnvFile -Key "AZURE_TENANT_ID"
    if (-not $TenantId) {
        $TenantId = Read-DotEnvKey -Path $backendExample -Key "AZURE_TENANT_ID"
    }
}
if (-not $TenantId -or $TenantId -match 'your-tenant') {
    Write-Error "Provide -TenantId or set AZURE_TENANT_ID in $BackendEnvFile"
    exit 1
}

$allRedirectUris = @($RedirectUri) + @($RedirectUris) | Where-Object { $_ } | Select-Object -Unique
if (-not $FrontendOrigin) {
    $FrontendOrigin = ($allRedirectUris -join ",")
}

$createClientSecret = -not $SkipClientSecret

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  Agent Identity Portal — Entra setup"
Write-Host "  Build:   $SetupPortalScriptBuild"
Write-Host "  Tenant:  $TenantId"
Write-Host "  Backend: $BackendDisplayName"
Write-Host "  SPA:     $SpaDisplayName"
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
    "DelegatedPermissionGrant.ReadWrite.All",
    "Directory.Read.All",
    "User.Read"
) -TenantId $TenantId -NoWelcome
Write-Host "   ✅ Connected"
Write-Host ""

# ── Backend app ────────────────────────────────────────────────────────────
Write-Host "🔧 Backend API app: $BackendDisplayName"
$backendApp = Get-ApplicationByDisplayName -DisplayName $BackendDisplayName
if ($backendApp) {
    $backendAppId = $backendApp.appId
    Write-Host "   ✅ Already exists appId=$backendAppId"
} else {
    $body = @{
        displayName    = $BackendDisplayName
        signInAudience = "AzureADMyOrg"
    } | ConvertTo-Json
    $created = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications" `
        -Body $body -ContentType "application/json"
    $backendAppId = $created.appId
    Write-Host "   ✅ Created appId=$backendAppId"
    Start-Sleep -Seconds 3
}

$backendObjectId = (Get-ApplicationByAppId -AppId $backendAppId).id
$backendSpId = Ensure-ServicePrincipal -AppId $backendAppId
Write-Host "   Backend objectId=$backendObjectId servicePrincipal=$backendSpId"
Write-Host ""

# App ID URI
$appIdUri = "api://$backendAppId"
Write-Host "🌐 Application ID URI → $appIdUri"
$appFull = Invoke-MgGraphRequest -Method GET `
    -Uri (Format-GraphResourceUri -Path "applications/$backendObjectId" -Query '$select=identifierUris')
$uris = @($appFull.identifierUris)
if ($uris -notcontains $appIdUri) {
    $uris += $appIdUri
    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/applications/$backendObjectId" `
        -Body (@{ identifierUris = $uris } | ConvertTo-Json) `
        -ContentType "application/json"
    Write-Host "   ✅ Set identifierUris"
} else {
    Write-Host "   ✅ Already set"
}
Write-Host ""

# access_as_user scope
Write-Host "🔑 Expose scope access_as_user..."
$appApi = Invoke-MgGraphRequest -Method GET `
    -Uri (Format-GraphResourceUri -Path "applications/$backendObjectId" -Query '$select=api')
$scopes = @()
if ($appApi.api -and $appApi.api.oauth2PermissionScopes) {
    $scopes = @($appApi.api.oauth2PermissionScopes)
}
$scopeEntry = $scopes | Where-Object { $_.value -eq "access_as_user" }
if ($scopeEntry) {
    $accessAsUserScopeId = $scopeEntry.id
    Write-Host "   ✅ Scope exists id=$accessAsUserScopeId"
} else {
    $accessAsUserScopeId = (New-Guid).ToString()
    $scopes += @{
        id                       = $accessAsUserScopeId
        adminConsentDescription  = "Allow the SPA to access the Agent Identity Portal API on behalf of the signed-in user"
        adminConsentDisplayName  = "Access portal API as user"
        isEnabled                = $true
        type                     = "User"
        userConsentDescription   = "Allow the SPA to access the portal API on your behalf"
        userConsentDisplayName   = "Access portal API as user"
        value                    = "access_as_user"
    }
    Invoke-MgGraphRequest -Method PATCH `
        -Uri "https://graph.microsoft.com/v1.0/applications/$backendObjectId" `
        -Body (@{ api = @{ oauth2PermissionScopes = $scopes } } | ConvertTo-Json -Depth 6) `
        -ContentType "application/json"
    Write-Host "   ✅ Added scope id=$accessAsUserScopeId"
}
Write-Host ""

# Graph application permissions on backend app registration
Write-Host "📋 Backend Microsoft Graph application permissions..."
$graphSp = Get-GraphServicePrincipal
$script:BackendGraphRolesToGrant = Resolve-BackendGraphRoleValues -GraphSpAppRoles $graphSp.appRoles

$resourceAccess = @()
foreach ($roleValue in $BackendGraphRolesToGrant) {
    $role = $graphSp.appRoles | Where-Object {
        $_.value -eq $roleValue -and $_.allowedMemberTypes -contains "Application"
    } | Select-Object -First 1
    if (-not $role) {
        Write-Host "   ⚠️  Graph app role not found in tenant: $roleValue"
        continue
    }
    $resourceAccess += [ordered]@{ id = [string]$role.id; type = "Role" }
}

if ($resourceAccess.Count -eq 0) {
    Write-Error "No Microsoft Graph application permissions could be resolved in this tenant. Check Agent ID / Graph preview."
    exit 1
}

$appRra = Invoke-MgGraphRequest -Method GET `
    -Uri (Format-GraphResourceUri -Path "applications/$backendObjectId" -Query '$select=requiredResourceAccess')
$mergedRra = Merge-RequiredResourceAccess -Existing $appRra.requiredResourceAccess `
    -ResourceAppId $GraphResourceAppId -NewAccess $resourceAccess
Invoke-ApplicationRequiredResourceAccessPatch -ApplicationObjectId $backendObjectId `
    -RequiredResourceAccess $mergedRra
Write-Host "   ✅ Patched requiredResourceAccess on backend app"
Write-Host ""

Write-Host "🛡️  Admin consent backend → Microsoft Graph (application permissions)..."
foreach ($roleValue in $BackendGraphRolesToGrant) {
    $role = $graphSp.appRoles | Where-Object {
        $_.value -eq $roleValue -and $_.allowedMemberTypes -contains "Application"
    } | Select-Object -First 1
    if ($role) {
        Grant-ApplicationPermission -ClientSpId $backendSpId -ResourceSpId $graphSp.id `
            -AppRoleId ([string]$role.id) -RoleValue $roleValue
    }
}
Write-Host ""

$clientSecretPlain = $null
if ($createClientSecret) {
    Write-Host "🔐 Creating backend client secret (local dev)..."
    $pwdBody = @{
        passwordCredential = @{
            displayName = "portal-local-dev"
        }
    } | ConvertTo-Json -Depth 3
    $pwdResult = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications/$backendObjectId/addPassword" `
        -Body $pwdBody -ContentType "application/json"
    $clientSecretPlain = $pwdResult.secretText
    Write-Host "   ✅ Secret created (displayName=portal-local-dev)"
    Write-Host ""
    Write-Host "   ⚠️  Save AZURE_CLIENT_SECRET now — it will not be shown again."
    Write-Host "      Length: $($clientSecretPlain.Length) characters"
    Write-Host ""
}

# ── SPA ────────────────────────────────────────────────────────────────────
Write-Host "🌐 SPA: $SpaDisplayName"
$spaCreatedThisRun = $false
$spaApp = Get-ApplicationByDisplayName -DisplayName $SpaDisplayName
if ($spaApp) {
    $spaAppId = $spaApp.appId
    Write-Host "   ✅ Already exists appId=$spaAppId"
} else {
    $spaCreatedThisRun = $true
    $body = New-GraphSpaApplicationCreateBody -DisplayName $SpaDisplayName -RedirectUris $allRedirectUris
    $created = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications" `
        -Body $body -ContentType "application/json"
    $spaAppId = $created.appId
    Write-Host "   ✅ Created appId=$spaAppId"
    Start-Sleep -Seconds 3
}

$spaObjectId = (Get-ApplicationByAppId -AppId $spaAppId).id
$spaSpId = Ensure-ServicePrincipal -AppId $spaAppId

# Merge redirect URIs
$spaFull = Invoke-MgGraphRequest -Method GET `
    -Uri (Format-GraphResourceUri -Path "applications/$spaObjectId" -Query '$select=spa')
$currentRedirects = @()
if ($spaFull.spa -and $spaFull.spa.redirectUris) {
    $currentRedirects = @($spaFull.spa.redirectUris)
}
$mergedRedirects = @($currentRedirects) + @($allRedirectUris) | ForEach-Object { [string]$_ } | Select-Object -Unique
if ($spaCreatedThisRun) {
    Write-Host "   ✅ Redirect URIs set at create: $($mergedRedirects -join ', ')"
} else {
    Invoke-ApplicationSpaRedirectUrisPatch -ApplicationObjectId $spaObjectId -RedirectUris $mergedRedirects
    Write-Host "   ✅ Redirect URIs: $($mergedRedirects -join ', ')"
}
Write-Host ""

# SPA → backend API permission
Write-Host "📝 SPA delegated permission → access_as_user..."
$apiSpId = Ensure-ServicePrincipal -AppId $backendAppId
$spaDelegated = @(
    [ordered]@{ id = [string]$accessAsUserScopeId; type = "Scope" }
)
$spaRraApp = Invoke-MgGraphRequest -Method GET `
    -Uri (Format-GraphResourceUri -Path "applications/$spaObjectId" -Query '$select=requiredResourceAccess')
$spaMergedRra = Merge-RequiredResourceAccess -Existing $spaRraApp.requiredResourceAccess `
    -ResourceAppId $backendAppId -NewAccess $spaDelegated
Invoke-ApplicationRequiredResourceAccessPatch -ApplicationObjectId $spaObjectId `
    -RequiredResourceAccess $spaMergedRra
Write-Host "   ✅ Patched SPA requiredResourceAccess"
Write-Host ""

Write-Host "🛡️  Admin consent SPA → backend API..."
Grant-SpaDelegatedConsent -SpaSpId $spaSpId -ApiSpId $apiSpId
Write-Host ""

# ── Update .env files ────────────────────────────────────────────────────────
if ($UpdateEnv) {
    Write-Host "📝 Updating .env files..."
    foreach ($path in @($BackendEnvFile, $FrontendEnvFile)) {
        $dir = Split-Path -Parent $path
        $example = Join-Path $dir ".env.example"
        if (-not (Test-Path -LiteralPath $path) -and (Test-Path -LiteralPath $example)) {
            Copy-Item -LiteralPath $example -Destination $path
            Write-Host "   Copied $example → $path"
        }
    }

    Set-DotEnvValue -Path $BackendEnvFile -Key "PROVISIONING_MODE" -Value "live"
    Set-DotEnvValue -Path $BackendEnvFile -Key "AZURE_TENANT_ID" -Value $TenantId
    Set-DotEnvValue -Path $BackendEnvFile -Key "AZURE_CLIENT_ID" -Value $backendAppId
    Set-DotEnvValue -Path $BackendEnvFile -Key "FRONTEND_CLIENT_ID" -Value $spaAppId
    Set-DotEnvValue -Path $BackendEnvFile -Key "FRONTEND_ORIGIN" -Value $FrontendOrigin
    if ($clientSecretPlain) {
        Set-DotEnvValue -Path $BackendEnvFile -Key "AZURE_CLIENT_SECRET" -Value $clientSecretPlain `
            -Comment "# Created by setup-portal-entra-apps.ps1 — do not commit"
    }

    Set-DotEnvValue -Path $FrontendEnvFile -Key "REACT_APP_AZURE_CLIENT_ID" -Value $spaAppId
    Set-DotEnvValue -Path $FrontendEnvFile -Key "REACT_APP_API_CLIENT_ID" -Value $backendAppId
    Set-DotEnvValue -Path $FrontendEnvFile -Key "REACT_APP_AZURE_TENANT_ID" -Value $TenantId
    Set-DotEnvValue -Path $FrontendEnvFile -Key "REACT_APP_AZURE_REDIRECT_URI" -Value $RedirectUri
    Set-DotEnvValue -Path $FrontendEnvFile -Key "REACT_APP_API_BASE_URL" -Value "http://localhost:8000"

    Write-Host "   ✅ Updated $BackendEnvFile"
    Write-Host "   ✅ Updated $FrontendEnvFile"
    Write-Host ""
}

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  ✅ Done"
Write-Host ""
Write-Host "  Backend app id (AZURE_CLIENT_ID):     $backendAppId"
Write-Host "  SPA app id (REACT_APP_AZURE_CLIENT_ID): $spaAppId"
Write-Host "  API resource (REACT_APP_API_CLIENT_ID): $backendAppId"
Write-Host "  API scope:                            $appIdUri/access_as_user"
Write-Host ""
if (-not $UpdateEnv) {
    Write-Host "  Re-run with -UpdateEnv to patch backend/frontend .env"
}
Write-Host ""
Write-Host "  Next:"
Write-Host "    cd ../backend  →  uvicorn main:app --reload --port 8000"
Write-Host "    cd ../frontend →  npm start"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
