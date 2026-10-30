# =============================================================================
# Register MSAL.js Client SPA App (PowerShell)
# =============================================================================
# Creates a regular Entra app registration for MSAL.js browser sign-in.
# The user signs into THIS app → gets Tc → sends Tc to agent → OBO exchange.
#
# Uses Connect-MgGraph with Application.ReadWrite.All.
# Uses Invoke-MgGraphRequest for all calls (avoids module version conflicts).
#
# Usage:
#   cd scripts
#   pwsh ./setup-obo-client-app.ps1
#
#   pwsh ./setup-obo-client-app.ps1 -TenantId '<tenant-guid>' -EnvFile '..\sidecar\dev\.env'
# =============================================================================

param(
    [string]$TenantId,
    [string]$EnvFile = (Join-Path $PSScriptRoot ".env"),
    [string]$AppDisplayName = "Agent Demo Client SPA",
    [string]$RedirectUri = "http://localhost:3003"
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

if (-not $TenantId) {
    if (-not (Test-Path -LiteralPath $EnvFile)) {
        Write-Error ".env not found at $EnvFile (set -TenantId or create the file with TENANT_ID=)"
        exit 1
    }
    foreach ($line in Get-Content -LiteralPath $EnvFile) {
        if ($line -match '^\s*TENANT_ID\s*=\s*(.+)\s*$') {
            $TenantId = $Matches[1].Trim().Trim('"').Trim("'")
            break
        }
    }
    if (-not $TenantId) {
        Write-Error "TENANT_ID not set in $EnvFile and -TenantId was not provided"
        exit 1
    }
}

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  Register MSAL.js Client SPA App"
Write-Host "  Tenant: $TenantId"
Write-Host "  .env:   $EnvFile"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host ""

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Error @"
Microsoft Graph PowerShell SDK is not installed.

Install it (CurrentUser scope is enough), then re-run this script:

  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery

Or install the full meta-module:

  Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery
"@
    exit 1
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Write-Host "🔐 Connecting to Microsoft Graph..."
Write-Host "   (A browser window will open for sign-in)"
Connect-MgGraph -Scopes @("Application.ReadWrite.All") -TenantId $TenantId -NoWelcome
Write-Host "   ✅ Connected"
Write-Host ""

# ── Create app (or find existing) ───────────────────────────────────────────
Write-Host "🌐 Creating app registration: $AppDisplayName"

$existing = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/applications?`$filter=displayName eq '$AppDisplayName'&`$select=id,appId,displayName"

$clientSpaAppId = $null
if ($existing.value -and $existing.value.Count -gt 0) {
    $clientSpaAppId = $existing.value[0].appId
    Write-Host "   ✅ Already exists: $clientSpaAppId"
} else {
    $body = @{
        displayName    = $AppDisplayName
        signInAudience = "AzureADMyOrg"
        spa            = @{
            redirectUris = @($RedirectUri)
        }
    } | ConvertTo-Json -Depth 5

    $created = Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/applications" `
        -Body $body -ContentType "application/json"

    $clientSpaAppId = $created.appId
    Write-Host "   ✅ Created: $clientSpaAppId"
    Start-Sleep -Seconds 5
}
Write-Host ""

# ── Set SPA redirect URI ──────────────────────────────────────────────────────
Write-Host "🔗 Setting SPA redirect URI → $RedirectUri"
$spaBody = @{
    spa = @{
        redirectUris = @($RedirectUri)
    }
} | ConvertTo-Json -Depth 5

Invoke-MgGraphRequest -Method PATCH `
    -Uri "https://graph.microsoft.com/v1.0/applications(appId='$clientSpaAppId')" `
    -Body $spaBody -ContentType "application/json"
Write-Host "   ✅ Done"
Write-Host ""

# ── Ensure service principal exists ───────────────────────────────────────────
Write-Host "👤 Ensuring service principal exists..."
$spResult = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '$clientSpaAppId'&`$select=id"

if ($spResult.value -and $spResult.value.Count -gt 0) {
    Write-Host "   ✅ Already exists"
} else {
    $spBody = @{ appId = $clientSpaAppId } | ConvertTo-Json
    Invoke-MgGraphRequest -Method POST `
        -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
        -Body $spBody -ContentType "application/json"
    Write-Host "   ✅ Created service principal"
}
Write-Host ""

# ── Save to .env ──────────────────────────────────────────────────────────────
Write-Host "📝 Saving to .env..."
Set-DotEnvValue -Path $EnvFile -Key "CLIENT_SPA_APP_ID" -Value $clientSpaAppId `
    -Comment "# Client SPA App for MSAL.js user sign-in (OBO)"
Write-Host "   ✅ Updated CLIENT_SPA_APP_ID in $EnvFile"
Write-Host ""

Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
Write-Host "  ✅ Done!"
Write-Host ""
Write-Host "  Client SPA App ID: $clientSpaAppId"
Write-Host "  Redirect URI:      $RedirectUri"
Write-Host ""
Write-Host "  Next:"
Write-Host "    pwsh ./setup-obo-blueprint.ps1 \"
Write-Host "      -TenantId $TenantId \"
Write-Host "      -BlueprintAppId '<BLUEPRINT_APP_ID>' \"
Write-Host "      -AgentAppId '<AGENT_CLIENT_ID>' \"
Write-Host "      -ClientSpaAppId '$clientSpaAppId'"
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
