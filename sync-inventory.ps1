# Syncs InventoryTransactionsF6.xlsx from SharePoint using Microsoft Graph API.
# No additional modules required. First run opens a browser for login;
# subsequent runs are silent using a cached refresh token (encrypted with Windows DPAPI).
#
# Usage:
#   .\sync-inventory.ps1                        # saves file next to this script
#   .\sync-inventory.ps1 -OutputFolder "C:\MyFolder"

param(
    [string]$OutputFolder = $PSScriptRoot
)

$clientId  = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$tenantId  = "common"
$scope     = "https://graph.microsoft.com/Files.Read.All offline_access"
$driveId   = "b!nk-4Mun_TEyfLR3l7DosWfA9_f8-NZhJvGYZCLDd6FiwiuwpSvUQRIcLpJFheuXb"
$itemId    = "01ZAYHC4XOU5UCHDOWCNG3NNVN7TJLNN7Y"
$localPath = Join-Path $OutputFolder "InventoryTransactionsF6.xlsx"
$tokenFile = "$env:APPDATA\sync-inventory-token.enc"
$logFile   = Join-Path $OutputFolder "sync-inventory.log"

function Log($msg) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $msg"
    Add-Content -Path $logFile -Value $line
    Write-Host $line
}

function Get-NewToken {
    $dc = Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" `
        -Body @{ client_id = $clientId; scope = $scope }

    Write-Host ""
    Write-Host $dc.message
    Write-Host ""

    $deadline = (Get-Date).AddSeconds($dc.expires_in)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $dc.interval
        try {
            return Invoke-RestMethod -Method POST `
                -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
                -Body @{
                    client_id   = $clientId
                    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                    device_code = $dc.device_code
                }
        } catch { }
    }
    throw "Authentication timed out."
}

function Update-Token($refreshToken) {
    return Invoke-RestMethod -Method POST `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -Body @{
            client_id     = $clientId
            grant_type    = "refresh_token"
            refresh_token = $refreshToken
            scope         = $scope
        }
}

function Save-Token($token) {
    $json      = $token | ConvertTo-Json
    $encrypted = $json | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
    Set-Content -Path $tokenFile -Value $encrypted
}

function Import-Token {
    if (-not (Test-Path $tokenFile)) { return $null }
    try {
        $ss    = Get-Content $tokenFile | ConvertTo-SecureString
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ss))
        return $plain | ConvertFrom-Json
    } catch { return $null }
}

# --- Auth ---
$saved = Import-Token
if ($saved -and $saved.refresh_token) {
    Log "Refreshing access token..."
    try {
        $token = Update-Token $saved.refresh_token
        Save-Token $token
    } catch {
        Log "Refresh failed, re-authenticating..."
        $token = Get-NewToken
        Save-Token $token
    }
} else {
    Log "First-time login required..."
    $token = Get-NewToken
    Save-Token $token
}

# --- Download ---
Log "Downloading InventoryTransactionsF6.xlsx to $OutputFolder ..."
try {
    Invoke-WebRequest `
        -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/content" `
        -Headers @{ Authorization = "Bearer $($token.access_token)" } `
        -OutFile $localPath
    Log "Done. Saved to $localPath"
} catch {
    Log "ERROR: $_"
    exit 1
}
