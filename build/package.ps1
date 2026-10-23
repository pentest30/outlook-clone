# Builds the Windows installer: releases\OutlookClone.Desktop-win-Setup.exe
#
#   pwsh build/package.ps1                 # version from src/Outlook.App/Outlook.App.csproj
#   pwsh build/package.ps1 -Version 0.6.0
#   pwsh build/package.ps1 -AzureSigningMetadata metadata.json   # signed (see docs/SIGNING.md)
#
# The installer is per-user (no admin rights), adds Start menu and desktop shortcuts and an
# uninstaller. The pack id is deliberately not "OutlookClone": Velopack installs to
# %LOCALAPPDATA%\<packId> and deletes that folder on uninstall, and %LOCALAPPDATA%\OutlookClone is
# where the app keeps the mail cache and saved passwords.
param(
    [string]$Version,
    [string]$Runtime = "win-x64",
    # Azure Trusted Signing: { "Endpoint", "CodeSigningAccountName", "CertificateProfileName" }.
    # Signs the app's files and the installer; credentials come from the environment
    # (AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET) or an Azure CLI login.
    [string]$AzureSigningMetadata
)
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$project = Join-Path $root "src/Outlook.App/Outlook.App.csproj"

if (-not $Version) {
    $Version = ([xml](Get-Content $project)).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
}
if (-not $Version) { throw "No <Version> in $project; pass -Version." }

$publish = Join-Path $root "publish/$Runtime"
$releases = Join-Path $root "releases"
if (Test-Path $publish) { Remove-Item $publish -Recurse -Force }

Write-Host "Publishing $Version for $Runtime..."
# Self-contained: users do not need .NET installed.
dotnet publish $project -c Release -r $Runtime --self-contained true -p:Version=$Version -o $publish
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed" }

Write-Host "Packing installer..."
Push-Location $root
try {
    dotnet tool restore | Out-Null
    $sign = @()
    if ($AzureSigningMetadata) {
        if (-not (Test-Path $AzureSigningMetadata)) { throw "Signing metadata not found: $AzureSigningMetadata" }
        Write-Host "Signing with Azure Trusted Signing"
        $sign = @("--azureTrustedSignFile", (Resolve-Path $AzureSigningMetadata).Path)
    }
    else {
        Write-Warning "Not signed: browsers and SmartScreen will warn about the installer."
    }
    dotnet vpk pack @sign `
        --packId OutlookClone.Desktop `
        --packVersion $Version `
        --packTitle "Outlook Clone" `
        --packAuthors "Netcom Group" `
        --packDir $publish `
        --mainExe OutlookClone.exe `
        --icon (Join-Path $root "src/Outlook.App/Assets/app.ico") `
        --outputDir $releases
    if ($LASTEXITCODE -ne 0) { throw "vpk pack failed" }
}
finally {
    Pop-Location
}
Write-Host "Done: $releases"
