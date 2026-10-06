<#
.SYNOPSIS
  Builds the Claude Mail Sorter Outlook add-ins on Windows.

.DESCRIPTION
  - Classic Outlook add-in (outlook-vsto): builds with MSBuild from Visual Studio, signs it with a
    self-signed certificate, and optionally installs it for the current user.
  - Office add-in (outlook): nothing to compile; the manifest is validated if Node.js is installed.

  Needs Visual Studio 2019/2022 (Community is fine) or Build Tools with the
  "Office/SharePoint development" workload, and an installed desktop Outlook.

.PARAMETER Install
  After building, copy the add-in to %LOCALAPPDATA%\ClaudeMailSorter, trust its certificate and register it
  with Outlook. Close Outlook first.

.PARAMETER Configuration
  Release (default) or Debug.
#>
[CmdletBinding()]
param(
    [switch]$Install,
    [ValidateSet("Release", "Debug")][string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$vstoDir = Join-Path $root "outlook-vsto"
$project = Join-Path $vstoDir "ClaudeMailSorter.csproj"
$certSubject = "CN=ClaudeMailSorter"
$addInName = "ClaudeMailSorter"

# Returns MSBuild.exe and the folder holding OfficeTools\Microsoft.VisualStudio.Tools.Office.targets.
# Checks every Visual Studio install, since the newest one may be a Build Tools install without the Office workload.
function Find-MSBuild {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        throw "Visual Studio not found. Install Visual Studio (Community) or Build Tools with the 'Office/SharePoint development' workload."
    }
    $installs = @(& $vswhere -all -products * -requires Microsoft.Component.MSBuild -property installationPath)
    if (-not $installs) { throw "No Visual Studio installation with MSBuild found." }

    $checked = @()
    foreach ($vs in $installs) {
        $msbuild = Join-Path $vs "MSBuild\Current\Bin\MSBuild.exe"
        if (-not (Test-Path $msbuild)) { continue }
        $searchRoots = @((Join-Path $vs "MSBuild\Microsoft\VisualStudio"),
                         (Join-Path ${env:ProgramFiles(x86)} "MSBuild\Microsoft\VisualStudio"))
        $targets = $searchRoots | Where-Object { Test-Path $_ } | ForEach-Object {
            Get-ChildItem $_ -Recurse -Filter "Microsoft.VisualStudio.Tools.Office.targets" -ErrorAction SilentlyContinue
        } | Select-Object -First 1
        if ($targets) {
            # ...\v17.0\OfficeTools\Microsoft.VisualStudio.Tools.Office.targets -> ...\v17.0
            return @{ MSBuild = $msbuild; VSToolsPath = $targets.Directory.Parent.FullName; InstallPath = $vs }
        }
        $checked += $vs
    }
    throw ("Found Visual Studio but not the VSTO build files (Microsoft.VisualStudio.Tools.Office.targets).`n" +
        "Checked: $($checked -join '; ')`n" +
        "In Visual Studio Installer choose Modify, then tick the 'Office/SharePoint development' workload " +
        "(or the individual component 'Visual Studio Tools for Office (VSTO)'), and run this again.")
}

# MSBuild can't find the VSTO runtime assemblies the Office build tasks load (they sit in Visual Studio's
# ReferenceAssemblies folder). Build from a private copy of the Office build files with those assemblies
# beside them, so nothing in the Visual Studio install is modified.
function Get-VstoToolsPath($tools) {
    $refs = Join-Path $tools.InstallPath "Common7\IDE\ReferenceAssemblies\v4.0"
    if (-not (Test-Path $refs)) { return $tools.VSToolsPath }
    $work = Join-Path $env:TEMP "ClaudeMailSorter\vstotools"
    if (Test-Path $work) { Remove-Item $work -Recurse -Force }
    $officeTools = Join-Path $work "OfficeTools"
    New-Item -ItemType Directory -Path $officeTools | Out-Null
    Copy-Item (Join-Path $tools.VSToolsPath "OfficeTools\*") $officeTools
    Copy-Item (Join-Path $refs "*.dll") $officeTools
    return $work
}

function Get-SigningCertificate {
    $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
        Where-Object { $_.Subject -eq $certSubject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
    if (-not $cert) {
        Write-Host "Creating a self-signed code-signing certificate ($certSubject)..."
        $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $certSubject `
            -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(5)
    }
    return $cert
}

function Build-ClassicAddIn {
    $tools = Find-MSBuild
    $cert = Get-SigningCertificate

    Write-Host "Building classic Outlook add-in ($Configuration)..."
    & $tools.MSBuild $project -restore "/p:Configuration=$Configuration" "/p:VSToolsPath=$(Get-VstoToolsPath $tools)" `
        "/p:SignManifests=true" `
        "/p:ManifestCertificateThumbprint=$($cert.Thumbprint)" "/v:minimal" "/nologo" | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "MSBuild failed (exit code $LASTEXITCODE)." }

    $out = Join-Path $vstoDir "bin\$Configuration"
    $vsto = Join-Path $out "ClaudeMailSorter.vsto"
    if (-not (Test-Path $vsto)) { throw "Build finished but $vsto was not produced." }
    Write-Host "Built: $vsto"
    return @{ Output = $out; Certificate = $cert }
}

function Install-ClassicAddIn($build) {
    if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) {
        throw "Close Outlook before installing."
    }
    $dest = Join-Path $env:LOCALAPPDATA "ClaudeMailSorter\addin"
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Path $dest | Out-Null
    Copy-Item (Join-Path $build.Output "*") $dest -Recurse

    # Outlook only loads VSTO add-ins signed by a publisher the user trusts.
    $cer = Join-Path $dest "ClaudeMailSorter.cer"
    Export-Certificate -Cert $build.Certificate -FilePath $cer | Out-Null
    Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
    Write-Host "Windows may now ask you to confirm trusting the certificate. Choose Yes."
    Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null

    $key = "HKCU:\Software\Microsoft\Office\Outlook\Addins\$addInName"
    New-Item -Path $key -Force | Out-Null
    $manifest = "file:///" + ((Join-Path $dest "ClaudeMailSorter.vsto") -replace "\\", "/") + "|vstolocal"
    New-ItemProperty -Path $key -Name Manifest -Value $manifest -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name FriendlyName -Value "Claude Mail Sorter" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name Description -Value "Uses Claude to sort email." -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
    Write-Host "Installed to $dest. Start Outlook and look for the Claude group on the Mail tab."
}

function Test-OfficeAddIn {
    if (-not (Get-Command npx -ErrorAction SilentlyContinue)) {
        Write-Host "Skipping Office add-in manifest check (Node.js not installed)."
        return
    }
    Write-Host "Validating Office add-in manifest..."
    Push-Location (Join-Path $root "outlook")
    try {
        & npx --yes office-addin-manifest validate manifest.xml
        if ($LASTEXITCODE -ne 0) { Write-Warning "Office add-in manifest did not validate." }
    }
    finally { Pop-Location }
}

$build = Build-ClassicAddIn
if ($Install) { Install-ClassicAddIn $build }
Test-OfficeAddIn
