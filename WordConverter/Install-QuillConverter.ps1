#Requires -RunAsAdministrator
<#
.SYNOPSIS
Installs the Quill text converter for Microsoft Word (64-bit).

.DESCRIPTION
Copies QuillCnv.cnv and QuillToDocx.exe to "Program Files\QuillCnv" and registers the
converter under HKLM\SOFTWARE\Microsoft\Shared Tools\Text Converters\Import\Quill and
...\Export\Quill. Word only reads converters from HKLM, so this needs an elevated PowerShell.
For Click-to-Run Office (Microsoft 365) the same entries are also written to Office's virtual
registry, where Word's own converters are registered.

Word then opens Sinclair QL Quill documents directly, including by double-click on a Quill
".doc", and can save them as Quill again. Run build.cmd first and close Word before
installing; Word reads converters when it starts. Undo with Uninstall-QuillConverter.ps1.
This is for development; the MSI is the way to install for use.
#>
param(
    [string]$SourceDir = (Join-Path $PSScriptRoot 'bin'),
    [string]$InstallDir = (Join-Path $env:ProgramFiles 'QuillCnv')
)

$ErrorActionPreference = 'Stop'

foreach ($file in 'QuillCnv.cnv', 'QuillToDocx.exe') {
    if (-not (Test-Path (Join-Path $SourceDir $file))) {
        throw "$file not found in $SourceDir. Run build.cmd first."
    }
}

New-Item -ItemType Directory -Force $InstallDir | Out-Null
foreach ($file in 'QuillCnv.cnv', 'QuillToDocx.exe') {
    Copy-Item (Join-Path $SourceDir $file) (Join-Path $InstallDir $file) -Force
}

$keys = @()
$clickToRun = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Shared Tools\Text Converters'
foreach ($direction in 'Import', 'Export') {
    $keys += "HKLM:\SOFTWARE\Microsoft\Shared Tools\Text Converters\$direction\Quill"
    if (Test-Path "$clickToRun\Import") {
        $keys += "$clickToRun\$direction\Quill"
    }
}

foreach ($key in $keys) {
    New-Item -Force $key | Out-Null
    Set-ItemProperty $key -Name Name -Value 'Sinclair QL Quill'
    Set-ItemProperty $key -Name Path -Value (Join-Path $InstallDir 'QuillCnv.cnv')
    Set-ItemProperty $key -Name Extensions -Value 'doc'
    "Registered $key"
}

"Installed to $InstallDir. Restart Word if it is running."
