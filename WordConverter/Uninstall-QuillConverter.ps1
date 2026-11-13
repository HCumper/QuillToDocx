#Requires -RunAsAdministrator
<#
.SYNOPSIS
Removes the Quill text converter installed by Install-QuillConverter.ps1.
Close Word first: Word keeps the converter loaded while it runs.
#>
param(
    [string]$InstallDir = (Join-Path $env:ProgramFiles 'QuillCnv')
)

$ErrorActionPreference = 'Stop'

$keys = foreach ($direction in 'Import', 'Export') {
    "HKLM:\SOFTWARE\Microsoft\Shared Tools\Text Converters\$direction\Quill"
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Shared Tools\Text Converters\$direction\Quill"
}
foreach ($key in $keys) {
    if (Test-Path $key) {
        Remove-Item $key
        "Removed $key"
    }
}

foreach ($file in 'QuillCnv.cnv', 'QuillToDocx.exe') {
    $path = Join-Path $InstallDir $file
    if (Test-Path $path) { Remove-Item $path }
}
if ((Test-Path $InstallDir) -and -not (Get-ChildItem $InstallDir)) {
    Remove-Item $InstallDir
}
"Removed the converter files from $InstallDir."
