[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$XrmToolBoxDirectory
)

$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot 'bin\Release\net48\XrmToolBox.DocumentTemplateDeploymentManager.dll'
if (-not (Test-Path -LiteralPath $source)) { throw "Build the Release configuration first. Missing: $source" }
$plugins = Join-Path (Resolve-Path -LiteralPath $XrmToolBoxDirectory).Path 'Plugins'
New-Item -ItemType Directory -Force -Path $plugins | Out-Null
$destination = Join-Path $plugins (Split-Path $source -Leaf)
Copy-Item -LiteralPath $source -Destination $destination -Force
Write-Host "Installed: $destination"
Write-Host 'Restart XrmToolBox to load the plugin.'
