#Requires -Version 7.0
<#
.SYNOPSIS
    Builds docs/index.html, a dashboard that replays an onboarding run step by step.
    Start-Onboarding.ps1 runs this automatically at the end of every run.

.EXAMPLE
    ./Build-Dashboard.ps1 -RunLog output/run_log.json
#>
param(
    [string]$RunLog = (Join-Path $PSScriptRoot 'output/run_log.json'),
    [string]$OutFile = (Join-Path $PSScriptRoot 'docs/index.html')
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $RunLog)) { throw "Run log not found: $RunLog. Run Start-Onboarding.ps1 first." }

$template = Get-Content (Join-Path $PSScriptRoot 'templates/dashboard_template.html') -Raw
# Re-serialize so the page gets compact JSON, and make it safe to embed in a <script> tag
$json = (Get-Content $RunLog -Raw | ConvertFrom-Json | ConvertTo-Json -Depth 8 -Compress) -replace '</', '<\/'
New-Item -ItemType Directory -Force -Path (Split-Path $OutFile) -WhatIf:$false | Out-Null
$template.Replace('/*__DATA__*/null', $json) | Set-Content -Path $OutFile -Encoding utf8 -WhatIf:$false
Write-Output $OutFile
