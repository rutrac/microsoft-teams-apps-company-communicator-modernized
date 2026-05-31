# Smoke-test for Invoke-SizingWizard. Loads helpers + the wizard from deploy.ps1
# without running the full deployment. Uses an in-memory $parameters object
# shaped like the parsed parameters.json so the wizard can mutate it.
#
# Usage:
#   pwsh ./test-wizard.ps1                # interactive
#   $env:CC_NONINTERACTIVE='1'; pwsh ./test-wizard.ps1   # non-interactive (no-op)

$ErrorActionPreference = 'Stop'

# Extract just the helper functions + Invoke-SizingWizard from deploy.ps1 using its AST.
$src = Get-Content -Raw "$PSScriptRoot/deploy.ps1"
$ast = [System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$null, [ref]$null)
$wanted = @('WriteI','WriteE','WriteW','WriteS','Invoke-SizingWizard')
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)) {
    Invoke-Expression $fn.Extent.Text
}

# Fake $parameters with just the fields the wizard touches.
$parameters = [pscustomobject]@{
    hostingPlanSku  = [pscustomobject]@{ Value = 'Standard' }
    hostingPlanSize = [pscustomobject]@{ Value = '2' }
}

Write-Host "BEFORE: tier=$($parameters.hostingPlanSku.Value)  size=$($parameters.hostingPlanSize.Value)"
Invoke-SizingWizard
Write-Host "AFTER:  tier=$($parameters.hostingPlanSku.Value)  size=$($parameters.hostingPlanSize.Value)"
