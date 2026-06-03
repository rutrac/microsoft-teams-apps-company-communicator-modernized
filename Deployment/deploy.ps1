param(
    [string]$DeployerIp = ""
)

# Quiet az CLI noise (Python tracebacks on transient Graph timeouts, credential warnings).
# Full traces still flow into the transcript log below and ~/.azure/*.log.
$env:AZURE_CORE_ONLY_SHOW_ERRORS = 'true'
$env:AZURE_CORE_NO_COLOR         = 'true'
$env:AZURE_HTTP_USER_AGENT       = 'cc-modernized-deploy/5.26'

# Full-fidelity transcript: every Write-Host / stdout / stderr from this session lands on disk.
$script:DeployLogDir  = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $script:DeployLogDir)) { New-Item -ItemType Directory -Path $script:DeployLogDir | Out-Null }
$script:DeployLogPath = Join-Path $script:DeployLogDir ("deploy_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
try { Start-Transcript -Path $script:DeployLogPath -Append -IncludeInvocationHeader | Out-Null } catch { }
Write-Host "Deploy log: $script:DeployLogPath" -ForegroundColor DarkGray

function IsValidateSecureUrl {
    param(
        [Parameter(Mandatory = $true)] [string] $url
    )
    # Url with https prefix REGEX matching
    return ($url -match "https:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)")
}

# write information
function WriteI{
    param(
        [parameter(mandatory = $true)]
        [string]$message
    )
    Write-Host $message -foregroundcolor white
}

# write error
function WriteE{
    param(
        [parameter(mandatory = $true)]
        [string]$message
    )
    Write-Host $message -foregroundcolor red -BackgroundColor black
}

# write warning
function WriteW{
    param(
        [parameter(mandatory = $true)]
        [string]$message
    )
    Write-Host $message -foregroundcolor yellow -BackgroundColor black
}

# write success
function WriteS{
    param(
        [parameter(mandatory = $true)]
        [string]$message
    )
    Write-Host $message -foregroundcolor green -BackgroundColor black
}

function IsValidGuid
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$ObjectGuid
    )

    # Define verification regex
    [regex]$guidRegex = '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$'

    # Check guid against regex
    return $ObjectGuid -match $guidRegex
}

function IsValidParam {
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        $param
    )

    return -not([string]::IsNullOrEmpty($param.Value)) -and ($param.Value -ne '<<value>>')
}

# Validate input parameters.
function ValidateParameters {
    $isValid = $true
    if (-not(IsValidParam($parameters.subscriptionId))) {
        WriteE -message "Invalid subscriptionId."
        $isValid = $false;
    }

    if (-not(IsValidParam($parameters.subscriptionTenantId)) -or -not(IsValidGuid -ObjectGuid $parameters.subscriptionTenantId.Value)) {
        WriteE -message "Invalid subscriptionTenantId. This should be a GUID."
        $isValid = $false;
    }

    if (-not (IsValidParam($parameters.resourceGroupName))) {
        WriteE -message "Invalid resourceGroupName."
        $isValid = $false;
    }

    if (-not (IsValidParam($parameters.region))) {
        WriteE -message "Invalid region."
        $isValid = $false;
    }

    if (-not (IsValidParam($parameters.baseResourceName))) {
        WriteE -message "Invalid baseResourceName."
        $isValid = $false;
    }

    if (-not(IsValidParam($parameters.tenantId)) -or -not(IsValidGuid -ObjectGuid $parameters.tenantId.Value)) {
        WriteE -message "Invalid tenantId. This should be a GUID."
        $isValid = $false;
    }

    if (-not(IsValidParam($parameters.senderUPNList))) {
        WriteE -message "Invalid senderUPNList."
        $isValid = $false;
    }

    if (-not (IsValidParam($parameters.customDomainOption))) {
        WriteE -message "Invalid customDomainOption."
        $isValid = $false;
    }

    if (-not(IsValidParam($parameters.companyName))) {
        WriteE -message "Invalid companyName."
        $isValid = $false;
    }

    if (-not(IsValidateSecureUrl($parameters.WebsiteUrl.Value))) {
        WriteE -message "Invalid websiteUrl. This should be an https url."
        $isValid = $false;
    }

    if (-not(IsValidateSecureUrl($parameters.PrivacyUrl.Value))) {
        WriteE -message "Invalid PrivacyUrl. This should be an https url."
        $isValid = $false;
    }

    if (-not(IsValidateSecureUrl($parameters.TermsOfUseUrl.Value))) {
        WriteE -message "Invalid TermsOfUseUrl. This should be an https url."
        $isValid = $false;
    }

    return $isValid
}

function validateresourcesnames {
    WriteI -message "Checking for resources availability..."

    $authorizationtoken = get-accesstokenfromcurrentuser -erroraction stop
    $resources = @(@{
            name               = $parameters.baseresourcename.value
            servicetype        = 'webapp'
            authorizationtoken = $authorizationtoken
        },
        @{
            name               = $parameters.baseresourcename.value + '-data-function'
            servicetype        = 'webapp'
            authorizationtoken = $authorizationtoken
        },
        @{
            name               = $parameters.baseresourcename.value + '-function'
            servicetype        = 'webapp'
            authorizationtoken = $authorizationtoken
        },
        @{
            name               = $parameters.baseresourcename.value + '-prep-function'
            servicetype        = 'webapp'
            authorizationtoken = $authorizationtoken
        },
        @{
            name        = $parameters.baseresourcename.value
            servicetype = 'applicationinsights'
        })

    $allresourcesavailable = $true
    foreach ($resource in $resources) {
        $isresourcenameavailable = validateresourcenames $resource -erroraction stop
        $allresourcesavailable = $allresourcesavailable -and $isresourcenameavailable
    }

    if (!$allresourcesavailable) {
        $confirmationtitle = "Some of the resource types names already exist. If you proceed, this will update the existing resources."
        $confirmationquestion = "Do you want to proceed?"
        $confirmationchoices = "&yes", "&no" # 0 = yes, 1 = no

        $updatedecision = $host.ui.promptforchoice($confirmationtitle, $confirmationquestion, $confirmationchoices, 1)
        return ($updatedecision -eq 0)
    } else {
        return $true
    }
}

function validateresourcenames {
    param(
        [parameter(mandatory = $true)] $resourceinfo
    )

    if ($resourceinfo.servicetype -eq "applicationinsights") {
        $aiExists = az monitor app-insights component list --subscription $parameters.subscriptionId.value -o json 2>$null | ConvertFrom-Json | Where-Object { $_.name -eq $resourceinfo.name }
        if ($null -eq $aiExists) {
            WriteS -message "Application Insights resource ($($resourceinfo.name)) is available."
            return $true
        } else {
            WriteW -message "Application Insights resource ($($resourceinfo.name)) is not available."
            return $false
        }
    } else {
        $availabilityresult = $null
        $availabilityresult = IsResourceNameAvailable @resourceinfo -erroraction stop

        if ($availabilityresult.available) {
            WriteS -message "resource: $($resourceinfo.name) of type $($resourceinfo.servicetype) is available."
            return $true
        } else {
            WriteW -message "resource $($resourceinfo.name) is not available."
            WriteW -message $availabilityresult.message
            return $false
        }
    }
}
# Get access token from the logged-in az CLI session.
function get-accesstokenfromcurrentuser {
    try {
        $token = az account get-access-token --query accessToken -o tsv
        if ($LASTEXITCODE -ne 0) { throw "az account get-access-token failed" }
        ('Bearer ' + $token.Trim())
    }
    catch {
        throw
    }
}
# Check if the name of resource is available.
function IsResourceNameAvailable {
    param(
        [parameter(mandatory = $true)] [string] $authorizationtoken,
        [parameter(mandatory = $true)] [string] $name,
        [parameter(mandatory = $true)] [validateset(
            'apimanagement', 'keyvault', 'managementgroup', 'sql', 'storageaccount', 'webapp', 'cognitiveservice')]
        $servicetype
    )

    $uribyservicetype = @{
        apimanagement    = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.apimanagement/checknameavailability?api-version=2019-01-01'
        keyvault         = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.keyvault/checknameavailability?api-version=2019-09-01'
        managementgroup  = 'https://management.azure.com/providers/microsoft.management/checknameavailability?api-version=2018-03-01-preview'
        sql              = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.sql/checknameavailability?api-version=2018-06-01-preview'
        storageaccount   = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.storage/checknameavailability?api-version=2019-06-01'
        webapp           = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.web/checknameavailability?api-version=2020-06-01'
        cognitiveservice = 'https://management.azure.com/subscriptions/{subscriptionid}/providers/microsoft.cognitiveservices/checkdomainavailability?api-version=2017-04-18'
    }

    $typebyservicetype = @{
        apimanagement    = 'microsoft.apimanagement/service'
        keyvault         = 'microsoft.keyvault/vaults'
        managementgroup  = '/providers/microsoft.management/managementgroups'
        sql              = 'microsoft.sql/servers'
        storageaccount   = 'microsoft.storage/storageaccounts'
        webapp           = 'microsoft.web/sites'
        cognitiveservice = 'microsoft.cognitiveservices/accounts'
    }

    $uri = $uribyservicetype[$servicetype] -replace ([regex]::escape('{subscriptionid}')), $parameters.subscriptionid.value
    if ($servicetype -eq 'cognitiveservice') {
        $nameproperty = "subdomainname"
    } else {
        $nameproperty = "name"
    }
    $body = '"{0}": "{1}", "type": "{2}"' -f $nameproperty, $name, $typebyservicetype[$servicetype]

    $response = (invoke-webrequest -uri $uri -method post -body "{$body}" -contenttype "application/json" -headers @{authorization = $authorizationtoken } -usebasicparsing).content
    $response | convertfrom-json |
    select-object @{n = 'name'; e = { $name } }, @{n = 'type'; e = { $servicetype } }, @{n = 'available'; e = { $_ | select-object -expandproperty *available } }, reason, message
}

# To get the Azure AD app detail (with retry for transient Graph API timeouts).
function GetAzureADApp {
    param ($appName)
    try {
        $json = Invoke-AzWithRetry -Label "ad app list '$appName'" -ScriptBlock { az ad app list --filter "displayName eq '$appName'" }
        return ($json | ConvertFrom-Json)
    } catch {
        WriteE -message $_.Exception.Message
        return $null
    }
}

# To get the Azure AD app detail with new secret.
function GetAzureADAppWithSecret {
    param ($appName)
    $app = GetAzureADApp $appName

    #Reset the app credentials to get the secret. The default validity of this secret will be for 1 year from the date its created.
    WriteI -message "Retreiving new app with secrets..."
    $appSecret = az ad app credential reset --id $app.appId --append | ConvertFrom-Json;

    return $appSecret
}

# Create/re-set Azure AD app.
function CreateAzureADApp {
    param(
        [Parameter(Mandatory = $true)] [string] $AppName,
		[Parameter(Mandatory = $false)] [bool] $ResetAppSecret = $true,
        [Parameter(Mandatory = $false)] [bool] $MultiTenant = $false,
        [Parameter(Mandatory = $false)] [bool] $AllowImplicitFlow
    )

    # Bot Service msaAppType is hard-coded SingleTenant in azuredeploy.json, so AAD apps must match.
    if ($MultiTenant) {
        $signInAudience = 'AzureADMultipleOrgs'
    } else {
        $signInAudience = 'AzureADMyOrg'
    }

    try {
        WriteI -message "`r`nCreating Azure AD App: $appName..."

        # Check if the app already exists - script has been previously executed
        $app = GetAzureADApp $appName

        if (-not ([string]::IsNullOrEmpty($app))) {

            # Update Azure AD app registration using CLI
            $confirmationTitle = "The Azure AD app '$appName' already exists. If you proceed, this will update the existing app configuration."
            $confirmationQuestion = "Do you want to proceed?"
            $confirmationChoices = "&Yes", "&No" # 0 = Yes, 1 = No

            $updateDecision = $Host.UI.PromptForChoice($confirmationTitle, $confirmationQuestion, $confirmationChoices, 1)
            if ($updateDecision -eq 0) {
                WriteI -message "Updating the existing app..."

                try {
                    Invoke-AzWithRetry -Label "ad app update '$appName'" -ScriptBlock {
                        az ad app update --id $app.appId --display-name $appName --sign-in-audience $signInAudience
                    } | Out-Null
                } catch {
                    WriteE -message "Failed to update Azure AD app '$appName': $($_.Exception.Message)"
                    return $null
                }

                WriteI -message "Waiting for app update to finish..."

                Start-Sleep -s 10

                WriteS -message "Azure AD App: $appName is updated."
            } else {
                WriteE -message "Deployment canceled. Please use a different name for the Azure AD app and try again."
                return $null
            }
        } else {
            # Create Azure AD app registration using CLI (with retry for transient Graph API timeouts)
            try {
                Invoke-AzWithRetry -Label "ad app create '$appName'" -ScriptBlock {
                    az ad app create --display-name $appName --sign-in-audience $signInAudience
                } | Out-Null
            } catch {
                WriteE -message "Failed to create Azure AD app '$appName': $($_.Exception.Message)"
                return $null
            }

            WriteI -message "Waiting for app creation to finish..."

            Start-Sleep -s 10

            WriteS -message "Azure AD App: $appName is created."
        }

        $app = GetAzureADApp $appName

        # Ensure the matching service principal exists. Without an SP, admin consent has nothing
        # to write grants against, bot connector auth returns 401, and Graph OBO fails (AADSTS65001).
        # 'az ad sp create' is idempotent: if the SP already exists it errors out, which is fine.
        if ($null -ne $app) {
            $spOk = $false
            for ($spAttempt = 1; $spAttempt -le 10; $spAttempt++) {
                $spOut = az ad sp create --id $app.appId 2>&1
                if ($LASTEXITCODE -eq 0) { $spOk = $true; break }
                if ($spOut -match 'already exists|Another object with the same value') { $spOk = $true; break }
                if ($spAttempt -lt 10) {
                    WriteW -message "SP create not ready (AAD propagation). Retrying in 6s ($spAttempt/10)..."
                    Start-Sleep -Seconds 6
                }
            }
            if (-not $spOk) {
                WriteE -message "Failed to create service principal for '$appName'. Admin consent and bot auth will not work."
                return $null
            }
        }

        $appSecret = $null;
        #Reset the app credentials to get the secret. The default validity of this secret will be for 1 year from the date its created.
        if ($ResetAppSecret) {
            if ($null -eq $app) {
                WriteE -message "Cannot create secret for '$appName': app lookup returned null."
                return $null
            }
            WriteI -message "Updating app secret..."
            try {
                $secretJson = Invoke-AzWithRetry -Label "credential reset '$appName'" -ScriptBlock {
                    az ad app credential reset --id $app.appId --append
                }
                $appSecret = $secretJson | ConvertFrom-Json
            } catch {
                WriteE -message "Failed to reset credential for '$appName': $($_.Exception.Message)"
            }
        }

        WriteS -message "Azure AD App: $appName registered successfully."
        return $appSecret
    }
    catch {
        $errorMessage = $_.Exception.Message
        WriteE -message "Failed to register/configure the Azure AD app. Error message: $errorMessage"
    }
    return $null
}

#to get the deployment log with the help of logged in user detail.
function CollectARMDeploymentLogs {
    $logsPath = '.\DeploymentLogs'
    $activityLogPath = "$logsPath\activity_log.log"
    $deploymentLogPath = "$logsPath\deployment_operation.log"

    $logsFolder = New-Item -ItemType Directory -Force -Path $logsPath

    az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploy --query "[?properties.provisioningState=='Failed'].properties.statusMessage.error" | Set-Content $deploymentLogPath

    $activityLog = $null
    $retryCount = 5
    DO {
        WriteI -message "Collecting deployment logs..."

        # Wait for async logs to persist
        Start-Sleep -s 30

        # Returns empty [] if logs are not available yet
        $activityLog = az monitor activity-log list -g $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --caller $userAlias --status Failed --offset 30m

        $retryCount--

    } While (($activityLog.Length -lt 3) -and ($retryCount -gt 0))

    $activityLog | Set-Content $activityLogPath

    # collect web apps deployment logs
    $activityLogErrors = ($activityLog | ConvertFrom-Json) | Where-Object { ($null -ne $_.resourceType) -and ($_.resourceType.value -eq "Microsoft.Web/sites/sourcecontrols") }
    $resourcesLookup = @($activityLogErrors | Select-Object resourceId, @{Name = "resourceName"; Expression = { GetResourceName $_.resourceId } })
    if ($resourcesLookup.length -gt 0) {
        foreach ($resourceInfo in $resourcesLookup) {
            if ($null -ne $resourceInfo.resourceName) {
                az webapp log download --ids $resourceInfo.resourceId --log-file "$logsPath\$($resourceInfo.resourceName).zip"
            }
        }
    }

    # Generate zip archive and delete folder
    $compressManifest = @{
        Path             = $logsPath
        CompressionLevel = "Fastest"
        DestinationPath  = "logs.zip"
    }
    Compress-Archive @compressManifest -Force
    Get-ChildItem -Path $logsPath -Recurse | Remove-Item -Force -Recurse -ErrorAction Continue
    Remove-Item $logsPath -Force -ErrorAction Continue

    WriteI -message "Deployment logs generation finished. Please share Deployment\logs.zip file with the app template team to investigate..."
}

function IsSourceControlTimeOut {
    $failedResourcesList = $null
    $failedResourcesList = az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploy --query "[?properties.provisioningState=='Failed']" | ConvertFrom-Json
    $nonCodeSyncErrors = $failedResourcesList | Where-Object {($null -ne $_.properties.targetResource -and 'Microsoft.Web/sites/sourcecontrols' -ne $_.properties.targetResource.resourceType)}
    return (0 -ne $failedResourcesList.length -and 0 -eq $nonCodeSyncErrors.length)
}

function WaitForCodeDeploymentSync {
    Param(
        [Parameter(Mandatory = $true)] $appServicesNames
    )

    $appserviceCodeSyncSuccess = $true
    while($appServicesNames.Count -gt 0)
    {
        WriteI -message "Checking source control deployment progress..."
        For ($i=0; $i -le $appServicesNames.Count; $i++) {
            $appService = $appServicesNames[$i]
            if($null -ne $appService){
                $deploymentResponse = az rest --method get --uri /subscriptions/$($parameters.subscriptionId.Value)/resourcegroups/$($parameters.resourceGroupName.Value)/providers/Microsoft.Web/sites/$appService/deployments?api-version=2019-08-01 | ConvertFrom-Json
                $deploymentsList = $deploymentResponse.value
                if($deploymentsList.length -eq 0 -or $deploymentsList[0].properties.complete){
                    $appserviceCodeSyncSuccess = $appserviceCodeSyncSuccess -and ($deploymentsList.length -eq 0 -or $deploymentsList[0].properties.status -ne 3) # 3 means sync fail
                    $appServicesNames.remove($appService)
                    $i--;
                }
            }
        }

        WriteI -message "Source control deployment is still in progress. Next check in 2 minutes."
        Start-Sleep -Seconds 120
    }
    if($appserviceCodeSyncSuccess){
        WriteI -message "Source control deployment is done."
    } else {
        WriteE -message "Source control deployment failed."
    }
    return $appserviceCodeSyncSuccess
}

# Build an ARM parameters file from $parameters + runtime AAD values, and invoke az.
# Using --parameters @file.json avoids the PowerShell 5.1 cmd.exe arg re-tokenization
# bug that emits a stray "<<" token when many quoted key=value args are passed inline.
function InvokeArmDeploymentWithParamsFile {
    Param(
        [Parameter(Mandatory = $true)] $graphappid,
        [Parameter(Mandatory = $true)] $authorappId,
        [Parameter(Mandatory = $true)] $userappId,
        [Parameter(Mandatory = $true)] $graphappsecret,
        [Parameter(Mandatory = $true)] $authorsecret,
        [Parameter(Mandatory = $true)] $usersecret
    )

    if ($null -ne $script:DeployerIpToUse) {
        $deployerIp = $script:DeployerIpToUse
    } else {
        $deployerIp = ''
    }

    $armParamsObj = [ordered]@{
        '$schema'      = 'https://schema.management.azure.com/schemas/2019-04-01/deploymentParameters.json#'
        contentVersion = '1.0.0.0'
        parameters     = [ordered]@{
            baseResourceName                   = @{ value = $parameters.baseResourceName.Value }
            authorClientId                     = @{ value = $authorappId }
            authorClientSecret                 = @{ value = $authorsecret }
            graphAppId                         = @{ value = $graphappid }
            graphAppSecret                     = @{ value = $graphappsecret }
            userClientId                       = @{ value = $userappId }
            userClientSecret                   = @{ value = $usersecret }
            senderUPNList                      = @{ value = $parameters.senderUPNList.Value }
            customDomainOption                 = @{ value = $parameters.customDomainOption.Value }
            appDisplayName                     = @{ value = $parameters.appDisplayName.Value }
            appDescription                     = @{ value = $parameters.appDescription.Value }
            appIconUrl                         = @{ value = $parameters.appIconUrl.Value }
            tenantId                           = @{ value = $parameters.tenantId.Value }
            hostingPlanSku                     = @{ value = $parameters.hostingPlanSku.Value }
            location                           = @{ value = $parameters.region.Value }
            ProactivelyInstallUserApp          = @{ value = [System.Convert]::ToBoolean($parameters.proactivelyInstallUserApp.Value) }
            UserAppExternalId                  = @{ value = $parameters.userAppExternalId.Value }
            DefaultCulture                     = @{ value = $parameters.defaultCulture.Value }
            SupportedCultures                  = @{ value = $parameters.supportedCultures.Value }
            serviceBusWebAppRoleNameGuid       = @{ value = $parameters.serviceBusWebAppRoleNameGuid.Value }
            serviceBusPrepFuncRoleNameGuid     = @{ value = $parameters.serviceBusPrepFuncRoleNameGuid.Value }
            serviceBusSendFuncRoleNameGuid     = @{ value = $parameters.serviceBusSendFuncRoleNameGuid.Value }
            serviceBusDataFuncRoleNameGuid     = @{ value = $parameters.serviceBusDataFuncRoleNameGuid.Value }
            storageAccountWebAppRoleNameGuid   = @{ value = $parameters.storageAccountWebAppRoleNameGuid.Value }
            storageAccountPrepFuncRoleNameGuid = @{ value = $parameters.storageAccountPrepFuncRoleNameGuid.Value }
            storageAccountDataFuncRoleNameGuid = @{ value = $parameters.storageAccountDataFuncRoleNameGuid.Value }
            TargetingEnabled                   = @{ value = [System.Convert]::ToBoolean($parameters.TargetingEnabled.Value) }
            MasterAdminUpns                    = @{ value = $parameters.MasterAdminUpns.Value }
            deployerIpAddress                  = @{ value = $deployerIp }
        }
    }

    $armParamsPath = Join-Path (Get-Location) 'armparams.runtime.json'
    ($armParamsObj | ConvertTo-Json -Depth 5) | Set-Content -Path $armParamsPath -Encoding utf8

    # Submit with --no-wait to avoid holding a long-lived TCP connection
    # (the default blocking call can be open for 60+ minutes and gets reset by firewalls).
    # Retry the submission itself on transient connection errors.
    $submitAttempts = 0
    $maxSubmitAttempts = 4
    do {
        $submitAttempts++
        az deployment group create `
            --name 'azuredeploy' `
            --resource-group $parameters.resourceGroupName.Value `
            --subscription $parameters.subscriptionId.Value `
            --template-file 'azuredeploy.json' `
            --parameters "@$armParamsPath" `
            --no-wait 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) { break }
        if ($submitAttempts -lt $maxSubmitAttempts) {
            WriteW -message "ARM deployment submission failed (attempt $submitAttempts/$maxSubmitAttempts). Retrying in 30s..."
            Start-Sleep -Seconds 30
        }
    } while ($submitAttempts -lt $maxSubmitAttempts)

    if ($LASTEXITCODE -ne 0) {
        WriteE -message "ARM deployment submission failed after $maxSubmitAttempts attempts."
        $global:LASTEXITCODE = 1
        return $null
    }

    # Poll until deployment reaches a terminal state. Each poll is a short HTTP call.
    WriteI -message "ARM deployment submitted. Polling every 60s for completion (this can take over an hour)..."
    $pollIntervalSec = 60
    $maxWaitSec = 14400  # 4 hours
    $elapsedSec = 0
    $deployState = ''
    while ($elapsedSec -lt $maxWaitSec) {
        Start-Sleep -Seconds $pollIntervalSec
        $elapsedSec += $pollIntervalSec
        $deployState = az deployment group show `
            --name 'azuredeploy' `
            --resource-group $parameters.resourceGroupName.Value `
            --subscription $parameters.subscriptionId.Value `
            --query "properties.provisioningState" -o tsv 2>$null
        if ($deployState -eq 'Succeeded' -or $deployState -eq 'Failed' -or $deployState -eq 'Canceled') { break }
        if (($elapsedSec % 300) -eq 0) {
            WriteI -message "ARM deployment in progress (state: $deployState, elapsed: $([int]($elapsedSec/60)) min)..."
        }
    }

    $deployResult = az deployment group show `
        --name 'azuredeploy' `
        --resource-group $parameters.resourceGroupName.Value `
        --subscription $parameters.subscriptionId.Value

    if ($deployState -eq 'Succeeded') {
        $global:LASTEXITCODE = 0
    } else {
        $global:LASTEXITCODE = 1
    }
    return $deployResult
}

function DeployARMTemplate {
    Param(
        [Parameter(Mandatory = $true)] $graphappid,
        [Parameter(Mandatory = $true)] $authorappId,
        [Parameter(Mandatory = $true)] $userappId,
        [Parameter(Mandatory = $false)] $graphappsecret,
        [Parameter(Mandatory = $false)] $authorsecret,
        [Parameter(Mandatory = $false)] $usersecret
    )
    try {
        if ((az group exists --name $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value) -eq $false) {
            WriteI -message "Creating resource group $($parameters.resourceGroupName.Value)..."
            az group create --name $parameters.resourceGroupName.Value --location $parameters.region.Value --subscription $parameters.subscriptionId.Value
        }

        $appServicesNames = [System.Collections.ArrayList]@($parameters.BaseResourceName.Value, #app-service
        "$($parameters.BaseResourceName.Value)-prep-function", #prep-function
        "$($parameters.BaseResourceName.Value)-function", #function
        "$($parameters.BaseResourceName.Value)-data-function" #data-function
        )

        $codeSynced = $false
        # Remove source control config if conflict detected
        if($parameters.isUpgrade.Value){
            foreach ($appService in $appServicesNames) {
                WriteI -message "Scan $appService source control configuration for conflicts"
                $deploymentConfig = az webapp deployment source show --name $appService --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value
                if($deploymentConfig){
                    $deploymentConfig = $deploymentConfig | ConvertFrom-Json
                    # conflicts in branches, clear old configuraiton
                    if(($deploymentConfig.branch -ne $parameters.gitBranch.Value) -or ($deploymentConfig.repoUrl -ne $parameters.gitRepoUrl.Value)){
                        WriteI -message "Remove $appService source control configuration"
                        az webapp deployment source delete --name $appService --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value
                        # code will be synced in ARM deployment stage
                        $codeSynced = $true
                    }
                }
                else {
                    # If command failed due to resource not exists, then screen colors is becoming red
                    [Console]::ResetColor()
                }
            }
        }

        # Deploy ARM templates
        WriteI -message "`nDeploying app services, Azure function, bot service, and other supporting resources... (this step can take over an hour)"
        $armDeploymentResult = InvokeArmDeploymentWithParamsFile $graphappid $authorappId $userappId $graphappsecret $authorsecret $usersecret

        $deploymentExceptionMessage = "ERROR: ARM template deployment error."
        if ($LASTEXITCODE -ne 0) {
            # If ARM template deployment failed for any reason, then screen colors is becoming red
            [Console]::ResetColor()

            WriteI -message "Fetching deployment status to check if deployment really failed..."

            # Check if deployment reached Azure despite the connection error
            $deployName = 'azuredeploy'
            $existingDeployState = az deployment group show `
                --name $deployName `
                --resource-group $parameters.resourceGroupName.Value `
                --subscription $parameters.subscriptionId.Value `
                --query "properties.provisioningState" -o tsv 2>$null

            if ($existingDeployState -eq 'Running') {
                WriteW -message "Connection was reset but ARM deployment is running in Azure. Waiting for it to complete (this can take over an hour)..."
                az deployment group wait `
                    --name $deployName `
                    --resource-group $parameters.resourceGroupName.Value `
                    --subscription $parameters.subscriptionId.Value `
                    --created --timeout 7200 2>$null | Out-Null
                $armDeploymentResult = az deployment group show `
                    --name $deployName `
                    --resource-group $parameters.resourceGroupName.Value `
                    --subscription $parameters.subscriptionId.Value
                $finalState = ($armDeploymentResult | ConvertFrom-Json).properties.provisioningState
                if ($finalState -ne 'Succeeded') {
                    CollectARMDeploymentLogs
                    Throw $deploymentExceptionMessage
                }
                WriteS -message "ARM deployment completed successfully."
            } elseif ($existingDeployState -eq 'Succeeded') {
                WriteI -message "ARM deployment had already completed successfully. Fetching outputs..."
                $armDeploymentResult = az deployment group show `
                    --name $deployName `
                    --resource-group $parameters.resourceGroupName.Value `
                    --subscription $parameters.subscriptionId.Value
            } elseif (IsSourceControlTimeOut) {
                # wait couple of minutes & check deployment status...
                $appserviceCodeSyncSuccess = WaitForCodeDeploymentSync $appServicesNames.Clone()

                if($appserviceCodeSyncSuccess){
                    WriteI -message "Re-running deployment to fetch output..."
                       $armDeploymentResult = InvokeArmDeploymentWithParamsFile $graphappid $authorappId $userappId $graphappsecret $authorsecret $usersecret
                } else{
                    CollectARMDeploymentLogs
                    Throw $deploymentExceptionMessage
                }
            } else {
                # Deployment never reached Azure — retry once
                WriteW -message "Deployment not found in Azure (connection reset before submission). Retrying ARM deployment..."
                if ($LASTEXITCODE -ne 0) {
                    CollectARMDeploymentLogs
                    Throw $deploymentExceptionMessage
                }
            }
        }
        else {
            # First-time sync wait removed: ARM no longer provisions sourcecontrols/web (the
            # ARM provider returned InternalServerError on git-based sourcecontrols with
            # IsManualIntegration:true). Source control is now configured below via
            # `az webapp deployment source config` after ARM completes successfully.
        }
        WriteS -message "Finished deploying resources. ARM template deployment succeeded."

        # Configure source control on each App Service (post-ARM, Microsoft-recommended pattern).
        WriteI -message "`nConfiguring source control on App Services from $($parameters.gitRepoUrl.Value) (branch: $($parameters.gitBranch.Value))..."
        foreach ($appService in $appServicesNames) {
            # Defensive: clear any stale source-control config left by a prior failed deploy.
            # Ignore errors here — when nothing is configured, delete returns 404.
            az webapp deployment source delete `
                --name $appService `
                --resource-group $parameters.resourceGroupName.Value `
                --subscription $parameters.subscriptionId.Value 2>$null | Out-Null
            [Console]::ResetColor()

            WriteI -message "Set source control for $appService"
            az webapp deployment source config `
                --name $appService `
                --resource-group $parameters.resourceGroupName.Value `
                --subscription $parameters.subscriptionId.Value `
                --repo-url $parameters.gitRepoUrl.Value `
                --branch $parameters.gitBranch.Value `
                --manual-integration 2>&1 | Out-Null
            # az CLI bug: returns exit code 1 when the API responds 200 OK (instead of 201/202).
            # Verify actual state via ARM API rather than trusting $LASTEXITCODE.
            $scVerify = az webapp deployment source show `
                --name $appService `
                --resource-group $parameters.resourceGroupName.Value `
                --subscription $parameters.subscriptionId.Value `
                -o json 2>$null | ConvertFrom-Json
            if (-not $scVerify -or -not $scVerify.repoUrl) {
                [Console]::ResetColor()
                WriteE -message "Failed to configure source control on $appService"
                CollectARMDeploymentLogs
                Throw "ERROR: Source control configuration failed on $appService"
            }
            WriteS -message "Source control verified on $appService (repo: $($scVerify.repoUrl))"
        }
        WriteI -message "Source control configured. Waiting for initial code sync to finish (this can take a while for the bot app due to npm/dotnet build)..."
        $appserviceCodeSyncSuccess = WaitForCodeDeploymentSync $appServicesNames.Clone()
        if (-not $appserviceCodeSyncSuccess) {
            CollectARMDeploymentLogs
            Throw "ERROR: Code sync did not complete successfully after source control configuration."
        }
        WriteS -message "Source control configured and code sync completed for all App Services."

        #get the output of current deployment
        $deploymentOutput = $null
            $deploymentOutput = az deployment group show --name azuredeploy --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value | ConvertFrom-Json

        # Sync only in upgrades & if no source branch conflict detected
        if($parameters.isUpgrade.Value -and (-not $codeSynced)){
            # sync app services code deployment (ARM deployment will not sync automatically)
            foreach ($appService in $appServicesNames) {
                WriteI -message "Sync $appService code from latest version"
                az webapp deployment source sync --name $appService --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value
            }
            # sync command is async. Wait for source control sync to finish
            $appserviceCodeSyncSuccess = WaitForCodeDeploymentSync $appServicesNames.Clone()
            if(-not $appserviceCodeSyncSuccess){
                CollectARMDeploymentLogs
                Throw $deploymentExceptionMessage
            }
        }

        return $deploymentOutput
    }
    catch {
        WriteE -message "Error occurred while deploying Azure resources."
        throw
    }
}


# Read a single Y/N keystroke with a timeout. Returns 'Y', 'N', or 'TIMEOUT'.
# Falls back to a blocking Read-Host when no interactive console is attached (CI).
function Read-YesNoWithTimeout {
    Param(
        [Parameter(Mandatory = $true)][int]$TimeoutSeconds,
        [Parameter(Mandatory = $true)][ValidateSet('Y','N')][string]$DefaultOnTimeout
    )

    # If no real console (redirected/non-interactive), can't poll keys — just return default.
    if ([Console]::IsInputRedirected) {
        return 'TIMEOUT'
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastShown = -1
    while ((Get-Date) -lt $deadline) {
        $remaining = [int][Math]::Ceiling(($deadline - (Get-Date)).TotalSeconds)
        if ($remaining -ne $lastShown) {
            Write-Host -NoNewline ("`r[Y/N] (default $DefaultOnTimeout in {0,2}s): " -f $remaining)
            $lastShown = $remaining
        }
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $ch = ([string]$key.KeyChar).ToUpper()
            if ($ch -eq 'Y' -or $ch -eq 'N') {
                Write-Host ("`r[Y/N] (default $DefaultOnTimeout): $ch                    ")
                return $ch
            }
        }
        Start-Sleep -Milliseconds 200
    }
    Write-Host ("`r[Y/N] (timeout — defaulting to $DefaultOnTimeout)            ")
    return 'TIMEOUT'
}

# Grant Admin consent
# Behaviour: 30s prompt, default = N (skip). On skip OR failure, records the appId in
# $script:AdminConsentPending so the end-of-deploy summary can surface the admin URL.
# Never throws — the deploy continues either way (consent can be granted post-deploy).
function GrantAdminConsent {
    Param(
        [Parameter(Mandatory = $true)] $graphAppId
        )

    Write-Host ""
    Write-Host "Admin consent is required for the Company Communicator bot app registration." -ForegroundColor Yellow
    Write-Host "Grant admin consent now? (requires you to be a Global Admin or Privileged Role Administrator)" -ForegroundColor Yellow
    $answer = Read-YesNoWithTimeout -TimeoutSeconds 30 -DefaultOnTimeout 'N'

    if ($answer -eq 'Y') {
        WriteI -message "Waiting for admin consent to finish..."
        $consentOk = $false
        try {
            Invoke-AzWithRetry -Label "admin-consent $graphAppId" -ScriptBlock {
                az ad app permission admin-consent --id $graphAppId
            } | Out-Null
            $consentOk = $true
        } catch {
            WriteE -message $_.Exception.Message
        }

        if ($consentOk) {
            WriteS -message "Admin consent has been granted."
            return
        }
        WriteW -message "Admin consent grant failed (caller may lack admin rights). Deploy will continue; consent must be granted manually."
    } else {
        WriteW -message "Admin consent skipped. Deploy will continue; consent must be granted manually before the app can be used."
    }

    if (-not $script:AdminConsentPending) { $script:AdminConsentPending = @() }
    $script:AdminConsentPending += [PSCustomObject]@{
        AppId = $graphAppId
        ConsentUrl = "https://login.microsoftonline.com/$($parameters.tenantId.value)/adminconsent?client_id=$graphAppId"
    }
}

# Azure AD app update. Assigning Admin-consent,RedirectUris,IdentifierUris,Optionalclaim etc.
# Rewritten to use az CLI + Microsoft Graph REST API (az rest) instead of the retired AzureAD PS module.

# Classify az/Graph stderr as transient (timeout / 429 / 5xx / connection reset) so we know whether to retry.
function Test-TransientAzError {
    Param([string]$ErrText)
    return ($ErrText -match 'ReadTimeout|Read timed out|WinError 10060|HTTPSConnectionPool|TooManyRequests|\b429\b|\b50[234]\b|RemoteDisconnected|ConnectionError|ServiceUnavailable|GatewayTimeout|Connection aborted')
}

# Invoke an `az` command with exponential backoff retry on transient Graph failures.
# Returns stdout (string). Throws after exhausting retries OR on non-transient errors (no wasted wait).
function Invoke-AzWithRetry {
    Param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$Label = 'az call',
        [int]$MaxAttempts = 5,
        [int]$InitialDelaySeconds = 10
    )
    $delay = $InitialDelaySeconds
    $errText = ''
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        $output = & $ScriptBlock 2>&1
        if ($LASTEXITCODE -eq 0) { return ($output -join "`n") }
        $errText = ($output | Out-String)
        if (-not (Test-TransientAzError $errText)) {
            throw "$Label failed (non-transient, attempt $i): $errText"
        }
        if ($i -lt $MaxAttempts) {
            WriteW -message "$Label transient failure (attempt $i/$MaxAttempts). Retrying in ${delay}s..."
            Start-Sleep -Seconds $delay
            $delay = [Math]::Min($delay * 2, 60)
        }
    }
    throw "$Label failed after $MaxAttempts attempts. Last error: $errText"
}

# Helper: write JSON body to a temp file and call az rest PATCH.
# Avoids PowerShell→external-command argument escaping with complex JSON bodies.
# Retries on transient failures (Graph ReadTimeout / TCP reset / throttling) and throws on final failure
# so callers cannot silently proceed past a failed write.
function AzRestPatch {
    Param(
        [Parameter(Mandatory=$true)] [string]$Url,
        [Parameter(Mandatory=$true)] [string]$Body,
        [int]$MaxAttempts = 5,
        [int]$InitialDelaySeconds = 10
    )
    $tmp = [System.IO.Path]::GetTempFileName()
    try {
        [System.IO.File]::WriteAllText($tmp, $Body, [System.Text.Encoding]::UTF8)
        Invoke-AzWithRetry -Label "Graph PATCH $Url" -MaxAttempts $MaxAttempts -InitialDelaySeconds $InitialDelaySeconds -ScriptBlock {
            az rest --method PATCH --url $Url --body "@$tmp" --headers "Content-Type=application/json"
        } | Out-Null
    } finally {
        Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
    }
}

# Pre-create diagnostic settings on every site BEFORE tenant 'deployIfNotExists' policies fire their
# remediation. Each policy remediation triggers an internal Microsoft.Web UpdateWebSite, which recycles
# the .NET-isolated worker. When the new worker loses the race to rebind local gRPC port 4001 to the
# prior worker, the host enters an unrecoverable 'Failed to start language worker process' loop and the
# function app silently wedges (no telemetry). Pre-creating the setting makes the policy's
# existenceCondition immediately compliant -> no remediation.
#
# Two modes (both run, complementary):
#   1. EXPLICIT: caller passes -SettingsSpec — array of @{ name=...; workspaceId=... } pairs sourced
#      from parameters.json:policyDiagnosticSettings. Best for tenants with known DINE policies.
#   2. AUTO-REPLICATE: any diagnostic setting present on a SUBSET of the 4 sites is copied to the missing
#      ones. Catches unknown-tenant policies that hit one site faster than another, reducing blast radius
#      from N wedges to 1. Always runs as a safety net.
function Set-PolicyDiagnosticSettings {
    Param(
        [Parameter(Mandatory=$true)][string]$ResourceGroup,
        [Parameter(Mandatory=$true)][string]$BaseResourceName,
        [Parameter(Mandatory=$true)][string]$SubscriptionId,
        # Array of [pscustomobject]@{ name=<string>; workspaceId=<armId> } (explicit known-tenant entries).
        $SettingsSpec = @()
    )

    $apiVersion = '2021-05-01-preview'
    $sites = @(
        @{ name = $BaseResourceName;                       kind = 'web'      },
        @{ name = "$BaseResourceName-function";            kind = 'function' },
        @{ name = "$BaseResourceName-prep-function";       kind = 'function' },
        @{ name = "$BaseResourceName-data-function";       kind = 'function' }
    )

    # --- Snapshot current diagnostic settings on every site ---
    WriteI -message "Scanning existing diagnostic settings across $($sites.Count) sites..."
    $perSite = @{}        # site name -> hashtable(settingName -> setting object)
    $catsBySite = @{}     # site name -> log-category names array (cached for the PUT phase)
    foreach ($s in $sites) {
        $perSite[$s.name] = @{}
        $url = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/$($s.name)/providers/Microsoft.Insights/diagnosticSettings?api-version=$apiVersion"
        $raw = az rest --method GET --url $url 2>$null
        if ($LASTEXITCODE -eq 0 -and $raw) {
            try {
                foreach ($d in (($raw | ConvertFrom-Json).value)) { $perSite[$s.name][$d.name] = $d }
            } catch { }
        }
    }

    # --- Build the union of (settingName -> workspaceId) targets ---
    # Priority: explicit SettingsSpec wins; otherwise infer from any site that already has the setting.
    $targets = @{}   # settingName -> workspaceId
    foreach ($spec in $SettingsSpec) {
        if ($spec -and $spec.name -and $spec.workspaceId) {
            $targets[[string]$spec.name] = [string]$spec.workspaceId
        }
    }
    foreach ($siteName in $perSite.Keys) {
        foreach ($settingName in $perSite[$siteName].Keys) {
            if (-not $targets.ContainsKey($settingName)) {
                $ws = $perSite[$siteName][$settingName].properties.workspaceId
                if ($ws) { $targets[$settingName] = $ws }
            }
        }
    }

    if ($targets.Count -eq 0) {
        WriteI -message "No explicit policyDiagnosticSettings configured and no existing diagnostic settings found. Skipping (safe no-op for tenants without diagnostic-setting policies)."
        return
    }

    # --- Helper: lazy-load log categories for a site ---
    $getCats = {
        Param([string]$siteName)
        if ($catsBySite.ContainsKey($siteName)) { return $catsBySite[$siteName] }
        $catsUrl = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/$siteName/providers/Microsoft.Insights/diagnosticSettingsCategories?api-version=$apiVersion"
        $catsRaw = az rest --method GET --url $catsUrl 2>$null
        if ($LASTEXITCODE -ne 0 -or -not $catsRaw) { return $null }
        $names = @()
        foreach ($c in (($catsRaw | ConvertFrom-Json).value)) {
            if ($c.properties.categoryType -eq 'Logs') { $names += $c.name }
        }
        $catsBySite[$siteName] = $names
        return $names
    }

    # --- For each target setting, PUT it on every site missing it ---
    foreach ($settingName in $targets.Keys) {
        $workspaceId = $targets[$settingName]
        WriteI -message "Setting '$settingName' (workspace: $workspaceId)"

        foreach ($s in $sites) {
            if ($perSite[$s.name].ContainsKey($settingName)) {
                WriteI -message "  [$($s.name)] already present — skipping."
                continue
            }

            $catNames = & $getCats $s.name
            if (-not $catNames -or $catNames.Count -eq 0) {
                WriteW -message "  [$($s.name)] could not list categories — skipping."
                continue
            }

            $primary = if ($s.kind -eq 'function') { 'FunctionAppLogs' } else { 'AppServiceHTTPLogs' }
            $logsArr = @()
            foreach ($cn in $catNames) {
                $logsArr += @{
                    category = $cn
                    enabled  = ($cn -eq $primary)
                    retentionPolicy = @{ enabled = $false; days = 0 }
                }
            }
            $metricsArr = @( @{ category = 'AllMetrics'; enabled = $false; retentionPolicy = @{ enabled = $false; days = 0 } } )

            $bodyObj  = @{ properties = @{ workspaceId = $workspaceId; metrics = $metricsArr; logs = $logsArr } }
            $bodyJson = $bodyObj | ConvertTo-Json -Depth 10 -Compress

            $putUrl = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroup/providers/Microsoft.Web/sites/$($s.name)/providers/Microsoft.Insights/diagnosticSettings/$settingName" + "?api-version=$apiVersion"
            $tmp = [System.IO.Path]::GetTempFileName()
            try {
                [System.IO.File]::WriteAllText($tmp, $bodyJson, [System.Text.Encoding]::UTF8)
                try {
                    Invoke-AzWithRetry -Label "PUT diagnostic setting '$settingName' on $($s.name)" -ScriptBlock {
                        az rest --method PUT --url $putUrl --body "@$tmp" --headers "Content-Type=application/json"
                    } | Out-Null
                    WriteS -message "  [$($s.name)] pre-created ($primary enabled)."
                } catch {
                    WriteW -message "  [$($s.name)] PUT failed: $($_.Exception.Message). Policy may fire its own remediation here."
                }
            } finally {
                Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# Defense-in-depth: enable Proactive Auto-Heal on every site so App Service automatically recycles
# instances that exceed memory thresholds or show signs of wedge. This is a built-in App Service
# feature (no autoHealRules to author) that complements the policy-hardening mitigation above — if
# a host ever does wedge for any reason, it self-recovers within minutes instead of hours.
# Only writes the setting when missing or false to avoid an unnecessary recycle on idempotent re-runs.
function Enable-ProactiveAutoHeal {
    Param(
        [Parameter(Mandatory=$true)][string]$ResourceGroup,
        [Parameter(Mandatory=$true)][string]$BaseResourceName
    )

    $appSettingName  = 'WEBSITE_PROACTIVE_AUTOHEAL_ENABLED'
    $appSettingValue = 'true'
    $sites = @(
        $BaseResourceName,
        "$BaseResourceName-function",
        "$BaseResourceName-prep-function",
        "$BaseResourceName-data-function"
    )

    foreach ($name in $sites) {
        $current = az webapp config appsettings list --name $name --resource-group $ResourceGroup --query "[?name=='$appSettingName'].value | [0]" -o tsv 2>$null
        if ($LASTEXITCODE -ne 0) {
            WriteW -message "  [$name] could not read appsettings — skipping."
            continue
        }
        if ($current -eq $appSettingValue) {
            WriteI -message "  [$name] $appSettingName already set to '$appSettingValue' — skipping."
            continue
        }

        try {
            az webapp config appsettings set --name $name --resource-group $ResourceGroup --settings "$appSettingName=$appSettingValue" --output none 2>$null
            if ($LASTEXITCODE -eq 0) {
                WriteS -message "  [$name] enabled Proactive Auto-Heal."
            } else {
                WriteW -message "  [$name] az webapp config appsettings set returned $LASTEXITCODE."
            }
        } catch {
            WriteW -message "  [$name] failed to set Proactive Auto-Heal: $($_.Exception.Message)"
        }
    }
}

# Resolve an application's object ID from its appId, retrying through AAD replication lag.
# After 'az ad app create' the new app may take several seconds to be queryable by appId; without this
# retry, downstream Graph PATCHes hit https://graph.microsoft.com/v1.0/applications/ (empty id) and 405.
function Resolve-AppObjectId {
    Param(
        [Parameter(Mandatory=$true)] [string]$AppId,
        [int]$MaxAttempts = 10,
        [int]$RetryDelaySeconds = 6
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        $raw = az ad app show --id $AppId 2>$null
        if ($LASTEXITCODE -eq 0 -and $raw) {
            try {
                $obj = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($obj -and $obj.id) { return $obj }
            } catch { }
        }
        if ($i -lt $MaxAttempts) {
            WriteW -message "az ad app show returned no objectId for appId=$AppId (attempt $i/$MaxAttempts). Retrying in ${RetryDelaySeconds}s..."
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
    throw "Resolve-AppObjectId: failed to resolve object id for appId=$AppId after $MaxAttempts attempts."
}

function ADAppUpdate {
    Param(
        [Parameter(Mandatory = $true)] $appdomainName,
        [Parameter(Mandatory = $true)] $appId
    )
    $configAppId = $appId
    $azureDomainBase = $appdomainName
    $configAppUrl = "https://$azureDomainBase"
    $RedirectUris = ($configAppUrl + '/signin-simple-end')
    # Teams SSO requires the resource URI to contain the iframe host so it can match against the tab origin.
    # Form: api://<appDomain>/<appId> (contains appId so it also satisfies tenant ID-URI policy).
    $IdentifierUris = "api://$azureDomainBase/$configAppId"

    # set subscription
    az account set --subscription $parameters.subscriptionId.Value

    # Assigning graph permissions FIRST so admin consent below has scopes to grant.
    # Previously consent ran before this and silently no-oped (no scopes on the app yet),
    # leaving the Teams web app to fail every Graph OBO call with AADSTS65001.
    $manifestPath = Join-Path $PSScriptRoot 'AadAppManifest.json'
    Invoke-AzWithRetry -Label "ad app update required-resource-accesses" -ScriptBlock {
        az ad app update --id $configAppId --required-resource-accesses $manifestPath
    } | Out-Null

    # Grant Admin consent AFTER required permissions are written to the app.
    GrantAdminConsent $configAppId

    # Get the object ID (required for Graph API calls; appId and objectId are different).
    # Retries through AAD replication lag — without it, a freshly created app sometimes returns empty,
    # producing PATCH https://graph.microsoft.com/v1.0/applications/ (405 Method Not Allowed).
    $appObjRaw = Resolve-AppObjectId -AppId $configAppId
    $applicationObjectId = $appObjRaw.id

    # Fetch full application object via Graph API
    $app = (Invoke-AzWithRetry -Label "GET application $applicationObjectId" -ScriptBlock { az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" }) | ConvertFrom-Json

    # Ensure web settings are correct (redirectUris + implicit id_token issuance for the popup auth flow).
    # PATCH on `web` replaces the whole complex type, so redirectUris and implicitGrantSettings MUST be set together —
    # patching only implicitGrantSettings would wipe redirectUris (the root cause of an earlier silent failure).
    $webOk = ($app.web.redirectUris -contains $RedirectUris) -and $app.web.implicitGrantSettings.enableIdTokenIssuance
    if (-not $webOk) {
        $webPatch = @{ web = @{
            redirectUris          = @($RedirectUris)
            implicitGrantSettings = @{ enableIdTokenIssuance = $true; enableAccessTokenIssuance = $false }
        } } | ConvertTo-Json -Depth 10 -Compress
        AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $webPatch
        $appAfterWeb = (Invoke-AzWithRetry -Label "GET application (web verify) $applicationObjectId" -ScriptBlock { az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" }) | ConvertFrom-Json
        if ($appAfterWeb.web.redirectUris -notcontains $RedirectUris) {
            throw "Failed to set web.redirectUris to $RedirectUris on the Graph app."
        }
        if (-not $appAfterWeb.web.implicitGrantSettings.enableIdTokenIssuance) {
            throw "Failed to enable implicit id_token issuance on the Graph app."
        }
        WriteI -message "App web settings set (redirectUris=$RedirectUris, implicit id_token=true)"
    }

    # Do nothing if the app has already been fully configured (identifier URIs + access_as_user scope)
    $existingUserScope = $app.api.oauth2PermissionScopes | Where-Object { $_.value -eq 'access_as_user' }
    if ($app.identifierUris.Count -gt 0 -and $null -ne $existingUserScope) {
        WriteS -message "Graph application is already configured."
        return
    }
    WriteI -message "`nUpdating graph app..."

    # Disable then remove the default user_impersonation scope (Graph API requires two-step: disable then delete)
    $existingScopes = $app.api.oauth2PermissionScopes
    if ($existingScopes.Count -gt 0) {
        $disabledScopes = $existingScopes | ForEach-Object { $_.isEnabled = $false; $_ }
        $disablePatch = @{ api = @{ oauth2PermissionScopes = $disabledScopes } } | ConvertTo-Json -Depth 10 -Compress
        AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $disablePatch
        AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body '{"api":{"oauth2PermissionScopes":[]}}'
    }

    # Set the single Application ID URI expected by the Teams manifest (api://<appId> form).
    $identifierUriPatch = ('{"identifierUris":["' + $IdentifierUris + '"]}')
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $identifierUriPatch

    $appAfterIdentifierUriUpdate = (Invoke-AzWithRetry -Label "GET application (uri verify) $applicationObjectId" -ScriptBlock { az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" }) | ConvertFrom-Json
    if ($appAfterIdentifierUriUpdate.identifierUris -notcontains $IdentifierUris) {
        throw "Failed to set Application ID URI to $IdentifierUris on the Graph app."
    }
    WriteI -message "App identifier URI set ($IdentifierUris)"

    # Set optionalClaims via Graph PATCH (replaces the unverified `az ad app update --optional-claims` CLI call,
    # which silently swallows transient Graph ReadTimeouts).
    $optionalClaimsPath = Join-Path $PSScriptRoot 'AadOptionalClaims.json'
    $optionalClaimsObj = Get-Content -Path $optionalClaimsPath -Raw | ConvertFrom-Json
    $optionalClaimsPatch = @{ optionalClaims = $optionalClaimsObj } | ConvertTo-Json -Depth 10 -Compress
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $optionalClaimsPatch
    WriteI -message "App optionalclaim set."

    # Create access_as_user scope via Graph API PATCH
    $scopeId = (New-Guid).ToString()
    $newScope = @{
        id                      = $scopeId
        value                   = "access_as_user"
        userConsentDisplayName  = "Access the API as the current logged-in user."
        userConsentDescription  = "Access the API as the current logged-in user."
        adminConsentDisplayName = "Access the API as the current logged-in user."
        adminConsentDescription = "Access the API as the current logged-in user."
        isEnabled               = $true
        type                    = "User"
    }
    $scopePatch = @{ api = @{ oauth2PermissionScopes = @($newScope) } } | ConvertTo-Json -Depth 10 -Compress
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $scopePatch
    WriteI -message "Scope access_as_user added."

    # Pre-authorize Teams mobile/desktop and web clients to access the API
    $preAuthApps = @(
        @{ appId = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'; delegatedPermissionIds = @($scopeId) }
        @{ appId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'; delegatedPermissionIds = @($scopeId) }
    )
    $preAuthPatch = @{ api = @{ preAuthorizedApplications = $preAuthApps } } | ConvertTo-Json -Depth 10 -Compress
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $preAuthPatch
    WriteI -message "Teams mobile/desktop and web clients applications pre-authorized."

    # Set requestedAccessTokenVersion = 2 so AAD issues v2 tokens for Teams SSO.
    # Required because Microsoft.Identity.Web disables inbound claim mapping and expects v2 JWT claim names.
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body '{"api":{"requestedAccessTokenVersion":2}}'
    WriteI -message "requestedAccessTokenVersion set to 2 (Teams SSO v2 JWT format)"
}

# update app name
# Rewritten to use az CLI directly (AzureAD module lookup was only needed to get the object ID, which is unused here).
function ADAppUpdateDisplayName{
    Param(
        [Parameter(Mandatory = $true)] $appId,
        [Parameter(Mandatory = $true)] $currentName,
        [Parameter(Mandatory = $true)] $newName
    )
    Invoke-AzWithRetry -Label "ad app update display-name '$newName'" -ScriptBlock {
        az ad app update --id $appId --display-name $newName
    } | Out-Null
}

# Removing existing access of app (used on upgrades to reset the authors app before re-configuring).
# Rewritten to use az CLI + Microsoft Graph REST API instead of the retired AzureAD PS module.
function FormatAADApp {
    Param(
        [Parameter(Mandatory = $true)] $appId,
        [Parameter(Mandatory = $true)] $appName
    )

    # Get the object ID (required for Graph API calls). Same propagation-safe retry as in ADAppUpdate.
    $appObjRaw = Resolve-AppObjectId -AppId $appId
    $applicationObjectId = $appObjRaw.id

    # Fetch full application object via Graph API
    $app = (Invoke-AzWithRetry -Label "GET application (FormatAADApp) $applicationObjectId" -ScriptBlock { az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" }) | ConvertFrom-Json

    # Do nothing if the app has already been reset
    if ($app.identifierUris.Count -eq 0) {
        WriteS -message "App already configured."
        return
    }

    WriteI -message "`nUpdating app..."

    # Disable then remove existing oauth2 permission scopes (two-step required by Graph API)
    $existingScopes = $app.api.oauth2PermissionScopes
    if ($existingScopes.Count -gt 0) {
        $disabledScopes = $existingScopes | ForEach-Object { $_.isEnabled = $false; $_ }
        $disablePatch = @{ api = @{ oauth2PermissionScopes = $disabledScopes } } | ConvertTo-Json -Depth 10 -Compress
        AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $disablePatch
        AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body '{"api":{"oauth2PermissionScopes":[]}}'
    }

    # Clear implicit grant, redirect URIs, and identifier URIs in one PATCH
    $resetPatch = '{"web":{"implicitGrantSettings":{"enableIdTokenIssuance":false},"redirectUris":[]},"identifierUris":[]}'
    AzRestPatch -Url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" -Body $resetPatch

    Invoke-AzWithRetry -Label "ad app update optional-claims (reset)" -ScriptBlock {
        az ad app update --id $appId --optional-claims './AadOptionalClaims_Reset.json'
    } | Out-Null
    Invoke-AzWithRetry -Label "ad app update remove requiredResourceAccess" -ScriptBlock {
        az ad app update --id $appId --remove requiredResourceAccess
    } | Out-Null
}
#update manifest file and create a .zip file.
function GenerateAppManifestPackage {
    Param(
        [Parameter(Mandatory = $true)] [ValidateSet('authors', 'users')] $manifestType,
        [Parameter(Mandatory = $true)] $appdomainName,
        [Parameter(Mandatory = $true)] $appId,
        [Parameter(Mandatory = $false)] $graphAppId
    )

        WriteI -message "`nGenerating package for $manifestType..."

        $azureDomainBase = $appdomainName
        $sourceManifestPath = "..\Manifest\manifest_$manifestType.json"
        $destManifestFilePath = '..\Manifest\manifest.json'
        $destinationZipPath = "..\manifest\CC-$manifestType.zip"

    if (!(Test-Path $sourceManifestPath)) {
        throw "$sourceManifestPath does not exist. Please make sure you download the full app template source."
    }

    copy-item -path $sourceManifestPath -destination $destManifestFilePath -Force

    # Replace merge fields with proper values in manifest file and save
        $buildVersion = "5.$((Get-Date).ToString('yy')).$([int]((Get-Date).ToString('Mdd')))"
        # webApplicationInfo.resource must target the SSO/main app and contain the iframe host so Teams SSO can match it against the tab origin.
        # Form: api://<appDomain>/<mainAppId>
        if ($graphAppId) {
        $ssoAppId = $graphAppId
    } else {
        $ssoAppId = $appId
    }
        $identifierUri = "api://$azureDomainBase/$ssoAppId"
        $mergeFields = @{
            '<<companyName>>'   = $parameters.companyName.Value
            '<<botId>>'         = $appId
            '<<appDomain>>'     = $azureDomainBase
            '<<websiteUrl>>'    = $parameters.websiteUrl.Value
            '<<privacyUrl>>'    = $parameters.privacyUrl.Value
            '<<termsOfUseUrl>>' = $parameters.termsOfUseUrl.Value
            '<<graphAppId>>'    = $ssoAppId
            '<<identifierUri>>' = $identifierUri
            '<<version>>'       = $buildVersion
        }
        $appManifestContent = Get-Content $destManifestFilePath
        foreach ($mergeField in $mergeFields.GetEnumerator()) {
            $appManifestContent = $appManifestContent.replace($mergeField.Name, $mergeField.Value)
        }
        $appManifestContent | Set-Content $destManifestFilePath -Force

    # Generate zip archive
        $compressManifest = @{
            LiteralPath      = "..\manifest\color.png", "..\manifest\outline.png", $destManifestFilePath
            CompressionLevel = "Fastest"
            DestinationPath  = $destinationZipPath
        }
        Compress-Archive @compressManifest -Force

        Remove-Item $destManifestFilePath -ErrorAction Continue

        WriteS -message "Package has been created under this path $(Resolve-Path $destinationZipPath)"
}

function logout {
    az logout | Out-Null
}

# ---------------------------------------------------------
# DEPLOYMENT SCRIPT
# ---------------------------------------------------------

# Check if Azure CLI is installed.
    WriteI -message "Checking if Azure CLI is installed."
    $localPath = [Environment]::GetEnvironmentVariable("ProgramFiles(x86)")
    if ($localPath -eq $null) {
        $localPath = "C:\Program Files (x86)"
    }

    $localPath = $localPath + "\Microsoft SDKs\Azure\CLI2"
    If (-not(Test-Path -Path $localPath)) {
        WriteW -message "Azure CLI is not installed!"
        $confirmationtitle      = "Please select YES to install Azure CLI."
        $confirmationquestion   = "Do you want to proceed?"
        $confirmationchoices    = "&yes", "&no" # 0 = yes, 1 = no

        $updatedecision = $host.ui.promptforchoice($confirmationtitle, $confirmationquestion, $confirmationchoices, 1)
        if ($updatedecision -eq 0) {
            WriteI -message "Installing Azure CLI ..."
            Invoke-WebRequest -Uri https://aka.ms/installazurecliwindows -OutFile .\AzureCLI.msi; Start-Process msiexec.exe -Wait -ArgumentList '/I AzureCLI.msi /quiet'; rm .\AzureCLI.msi
            WriteS -message "Azure CLI is installed! Please close this PowerShell window and re-run this script in a new PowerShell session."
            EXIT
        } else {
            WriteE -message "Azure CLI is not installed.`nPlease install the CLI from https://docs.microsoft.com/en-us/cli/azure/install-azure-cli?view=azure-cli-latest and re-run this script in a new PowerShell session"
            EXIT
        }
    } else {
        WriteS -message "Azure CLI is installed."
    }

# Check optional modules
    if ((Get-Module -ListAvailable -Name "Az.*")) {
        WriteI -message "Az module is available."
    } else {
        WriteI -message "Az module is not installed. (optional)."
    }

# Load Parameters from JSON meta-data file
    $parametersListContent = Get-Content '.\parameters.json' -ErrorAction Stop

# Validate all the parameters.
    WriteI -message "Validating all the parameters from parameters.json."
    $parameters = $parametersListContent | ConvertFrom-Json

# Resolve deployer IP (used for storage/KV firewall and SCM access restrictions).
    if ($DeployerIp) {
        $script:DeployerIpToUse = $DeployerIp.Trim()
        WriteI -message "Deployer IP (from -DeployerIp): $script:DeployerIpToUse"
    } elseif ($parameters.PSObject.Properties.Match('deployerIpAddress') -and $parameters.deployerIpAddress.Value) {
        $script:DeployerIpToUse = ([string]$parameters.deployerIpAddress.Value).Trim()
        WriteI -message "Deployer IP (from parameters.json): $script:DeployerIpToUse"
    } else {
        try {
            $script:DeployerIpToUse = ((Invoke-RestMethod -Uri 'https://api.ipify.org' -TimeoutSec 10) -as [string]).Trim()
            WriteI -message "Deployer IP (auto-detected via api.ipify.org): $script:DeployerIpToUse"
        } catch {
            $script:DeployerIpToUse = ''
            WriteW -message "Could not auto-detect deployer IP ($($_.Exception.Message)). Continuing without operator IP allow rule; you will not be able to reach storage/KV/SCM from this machine."
        }
    }

    if (-not(ValidateParameters)) {
        WriteE -message "Invalid parameters found. Please update the parameters in the parameters.json with valid values and re-run the script."
        EXIT
    }

# Start Deployment.
    $buildVersion = "5.$((Get-Date).ToString('yy')).$([int]((Get-Date).ToString('Mdd')))"
    Write-Host ""
    Write-Host "===== Company Communicator v$buildVersion =====" -ForegroundColor Magenta
    Write-Host ""
    WriteI -message "Starting deployment..."

# Initialize connections - Azure CLI
    WriteI -message "Login with your Azure subscription account. Launching Azure sign-in window..."
    $user = az login --tenant $parameters.subscriptionTenantId.value
    if ($LASTEXITCODE -ne 0) {
        WriteE -message "Login failed for user..."
        EXIT
    }
    az account set --subscription $parameters.subscriptionId.value | Out-Null
    # Graph API calls use 'az rest' which reuses this az login session.
    $userAlias = (($user | ConvertFrom-Json) | where {$_.id -eq $parameters.subscriptionId.Value}).user.name


# Validate the name of resources to be created.
    if (-not(validateresourcesnames)) {
        WriteE -message "Please choose a different baseResourceName in the parameters.json and re-run the script. Exiting..."
        logout
        EXIT
    }

# Create or Update User App
	$usersApp = $parameters.baseresourcename.Value + '-users'
	$userAppCred = $null

	if($parameters.isUpgrade.Value){
        $currentAppName = $parameters.baseresourcename.Value
		$userAppCred = GetAzureADAppWithSecret $currentAppName
		ADAppUpdateDisplayName $userAppCred.appId $currentAppName $usersApp
	}
	else
	{
		# SingleTenant: must match azuredeploy.json msaAppType for bot connector auth.
		$userAppCred = CreateAzureADApp -AppName $usersApp -ResetAppSecret $True -MultiTenant $False
		if ($null -eq $userAppCred) {
			WriteE -message "Failed to create or update User app in Azure Active Directory. Exiting..."
			logout
			Exit
		}
	}

# Create Author App
    $authorsApp = $parameters.baseResourceName.Value + '-authors'
	$authorAppCred = $null
	if($parameters.isUpgrade.Value){
		$authorAppCred = GetAzureADAppWithSecret $authorsApp
	}
	else
	{
		# SingleTenant: must match azuredeploy.json msaAppType for bot connector auth.
		$authorAppCred = CreateAzureADApp -AppName $authorsApp -ResetAppSecret $True -MultiTenant $False
		if ($null -eq $authorAppCred) {
        WriteE -message "Failed to create or update the Author app in Azure Active Directory. Exiting..."
        logout
        Exit
		}
	}

# Create Company Communicator App
	$graphApp = $parameters.baseResourceName.Value
	$graphAppCred = CreateAzureADApp -AppName $graphApp -ResetAppSecret $True -MultiTenant $False
	if ($null -eq $graphAppCred) {
		WriteE -message "Failed to create or update the main app in Azure Active Directory. Exiting..."
		logout
		Exit
	}

# Function call to Deploy ARM Template
    $deploymentOutput = $null
    $appDisplayName = $null
        $deploymentOutput = DeployARMTemplate $graphAppCred.appId $authorAppCred.appId $userAppCred.appId $graphAppCred.password $authorAppCred.password $userAppCred.password

        # Reading the deployment output.
        WriteI -message "Reading deployment outputs..."
        if(($null -eq $deploymentOutput) -or ($null -eq $deploymentOutput.properties) -or ($null -eq $deploymentOutput.properties.Outputs) -or ($null -eq $deploymentOutput.properties.Outputs.keyVaultName) -or ($null -eq $deploymentOutput.properties.Outputs.keyVaultName.Value))
        {
            $keyVaultName = $parameters.BaseResourceName.Value + 'vault'
            if($parameters.customDomainOption.Value -eq 'Azure Front Door')
            {
               WriteW -message "ARM deployment outputs are missing. The AFD Standard endpoint hostname is Azure-generated (format <name>-<hash>.z01.azurefd.net) and cannot be derived from the base resource name like AFD Classic could."
               WriteW -message "Retrieve the real hostname from the Azure portal (Front Door profile -> endpoint -> Endpoint hostname) or run: az afd endpoint show -n $($parameters.BaseResourceName.Value) --profile-name $($parameters.BaseResourceName.Value) -g <resourceGroup> --query hostName -o tsv"
               $appdomainName = '<retrieve-from-azure-portal-afd-standard-endpoint-hostname>'
            }
            else
            {
                $appdomainName = 'Please create a custom domain name for ' + $parameters.BaseResourceName.Value + ' and use that in the manifest'
            }
        }
        else
        {
            # Assigning return values to variable.
            $appdomainName = $deploymentOutput.properties.Outputs.appDomain.Value
        }
    if ($null -eq $deploymentOutput) {
        WriteE -message "Encountered an error during ARM template deployment. Exiting..."
        logout
        Exit
    }

# Tenant-policy hardening: pre-create diagnostic settings on every site BEFORE any tenant
# 'deployIfNotExists' policy can fire its remediation (which triggers UpdateWebSite -> host recycle
# race -> .NET-isolated worker wedge). Combines an explicit list from parameters.json with an
# auto-replicate safety net that copies any setting already on one site to its siblings.
    WriteI -message "Applying tenant-policy hardening (diagnostic-setting pre-creation)..."
    try {
        $diagSpec = @()
        if ($parameters.PSObject.Properties.Match('policyDiagnosticSettings') -and $parameters.policyDiagnosticSettings.Value) {
            $diagSpec = @($parameters.policyDiagnosticSettings.Value)
        }
        Set-PolicyDiagnosticSettings -ResourceGroup $parameters.resourceGroupName.Value -BaseResourceName $parameters.baseResourceName.Value -SubscriptionId $parameters.subscriptionId.Value -SettingsSpec $diagSpec
    } catch {
        WriteW -message "Diagnostic-setting pre-creation step failed: $($_.Exception.Message). Continuing — tenant policy may perform its own remediation (may cause one-time worker recycle)."
    }

# Defense-in-depth: enable Proactive Auto-Heal on every site so any future wedge self-recovers.
    WriteI -message "Enabling Proactive Auto-Heal across all sites..."
    try {
        Enable-ProactiveAutoHeal -ResourceGroup $parameters.resourceGroupName.Value -BaseResourceName $parameters.baseResourceName.Value
    } catch {
        WriteW -message "Enable-ProactiveAutoHeal step failed: $($_.Exception.Message). Continuing."
    }

# Function call to update reply-urls and uris for registered app.
    WriteI -message "Updating required parameters and urls..."
	if($parameters.isUpgrade.Value){
    FormatAADApp $authorAppCred.appId $authorsApp
	}
    ADAppUpdate $appdomainName $graphAppCred.appId

# Log out to avoid tokens caching
    logout

# Function call to generate manifest.zip folder for User and Author.
    GenerateAppManifestPackage 'authors' $appdomainName $authorAppCred.appId $graphAppCred.appId
    GenerateAppManifestPackage 'users' $appdomainName $userAppCred.appId $graphAppCred.appId

# Open manifest folder
    Invoke-Item ..\Manifest\

# Deployment completed.
    Write-Host ""
    Write-Host "===== DEPLOYMENT COMPLETED =====" -ForegroundColor Green
    Write-Host ""

    if ($script:AdminConsentPending -and $script:AdminConsentPending.Count -gt 0) {
        Write-Host ""
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host " ADMIN CONSENT NOT GRANTED" -ForegroundColor Yellow
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host "The deploy completed, but the Company Communicator app will not" -ForegroundColor Yellow
        Write-Host "work until a Global Admin (or Privileged Role Administrator)" -ForegroundColor Yellow
        Write-Host "grants admin consent to the bot app registration." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Open this URL signed in as an AAD admin:" -ForegroundColor Yellow
        foreach ($p in $script:AdminConsentPending) {
            Write-Host ("  $($p.ConsentUrl)") -ForegroundColor Cyan
        }
        Write-Host "================================================================" -ForegroundColor Yellow
        Write-Host ""
    }

try { Stop-Transcript | Out-Null } catch { }
