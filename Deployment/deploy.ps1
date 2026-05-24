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
    $nameproperty = if ($servicetype -eq 'cognitiveservice') { "subdomainname" } else { "name" }
    $body = '"{0}": "{1}", "type": "{2}"' -f $nameproperty, $name, $typebyservicetype[$servicetype]

    $response = (invoke-webrequest -uri $uri -method post -body "{$body}" -contenttype "application/json" -headers @{authorization = $authorizationtoken } -usebasicparsing).content
    $response | convertfrom-json |
    select-object @{n = 'name'; e = { $name } }, @{n = 'type'; e = { $servicetype } }, @{n = 'available'; e = { $_ | select-object -expandproperty *available } }, reason, message
}

# To get the Azure AD app detail (with retry for transient Graph API timeouts).
function GetAzureADApp {
    param ($appName)
    $maxAttempts = 4
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $app = az ad app list --filter "displayName eq '$appName'" | ConvertFrom-Json
        if ($LASTEXITCODE -eq 0) { return $app }
        if ($attempt -lt $maxAttempts) {
            WriteW -message "Graph API timeout looking up '$appName'. Retrying in 15s ($attempt/$maxAttempts)..."
            Start-Sleep -Seconds 15
        }
    }
    WriteE -message "Failed to look up Azure AD app '$appName' after $maxAttempts attempts."
    return $null
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
        [Parameter(Mandatory = $false)] [bool] $MultiTenant = $true,
        [Parameter(Mandatory = $false)] [bool] $AllowImplicitFlow
    )

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

                $updateSuccess = $false
                for ($updateAttempt = 1; $updateAttempt -le 4; $updateAttempt++) {
                    az ad app update --id $app.appId --display-name $appName --sign-in-audience AzureADMultipleOrgs | Out-Null
                    if ($LASTEXITCODE -eq 0) { $updateSuccess = $true; break }
                    if ($updateAttempt -lt 4) {
                        WriteW -message "App update timed out. Retrying in 15s ($updateAttempt/4)..."
                        Start-Sleep -Seconds 15
                    }
                }
                if (-not $updateSuccess) {
                    WriteE -message "Failed to update Azure AD app '$appName' after 4 attempts."
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
            $createSuccess = $false
            for ($createAttempt = 1; $createAttempt -le 4; $createAttempt++) {
                az ad app create --display-name $appName --sign-in-audience AzureADMultipleOrgs | Out-Null
                if ($LASTEXITCODE -eq 0) { $createSuccess = $true; break }
                if ($createAttempt -lt 4) {
                    WriteW -message "App creation timed out. Retrying in 15s ($createAttempt/4)..."
                    Start-Sleep -Seconds 15
                }
            }
            if (-not $createSuccess) {
                WriteE -message "Failed to create Azure AD app '$appName' after 4 attempts."
                return $null
            }

            WriteI -message "Waiting for app creation to finish..."

            Start-Sleep -s 10

            WriteS -message "Azure AD App: $appName is created."
        }

        $app = GetAzureADApp $appName

        $appSecret = $null;
        #Reset the app credentials to get the secret. The default validity of this secret will be for 1 year from the date its created.
        if ($ResetAppSecret) {
            if ($null -eq $app) {
                WriteE -message "Cannot create secret for '$appName': app lookup returned null."
                return $null
            }
            WriteI -message "Updating app secret..."
            $maxSecretAttempts = 4
            for ($secretAttempt = 1; $secretAttempt -le $maxSecretAttempts; $secretAttempt++) {
                $appSecret = az ad app credential reset --id $app.appId --append | ConvertFrom-Json;
                if ($LASTEXITCODE -eq 0 -and $null -ne $appSecret) { break }
                if ($secretAttempt -lt $maxSecretAttempts) {
                    WriteW -message "Secret creation timed out. Retrying in 15s ($secretAttempt/$maxSecretAttempts)..."
                    Start-Sleep -Seconds 15
                }
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

    if($parameters.useCertificate.value)
    {
    az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploywithcert --query "[?properties.provisioningState=='Failed'].properties.statusMessage.error" | Set-Content $deploymentLogPath
    }
    else{
    az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploy --query "[?properties.provisioningState=='Failed'].properties.statusMessage.error" | Set-Content $deploymentLogPath
    }

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
    if($parameters.useCertificate.Value){
    $failedResourcesList = az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploywithcert --query "[?properties.provisioningState=='Failed']" | ConvertFrom-Json
    }
    else{
    $failedResourcesList = az deployment operation group list --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --name azuredeploy --query "[?properties.provisioningState=='Failed']" | ConvertFrom-Json
    }
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
            hostingPlanSize                    = @{ value = $parameters.hostingPlanSize.Value }
            location                           = @{ value = $parameters.region.Value }
            gitRepoUrl                         = @{ value = $parameters.gitRepoUrl.Value }
            gitBranch                          = @{ value = $parameters.gitBranch.Value }
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

    $global:LASTEXITCODE = if ($deployState -eq 'Succeeded') { 0 } else { 1 }
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
        if($parameters.useCertificate.Value){
        $armDeploymentResult = az deployment group create --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --template-file 'azuredeploywithcert.json' --parameters "baseResourceName=$($parameters.baseResourceName.Value)" "authorClientId=$authorappId" "authorAppCertName=$($parameters.authorAppCertName.Value)" "graphAppId=$graphappid" "graphAppCertName=$($parameters.graphAppCertName.Value)" "userClientId=$userappId" "userAppCertName=$($parameters.userAppCertName.Value)" "senderUPNList=$($parameters.senderUPNList.Value)" "customDomainOption=$($parameters.customDomainOption.Value)" "appDisplayName=$($parameters.appDisplayName.Value)" "appDescription=$($parameters.appDescription.Value)" "appIconUrl=$($parameters.appIconUrl.Value)" "tenantId=$($parameters.tenantId.Value)" "hostingPlanSku=$($parameters.hostingPlanSku.Value)" "hostingPlanSize=$($parameters.hostingPlanSize.Value)" "location=$($parameters.region.Value)" "gitRepoUrl=$($parameters.gitRepoUrl.Value)" "gitBranch=$($parameters.gitBranch.Value)" "ProactivelyInstallUserApp=$($parameters.proactivelyInstallUserApp.Value)" "objectId=$($parameters.UserObjectId.Value)" "UserAppExternalId=$($parameters.userAppExternalId.Value)" "DefaultCulture=$($parameters.defaultCulture.Value)" "SupportedCultures=$($parameters.supportedCultures.Value)" "serviceBusWebAppRoleNameGuid=$($parameters.serviceBusWebAppRoleNameGuid.Value)" "serviceBusPrepFuncRoleNameGuid=$($parameters.serviceBusPrepFuncRoleNameGuid.Value)" "serviceBusSendFuncRoleNameGuid=$($parameters.serviceBusSendFuncRoleNameGuid.Value)" "serviceBusDataFuncRoleNameGuid=$($parameters.serviceBusDataFuncRoleNameGuid.Value)" "storageAccountWebAppRoleNameGuid=$($parameters.storageAccountWebAppRoleNameGuid.Value)" "storageAccountPrepFuncRoleNameGuid=$($parameters.storageAccountPrepFuncRoleNameGuid.Value)" "storageAccountDataFuncRoleNameGuid=$($parameters.storageAccountDataFuncRoleNameGuid.Value)" "TargetingEnabled=$($parameters.TargetingEnabled.Value)" "MasterAdminUpns=$($parameters.MasterAdminUpns.Value)"
        }
        else{
        $armDeploymentResult = InvokeArmDeploymentWithParamsFile $graphappid $authorappId $userappId $graphappsecret $authorsecret $usersecret
        }

        $deploymentExceptionMessage = "ERROR: ARM template deployment error."
        if ($LASTEXITCODE -ne 0) {
            # If ARM template deployment failed for any reason, then screen colors is becoming red
            [Console]::ResetColor()

            WriteI -message "Fetching deployment status to check if deployment really failed..."

            # Check if deployment reached Azure despite the connection error
            $deployName = if ($parameters.useCertificate.Value) { 'azuredeploywithcert' } else { 'azuredeploy' }
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
                    if($parameters.useCertificate.Value){
                        $armDeploymentResult = az deployment group create --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --template-file 'azuredeploywithcert.json' --parameters "baseResourceName=$($parameters.baseResourceName.Value)" "authorClientId=$authorappId" "authorAppCertName=$($parameters.authorAppCertName.Value)" "graphAppId=$graphappid" "graphAppCertName=$($parameters.graphAppCertName.Value)" "userClientId=$userappId" "userAppCertName=$($parameters.userAppCertName.Value)" "senderUPNList=$($parameters.senderUPNList.Value)" "customDomainOption=$($parameters.customDomainOption.Value)" "appDisplayName=$($parameters.appDisplayName.Value)" "appDescription=$($parameters.appDescription.Value)" "appIconUrl=$($parameters.appIconUrl.Value)" "tenantId=$($parameters.tenantId.Value)" "hostingPlanSku=$($parameters.hostingPlanSku.Value)" "hostingPlanSize=$($parameters.hostingPlanSize.Value)" "location=$($parameters.region.Value)" "gitRepoUrl=$($parameters.gitRepoUrl.Value)" "gitBranch=$($parameters.gitBranch.Value)" "ProactivelyInstallUserApp=$($parameters.proactivelyInstallUserApp.Value)" "objectId=$($parameters.UserObjectId.Value)" "UserAppExternalId=$($parameters.userAppExternalId.Value)" "DefaultCulture=$($parameters.defaultCulture.Value)" "SupportedCultures=$($parameters.supportedCultures.Value)"  "serviceBusWebAppRoleNameGuid=$($parameters.serviceBusWebAppRoleNameGuid.Value)" "serviceBusPrepFuncRoleNameGuid=$($parameters.serviceBusPrepFuncRoleNameGuid.Value)" "serviceBusSendFuncRoleNameGuid=$($parameters.serviceBusSendFuncRoleNameGuid.Value)" "serviceBusDataFuncRoleNameGuid=$($parameters.serviceBusDataFuncRoleNameGuid.Value)" "storageAccountWebAppRoleNameGuid=$($parameters.storageAccountWebAppRoleNameGuid.Value)" "storageAccountPrepFuncRoleNameGuid=$($parameters.storageAccountPrepFuncRoleNameGuid.Value)" "storageAccountDataFuncRoleNameGuid=$($parameters.storageAccountDataFuncRoleNameGuid.Value)" "TargetingEnabled=$($parameters.TargetingEnabled.Value)" "MasterAdminUpns=$($parameters.MasterAdminUpns.Value)"
                    }
                    else{
                       $armDeploymentResult = InvokeArmDeploymentWithParamsFile $graphappid $authorappId $userappId $graphappsecret $authorsecret $usersecret
                    }
                } else{
                    CollectARMDeploymentLogs
                    Throw $deploymentExceptionMessage
                }
            } else {
                # Deployment never reached Azure — retry once
                WriteW -message "Deployment not found in Azure (connection reset before submission). Retrying ARM deployment..."
                if ($parameters.useCertificate.Value) {
                    $armDeploymentResult = az deployment group create --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value --template-file 'azuredeploywithcert.json' --parameters "baseResourceName=$($parameters.baseResourceName.Value)" "authorClientId=$authorappId" "authorAppCertName=$($parameters.authorAppCertName.Value)" "graphAppId=$graphappid" "graphAppCertName=$($parameters.graphAppCertName.Value)" "userClientId=$userappId" "userAppCertName=$($parameters.userAppCertName.Value)" "senderUPNList=$($parameters.senderUPNList.Value)" "customDomainOption=$($parameters.customDomainOption.Value)" "appDisplayName=$($parameters.appDisplayName.Value)" "appDescription=$($parameters.appDescription.Value)" "appIconUrl=$($parameters.appIconUrl.Value)" "tenantId=$($parameters.tenantId.Value)" "hostingPlanSku=$($parameters.hostingPlanSku.Value)" "hostingPlanSize=$($parameters.hostingPlanSize.Value)" "location=$($parameters.region.Value)" "gitRepoUrl=$($parameters.gitRepoUrl.Value)" "gitBranch=$($parameters.gitBranch.Value)" "ProactivelyInstallUserApp=$($parameters.proactivelyInstallUserApp.Value)" "objectId=$($parameters.UserObjectId.Value)" "UserAppExternalId=$($parameters.userAppExternalId.Value)" "DefaultCulture=$($parameters.defaultCulture.Value)" "SupportedCultures=$($parameters.supportedCultures.Value)"  "serviceBusWebAppRoleNameGuid=$($parameters.serviceBusWebAppRoleNameGuid.Value)" "serviceBusPrepFuncRoleNameGuid=$($parameters.serviceBusPrepFuncRoleNameGuid.Value)" "serviceBusSendFuncRoleNameGuid=$($parameters.serviceBusSendFuncRoleNameGuid.Value)" "serviceBusDataFuncRoleNameGuid=$($parameters.serviceBusDataFuncRoleNameGuid.Value)" "storageAccountWebAppRoleNameGuid=$($parameters.storageAccountWebAppRoleNameGuid.Value)" "storageAccountPrepFuncRoleNameGuid=$($parameters.storageAccountPrepFuncRoleNameGuid.Value)" "storageAccountDataFuncRoleNameGuid=$($parameters.storageAccountDataFuncRoleNameGuid.Value)" "TargetingEnabled=$($parameters.TargetingEnabled.Value)" "MasterAdminUpns=$($parameters.MasterAdminUpns.Value)"
                } else {
                    $armDeploymentResult = InvokeArmDeploymentWithParamsFile $graphappid $authorappId $userappId $graphappsecret $authorsecret $usersecret
                }
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
        if($parameters.useCertificate.value)
        {
            $deploymentOutput = az deployment group show --name azuredeploywithcert --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value | ConvertFrom-Json
        }
        else
        {
            $deploymentOutput = az deployment group show --name azuredeploy --resource-group $parameters.resourceGroupName.Value --subscription $parameters.subscriptionId.Value | ConvertFrom-Json
        }

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

function CreateCertificateInKeyVault {
    Param(
        [Parameter(Mandatory= $true)] $certificateName,
        [Parameter(Mandatory= $true)] $keyVaultName,
        [Parameter(Mandatory= $true)] $domainName
    )

     #Get existing Azure Key Vault information
    $azKeyVault = Get-AzKeyVault -Name $keyVaultName -ErrorAction SilentlyContinue
    if ($null -eq $azKeyVault)
    {
        Write-Host "Didn't find Key Vault with name Azure:- $keyVaultName" -BackgroundColor DarkRed
        break
    }
    else
    {
        Write-Host "Found Key Vault Name:- $keyVaultName" -BackgroundColor DarkGreen
    }
     #Generate new Azure Key Vault Certificate
    Write-Host "Processing creation of Azure Key Vault Certificate" -ForegroundColor Yellow
    $certSubjectName = 'cn=' + $domainName
    $azKeyVaultCertPolicy = New-AzKeyVaultCertificatePolicy -SecretContentType "application/x-pkcs12" -SubjectName $certSubjectName -IssuerName "Self" -ValidityInMonths 24 -ReuseKeyOnRenewal
    $azKeyVaultCertStatus = Add-AzKeyVaultCertificate -VaultName $keyVaultName -Name $CertificateName -CertificatePolicy $azKeyVaultCertPolicy

    #Wait for certificate to generate
    $counter = 1
    While ($azKeyVaultCertStatus.Status -eq 'inProgress') {
        Start-Sleep -Milliseconds 50
        Write-Host "`r$counter% creation in progress" -NoNewline -ForegroundColor Yellow
        $azKeyVaultCertStatus = Get-AzKeyVaultCertificateOperation -VaultName $keyVaultName -Name $CertificateName
        $counter++
    }
    Write-Host "`r100% Completed. Checking status... " -ForegroundColor Yellow
    if ($azKeyVaultCertStatus.Status -ne 'completed') {
        Write-Host $($azKeyVaultCertStatus.StatusDetails) -ForegroundColor Magenta
    }
    else {
        Write-Host "Generated Key Vault Certificate successfully" -BackgroundColor DarkGreen
        Write-Output $azKeyVaultCertStatus
    }
}

function UpdateAadAppWithCertificate {
        Param(
        [Parameter(Mandatory = $true)] $appId,
        [Parameter(Mandatory = $true)] $keyVaultName,
        [Parameter(Mandatory = $true)] $certificateName
        )
        # Update AAD app with keyvault certificate
        az ad app credential reset --id $appId --keyvault $keyVaultName --cert $certificateName --append
}

function ImportKeyVaultCertificate{
    Param(
        [Parameter(Mandatory = $true)] $keyVaultName,
        [Parameter(Mandatory = $true)] $appName
        )
        Set-AzKeyVaultAccessPolicy -VaultName $keyVaultName -ServicePrincipalName abfa0a7c-a6b6-4736-8310-5855508787cd -PermissionsToSecrets get
        az webapp config ssl import --resource-group $parameters.resourceGroupName.Value --name $appName --subscription $parameters.subscriptionId.Value --key-vault $keyVaultName --key-vault-certificate-name $parameters.authorAppCertName.Value
        az webapp config ssl import --resource-group $parameters.resourceGroupName.Value --name $appName --subscription $parameters.subscriptionId.Value --key-vault $keyVaultName --key-vault-certificate-name $parameters.userAppCertName.Value
        az webapp config ssl import --resource-group $parameters.resourceGroupName.Value --name $appName --subscription $parameters.subscriptionId.Value --key-vault $keyVaultName --key-vault-certificate-name $parameters.graphAppCertName.Value
}

# Grant Admin consent
function GrantAdminConsent {
    Param(
        [Parameter(Mandatory = $true)] $graphAppId
        )

    $confirmationTitle = "Admin consent permissions is required for app registration using CLI"
    $confirmationQuestion = "Do you want to proceed?"
    $confirmationChoices = "&Yes", "&No" # 0 = Yes, 1 = No
    $consentErrorMessage = "Current user does not have the privilege to consent the below permissions on this app.
    * AppCatalog.Read.All(Delegated)
    * GroupMember.Read.All(Delegated)
    * GroupMember.Read.All(Application)
    * TeamsAppInstallation.ReadWriteForUser.All(Application)
    * User.Read.All(Delegated)
    * User.Read(Application)
    Please ask the tenant's global administrator to consent."

    $updateDecision = $Host.UI.PromptForChoice($confirmationTitle, $confirmationQuestion, $confirmationChoices, 1)
    if ($updateDecision -eq 0) {
        # Grant admin consent for app registration required permissions using CLI
        WriteI -message "Waiting for admin consent to finish..."
        az ad app permission admin-consent --id $graphAppId

        if ($LASTEXITCODE -ne 0) {
            WriteE -message $consentErrorMessage
            WriteW -message "`nPlease inform the global admin to consent the app permissions from this link`nhttps://login.microsoftonline.com/$($parameters.tenantId.value)/adminconsent?client_id=$graphAppId"
        } else {
            WriteS -message "Admin consent has been granted."
        }
    } else {
        WriteW -message "`nPlease inform the global admin to consent the app permissions from this link`nhttps://login.microsoftonline.com/$($parameters.tenantId.value)/adminconsent?client_id=$graphAppId"
    }
}

# Azure AD app update. Assigning Admin-consent,RedirectUris,IdentifierUris,Optionalclaim etc.
# Rewritten to use az CLI + Microsoft Graph REST API (az rest) instead of the retired AzureAD PS module.
function ADAppUpdate {
    Param(
        [Parameter(Mandatory = $true)] $appdomainName,
        [Parameter(Mandatory = $true)] $appId
    )
    $configAppId = $appId
    $azureDomainBase = $appdomainName
    $configAppUrl = "https://$azureDomainBase"
    $RedirectUris = ($configAppUrl + '/signin-simple-end')
    $IdentifierUris = "api://$azureDomainBase/$configAppId"

    # Grant Admin consent
    GrantAdminConsent $configAppId

    # set subscription
    az account set --subscription $parameters.subscriptionId.Value

    # Assigning graph permissions
    az ad app update --id $configAppId --required-resource-accesses './AadAppManifest.json'

    # Get the object ID (required for Graph API calls; appId and objectId are different)
    $appObjRaw = az ad app show --id $configAppId | ConvertFrom-Json
    $applicationObjectId = $appObjRaw.id

    # Fetch full application object via Graph API
    $app = az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" | ConvertFrom-Json

    # Do nothing if the app has already been configured
    if ($app.identifierUris.Count -gt 0) {
        WriteS -message "Graph application is already configured."
        return
    }
    WriteI -message "`nUpdating graph app..."

    # Disable then remove the default user_impersonation scope (Graph API requires two-step: disable then delete)
    $existingScopes = $app.api.oauth2PermissionScopes
    if ($existingScopes.Count -gt 0) {
        $disabledScopes = $existingScopes | ForEach-Object { $_.isEnabled = $false; $_ }
        $disablePatch = @{ api = @{ oauth2PermissionScopes = $disabledScopes } } | ConvertTo-Json -Depth 10 -Compress
        az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
            --body $disablePatch --headers "Content-Type=application/json" | Out-Null
        az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
            --body '{"api":{"oauth2PermissionScopes":[]}}' --headers "Content-Type=application/json" | Out-Null
    }

    # Set both identifier URIs required for Teams SSO:
    # - api://domain/appId  (used by Teams SSO v2 token requests via getAuthToken)
    # - api://domain        (Teams also validates this form; must be registered for AADSTS500011 to not occur)
    $identifierUriPatch = ('{"identifierUris":["api://' + $azureDomainBase + '","api://' + $azureDomainBase + '/' + $configAppId + '"]}')
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
        --body $identifierUriPatch --headers "Content-Type=application/json" | Out-Null
    WriteI -message "App identifier URIs set (api://domain and api://domain/appId)"

    az ad app update --id $configAppId --web-redirect-uris $RedirectUris
    WriteI -message "App reply-urls set"

    az ad app update --id $configAppId --optional-claims './AadOptionalClaims.json'
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
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
        --body $scopePatch --headers "Content-Type=application/json" | Out-Null
    WriteI -message "Scope access_as_user added."

    # Pre-authorize Teams mobile/desktop and web clients to access the API
    $preAuthApps = @(
        @{ appId = '1fec8e78-bce4-4aaf-ab1b-5451cc387264'; delegatedPermissionIds = @($scopeId) }
        @{ appId = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'; delegatedPermissionIds = @($scopeId) }
    )
    $preAuthPatch = @{ api = @{ preAuthorizedApplications = $preAuthApps } } | ConvertTo-Json -Depth 10 -Compress
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
        --body $preAuthPatch --headers "Content-Type=application/json" | Out-Null
    WriteI -message "Teams mobile/desktop and web clients applications pre-authorized."

    # Set requestedAccessTokenVersion = 2 so AAD issues v2 tokens for Teams SSO.
    # Required because Microsoft.Identity.Web disables inbound claim mapping and expects v2 JWT claim names.
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
        --body '{"api":{"requestedAccessTokenVersion":2}}' --headers "Content-Type=application/json" | Out-Null
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
    az ad app update --id $appId --display-name $newName
}

# Removing existing access of app (used on upgrades to reset the authors app before re-configuring).
# Rewritten to use az CLI + Microsoft Graph REST API instead of the retired AzureAD PS module.
function FormatAADApp {
    Param(
        [Parameter(Mandatory = $true)] $appId,
        [Parameter(Mandatory = $true)] $appName
    )

    # Get the object ID (required for Graph API calls)
    $appObjRaw = az ad app show --id $appId | ConvertFrom-Json
    $applicationObjectId = $appObjRaw.id

    # Fetch full application object via Graph API
    $app = az rest --method GET --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" | ConvertFrom-Json

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
        az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
            --body $disablePatch --headers "Content-Type=application/json" | Out-Null
        az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
            --body '{"api":{"oauth2PermissionScopes":[]}}' --headers "Content-Type=application/json" | Out-Null
    }

    # Clear implicit grant, redirect URIs, and identifier URIs in one PATCH
    $resetPatch = '{"web":{"implicitGrantSettings":{"enableIdTokenIssuance":false},"redirectUris":[]},"identifierUris":[]}'
    az rest --method PATCH --url "https://graph.microsoft.com/v1.0/applications/$applicationObjectId" `
        --body $resetPatch --headers "Content-Type=application/json" | Out-Null

    az ad app update --id $appId --optional-claims './AadOptionalClaims_Reset.json'
    az ad app update --id $appId --remove requiredResourceAccess
}
#update manifest file and create a .zip file.
function GenerateAppManifestPackage {
    Param(
        [Parameter(Mandatory = $true)] [ValidateSet('authors', 'users')] $manifestType,
        [Parameter(Mandatory = $true)] $appdomainName,
        [Parameter(Mandatory = $true)] $appId
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
        $mergeFields = @{
            '<<companyName>>'   = $parameters.companyName.Value
            '<<botId>>'         = $appId
            '<<appDomain>>'     = $azureDomainBase
            '<<websiteUrl>>'    = $parameters.websiteUrl.Value
            '<<privacyUrl>>'    = $parameters.privacyUrl.Value
            '<<termsOfUseUrl>>' = $parameters.termsOfUseUrl.Value
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

# Installing required modules
    WriteI -message "Checking if the required modules are installed..."
    $isAvailable = $true

    if ((Get-Module -ListAvailable -Name "WriteAscii")) {
        WriteI -message "WriteAscii module is available."
    } else {
        WriteW -message "WriteAscii module is missing."
        $isAvailable = $false
    }

    # Az module is only required when useCertificate=true (Key Vault certificate operations).
    if ((Get-Module -ListAvailable -Name "Az.*")) {
        WriteI -message "Az module is available."
    } else {
        WriteI -message "Az module is not installed. It is only required when useCertificate=true."
    }

    if (-not $isAvailable)
    {
        $confirmationTitle = WriteI -message "The script requires the following module to deploy: `n 1.WriteAscii module`nIf you proceed, the script will install the missing module."
        $confirmationQuestion = "Do you want to proceed?"
        $confirmationChoices = "&Yes", "&No" # 0 = Yes, 1 = No

        $updateDecision = $Host.UI.PromptForChoice($confirmationTitle, $confirmationQuestion, $confirmationChoices, 1)
            if ($updateDecision -eq 0) {
                if (-not (Get-Module -ListAvailable -Name "WriteAscii")) {
                    WriteI -message "Installing WriteAscii module..."
                    Install-Module WriteAscii -Scope CurrentUser -Force
                }
            } else {
                WriteE -message "You may install the modules manually by following the below link. Please re-run the script after the modules are installed. `nhttps://docs.microsoft.com/en-us/powershell/module/powershellget/install-module?view=powershell-7"
                EXIT
            }
    } else {
        WriteS -message "All the modules are available!"
    }

# Load Parameters from JSON meta-data file
    $parametersListContent = Get-Content '.\parameters.json' -ErrorAction Stop

# Validate all the parameters.
    WriteI -message "Validating all the parameters from parameters.json."
    $parameters = $parametersListContent | ConvertFrom-Json
    if (-not(ValidateParameters)) {
        WriteE -message "Invalid parameters found. Please update the parameters in the parameters.json with valid values and re-run the script."
        EXIT
    }

# Start Deployment.
    Write-Ascii -InputObject "Company Communicator v5.0" -ForegroundColor Magenta
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
		$userAppCred = CreateAzureADApp $usersApp
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
		$authorAppCred = CreateAzureADApp $authorsApp
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
    if($parameters.useCertificate.Value){
        $deploymentOutput = DeployARMTemplate $graphAppCred.appId $authorAppCred.appId $userAppCred.appId

        # Reading the deployment output.
        WriteI -message "Reading deployment outputs..."
        if(($null -eq $deploymentOutput) -or ( $null -eq $deploymentOutput.properties) -or ($null -eq $deploymentOutput.properties.Outputs) -or ($deploymentOutput.properties.Outputs.keyVaultName) -or ($deploymentOutput.properties.Outputs.keyVaultName.Value))
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
         $keyVaultName = $deploymentOutput.properties.Outputs.keyVaultName.Value
         $appdomainName = $deploymentOutput.properties.Outputs.appDomain.Value
        }

        CreateCertificateInKeyVault $parameters.authorAppCertName.value $keyVaultName $appdomainName
        CreateCertificateInKeyVault $parameters.userAppCertName.value $keyVaultName $appdomainName
        CreateCertificateInKeyVault $parameters.graphAppCertName.value $keyVaultName $appdomainName

        ImportKeyVaultCertificate $keyVaultName $parameters.BaseResourceName.Value
        ImportKeyVaultCertificate $keyVaultName "$($parameters.BaseResourceName.Value)-prep-function"
        ImportKeyVaultCertificate $keyVaultName "$($parameters.BaseResourceName.Value)-function"
        ImportKeyVaultCertificate $keyVaultName "$($parameters.BaseResourceName.Value)-data-function"

        UpdateAadAppWithCertificate $authorAppCred.appId $keyVaultName $parameters.authorAppCertName.value
        UpdateAadAppWithCertificate $userAppCred.appId $keyVaultName $parameters.userAppCertName.value
        UpdateAadAppWithCertificate $graphAppCred.appId $keyVaultName $parameters.graphAppCertName.value
    }
    else
    {
        $deploymentOutput = DeployARMTemplate $graphAppCred.appId $authorAppCred.appId $userAppCred.appId $graphAppCred.password $authorAppCred.password $userAppCred.password

        # Reading the deployment output.
        WriteI -message "Reading deployment outputs..."
        if(($null -eq $deploymentOutput) -or ( $null -eq $deploymentOutput.properties) -or ($null -eq $deploymentOutput.properties.Outputs) -or ($deploymentOutput.properties.Outputs.keyVaultName) -or ($deploymentOutput.properties.Outputs.keyVaultName.Value))
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
    }
    if ($null -eq $deploymentOutput) {
        WriteE -message "Encountered an error during ARM template deployment. Exiting..."
        logout
        Exit
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
    GenerateAppManifestPackage 'authors' $appdomainName $authorAppCred.appId
    GenerateAppManifestPackage 'users' $appdomainName $userAppCred.appId

# Open manifest folder
    Invoke-Item ..\Manifest\

# Deployment completed.
    Write-Ascii -InputObject "DEPLOYMENT COMPLETED." -ForegroundColor Green
