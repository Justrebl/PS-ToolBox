#=======================================================================================
#region Invoke AAD and Query Microsoft Graph
#=======================================================================================
<#
Disclaimer: The sample scripts are not supported under any Microsoft standard support program or service. 
The sample scripts are provided AS IS without warranty of any kind. Microsoft further disclaims all implied 
warranties including, without limitation, any implied warranties of merchantability or of fitness for a 
particular purpose. The entire risk arising out of the use or performance of the sample scripts and 
documentation remains with you. In no event shall Microsoft, its authors, or anyone else involved in the 
creation, production, or delivery of the scripts be liable for any damages whatsoever (including, without 
limitation, damages for loss of business profits, business interruption, loss of business information, or 
other pecuniary loss) arising out of the use of or inability to use the sample scripts or documentation, 
even if Microsoft has been advised of the possibility of such damages. 
#>
#endregion
#=======================================================================================

#=======================================================================================
#region Script configuration - Update Mandatory
#=======================================================================================

$resourceAppIdURI = "https://graph.microsoft.com"
$ClientID         = "xxx"                                                            #Application ID you created in AAD
$TenantName       = "xxx.onmicrosoft.com"                                            #Your Tenant Name
$CredPrompt       = "Auto"                                                           #Prompt for AAD Login - Auto, Always, Never, RefreshSession
$redirectUri      = "https://login.microsoftonline.com/common/oauth2/nativeclient"   #Your Application's Redirect URI if needed -> Here is default value
$ApiUri           = "https://graph.microsoft.com/v1.0/planner/plans"                 #The Graph Api URL to invoque REST Query on to create Plan
$Method           = "POST"                                                           #POST Operation to create/override the default Planner

#endregion
#=======================================================================================

#=======================================================================================
#region AAD Auth and Invoke REST API Functions - Do Not Modify
#=======================================================================================
Function Get-AccessToken ($TenantName, $ClientID, $redirectUri, $resourceAppIdURI, $CredPrompt){
    Write-Host "Checking for AzureAD module..."
    if (!$CredPrompt){$CredPrompt = 'Auto'}
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
    if ($AadModule -eq $null) {$AadModule = Get-Module -Name "AzureADPreview" -ListAvailable}
    if ($AadModule -eq $null) {write-host "AzureAD Powershell module is not installed. The module can be installed by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt. Stopping." -f Yellow;exit}
    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
        $aadModule      = $AadModule | ? { $_.version -eq $Latest_Version.version
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
        }
    } else {
        $adal           = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms      = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
    $authority          = "https://login.microsoftonline.com/$TenantName"
    $authContext        = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters"    -ArgumentList $CredPrompt
    $authResult         = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters).Result
    return $authResult
    }

Function Invoke-MSGraphQuery($AccessToken, $Uri, $Method, $Body){
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -CurrentOperation "Invoking MS Graph API"
    $Header = @{
        'Content-Type'  = 'application\json'
        'Authorization' = $AccessToken.CreateAuthorizationHeader()
        }
    $QueryResults = @()
    if($Method -eq "Get"){
        do{
            $Results =  Invoke-RestMethod -Headers $Header -Uri $Uri -UseBasicParsing -Method $Method -ContentType "application/json"
            if ($Results.value -ne $null){$QueryResults += $Results.value}
            else{$QueryResults += $Results}
            write-host "Method: $Method | URI $Uri | Found:" ($QueryResults).Count
            $uri = $Results.'@odata.nextlink'
            }until ($uri -eq $null)
        }
    else{
        $Results =  Invoke-RestMethod -Headers $Header -Uri $Uri -Method $Method -ContentType "application/json" -Body $Body
        write-host "Method: $Method | URI $Uri | Executing"
        }
    Write-Progress -Id 1 -Activity "Executing query: $Uri" -Completed
    Return $QueryResults
    }
#endregion
#=======================================================================================

#=======================================================================================
#region Main Program - No Update needed
#=======================================================================================

#Connection to AAD with Credential Prompt (needed to contact Graph API)
$AccessToken      = Get-AccessToken -TenantName $TenantName -ClientID $ClientID -redirectUri $redirectUri -resourceAppIdURI $resourceAppIdURI -CredPrompt $CredPrompt

#Will prompt once more to get the credentials to connect to the O365Tenant with the user credential
#Need to be tuned to use the AAD AccessToken from the Cmd line above
Connect-MicrosoftTeams 

Write-Host "What's the name you want to give to your new Teams/group ?"
$TeamsName = Read-Host "Teams Name"

#Will create a New Team with specified settings 
$newTeam = New-Team -DisplayName $TeamsName -Description "This is a test Team created with its O365 Group" -Visibility "Private"

#Display the newly created Team settings
$newTeam

#Set O365GroupId to send to Graph API
$groupId = $newTeam | Select -ExpandProperty GroupId

#Set the mandatory parameters of the Rest Body MSGraph API Query 
$JSON = @"
    {
        owner : "$groupId",
        title : "$TeamsName"
    }
"@

#Contact Graph API Url (Planner/Plans) to create/override the Default Plan created with O365 group 
Invoke-MSGraphQuery -AccessToken $AccessToken -Uri $ApiUri -Method $Method -Body $JSON

#endregion
#=======================================================================================