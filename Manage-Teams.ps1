<#
.DESCRIPTION

###############Disclaimer#####################################################
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###############Disclaimer#####################################################

Script to generate Teams reports. 

What this script does: 
    0) Check Script Pre-requisites
    1) Connect to O365
    2) Get Teams
        Properties: "GroupId","GroupName","TeamsEnabled","Provider","ManagedBy","WhenCreated","PrimarySMTPAddress","GroupGuestSetting","GroupAccessType","GroupClassification","GroupMemberCount","GroupExtMemberCount","SPOSiteUrl","SPOStorageUsed","SPOtorageQuota","SPOSharingSetting"
    3) Get Teams Membership
        Properties: "GroupID","GroupName","TeamsEnabled","Member","Name","RecipientType","Membership"
    4) Get Teams That Are Not Active
        Properties: "GroupID","Name","TeamsEnabled","PrimarySMTPAddress","MailboxStatus","LastConversationDate","NumOfConversations","SPOStatus","LastContentModified","StorageUsageCurrent"
    5) Get Users That Are Allowed To Create Teams
        Properties: "ObjectID","DisplayName","UserPrincipalName","UserType" 
    6) Get Teams Tenant Settings
        Settings captured: Azure AD Group Settings, Who's Allowed to Create Teams, Guest Access, Expiration Policy
    7) Get Groups/Teams Without Owner(s)
        Properties: "GroupID","GroupName","HasOwners","ManagedBy"
    8) Get All Above Reports
    9) Get Teams By User
        Properties: "User","GroupId","GroupName","TeamsEnabled","IsOwner"
    10) Exit Script

REQUIREMENTS: 
    -Powershell v3+ 
    -Azureadpreview Module - https://www.powershellgallery.com/packages/AzureADPreview/2.0.0.17
    -Teams module - https://www.powershellgallery.com/packages/MicrosoftTeams/0.9.0
    -SPO module - https://www.microsoft.com/en-us/download/details.aspx?id=35588 
    -SFBO module - https://www.microsoft.com/en-us/download/details.aspx?id=39366
    -EXO/SCC Click Once App for MFA - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell?view=exchange-ps
    -PNP Module - https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps

VERSION:
    06112018: Update MFA module
    06072018: Update Get-Teams Logic to use Microsoft Graph
    05302018: Added MFA Logon Capability, Teams by User, Teams without Owners
    02072018: v1

AUTHOR(S): 
    Alejandro Lopez - Alejanl@Microsoft.com

.EXAMPLE
#Run the script with no switches and it will provide you a menu of what reports to run. 
.\Manage-Teams.ps1

#>

#region Functions

#Check installed modules
Function Check-Modules{
    Write-LogEntry -LogName:$Log -LogEntryText "Pre-Flight Check" -ForegroundColor White        
    $needPowershellUpdated = $false
    $needAADPModuleInstall = $false
    $needTeamsModuleInstall = $false
    $needSPOModule = $false
    $needMSOAssistant = $false
    $needSkypeModule = $false
    $needEXOMFAModule = $false
    $needPNPModule = $false
    ""
    $acctHasMFA = Read-Host "Is your account enabled for MFA? (Y/N)"
    ""

    If($host.version.major -lt 3){
        Write-LogEntry -LogName:$Log -LogEntryText "Powershell V3+: Missing" -ForegroundColor Yellow        
        $needPowershellUpdated = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "Powershell V3+" -ForegroundColor Green        
    }

    $aadmodule = get-module -listavailable azureadpreview
    If(!$aadmodule){
        #Need at least AzureADPreview 2.0.0.137 for Get-AzureADMSGroupLifecyclePolicy
        Write-LogEntry -LogName:$Log -LogEntryText "AzureADPreview Missing" -ForegroundColor Yellow 
        $needAADPModuleInstall = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "AzureADPreview" -ForegroundColor Green
    }

    If(!(get-module -listavailable MicrosoftTeams)){
        Write-LogEntry -LogName:$Log -LogEntryText "MicrosoftTeams Module: Missing" -ForegroundColor Yellow
        $needTeamsModuleInstall = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "MicrosoftTeams Module" -ForegroundColor Green
    }

    If(!(get-module -listavailable microsoft.online.sharepoint.powershell)){
        Write-LogEntry -LogName:$Log -LogEntryText "SharePoint Online Module: Missing" -ForegroundColor Yellow
        $needSPOModule = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "SharePoint Online Module" -ForegroundColor Green
    }

    $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
    If ($CheckForSignInAssistant -eq $false) {
        Write-LogEntry -LogName:$Log -LogEntryText "Microsoft Online Services Sign-in Assistant: Missing" -ForegroundColor Yellow
        $needMSOAssistant = $true
    }
    else{
        Write-LogEntry -LogName:$Log -LogEntryText "Microsoft Online Services Sign-in Assistant" -ForegroundColor Green
    }

    If(!(get-module -listavailable SkypeOnlineConnector)){
        Write-LogEntry -LogName:$Log -LogEntryText "Skype for Business Online Module: Missing" -ForegroundColor Yellow
        $needSkypeModule = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "Skype for Business Online Module" -ForegroundColor Green
    }

    #Check for EXO/SCC Click Once Application required for MFA
    If($acctHasMFA -eq "Y" -or $acctHasMFA -eq "y" ){
        if ((Test-ClickOnce -ApplicationName "Microsoft Exchange Online Powershell Module" ) -eq $false)  {
            Write-LogEntry -LogName:$Log -LogEntryText "Exchange Online MFA Module: Missing" -ForegroundColor Yellow
            $needEXOMFAModule = $true
        }
        Else{
            Write-LogEntry -LogName:$Log -LogEntryText "Exchange Online MFA Module" -ForegroundColor Green
        }
        
    }

    #Check for PNP Module
    If(!(get-module -listavailable sharepointpnppowershellonline)){
        Write-LogEntry -LogName:$Log -LogEntryText "PNP Module: Missing" -ForegroundColor Yellow
        $needPNPModule = $true
    }
    Else{
        Write-LogEntry -LogName:$Log -LogEntryText "PNP Module" -ForegroundColor Green
    }

    ""
    ""

    If($needEXOMFAModule -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "Please install the EXO/SCC Click Once App and re-run pre-flight check: https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application" -ForegroundColor Yellow
    }
    ElseIf($acctHasMFA -eq "Y" -or $acctHasMFA -eq "y" ){
        Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName| ?{$_ -notmatch "_none_"} | select -First 1)
    }

    If($needAADPModuleInstall -eq $true){
        $check = Read-Host "Would you like to install the required AzureADPreview module? (Y/N)"
        If($check -eq "Y" -or $check -eq "y"){
            try{
                Write-LogEntry -LogName:$Log -LogEntryText "Installing latest version of AzureADPreview Module..." -ForegroundColor White
                Remove-module AzureADPreview -ErrorAction SilentlyContinue
                Install-Module AzureADPreview -Force
                Write-LogEntry -LogName:$Log -LogEntryText "Successfully installed AzureADPreview Module." -ForegroundColor Green
            }
            catch{
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to install the AzureADPreview Module. Please install manually from here: https://www.powershellgallery.com/packages/AzureADPreview/" -ForegroundColor Yellow
                Write-LogEntry -LogName:$Log -LogEntryText "$_" -ForegroundColor Red
            }
        }
        Else{
             Write-LogEntry -LogName:$Log -LogEntryText "Please install AzureADPreview to move forward: https://www.powershellgallery.com/packages/AzureADPreview/" -ForegroundColor Yellow
        }
    }
    Else{
        Import-module -Name AzureADPreview
    }

    If($needTeamsModuleInstall -eq $true){
        $check = Read-Host "Would you like to install the Microsoft Teams module? (Y/N)"
        If($check -eq "Y" -or $check -eq "y"){
            try{
                Install-Module MicrosoftTeams
                Write-LogEntry -LogName:$Log -LogEntryText "Successfully installed Microsoft Teams Module." -ForegroundColor Green        
            }
            catch{
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to install the Microsoft Teams Module. Please install from here: https://www.powershellgallery.com/packages/MicrosoftTeams/" -ForegroundColor Yellow
                Write-LogEntry -LogName:$Log -LogEntryText "$_" -ForegroundColor Red
            }     
        }
        Else{
             Write-LogEntry -LogName:$Log -LogEntryText "Please install the Microsoft Teams module to move forward: https://www.powershellgallery.com/packages/MicrosoftTeams/" -ForegroundColor Yellow
        }
    }
    Else{
        Import-module -Name MicrosoftTeams
    }
    
    If($needSPOModule -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "SharePoint Online module missing - https://www.microsoft.com/en-us/download/details.aspx?id=35588. Please install and re-run." -ForegroundColor Yellow
    }
    Else{
        Import-Module "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.dll" -DisableNameChecking
    }

    If($needMSOAssistant -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "Microsoft Online Services Sign-in Assistant missing - https://go.microsoft.com/fwlink/p/?LinkId=286152. Please install and re-run." -ForegroundColor Yellow
    }
    
    If($needSkypeModule -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "Skype for Business online module missing - https://www.microsoft.com/en-us/download/details.aspx?id=39366. Please install, restart powershell session, and re-run." -ForegroundColor Yellow
    }
    else{
        Import-module -Name SkypeOnlineConnector
    }

    If($needPNPModule -eq $true){
        $check = Read-Host "Would you like to install the required PNP module? (Y/N)"
        If($check -eq "Y" -or $check -eq "y"){
            try{
                Write-LogEntry -LogName:$Log -LogEntryText "Installing latest version of PNP Module..." -ForegroundColor White
                Install-Module SharePointPnPPowerShellOnline -SkipPublisherCheck -AllowClobber
                Write-LogEntry -LogName:$Log -LogEntryText "Successfully installed PNP Module." -ForegroundColor Green
            }
            catch{
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to install the PNP Module. Please install manually from here: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps" -ForegroundColor Yellow
                Write-LogEntry -LogName:$Log -LogEntryText "$_" -ForegroundColor Red
            }
        }
        Else{
             Write-LogEntry -LogName:$Log -LogEntryText "Please install AzureADPreview to move forward: https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps" -ForegroundColor Yellow
        }
    }
    Else{
        Import-module -Name SharePointPnPPowerShellOnline -DisableNameChecking -ErrorAction SilentlyContinue
    }

    Write-LogEntry -LogName:$Log -LogEntryText "Pre-Flight Done" -ForegroundColor Green
}

#Check for EXO & SCC click once application = Get-ClickOnce and Test-ClickOnce
#https://www.powershellgallery.com/packages/Load-ExchangeMFA/1.1/Content/Load-ExchangeMFA.ps1
function Get-ClickOnce {
    [CmdletBinding()]  
    Param(
        $ApplicationName = "Microsoft Exchange Online Powershell Module"
    )
        $InstalledApplicationNotMSI = Get-ChildItem HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall | foreach-object {Get-ItemProperty $_.PsPath}
        return $InstalledApplicationNotMSI | ? { $_.displayname -match $ApplicationName } | Select-Object -First 1
    }
    
    
Function Test-ClickOnce {
[CmdletBinding()] 
Param(
    $ApplicationName = "Microsoft Exchange Online Powershell Module"
)
    return ( (Get-ClickOnce -ApplicationName $ApplicationName) -ne $null) 
}

#Logging Function
Function Write-LogEntry {
    param(
        [string] $LogName ,
        [string] $LogEntryText,
        [string] $ForegroundColor
    )
    if ($LogName -NotLike $Null) {
        # log the date and time in the text file along with the data passed
        "$([DateTime]::Now.ToShortDateString()) $([DateTime]::Now.ToShortTimeString()) : $LogEntryText" | Out-File -FilePath $LogName -append;
        if ($ForeGroundColor -NotLike $null) {
            # for testing i pass the ForegroundColor parameter to act as a switch to also write to the shell console
            write-host $LogEntryText -ForegroundColor $ForeGroundColor
        }
    }
}

#Function to connect to O365
Function Logon-O365MFA {
    #You can't use the Exchange Online Remote PowerShell Module to connect to Exchange Online PowerShell and Security & Compliance Center PowerShell in the same session (window). You need to use separate sessions of the Exchange Online Remote PowerShell Module. 
    #https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell?view=exchange-ps

    Write-LogEntry -LogName:$Log -LogEntryText "For MFA, we need to prompt for credentials for each service. Multiple prompts are expected." -ForegroundColor Yellow
    start-sleep 2
    $gotError = $false

    #SPO
    try{$testSPO = get-spotenant -erroraction silentlycontinue}
    catch{}
    If($testSPO -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
    }
    Else{
        try{
            Import-Module "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.dll" -DisableNameChecking
            $domainHost = Read-Host "Enter tenant name, such as contoso for contoso.onmicrosoft.com"
            If($domainHost -like "*.onmicrosoft.com"){
                $split = $domainHost.split(".")
                $domainHost = $split[0]
            } 
            $spoAdminUrl = "https://$domainHost-admin.sharepoint.com"
            Connect-SPOService -Url $spoAdminUrl 
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to SharePoint Online: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }

    #PNP
    try{$testPNP = Get-PnPAccessToken}
    catch{}
    If($testPNP -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to PNP Online" -ForegroundColor Green
        $Global:authToken = Get-PNPAccessToken
        $Global:authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authToken               
        }
        write-host "authtoken: $authtoken"
        write-host "authHeader: $authHeader"
    }
    Else{
        try{
            Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
            Connect-PnPOnline -Scopes "Group.ReadWrite.All"
            $Global:authToken = Get-PNPAccessToken
            $Global:authHeader = @{
                'Content-Type'='application/json'
                'Authorization'="Bearer " + $authToken               
            }
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to PNP Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to PNP Online: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }

    #TEAMS
    try{
        $testTeams = get-team
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green
    } 
    #catch [Microsoft.Open.Teams.CommonLibrary.AadNeedAuthenticationException]{
    catch{
        try{
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams 
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Microsoft Teams: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }

    #EXO
    try{$testEXO = Get-OrganizationConfig | ?{$_.identity -like "*.onmicrosoft.com"}}
    catch{}
    If($testEXO -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
    }
    Else{
        try{
            #Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse).FullName|?{$_ -notmatch "_none_"}|select -First 1)
            Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName| ?{$_ -notmatch "_none_"} | select -First 1)
            $EXOSession = New-ExoPSSession
            Import-PSSession $EXOSession | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Exchange Online: $_" -ForegroundColor Red
            $gotError = $true
        }
    }
    
    #Security and Compliance Center
    #For MFA - it's not possible to connect to EXO Powershell and Security & Compliance center in the same session
    #https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell?view=exchange-ps
    #https://techcommunity.microsoft.com/t5/Windows-PowerShell/Can-I-Connect-to-O365-Security-amp-Compliance-center-via/td-p/68898
    #Potential workaround: https://gallery.technet.microsoft.com/Office-365-Connection-47e03052
    <#
    try{$testSCC = Get-DlpSensitiveInformationType}
    catch{}
    If($testSCC -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Security and Compliance Center" -ForegroundColor Green
    }
    Else{
        try{
            Connect-IPPSSession
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Security and Compliance Center" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Security and Compliance Center: $_" -ForegroundColor Red
            $gotError = $true
        }
    }
    #>

    #Azure AD 
    try{$testAAD = Get-AzureADCurrentSessionInfo -erroraction silentlycontinue}
    catch{}
    If($testAAD -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
    }
    Else{
        try{
            #Need AzureADPreview 2.0.0.137 for Get-AzureADMSGroupLifecyclePolicy 
            #If AzureAD module (Not AzureADPreview) is also available, then the AzureADPreview module is not loaded
            $checkAzureADModule = get-module -name "AzureAD"
            If($checkforAzureADModule -ne $null){
                Remove-Module -Name "AzureAD"
            }
            Import-module -Name AzureADPreview
            Connect-AzureAD  | out-null #https://github.com/itnetxbe/Feedback/issues/15 - login issues sporadically
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Azure AD: $_" -ForegroundColor Red
            $gotError = $true
        }
    }

    #SFBO
    try{$CSTenant = (Get-CsTenant).DisplayName}
    catch{}
    If ($CSTenant -ne $null){    
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Skype for Business Online" -ForegroundColor Green
    }
    Else {
        try{
            Import-Module SkypeOnlineConnector
            $sfbSession = New-CsOnlineSession 
            Import-PSSession $sfbSession | out-null    
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Skype for Business Online: $_" -ForegroundColor Red
            $gotError = $true
        }
    }
    
    If($gotError -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "There was an error connecting to one of the services. Re-run Step 1 and try again." -ForegroundColor Yellow
    }

}

Function Logon-O365{
    $gotError = $false

    #SPO
    try{$testSPO = get-spotenant -erroraction silentlycontinue}
    catch{}
    If($testSPO -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
    }
    Else{
        try{
            Import-Module "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.dll" -DisableNameChecking
            $domainHost = Read-Host "Enter tenant name, such as contoso for contoso.onmicrosoft.com"
            If($domainHost -like "*.onmicrosoft.com"){
                $split = $domainHost.split(".")
                $domainHost = $split[0]
            }  
            $spoAdminUrl = "https://$domainHost-admin.sharepoint.com"
            Connect-SPOService -Url $spoAdminUrl #not passing credentials due to known issue with connect-sposervice
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to SharePoint Online: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }

    #PNP
    try{$testPNP = Get-PnPAccessToken}
    catch{}
    If($testPNP){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to PNP Online" -ForegroundColor Green
        $Global:authToken = Get-PNPAccessToken
        $Global:authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authToken               
        }
    }
    Else{
        try{
            Import-Module SharePointPnPPowerShellOnline -DisableNameChecking
            Connect-PnPOnline -Scopes "Group.ReadWrite.All"
            $Global:authToken = Get-PNPAccessToken
            $Global:authHeader = @{
                'Content-Type'='application/json'
                'Authorization'="Bearer " + $authToken               
            }
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to PNP Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to PNP Online: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }

    #EXO
    $session = Get-PSSession | where {($_.ComputerName -eq "outlook.office365.com") -and ($_.State -eq "Opened")}
    If ($session -ne $null) {
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
    }
    Else{
        try{
            If ($Credential -eq $null){
                $Credential = get-credential -credential $null 
            }
            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue 
            #Import-PSSession -session $exchangeSession -DisableNameChecking -AllowClobber | out-null
            Import-Module (Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Exchange Online: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }    

    #TEAMS
    try{
        $testTeams = get-team
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green
    } 
    #catch [Microsoft.Open.Teams.CommonLibrary.AadNeedAuthenticationException]{
    catch{
        try{
            Import-Module MicrosoftTeams
            If ($Credential -eq $null){
                $Credential = get-credential -credential $null
            }
            $connectTeams = Connect-MicrosoftTeams -Credential $Credential 
            If($connectTeams){Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green}
            Else{Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Teams. Hit Ctrl+C, Run Connect-MicrosoftTeams, then re-run script." -ForegroundColor Green}
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Microsoft Teams: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }


    #SFBO
    try{$CSTenant = (Get-CsTenant).DisplayName}
    catch{}
    If ($CSTenant -ne $null){    
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Skype for Business Online" -ForegroundColor Green
    }
    Else {
        try{
            Import-Module SkypeOnlineConnector
            $sfbSession = New-CsOnlineSession #not passing the $credential beause of AADSTS90014 error
            Import-PSSession $sfbSession | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Skype for Business Online" -ForegroundColor Green   
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Skype for Business Online: $_" -ForegroundColor Red
            $gotError = $true
        }
    }

    #Compliance Center
    #Getting Access Denied errors. Commenting out for now
    <#
    If (Get-PSSession | where {($_.ComputerName -eq "ps.compliance.protection.outlook.com") -and ($_.State -eq "Opened")}) {
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Compliance Center" -ForegroundColor Green    
    }
    Else{
        try{
            If ($Credential -eq $null){
                $Credential = get-credential -credential $null
            }
            $ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
            Import-PSSession -session $ccSession -Prefix cc -DisableNameChecking -AllowClobber | out-null
            #Import-Module (Import-PSSession $ccSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Compliance Center" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Compliance Center: $_" -ForegroundColor Red
            $gotError = $true 
        }
    }
    #>

    #Azure AD 
    try{$testAAD = Get-AzureADCurrentSessionInfo -erroraction silentlycontinue}
    catch{}
    If($testAAD -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
    }
    Else{
        try{
            #Need AzureADPreview 2.0.0.137 for Get-AzureADMSGroupLifecyclePolicy 
            #If AzureAD module (Not AzureADPreview) is also available, then the AzureADPreview module is not loaded
            $checkAzureADModule = get-module -name "AzureAD"
            If($checkforAzureADModule -ne $null){
                Remove-Module -Name "AzureAD"
            }
            Import-module -Name AzureADPreview
            Connect-AzureAD | out-null #prompt for credential due to issue: AADSTS90014
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Azure AD: $_" -ForegroundColor Red
            $gotError = $true
        }
    }
    ""
    If($gotError -eq $true){
        Write-LogEntry -LogName:$Log -LogEntryText "There was an error connecting to one of the services. Re-run Step 1 and try again." -ForegroundColor Yellow
    }
}


# Gets AllowBlockedList from SPO
function GetSPOPolicy {
    try
    {
        $SPOTenantSettings = Get-SPOTenant
    }
    catch [System.InvalidOperationException]
    {
        Write-Error "You must call Connect-SPOService cmdlet before using this parameter."
        Exit;
    }

    # Return JSON for Allow\Block domain list in SPO
    switch($SPOTenantSettings.SharingDomainRestrictionMode)
    {
        "AllowList"
        {
            #Write-Host "`nSPO Allowed DomainList:" $SPOTenantSettings.SharingAllowedDomainList
            $AllowDomainsList = $SPOTenantSettings.SharingAllowedDomainList.Split(' ')
            return  GetJSONForAllowBlockDomainPolicy -AllowDomains $AllowDomainsList
            break;
        }
        "BlockList"
        {
            #Write-Host "`nSPO Blocked DomainList:" $SPOTenantSettings.SharingBlockedDomainList
            $BlockDomainsList = $SPOTenantSettings.SharingBlockedDomainList.Split(' ')
            return GetJSONForAllowBlockDomainPolicy -BlockedDomains $BlockDomainsList
            break;
        }
        "None"
        {
            #Write-Error "There is no AllowBlockDomainList policy set for this SPO tenant."
            return $null
        }
    }
}

# Converts Json to Object since ConvertFrom-Json does not support the depth parameter.
function GetObjectFromJson([string] $JsonString) {
    ConvertFrom-Json -InputObject $JsonString |
        ForEach-Object {
            foreach ($property in ($_ | Get-Member -MemberType NoteProperty)) 
                {
                    $_.$($property.Name) | Add-Member -MemberType NoteProperty -Name 'Name' -Value $property.Name -PassThru
                }
        }
}

# Gets Json for the policy with given Allowed and Blocked Domain List
function GetJSONForAllowBlockDomainPolicy([string[]] $AllowDomains = @(), [string[]] $BlockedDomains = @()){
    # Remove any duplicate domains from Allowed or Blocked domains specified.
    $AllowDomains = $AllowDomains | select -uniq
    $BlockedDomains = $BlockedDomains | select -uniq

    return @{B2BManagementPolicy=@{InvitationsAllowedAndBlockedDomainsPolicy=@{AllowedDomains=@($AllowDomains); BlockedDomains=@($BlockedDomains)}}} | ConvertTo-Json -Depth 3 -Compress
}

# Get existing B2B management policy
function GetExistingPolicy{
    $currentpolicy = Get-AzureADPolicy | ?{$_.Type -eq 'B2BManagementPolicy'} | select -First 1

    return $currentpolicy;
}

# Gets AllowDomainList from the existing policy
function GetExistingAllowedDomainList(){
    $policy = GetExistingPolicy

    if($policy -ne $null)
    {
        $policyObject = GetObjectFromJson $policy.Definition[0];

        if($policyObject.InvitationsAllowedAndBlockedDomainsPolicy -ne $null -and $policyObject.InvitationsAllowedAndBlockedDomainsPolicy.AllowedDomains -ne $null)
        {
            return $policyObject.InvitationsAllowedAndBlockedDomainsPolicy.AllowedDomains;
        }
    }

    return $null
}

# Gets BlockDomainList from the existing policy
function GetExistingBlockedDomainList(){
    $policy = GetExistingPolicy

    if($policy -ne $null)
    {
        $policyObject = GetObjectFromJson $policy.Definition[0];

        if($policyObject.InvitationsAllowedAndBlockedDomainsPolicy -ne $null -and $policyObject.InvitationsAllowedAndBlockedDomainsPolicy.BlockedDomains -ne $null)
        {
            return $policyObject.InvitationsAllowedAndBlockedDomainsPolicy.BlockedDomains;
        }
    }

    return $null
}

#Get list of Teams
Function Get-Teams{
    #Using David Whitney (dawhitne@microsoft.com) method of identifying Teams based on http status code
    param (
        [switch]$ExportToCSV
    )

    if (-not (Get-PSSession | where {($_.ComputerName -eq "outlook.office365.com") -and ($_.State -eq "Opened")})) {
        Write-LogEntry -LogName:$Log -LogEntryText "You must connect to Exchange Online Remote PowerShell..." -ForegroundColor Yellow
        break
    }
    $testSPO = get-spotenant
    if (!$testSPO){
        Write-LogEntry -LogName:$Log -LogEntryText "You must connect to SharePoint Online PowerShell..." -ForegroundColor Yellow
        break
    }

    try{$testTeams = get-team} #using try catch since the admin may not be a member of any Teams, so not a valid way to test connection
    #catch [Microsoft.Open.Teams.CommonLibrary.AadNeedAuthenticationException]{
    catch{
        Write-LogEntry -LogName:$Log -LogEntryText "You must connect to Microsoft Teams PowerShell..." -ForegroundColor Yellow
        break
    }

    Write-LogEntry -LogName:$Log -LogEntryText "Getting Teams report..." -ForegroundColor Yellow
    $o365groups = Get-UnifiedGroup -ResultSize Unlimited | where-object{$_.sharepointsiteurl -ne $null}

    $global:ListOfGroupsTeams = New-Object System.Collections.ArrayList
    #$storageHash = @{"test"=@{"storage"="1";"quota"="2"}}

    $count = $o365groups.count
    $i = 0
    Write-LogEntry -LogName:$Log -LogEntryText "Found $count O365 Groups. Checking how many are also Teams..." -ForegroundColor White
    foreach ($o365group in $o365groups) {
        Write-Progress -Activity "Getting Teams Info..." -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        $spoSite = Get-SPOSite -Identity $o365group.SharePointSiteUrl
        $spoStorageQuota =  "$(($spoSite).StorageQuota)" + "MB"
        $spoStorageUsed = "$(($spoSite).StorageUsageCurrent)" + "MB"
        $spoSharingSetting = ($spoSite).SharingCapability
        
        #Microsoft Graph Query for Teams: https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_endpoints

        try {
            $GroupsUri = "https://graph.microsoft.com/beta/groups/$($o365group.ExternalDirectoryObjectId)/endpoints"
            $groupDetails = (Invoke-RestMethod -Uri $GroupsUri -Headers $authHeader -Method Get).value

            If($groupDetails){
                If($groupDetails.providerName -eq "Microsoft Teams"){
                    $GroupTeam = [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; 
                        GroupName = $o365group.DisplayName;
                        TeamsEnabled = $true;
                        Provider = $groupDetails.providerName;
                        ManagedBy = $o365group.ManagedBy 
                        WhenCreated = $o365group.WhenCreated;
                        PrimarySMTPAddress = $o365group.PrimarySMTPAddress;
                        GroupGuestSetting = $o365group.AllowAddGuests;
                        GroupAccessType = $o365group.AccessType;
                        GroupClassification = $o365group.Classification;
                        GroupMemberCount = $o365group.GroupMemberCount;
                        GroupExtMemberCount = $o365group.GroupExternalMemberCount; 
                        SPOSiteUrl =  $o365group.SharePointSiteUrl;
                        SPOStorageUsed = $spoStorageUsed;
                        SPOtorageQuota = $spoStorageQuota ;
                        SPOSharingSetting = $spoSharingSetting;
                    }
                    $ListOfGroupsTeams.add($GroupTeam) | out-null
                }
                Else{
                    $GroupTeam = [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; 
                        GroupName = $o365group.DisplayName;
                        TeamsEnabled = $false;
                        Provider = $groupDetails.providerName;
                        ManagedBy = $o365group.ManagedBy 
                        WhenCreated = $o365group.WhenCreated;
                        PrimarySMTPAddress = $o365group.PrimarySMTPAddress;
                        GroupGuestSetting = $o365group.AllowAddGuests;
                        GroupAccessType = $o365group.AccessType;
                        GroupClassification = $o365group.Classification;
                        GroupMemberCount = $o365group.GroupMemberCount;
                        GroupExtMemberCount = $o365group.GroupExternalMemberCount; 
                        SPOSiteUrl =  $o365group.SharePointSiteUrl;
                        SPOStorageUsed = $spoStorageUsed;
                        SPOtorageQuota = $spoStorageQuota ;
                        SPOSharingSetting = $spoSharingSetting;
                    }
                    $ListOfGroupsTeams.add($GroupTeam) | out-null
                }

            }
            Else{
                $GroupTeam = [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; 
                    GroupName = $o365group.DisplayName;
                    TeamsEnabled = $false;
                    Provider = "";
                    ManagedBy = $o365group.ManagedBy 
                    WhenCreated = $o365group.WhenCreated;
                    PrimarySMTPAddress = $o365group.PrimarySMTPAddress;
                    GroupGuestSetting = $o365group.AllowAddGuests;
                    GroupAccessType = $o365group.AccessType;
                    GroupClassification = $o365group.Classification;
                    GroupMemberCount = $o365group.GroupMemberCount;
                    GroupExtMemberCount = $o365group.GroupExternalMemberCount; 
                    SPOSiteUrl =  $o365group.SharePointSiteUrl;
                    SPOStorageUsed = $spoStorageUsed;
                    SPOtorageQuota = $spoStorageQuota ;
                    SPOSharingSetting = $spoSharingSetting;
                }
                $ListOfGroupsTeams.add($GroupTeam) | out-null
            }
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Error with Group: $($o365group.PrimarySMTPAddress) :: $_" -ForegroundColor White
        }
        $i++
    }

    IF($ExportToCSV){
        Write-LogEntry -LogName:$Log -LogEntryText "Found $($ListOfGroupsTeams.count) Teams" -ForegroundColor White
        $ListOfGroupsTeams | Export-CSV -Path $TeamsCSV -NoTypeInformation
    }
}

#Get Teams Membership
Function Get-TeamsMembersGuests(){
    If(!$ListOfGroupsTeams){
        Get-Teams
    }

    #Check for previous report, delete if found to avoid mixed results
    if (Test-Path $TeamsMemberGuestCSV) {
        Remove-Item $TeamsMemberGuestCSV
    }

    $count = $ListOfGroupsTeams.count
    $i = 0
    Write-LogEntry -LogName:$Log -LogEntryText "Getting Teams Membership Report..." -ForegroundColor Yellow
    Write-LogEntry -LogName:$Log -LogEntryText "Processing $count Teams..." -ForegroundColor White
    foreach ($team in $ListOfGroupsTeams){
        Write-Progress -Activity "Getting Team Membership..." -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        $membership = New-Object System.Collections.ArrayList
        try{
            $owners = Get-UnifiedGroupLinks -Identity $team.GroupID -linktype Owners
            foreach($owner in $owners){
                $record = [pscustomobject]@{GroupID = $team.GroupID;
                        GroupName = $team.GroupName;
                        TeamsEnabled = $team.TeamsEnabled;
                        Member = $owner.PrimarySMTPAddress;
                        Name = $owner.Name;
                        RecipientType = $owner.RecipientType;
                        Membership = "Owner"}
                $membership.add($record) | out-null
            }
            #$members = Get-UnifiedGroupLinks -Identity $team.GroupID -linktype Members | where-object {($membership.Member -notcontains $_.PrimarySMTPAddress)}
            $members = Get-UnifiedGroupLinks -Identity $team.GroupID -linktype Members 
            foreach($MemberOrGuest in $members){
                If($MemberOrGuest.Name -like "*#EXT#*"){
                    $record = [pscustomobject]@{GroupID = $team.GroupID;
                        GroupName = $team.GroupName;
                        TeamsEnabled = $team.TeamsEnabled;
                        Member = $MemberOrGuest.PrimarySMTPAddress;
                        Name = $MemberOrGuest.Name;
                        RecipientType = $MemberOrGuest.RecipientType;
                        Membership = "Guest"}
                    $membership.add($record) | out-null
                }
                Else{
                    $record = [pscustomobject]@{GroupID = $team.GroupID;
                        GroupName = $team.GroupName;
                        TeamsEnabled = $team.TeamsEnabled;
                        Member = $MemberOrGuest.PrimarySMTPAddress;
                        Name = $MemberOrGuest.Name;
                        RecipientType = $MemberOrGuest.RecipientType;
                        Membership = "Member"}
                    $membership.add($record) | out-null
                }
            }
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Membership report error with: $team" 
        }
        $i++

        #Flush membership after every team to maintain low memory usage
        $membership | Export-CSV -Path $TeamsMemberGuestCSV -append -NoTypeInformation
    }    
}

#Get Teams Settings
Function Get-TeamsSettings{
    Write-LogEntry -LogName:$Log -LogEntryText "Getting Teams Tenant Settings Report..." -ForegroundColor Yellow

    #pre-flight
    try{
        Import-module -Name AzureADPreview -force
        Get-AzureADDirectorySettingTemplate | out-null
    }
    catch{
        write-host "Unable to collect Teams Tenant Settings Report. You must connect to Azure AD Preview PowerShell to gather Azure AD Groups information"
        break;
    }

    #variables
    $sb = New-Object -TypeName "System.Text.StringBuilder"

    #Log header
    $sb.appendline("Report: $(Get-Date)") | out-null
    $sb.appendline("") | out-null  
    $sb.appendline("********************************************TEAMS GROUP SETTINGS********************************************") | out-null
    
    #Get tenant O365 Group Setting
    $Template = Get-AzureADDirectorySettingTemplate | Where-Object {$_.DisplayName -eq 'Group.Unified'}
    $Setting = $Template.CreateDirectorySetting()
    #create setting if non-existent: https://support.office.com/en-us/article/Manage-who-can-create-Office-365-Groups-4c46c8cb-17d0-44b5-9776-005fced8e618
    Try{New-AzureADDirectorySetting -DirectorySetting $Setting}
    catch{}
    $setting = Get-AzureADDirectorySetting | where-object {$_.displayname -eq "Group.Unified"}
    $string = $setting.values | out-string 
    $sb.appendline($string) | out-null

    #TEAM CREATION
    $sb.appendline("********************************************TEAMS CREATION********************************************") | out-null
    $creationSettings = $setting.values | ?{$_.name -like "EnableGroupCreation" -or $_.name -like "GroupCreationAllowedGroupID"} | out-string
    $sb.appendline($creationSettings) | out-null
    $sb.appendline("") | out-null 
    $orgAllowedToCreateGroups = $setting.values | ?{$_.Name -like "EnableGroupCreation"} | select -ExpandProperty value
    $groupAllowedToCreateGroups = $setting.values | ?{$_.Name -like "GroupCreationAllowedGroupId"} | select -ExpandProperty value
    If($orgAllowedToCreateGroups -eq $true){
        $sb.appendline("Everyone in the organization is allowed to create Teams.") | out-null 
    }
    Elseif($groupAllowedToCreateGroups -ne $null){
        $groupNameAllowedToCreateGroups = Get-AzureADGroup -ObjectId $groupAllowedToCreateGroups
        $sb.appendline("Only members of the below group are allowed to create Teams: ") | out-null
        $sb.appendline("$($groupNameAllowedToCreateGroups)") | out-null  
    }
    Else{
        $sb.appendline("No one is allowed to create Teams.") | out-null 
    }
    $sb.appendline("") | out-null

    #GUEST ACCESS  
    $sb.appendline("********************************************GUEST ACCESS********************************************") | out-null
    $guestSettings =  $setting.values | ?{$_.name -like "AllowGuestsToAccessGroups" -or $_.name -like "AllowToAddGuests"} | out-string
    $sb.appendline($guestSettings) | out-null
    $sb.appendline("") | out-null 
    $guestAccessToAllGroups = $setting["AllowGuestsToAccessGroups"]
    $guestCanBeAddedToGroups = $setting["AllowToAddGuests"]
    If($guestAccessToAllGroups -eq $false){
        $sb.appendline("Guest Access Restricted to both new and existing guest users.") | out-null
    }
    ElseIF($guestAccessToAllGroups -eq $true -and $guestCanBeAddedToGroups -eq $false){
        $sb.appendline("Guest Access Allowed for existing guest users, but Restricted to new guest users.") | out-null
    }
    ElseIF($guestAccessToAllGroups -eq $true -and $guestCanBeAddedToGroups -eq $true){
        $sb.appendline("Guest Access Allowed for both existing and new guest users.") | out-null
    }
    $sb.appendline("") | out-null

    #GUEST ACCESS - Allow/Block Domain Policy - AAD Premium feature
    $sb.appendline("SPO Allow/Blocked Domain Setting:") | out-null
    $string = GetSPOPolicy | out-string 
    $sb.appendline($string) | out-null

    $sb.appendline("Azure AD B2B ALLOW Setting") | out-null
    $string = GetExistingAllowedDomainList | out-string
    If($string){
        $sb.appendline($string) | out-null
    }
    Else{
        $sb.appendline("None") | out-null
    }
    $sb.appendline("") | out-null

    $sb.appendline("Azure AD B2B BLOCK Setting") | out-null
    $string = GetExistingBlockedDomainList | out-string
    If($string){
        $sb.appendline($string) | out-null
    }
    Else{
        $sb.appendline("None") | out-null
    }
    $sb.appendline("") | out-null

    #EXPIRATION POLICY
    $sb.appendline("********************************************EXPIRATION POLICY********************************************") | out-null
    $policy = Get-AzureADMSGroupLifecyclePolicy
    $string = $policy | fl | out-string 
    $sb.appendline($string) | out-null 

    If(!$policy){
        $sb.appendline("None") | out-null                
    }
    Else{
        If($policy.ManagedGroupTypes -eq "All"){
            $sb.appendline("All Teams Subject To Expiration Policy of $($policy.GroupLifeTimeInDays) Days.") | out-null
        }
        ElseIf($policy.ManagedGroupTypes -eq "Selected"){
            $sb.appendline("Only the below groups are subject to the Group Expiration Policy of $($policy.GroupLifeTimeInDays) Days.") | out-null
            If(!$ListOfGroupsTeams){
                Write-LogEntry -LogName:$Log -LogEntryText "Need to get list of Teams to get all settings..." -ForegroundColor White
                Get-Teams
            }
            foreach($team in $ListOfGroupsTeams){
                #Since only selected Groups are subject to expiration policy. We need to loop and find which ones were selected. 
                $check = get-azureadmslifecyclepolicygroup -id $team.GroupID
                If($check){
                    #$record = [pscustomobject]@{ObjectID = $team.GroupID;Name = $team.GroupName; PrimarySMTPAddress = $team.PrimarySMTPAddress} 
                    $sb.appendline("+ $($team.PrimarySMTPAddress)") | out-null
                }
            }
        }
        $sb.appendline("") | out-null
        $sb.appendline("For more info on group expiration policies: https://docs.microsoft.com/en-us/azure/active-directory/active-directory-groups-lifecycle-azure-portal") | out-null
    }
    $sb.ToString() > $teamsSettingsOut
}

#Get Teams not being used
Function Get-InactiveTeams(){
    $inactiveDays = -90

    $WarningDate = (Get-Date).AddDays($inactiveDays) #90 days
    $Today = (Get-Date)
    $Date = $Today.ToShortDateString()

    Write-LogEntry -LogName:$Log -LogEntryText "Getting Inactive Teams Report Based on Inactivity for $inactiveDays Days..." -ForegroundColor Yellow

    If(!$ListOfGroupsTeams){
        Write-LogEntry -LogName:$Log -LogEntryText "List of Teams Not Found, Getting That Report First..." -ForegroundColor White
        Get-Teams
    }

    $inactiveTeams = New-Object System.Collections.ArrayList
    
    $count = $ListOfGroupsTeams.count
    $i=0
    foreach($team in $ListOfGroupsTeams){
        Write-Progress -Activity "Getting InActive Teams Info..." -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        # Fetch information about activity in the Inbox folder of the group mailbox  
        $Data = (Get-MailboxFolderStatistics -Identity $team.PrimarySMTPAddress -IncludeOldestAndNewestITems -FolderScope Inbox)
        $LastConversation = $Data.NewestItemReceivedDate
        $NumberConversations = $Data.ItemsInFolder
        $MailboxStatus = "Normal"
        
        If ($Data.NewestItemReceivedDate -le $WarningDate){
            #days since any activity in the group mailbox
            $MailboxStatus = "Inactive"
        }
        <# If want to check for items in folder
            If ($Data.ItemsInFolder -lt 20){
                #Less than 20 conversations
                $MailboxStatus = "Inactive"
            }
        #>

        $SPOLastContentModified = (get-sposite $team.SPOSiteUrl).LastContentModifiedDate
        <# If want to check for usage
            #$SPOStorageUsageCurrent = (get-sposite $team.SPOSiteUrl).StorageUsageCurrent
        #>
        $SPOStatus = "Active"

        If ($SPOLastContentModified -le $WarningDate){
            $SPOStatus = "Inactive"
        }

        $record = [pscustomobject]@{GroupID = $team.GroupID;
            Name = $team.GroupName;
            TeamsEnabled = $team.TeamsEnabled;
            PrimarySMTPAddress = $team.PrimarySMTPAddress;
            MailboxStatus = $MailboxStatus;
            LastConversationDate = $Data.NewestItemReceivedDate;
            NumOfConversations = $Data.ItemsInFolder;
            SPOStatus = $SPOStatus;
            LastContentModified = $SPOLastContentModified;
            StorageUsageCurrent = $SPOStorageUsageCurrent;
        }
        $inactiveTeams.add($record) | out-null
        $i++
    }

    $inactiveTeams | Export-CSV -Path $InactiveTeamsCSV -NoTypeInformation

}

#Get Teams That User X Owns
Function Get-TeamsByUser(){
    $user = Read-host "Enter the UPN of the user, e.g. userA@contoso.com"
    $hashOfTeamsUserOwns = @{}
    $listOfTeamsByUser = New-Object -TypeName System.Collections.ArrayList

    if (-not (Get-PSSession | where {($_.ComputerName -eq "outlook.office365.com") -and ($_.State -eq "Opened")})) {
        throw "You must connect to Exchange Online Remote PowerShell to gather Groups information"
    }

    Write-LogEntry -LogName:$Log -LogEntryText "Searching for groups by user..." -ForegroundColor Yellow

    $userobject = Get-User -Identity $User
    $o365groups = Get-UnifiedGroup -Filter ("Members -eq '{0}'" -f $userobject.DistinguishedName)

    Write-LogEntry -LogName:$Log -LogEntryText "Found $($o365groups.Count) groups for user $($User)" -ForegroundColor Yellow
    
    $count = $o365groups.count
    $i=0
    foreach($o365group in $o365groups){
        Write-Progress -Activity "Getting Group/Team Owners" -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        $owners = $o365group | Get-UnifiedGroupLinks -linktype "Owners"
        If($owners.PrimarySMTPAddress -contains $user){
            $hashOfTeamsUserOwns.add($o365group.ExternalDirectoryObjectId,$true) | Out-Null
        }
        Else{
            $hashOfTeamsUserOwns.add($o365group.ExternalDirectoryObjectId,$false) | Out-Null
        }
        $i++
    }

    $count = $o365groups.count
    $i=0
    foreach ($o365group in $o365groups) {
        Write-Progress -Activity "Building collection of Teams/Groups by User" -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        try {
            $GroupsUri = "https://graph.microsoft.com/beta/groups/$($o365group.ExternalDirectoryObjectId)/endpoints"
            $groupDetails = (Invoke-RestMethod -Uri $GroupsUri -Headers $authHeader -Method Get).value

            If($groupDetails){
                If($groupDetails.providerName -eq "Microsoft Teams"){
                    $groupID = $o365group.ExternalDirectoryObjectId
                    $addGroup = [pscustomobject]@{User = $user; GroupId = $groupID; GroupName = $o365group.DisplayName; TeamsEnabled = $true; IsOwner = $hashOfTeamsUserOwns[$groupID]} 
                    $listOfTeamsByUser.add($addGroup) | Out-Null
                }
                Else{
                    $groupID = $o365group.ExternalDirectoryObjectId
                    $addGroup = [pscustomobject]@{User = $user; GroupId = $groupID; GroupName = $o365group.DisplayName; TeamsEnabled = $false; IsOwner = $hashOfTeamsUserOwns[$groupID]} 
                    $listOfTeamsByUser.add($addGroup) | Out-Null
                }
            }
            Else{
                $groupID = $o365group.ExternalDirectoryObjectId
                $addGroup = [pscustomobject]@{User = $user; GroupId = $groupID; GroupName = $o365group.DisplayName; TeamsEnabled = $false; IsOwner = $hashOfTeamsUserOwns[$groupID]} 
                $listOfTeamsByUser.add($addGroup) | Out-Null
            }
        } 
        catch {
            Write-LogEntry -LogName:$Log -LogEntryText "Error with Group: $($o365group.PrimarySMTPAddress) :: $_" -ForegroundColor White
        }
        $i++
    }

    $listOfTeamsByUser | Export-CSV -Path $TeamsByUserCSV -NoTypeInformation
}

#Get Users That Can Create Teams
Function Get-UsersCanCreateTeams(){
    Write-LogEntry -LogName:$Log -LogEntryText "Getting Users-Can-Create-Teams Report..." -ForegroundColor Yellow

    #pre-flight
    try{
        Import-module -Name AzureADPreview -force
        Get-AzureADDirectorySettingTemplate | out-null
    }
    catch{
        write-host "Unable to collect Users-Can-Create-Teams Report. You must connect to Azure AD Preview PowerShell to gather Azure AD Groups information"
        break;
    }

    #Get tenant O365 Group Setting
    $Template = Get-AzureADDirectorySettingTemplate | Where-Object {$_.DisplayName -eq 'Group.Unified'}
    $Setting = $Template.CreateDirectorySetting() | out-null
    #create setting if non-existent: https://support.office.com/en-us/article/Manage-who-can-create-Office-365-Groups-4c46c8cb-17d0-44b5-9776-005fced8e618
    Try{New-AzureADDirectorySetting -DirectorySetting $Setting | out-null}
    catch{}
    $setting = Get-AzureADDirectorySetting | where-object {$_.displayname -eq "Group.Unified"}

    $orgAllowedToCreateGroups = $setting.values | ?{$_.Name -like "EnableGroupCreation"} | select -ExpandProperty value
    $groupIDAllowedToCreateGroups = $setting.values | ?{$_.Name -like "GroupCreationAllowedGroupId"} | select -ExpandProperty value
    If($orgAllowedToCreateGroups -eq $true){
        [pscustomobject]@{ObjectID = "Everyone in the organization is allowed to create Teams.";DisplayName="";UserPrincipalName="";UserType=""} | Export-CSV -Path $UsersCanCreateCSV -NoTypeInformation  
    }
    Elseif($groupIDAllowedToCreateGroups){
        $groupAllowedToCreateGroups = Get-AzureADGroup -ObjectId $groupIDAllowedToCreateGroups
        $members = $groupAllowedToCreateGroups | get-azureadgroupmember | %{[pscustomobject]@{ObjectID = $_.ObjectID;
            DisplayName = $_.DisplayName;
            UserPrincipalName = $_.UserPrincipalName;
            UserType = $_.UserType}}
        $members | Export-CSV -Path $UsersCanCreateCSV -NoTypeInformation  
    }
    Elseif($orgAllowedToCreateGroups -eq $false){
        [pscustomobject]@{ObjectID = "No one is allowed to create Teams.";DisplayName="";UserPrincipalName="";UserType=""} | Export-CSV -Path $UsersCanCreateCSV -NoTypeInformation  
    }
}

Function Get-GroupsWithoutOwners(){
    $listOfTeamsWithoutOwners = New-Object -TypeName System.Collections.ArrayList

    Write-LogEntry -LogName:$Log -LogEntryText "Getting Teams without Owners Report..." -ForegroundColor Yellow

    If(!$ListOfGroupsTeams){
        Write-LogEntry -LogName:$Log -LogEntryText "List of Teams Not Found, Getting That Report First..." -ForegroundColor White
        Get-Teams
    }

    $count = $ListOfGroupsTeams.count
    $i=0
    foreach($team in $ListOfGroupsTeams){
        Write-Progress -Activity "Getting Teams without Owners Report..." -Status "Processed $i of $count " -PercentComplete ($i/$count*100);
        If ($team.ManagedBy -Ne $Null){
            $groupInfo = [pscustomobject]@{GroupID = $team.GroupID; GroupName = $team.GroupName;HasOwners="True";ManagedBy = $team.managedby}
            $listOfTeamsWithoutOwners.add($groupInfo) | out-null
        }
        Else{
            $groupInfo = [pscustomobject]@{GroupID = $team.GroupID; GroupName = $team.GroupName;HasOwners="False";ManagedBy = $team.managedby}
            $listOfTeamsWithoutOwners.add($groupInfo) | out-null
        }
    }

    $listOfTeamsWithoutOwners | Export-CSV -Path $groupsNoOwnersCSV -NoTypeInformation

}

#endregion Functions

#region MAIN

Clear-Host

#region Variables
$yyyyMMdd = Get-Date -Format 'yyyyMMdd'
$computer = $env:COMPUTERNAME
$user = $env:USERNAME
$version = "06112018"
$log = "$PSScriptRoot\Manage-Teams-$yyyyMMdd.log"
$output = $PSScriptRoot
$TeamsCSV = "$($output)\ListOfTeams.csv"
$TeamsMemberGuestCSV = "$($output)\ListOfMembers.csv"
$InactiveTeamsCSV = "$($output)\ListOfInactiveTeams.csv"
$UsersCanCreateCSV = "$($output)\ListOfUsersThatCanCreateTeams.csv"
$teamsSettingsOut = "$($output)\ListOfTeamsSettings.txt"
$TeamsByUserCSV = "$($output)\ListOfTeamsByUser.csv"
$groupsNoOwnersCSV = "$($output)\ListOfTeamsWithoutOwners.csv"
Write-LogEntry -LogName:$Log -LogEntryText "User: $user Computer: $computer Version: $version" -foregroundcolor Yellow
#endregion Variables

[string] $menu = @'

    ******************************************************************
	                    Manage Microsoft Teams
    ******************************************************************
	
    Please select an option from the list below:

        0) Check Script Pre-requisites
        1) Connect to O365
        2) Get Teams
        3) Get Teams Membership
        4) Get Teams That Are Not Active
        5) Get Users That Are Allowed To Create Teams
        6) Get Teams Tenant Settings
        7) Get Groups/Teams Without Owner(s)
        8) Get All Above Reports
        9) Get Teams By User
       10) Exit Script

Select an option.. [0-10]?
'@

Do { 	
	if ($opt) {"";Write-LogEntry -LogName:$Log -LogEntryText "Last command: $opt" -foregroundcolor White}	
	$opt = Read-Host $menu

	switch ($opt)    {
    			
	  	0 { # Logon to required services
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 0" -ForegroundColor Yellow
            Check-Modules
        }

        1 { # Logon to required services
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 1" -ForegroundColor Yellow
            $acctHasMFA = Read-Host "Is your account enabled for MFA? (Y/N)"
            If($acctHasMFA -eq "Y" -or $acctHasMFA -eq "y"){
                Logon-O365MFA
            }
            ElseIf($acctHasMFA -eq "N" -or $acctHasMFA -eq "n"){
                Logon-O365
            }
            Else{
                Write-LogEntry -LogName:$Log -LogEntryText "Please type 'Y' for Yes or 'N' for No" -ForegroundColor Yellow
            }
        }

        2 { # Get Teams
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 2" -ForegroundColor Yellow
            Get-Teams -ExportToCSV
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($TeamsCSV) " -ForegroundColor Green
        }

        3 { # Get Teams Members and Guests
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 3" -ForegroundColor Yellow
            Get-TeamsMembersGuests
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($TeamsMemberGuestCSV)" -ForegroundColor Green
        }

        4 { # Get Teams that are not active 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 4" -ForegroundColor Yellow
            Get-InactiveTeams
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($InactiveTeamsCSV)" -ForegroundColor Green
        }

        5 { # Get Users That Are Allowed to Create Teams 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 5" -ForegroundColor Yellow
            Get-UsersCanCreateTeams
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($UsersCanCreateCSV)" -ForegroundColor Green
        }

        6 { # Get Teams Tenant Settings 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 6" -ForegroundColor Yellow
            Get-TeamsSettings
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($teamsSettingsOut)" -ForegroundColor Green
        }

        7 { # Get List of Groups/Teams without Owner
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 7" -ForegroundColor Yellow
            Get-GroupsWithoutOwners
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($groupsNoOwnersCSV)" -ForegroundColor Green
        }

        8 { # Get All Reports Above
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 8" -ForegroundColor Yellow
            Get-Teams -ExportToCSV
            Get-TeamsMembersGuests
            Get-InactiveTeams
            Get-UsersCanCreateTeams
            Get-GroupsWithoutOwners
            Get-TeamsSettings
            Write-LogEntry -LogName:$Log -LogEntryText "Reports location: $($output)" -ForegroundColor Green
        }

        9 { # Get Teams that a User (input) Owns
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 9" -ForegroundColor Yellow
            Get-TeamsByUser
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($TeamsByUserCSV)" -ForegroundColor Green
        }

		10 { # Remove sessions and exit
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 10" -ForegroundColor Yellow
            try{Disconnect-AzureAD -erroraction silentlycontinue}
            catch{}
            try{Remove-PSSession $sfbSession -ErrorAction SilentlyContinue}
            catch{}
            try{Remove-PSSession $exchangeSession -erroraction silentlycontinue}
            catch{}
            try{Remove-PSSession $ccSession -erroraction silentlycontinue}
            catch{}
            try{Disconnect-SPOService -erroraction silentlycontinue}
            catch{}
            try{Disconnect-PnPOnline -erroraction silentlycontinue}
            catch{}
            try{Disconnect-MicrosoftTeams -erroraction silentlycontinue}
            catch{}
            try{Get-PSSession | Remove-PSSession}
            catch{}
            Write-Host "Exiting..."
		}
		
        default {Write-Host "You haven't selected any of the available options."}
	}
} while ($opt -ne 10)

#endregion MAIN
