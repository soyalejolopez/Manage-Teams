<#
.DESCRIPTION

###############Disclaimer#####################################################
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###############Disclaimer#####################################################

Script to generate Teams reports. 

What this script does: 
    0) Connect to AzureAD and Office 365
    1) Get Teams
        Properties: "GroupId","GroupName","WhenCreated","PrimarySMTPAddress","GroupGuestSetting","GroupAccessType","GroupClassification","GroupMemberCount","GroupExtMemberCount","SPOSiteUrl","SPOStorageUsed","SPOtorageQuota","SPOSharingSetting"
    2) Get Teams Membership
        Properties: "TeamID","TeamName","Member","Name","RecipientType","Membership"
    3) Get Teams That Are Not Active
        Properties: "TeamID","TeamName","PrimarySMTPAddress","MailboxStatus","LastConversationDate","NumOfConversations","SPOStatus","LastContentModified","StorageUsageCurrent" 
    4) Get Users That Are Allowed To Create Teams
        Properties: "ObjectID","DisplayName","UserPrincipalName","UserType" 
    5) Get Teams Tenant Settings
        Settings captured: Azure AD Group Settings, Who's Allowed to Create Teams, Guest Access, Expiration Policy
    6) Get All Above Reports 
    7) Exit Script

CREDIT: 
    Built on top of the great work from the following individuals:
        Get-Teams function - David Whitney (dawhitne@microsoft.com)
        Get-ObseleteGroup function - Tony Redmond (https://gallery.technet.microsoft.com/Check-for-obsolete-Office-c0020a42)
VERSION:
    02072018 - v1
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

    If($host.version.major -lt 3){
        throw "Powershell V3+ is required. Contact your administrator"
    }

    $aadmodule = get-module -listavailable azureadpreview
    If($aadmodule){
        #Need AzureADPreview 2.0.0.137 for Get-AzureADMSGroupLifecyclePolicy
        $reqAADPMinRevision = 137 
        $checkaadpversion = (get-module -listavailable azureadpreview | select -expandproperty version).revision
        If($reqAADPMinRevision -gt $checkaadpversion){
            $needAADPModuleInstall = $true
        }
    }
    Else{$needAADPModuleInstall = $true}

    If($needAADPModuleInstall -eq $true){
        $check = Read-Host "Didn't find the required AzureADPreview module version installed - https://www.powershellgallery.com/packages/AzureADPreview/. This is required to proceed. Install? (Y/N)"
        If($check -eq "Y" -or $check -eq "y"){
            try{
                Write-LogEntry -LogName:$Log -LogEntryText "Installing latest version of AzureADPreview Module..." -ForegroundColor White
                Remove-module AzureADPreview -ErrorAction SilentlyContinue
                Install-Module AzureADPreview -Force
                Write-LogEntry -LogName:$Log -LogEntryText "Successfully installed AzureADPreview Module." -ForegroundColor Green
            }
            catch{
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to install the AzureADPreview Module. Please install from here: https://www.powershellgallery.com/packages/AzureADPreview/" -ForegroundColor Red
                Write-LogEntry -LogName:$Log -LogEntryText "$_" -ForegroundColor Red
                exit        
            }
        }
        Else{
             Write-LogEntry -LogName:$Log -LogEntryText "Please install AzureADPreview to move forward: https://www.powershellgallery.com/packages/AzureADPreview/" -ForegroundColor White
             exit
        }
    }
    #If AzureAD module (Not AzureADPreview) is also available, then the AzureADPreview module is not loaded
    $checkAzureADModule = get-module -name "AzureAD"
    If($checkforAzureADModule -ne $null){
        Remove-Module -Name "AzureAD"
    }
    Import-module -Name AzureADPreview

    If(!(get-module -listavailable MicrosoftTeams)){
        $check = Read-Host "Didn't find MicrosoftTeams module installed - https://www.powershellgallery.com/packages/MicrosoftTeams/. This is required to proceed.` Install? (Y/N)"
        If($check -eq "Y" -or $check -eq "y"){
            try{
                Install-Module MicrosoftTeams
                Write-LogEntry -LogName:$Log -LogEntryText "Successfully installed Microsoft Teams Module." -ForegroundColor Green        
            }
            catch{
                Write-LogEntry -LogName:$Log -LogEntryText "Unable to install the Microsoft Teams Module. Please install from here: https://www.powershellgallery.com/packages/MicrosoftTeams/" -ForegroundColor Red
                Write-LogEntry -LogName:$Log -LogEntryText "$_" -ForegroundColor Red
                exit        
            }     
        }
        Else{
             Write-LogEntry -LogName:$Log -LogEntryText "Please install AzureADPreview to move forward: https://www.powershellgallery.com/packages/MicrosoftTeams/" -ForegroundColor White
             exit
        }
    }
    Import-module -Name MicrosoftTeams

    If(!(get-module -listavailable microsoft.online.sharepoint.powershell)){
        Write-LogEntry -LogName:$Log -LogEntryText "SPO management shell not found. If you recently installed it, make sure to close and re-open your powershell window. Please install and re-run script." -ForegroundColor Red
        Write-LogEntry -LogName:$Log -LogEntryText "Install Link: https://www.microsoft.com/en-us/download/details.aspx?id=35588" -ForegroundColor Red
        exit
    }
    Import-Module "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.Online.SharePoint.PowerShell.dll" -DisableNameChecking

    $CheckForSignInAssistant = Test-Path "HKLM:\SOFTWARE\Microsoft\MSOIdentityCRL"
    If ($CheckForSignInAssistant -eq $false) {
        Write-LogEntry -LogName:$Log -LogEntryText "Microsoft Online Services Sign-in Assistant not found. If you recently installed it, make sure to close and re-open your powershell window. Please install and re-run script." -ForegroundColor Red
        Write-LogEntry -LogName:$Log -LogEntryText "Install Link: https://go.microsoft.com/fwlink/p/?LinkId=286152" -ForegroundColor Red
        exit
    }

    Write-LogEntry -LogName:$Log -LogEntryText "Pre-Flight Done" -ForegroundColor Green
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
Function Logon-O365 {
    #https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-all-office-365-services-in-a-single-windows-powershell-window
    ""
    $domainHost = Read-Host "Enter tenant name, such as contoso for contoso.onmicrosoft.com" 
    $spoAdminUrl = "https://$domainHost-admin.sharepoint.com"
    $global:credential = Get-Credential
    
    try{$testSPO = get-spotenant -erroraction silentlycontinue}
    catch{}
    If($testSPO -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
    }
    Else{
        try{
            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
            Connect-SPOService -Url $spoAdminUrl -credential $credential
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to SharePoint Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to SharePoint Online. Re-running the script should help." -ForegroundColor Yellow 
        }
    }

    start-sleep 2
    try{$testTeams = Get-Team -erroraction silentlycontinue}
    catch{}
    If($testTeams -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green
    }
    Else{
        try{
            Import-Module MicrosoftTeams
            Connect-MicrosoftTeams -credential $credential | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Microsoft Teams" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Microsoft Teams. Re-running the script should help." -ForegroundColor Yellow 
        }
    }
           
    #need to wait a bit before connecting to EXO
    start-sleep 2
    $session = Get-PSSession | where {($_.ComputerName -eq "outlook.office365.com") -and ($_.State -eq "Opened")}
    If ($session -ne $null) {
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
    }
    Else{
        try{
            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue 
            #Import-PSSession -session $exchangeSession -DisableNameChecking -AllowClobber | out-null
            Import-Module (Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Exchange Online" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Exchange Online. Re-running the script should help." -ForegroundColor Yellow
        }
    }    
    #need to wait a bit before connecting to EXO
    start-sleep 2
    If (Get-PSSession | where {($_.ComputerName -eq "ps.compliance.protection.outlook.com") -and ($_.State -eq "Opened")}) {
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Compliance Center" -ForegroundColor Green    
    }
    Else{
        try{
            $ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection -WarningAction SilentlyContinue
            #Import-PSSession -session $ccSession -Prefix cc -DisableNameChecking -AllowClobber | out-null
            Import-Module (Import-PSSession $ccSession -DisableNameChecking -AllowClobber) -Global -DisableNameChecking | out-null
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Compliance Center" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Compliance Center. Re-running the script should help." -ForegroundColor Yellow
        }
    }
    start-sleep 2
    try{$testAAD = Get-AzureADCurrentSessionInfo -erroraction silentlycontinue}
    catch{}
    If($testAAD -ne $null){
        Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
    }
    Else{
        try{
            #Need AzureADPreview 2.0.0.137 for Get-AzureADMSGroupLifecyclePolicy 
            Import-Module AzureADPreview
            start-sleep 2
            Connect-AzureAD -credential $credential | out-null #https://github.com/itnetxbe/Feedback/issues/15 - login issues sporadically
            Write-LogEntry -LogName:$Log -LogEntryText "Connected to Azure AD" -ForegroundColor Green
        }
        catch{
            Write-LogEntry -LogName:$Log -LogEntryText "Unable to connect to Azure AD. Re-running the script should help." -ForegroundColor Yellow
        }
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
    param (
        [switch]$ExportToCSV
    )

    if (-not (Get-PSSession | where {($_.ComputerName -eq "outlook.office365.com") -and ($_.State -eq "Opened")})) {
        throw "You must connect to Exchange Online Remote PowerShell..."
    }
    $testSPO = get-spotenant
    if (!$testSPO){
        throw "You must connect to SharePoint Online PowerShell..."
    }
    $testTeams = get-team
    if(!$testTeams){
        throw "You must connect to Microsoft Teams PowerShell..."
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
        try {
            $teamschannels = Get-TeamChannel -GroupId $o365group.ExternalDirectoryObjectId
            $GroupTeam = [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; 
                GroupName = $o365group.DisplayName; 
                WhenCreated = $o365group.WhenCreated;
                PrimarySMTPAddress = $o365group.PrimarySMTPAddress;
                GroupGuestSetting = $o365group.AllowAddGuests;
                GroupAccessType = $o365group.AccessType;
                GroupClassification = $o365group.Classification;
                GroupMemberCount = $o365group.GroupMemberCount;
                GroupExtMemberCount = $o365group.GroupExternalMemberCount; 
                SPOSiteUrl =  $o365group.SharePointSiteUrl;
                SPOStorageUsed = $spoStorageQuota;
                SPOtorageQuota = $spoStorageUsed;
                SPOSharingSetting = $spoSharingSetting;
            }
            $ListOfGroupsTeams.add($GroupTeam) | out-null
        } catch {
            $Exception = $_.Exception
            if ($Exception.Message -like "*Connect-MicrosoftTeams*") {
                throw $Exception
            }

            $ErrorCode = $Exception.ErrorCode
            switch ($ErrorCode) {
                "404" {
                    break;
                }
                "403" {
                    $GroupTeam = [pscustomobject]@{GroupId = $o365group.ExternalDirectoryObjectId; 
                        GroupName = $o365group.DisplayName; 
                        WhenCreated = $o365group.WhenCreated;
                        PrimarySMTPAddress = $o365group.PrimarySMTPAddress;
                        GroupGuestSetting = $o365group.AllowAddGuests;
                        GroupAccessType = $o365group.AccessType;
                        GroupClassification = $o365group.Classification; 
                        SPOSiteUrl =  $o365group.SharePointSiteUrl;
                        SPOStorageUsed = $spoStorageQuota;
                        SPOtorageQuota = $spoStorageUsed;
                        SPOSharingSetting = $spoSharingSetting;
                    }
                    $ListOfGroupsTeams.add($GroupTeam) | out-null
                    break;
                }
                default {
                    Write-LogEntry -LogName:$Log -LogEntryText "Unknown ErrorCode trying to 'Get-TeamChannel -GroupId $($o365group)' :: $($ErrorCode)" -ForegroundColor Red
                    $Errorcatch
                }
            }
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
                $record = [pscustomobject]@{TeamID = $team.GroupID;
                        TeamName = $team.GroupName;
                        Member = $owner.PrimarySMTPAddress;
                        Name = $owner.Name;
                        RecipientType = $owner.RecipientType;
                        Membership = "Owner"}
                $membership.add($record) | out-null
            }
            $members = Get-UnifiedGroupLinks -Identity $team.GroupID -linktype Members | where-object {($membership.Member -notcontains $_.PrimarySMTPAddress)}
            foreach($MemberOrGuest in $members){
                If($MemberOrGuest.Name -like "*#EXT#*"){
                    $record = [pscustomobject]@{TeamID = $team.GroupID;
                        TeamName = $team.GroupName;
                        Member = $MemberOrGuest.PrimarySMTPAddress;
                        Name = $MemberOrGuest.Name;
                        RecipientType = $MemberOrGuest.RecipientType;
                        Membership = "Guest"}
                    $membership.add($record) | out-null
                }
                Else{
                    $record = [pscustomobject]@{TeamID = $team.GroupID;
                        TeamName = $team.GroupName;
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
    try{Get-AzureADDirectorySettingTemplate | out-null}
    catch{
        throw "You must connect to Azure AD Preview PowerShell to gather Azure AD Groups information"
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
    $WarningDate = (Get-Date).AddDays(-90) #90 days
    $Today = (Get-Date)
    $Date = $Today.ToShortDateString()
    $SPOWarningStorageUsage = "10" #MB

    Write-LogEntry -LogName:$Log -LogEntryText "Getting Inactive Teams Report..." -ForegroundColor Yellow

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
            #90 days since any activity in the group mailbox
            $MailboxStatus = "Inactive"
        }
        If ($Data.ItemsInFolder -lt 20){
            #Less than 20 conversations
            $MailboxStatus = "Inactive"
        }

        $SPOLastContentModified = (get-sposite $team.SPOSiteUrl).LastContentModifiedDate
        $SPOStorageUsageCurrent = (get-sposite $team.SPOSiteUrl).StorageUsageCurrent
        $SPOStatus = "Active"

        If ($SPOLastContentModified -le $WarningDate){
            #no activity in the last 90 days
            $SPOStatus = "Inactive"
        }
        If($SPOStorageUsageCurrent -le $SPOWarningStorageUsage){
            #less than 10 MB of usage
            $SPOStatus = "Inactive"
        }

        $record = [pscustomobject]@{TeamID = $team.GroupID;
            TeamName = $team.GroupName;
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

#Get Users That Can Create Teams
Function Get-UsersCanCreateTeams(){
    Write-LogEntry -LogName:$Log -LogEntryText "Getting Users-Can-Create-Teams Report..." -ForegroundColor Yellow

    #pre-flight
    try{Get-AzureADDirectorySettingTemplate | out-null}
    catch{
        write-host "You must connect to Azure AD Preview PowerShell to gather Azure AD Groups information"
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

#endregion Functions

#region MAIN

Clear-Host

#region Variables
$yyyyMMdd = Get-Date -Format 'yyyyMMdd'
$computer = $env:COMPUTERNAME
$user = $env:USERNAME
$log = "$PSScriptRoot\Manage-Teams-$yyyyMMdd.log"
$output = $PSScriptRoot
$TeamsCSV = "$($output)\ListOfTeams.csv"
$TeamsMemberGuestCSV = "$($output)\ListOfMembers.csv"
$InactiveTeamsCSV = "$($output)\ListOfInactiveTeams.csv"
$UsersCanCreateCSV = "$($output)\ListOfUsersThatCanCreateTeams.csv"
$teamsSettingsOut = "$($output)\ListOfTeamsSettings.txt"
Write-LogEntry -LogName:$Log -LogEntryText "User: $user Computer: $computer" -foregroundcolor Yellow
#endregion Variables

#Preflight Check
Check-Modules

[string] $menu = @'

    ******************************************************************
	                    Manage Microsoft Teams
    ******************************************************************
	
    Please select an option from the list below:

        0) Connect to AzureAD and Office 365
        1) Get Teams
        2) Get Teams Membership
        3) Get Teams That Are Not Active 
        4) Get Users That Are Allowed To Create Teams
        5) Get Teams Tenant Settings
        6) Get All Above Reports 
        7) Exit Script

Select an option.. [0-7]?
'@

Do { 	
	if ($opt) {"";Write-LogEntry -LogName:$Log -LogEntryText "Last command: $opt" -foregroundcolor White}	
	$opt = Read-Host $menu

	switch ($opt)    {
    			
	  	0 { # Logon to required services
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 0" 
            Logon-O365
        }

        1 { # Get Teams
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 1" 
            Get-Teams -ExportToCSV
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($TeamsCSV) " -ForegroundColor Green
        }

        2 { # Get Teams Members and Guests
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 2"
            Get-TeamsMembersGuests
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($TeamsMemberGuestCSV)" -ForegroundColor Green
        }

        3 { # Get Teams that are not active 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 3"
            Get-InactiveTeams
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($InactiveTeamsCSV)" -ForegroundColor Green
        }

        4 { # Get Users That Are Allowed to Create Teams 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 4"
            Get-UsersCanCreateTeams
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($UsersCanCreateCSV)" -ForegroundColor Green
        }

        5 { # Get Teams Tenant Settings 
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 5"
            Get-TeamsSettings
            Write-LogEntry -LogName:$Log -LogEntryText "Report location: $($teamsSettingsOut)" -ForegroundColor Green
        }

        6 { # Get All Above Reports
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 6"
            Get-Teams -ExportToCSV
            Get-TeamsMembersGuests
            Get-InactiveTeams
            Get-UsersCanCreateTeams
            Get-TeamsSettings
            Write-LogEntry -LogName:$Log -LogEntryText "Reports location: $($output)" -ForegroundColor Green
        }

		7 { # Remove sessions and exit
            Write-LogEntry -LogName:$Log -LogEntryText "Selected option 7"
            try{Disconnect-AzureAD -erroraction silentlycontinue}
            catch{}
            try{Remove-PSSession $exchangeSession -erroraction silentlycontinue}
            catch{}
            try{Remove-PSSession $ccSession -erroraction silentlycontinue}
            catch{}
            try{Disconnect-SPOService -erroraction silentlycontinue}
            catch{}
            try{Disconnect-MicrosoftTeams -erroraction silentlycontinue}
            catch{}
            Write-Host "Exiting..."
		}
		
        default {Write-Host "You haven't selected any of the available options."}
	}
} while ($opt -ne 7)

#endregion MAIN
