# Manage-Teams
Powershell script to generate Teams reports

# DISCLAIMER

###############Disclaimer#####################################################
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###############Disclaimer#####################################################

# OPTIONS  
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

# EXAMPLE
#Run the script with no switches and it will provide you a menu of what reports to run.     
.\Manage-Teams.ps1  

# CREDIT
    Built leveraging the great work from the following individuals:    
        Get-Teams function - David Whitney (dawhitne@microsoft.com)    
        Get-ObseleteGroup function - Tony Redmond (https://gallery.technet.microsoft.com/Check-for-obsolete-Office-c0020a42)  

# QUESTIONS
    Alejandro Lopez - Alejanl@Microsoft.com  
