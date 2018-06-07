# Manage-Teams
Powershell script to generate Teams reports

# DISCLAIMER

###############Disclaimer#####################################################
THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
###############Disclaimer#####################################################

# OPTIONS  
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

# EXAMPLE
#Run the script with no switches and it will provide you a menu of what reports to run.     
.\Manage-Teams.ps1  

# QUESTIONS
    Alejandro Lopez - Alejanl@Microsoft.com  
