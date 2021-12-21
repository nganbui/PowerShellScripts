@{
    Path                          = @{        
        Data = "D:\Scripting\O365DevOps\Common\Data" # not in use
        Log  = "D:\Scripting\O365DevOps\Logs" # not in use
        Cred = "D:\Scripting\O365DevOps\Common\Creds"
    }
    #----------OAuth-new file structure scripting D:\Scripting\O365DevOps--------#    
    AppConfigOperationSupport     = @{
        AppId      = 'ce2ff62c-3d76-4c62-9754-9ef2306a0de2'        
        Thumbprint = 'D0CE39CFCDBE72E92A3233A4C3BB891DC6045431' 
        AppName    = "SPO-M365 Operations Support"
        CertName   = "SPO-M365 AdminPortal Operations"
    }
    AppConfigAdminPortalOperation = @{
        AppId      = '8f3e09ce-dd5b-461d-becf-adea7f883e34'        
        Thumbprint = 'D0CE39CFCDBE72E92A3233A4C3BB891DC6045431' 
        AppName    = "SPO-Sync Operations"
        CertName   = "SPO-M365 AdminPortal Operations"
    }
    AppConfigUsageReport          = @{
        AppId      = '5a85c310-68b1-45dd-b76b-1681a731aac0'        
        Thumbprint = 'D0CE39CFCDBE72E92A3233A4C3BB891DC6045431' 
        AppName    = "SPO-MC and Usage Reports"
        CertName   = "SPO-M365 AdminPortal Operations"
        Resource   = "https://manage.office.com"
    }
    #----------New Azure App for PowerBI Workspace--------# 
    AppConfigPowerBIWorkspace          = @{
        AppId      = 'c8129a8a-6f86-493b-aaf1-369e6e153a50'        
        Thumbprint = 'E1158E188BF714332576160E148E990A4F8EBCD4' 
        AppName    = "SPO-PowerBI.Workspaces"
        CertName   = "SPO-Sync Operations"        
    }

    #----------New Azure App for reference--------# 
    AppConfigNIHTeamsPolicy       = @{
        AppId      = 'b05f1912-0e35-4db0-9600-d85c425a0c22'        
        AppSecret  = 'NIHTeamsPolicyManagement'
        Thumbprint = 'CFAB6381A91413CF33347370A6F14130E4FF5618' 
        AppName    = "SPO - NIH Teams Policy Management"
    }
    AppConfigMC                   = @{
        AppId     = '0f6a6412-8f62-4b7a-a929-664174baf961'        
        AppSecret = 'O365AppMCPwd'        
        Resource  = "https://manage.office.com"   
        AppName   = "SPO - Message Center and Health Reader"
    }
    #----------Tenant Config--------# 
    TenantConfig                  = @{
        Id                   = "c684c0c2-23d6-4d84-ba9d-306a5b0522d8"
        Name                 = "citspdev.onmicrosoft.com"
        AdminCenterUrl       = "https://citspdev-admin.sharepoint.com"
        RootSiteUrl          = "https://citspdev.sharepoint.com"
        O365TenantAdmin      = "spoadmportalsvc@citspdev.onmicrosoft.com"
        CloudSvcForProvision = "SPOADMSVC@citspdev.onmicrosoft.com"
        #CloudSvcForProvision   = "test@citspdev.onmicrosoft.com"        
    }	    
    DBConfig                      = @{
        DBServer  = "NIHSPSQL19D1"
        DBName    = "O365_SelfServicePortal_Dev"
        DBUser    = "icadminportal-dev"
        DBPwdFile = "O365DBPwd"
        <#DBServer  = "NIHSPSQL19P1"
        DBName    = "O365_SelfServicePortal"
        DBUser    = "O365_Sql_Admin"
        DBPwdFile = "O365DBProdPwd"#>   
    }    
    SNConfig                      = @{
        ServiceUrl      = "https://soadev.nih.gov/RemedyService/RemedyServiceDesk.serviceagent/RemedyServiceDeskEndpoint"
        AdminAccount    = "api_o365adminportal"
        PwdFile         = "ServiceNowPwd"
        Group           = "CIT - SharePoint Online Support"
        DefaultAssignee = "SPO Admin"
    }
    EmailConfig                   = @{
        SmtpServer = "mailfwd.nih.gov"
        NoReply    = "citspmail-noreply@nih.gov"
        From       = "NIH IC Admin Portal"
        #Admin      = "CITM365CollabDevOps@mail.nih.gov"
        Admin      = "ngan.bui@nih.gov"
        Support    = "ngan.bui@Nih.gov"        
    }
}