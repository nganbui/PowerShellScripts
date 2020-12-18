@{
    Path         = @{        
        Cred = "D:\Scripting\O365DevOps\Common\Creds"
    }
    #----------Legacy scripting D:\Scripting\O365-------------#
    AppConfig    = @{
        AppId      = 'e915ab9f-d0d5-4448-b448-b8b025338098'        
        AppSecret  = 'O365AppPwd' 
        AppName    = "SPO - NIH Collaboration Admin Portal"      
        TenantName = "14b77578-9773-42d5-8507-251ca2dc2b06"      
    }
    #----------OAuth-new file structure scripting D:\Scripting\O365DevOps--------#          
    AppConfigOperationSupport    = @{
        AppId      = '9624e216-9e73-4513-9251-4d4382950420'        
        Thumbprint = '1C9696EB9152228A42DAEB5C7075699795311662' 
        AppName    = "SPO-M365 Operations Support"
        CertName   = "SPO-M365 Operations Support"
    }
    AppConfigAdminPortalOperation    = @{
        AppId      = '497e07ac-d6f7-4d40-9d70-54ebb507ef39'        
        Thumbprint = '7BA7CBA81EDC57BF8446C549294148FB8490AD5B' 
        AppName    = "SPO-Sync Operations"
        CertName   = "SPO-Sync Operations"
    }
    AppConfigEXOV2    = @{
        AppId      = '1518f6e9-143b-4162-be08-db5eb9f78a28'        
        Thumbprint = 'EAADC1737F06BB0104C8E6932EB814AB0CD802E6'        
        AppName    = "SPO-GuestUsersMembershipReport"
    }
    AppConfigMC  = @{
        AppId      = '0f6a6412-8f62-4b7a-a929-664174baf961'        
        AppSecret  = 'O365AppMCPwd'
        Resource   = "https://manage.office.com"   
        AppName    = "SPO - Message Center and Health Reader"
    }
    #----------Tenant Config--------#      
    TenantConfig = @{
        Id   = "14b77578-9773-42d5-8507-251ca2dc2b06"
        Name = "nih.onmicrosoft.com"
        AdminCenterUrl    = "https://nih-admin.sharepoint.com"
        RootSiteUrl       = "https://nih.sharepoint.com"
        O365TenantAdmin   = "spoadmsvc@nih.gov"                          
    }
    #----------DB Config--------# 	    
    DBConfig     = @{
        DBServer  = "NIHSPSQL16D2"
        DBName    = "O365_SelfServicePortal_UAT"
        DBUser    = "admportal-uat"
        DBPwdFile = "O365DBPwd"
        <#DBServer  = "NIHSPSQL19P1"
        DBName    = "O365_SelfServicePortal"
        DBUser    = "O365_Sql_Admin"
        DBPwdFile = "O365DBPwd"#>
    }
    #----------Email Config--------#    
    EmailConfig  = @{
        SmtpServer = "mailfwd.nih.gov"
        NoReply    = "citspmail-noreply@nih.gov"
        <#Admin      = "CITHSSSharePointOnlineAdmins@mail.nih.gov"#>
        Admin      = "ngan.bui@Nih.gov"
        Support    = "ngan.bui@Nih.gov"        
    }
}