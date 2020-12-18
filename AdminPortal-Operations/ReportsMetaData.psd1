@{ 
    UsageReport = ( 
         @{
             ReportEndpoint = "getTeamsUserActivityUserDetail"
             ReportName     = "TeamsUserActivity"
         },
         @{
             ReportEndpoint = "getTeamsDeviceUsageUserDetail"
             ReportName     = "TeamsDeviceUsage"
         },
         @{
            ReportEndpoint = "getSharePointSiteUsageDetail"
            ReportName     = "SharePointSiteUsage"
        },
        @{
            ReportEndpoint = "getSharePointActivityUserDetail"
            ReportName     = "SharePointActivity"
        },
        @{
            ReportEndpoint = "getOneDriveUsageAccountDetail"
            ReportName     = "OneDriveUsage"
        },
        @{
            ReportEndpoint = "getOneDriveActivityUserDetail"
            ReportName     = "OneDriveActivity"
        },
        #manually download
        @{
            ReportEndpoint = ""            
            ReportName     = "TeamsUsageActivity"
            DropLocation = "D:\Reports\TeamsUsageActivity.csv"            
        },
        @{
            ReportEndpoint = ""            
            ReportName     = "FormsUserActivity"
            DropLocation = "D:\Reports\FormsUserActivity.csv"            
        }, 
        @{
            ReportEndpoint = ""            
            ReportName     = "PSTNCallUsage"
            DropLocation = "D:\Reports\PSTNCallUsage.csv"            
        },
        @{
            ReportEndpoint = ""            
            ReportName     = "TeamsOnlyUsers"
            DropLocation = "D:\Reports\TeamsOnlyUsers.csv"
        },       
        @{
            ReportEndpoint = ""            
            ReportName     = "BoxUsage"
            DropLocation = "D:\Reports\BoxUsage.csv"
        }
                     
    )
    ReportConfig = ( 
         @{
             StoredProc = "Reports_Usage"
             FileName   = "Reports_UsageReport"
         },
         @{
             StoredProc = "Reports_TeamsUserActivity"
             FileName   = "Reports_TeamsUserActivity"
         },
         @{
             StoredProc = "Reports_TeamsDeviceUsage"
             FileName   = "Reports_TeamsDeviceUsage"
         },
         @{
             StoredProc = "Reports_TeamsUsageActivity"
             FileName   = "Reports_TeamsUsageActivity"
         }, 
         @{
             StoredProc = "Reports_PSTNCallUsage"
             FileName   = "Reports_PSTNCallUsage"
         },        
         @{
             StoredProc = "Reports_SharePointActivity"
             FileName   = "Reports_SharePointActivity"
         },
         @{
             StoredProc = "Reports_SharePointSiteUsage"
             FileName   = "Reports_SharePointSiteUsage"
         },
         @{
             StoredProc = "Reports_OneDriveActivity"
             FileName   = "Reports_OneDriveActivity"
         },
         @{
             StoredProc = "Reports_OneDriveUsage"
             FileName   = "Reports_OneDriveUsage"
         },
          @{
             StoredProc = "Reports_FormsUserActivity"
             FileName   = "Reports_FormsUserActivity"
         },
         @{
             StoredProc = "Reports_BoxUsage"
             FileName   = "Reports_BoxUsage"
         },
         @{
             StoredProc = "Reports_Guests"
             FileName   = "Reports_Guests"             
         },
         @{
             StoredProc = "Reports_TeamsOnlyModeUsersByIC"
             FileName   = "Reports_TeamsOnlyMode"
             Baseline   = "Yes"
         },
         @{
             StoredProc = "Reports_Teams"
             FileName   = "Reports_Teams"
             Baseline   = "Yes"
         },
         @{
             StoredProc = "Reports_Groups"
             FileName   = "Reports_Groups"
             Baseline   = "Yes"
         },
         @{
             StoredProc = "Reports_Users"
             FileName   = "Reports_Users"
             Baseline   = "Yes"
         },                  
         @{
             StoredProc = "Reports_PersonalSites"
             FileName   = "Reports_PersonalSites"
             Baseline   = "Yes"
         }           
    )      
}
