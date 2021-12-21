Function Update_SNIncident() {
    param($IncidenttID,
        [ValidateSet("Provision", "Decommission", "Custom")]$IncidentType,
        [ValidateSet("SPO", "PowerBI")]$ServiceType = "SPO",
        [ValidateSet("Work In Progress", "Resolved")]$IncidentStatus,
        $SiteURL
    )
    try {
        switch ($IncidentType) {
            "Provision" {
                if ($IncidentStatus -eq "Resolved") {                    
                    $WorkLog = "This Site with URL: $SiteURL is now provisioned"
                    if ($ServiceType -eq "PowerBI"){
                        $WorkLog = "The Power BI Workspace with name : $SiteURL is now provisioned"
                    }
                    $Resolution = $WorkLog
                }
                else {
                    $WorkLog = "The request to provision a new site for the Site with URL: $($SiteURL) is being processed"
                    $Resolution = ""
                }
            }
            "Decommission" {
                if ($IncidentStatus -eq "Resolved") {
                    $WorkLog = "This Site with URL: $($SiteURL), is now Decommissioned"
                    $Resolution = $WorkLog
                }
                else {
                    $WorkLog = "The request to decommission a Site with URL: $($SiteURL), is being processed"
                    $Resolution = ""
                }
            }
            "Custom" {
            }
        }
        #This has be a non null value
        $TimeSpent = "30"

        #Create soap information
        $SOAPRequest = "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:urn=""urn:TIBCO_ITIL_ServiceDesk"">" +
        "   <soapenv:Header/>" +
        "   <soapenv:Body>" +
        "      <urn:OpSet>" +
        "         <urn:AuthenticationInfo>$($script:SNAuthInfo)</urn:AuthenticationInfo>" +
        "         <urn:Case_ID>" + $IncidenttID + "</urn:Case_ID>" +
        "         <urn:Group>" + $script:SNGroup + "</urn:Group>" +
        "         <urn:Assignee>" + $script:SNDefaultAssignee + "</urn:Assignee>" +
        "         <urn:Status>" + $IncidentStatus + "</urn:Status>" +
        "         <urn:Work_Log>" + $WorkLog + "</urn:Work_Log>" +
        "         <urn:Time_Spent>" + $TimeSpent + "</urn:Time_Spent>" +
        "         <urn:Incident_Resolution>" + $Resolution + "</urn:Incident_Resolution>" +
        "      </urn:OpSet>" +
        "   </soapenv:Body>" +
        " </soapenv:Envelope>"

        write-output $SOAPRequest

        #create XMLHTTP object for SOAP call
        $oXmlHTTP = new-object -comobject Microsoft.XMLHTTP;

        $oXmlHTTP.open("POST", $script:SNServcieUrl, $false); 

        $oXmlHTTP.setRequestHeader("Content-Type", "text/xml;charset=UTF-8");

        $oXmlHTTP.setRequestHeader("SOAPAction", "/RemedyService/RemedyServiceDesk.serviceagent/RemedyServiceDeskPortTypeEndpoint/OpSet");

        $oXmlHTTP.send($SOAPRequest);

        #create XMLDOM object for parsing XML result string
        $oXmlDoc = new-object -comobject Microsoft.XMLDOM 
        $oXmlDoc.async = "false"

        if ( ($oXmlHTTP.status -eq 200) -and ($oXmlHTTP.statusText -eq "OK") ) {
            if ($oXmlDoc.loadXML($oXmlHTTP.responsexml.xml)) {	
                $x = $oXmlDoc.getElementsByTagName("ns0:Case_ID");		
                #$x.length
                $caseID = "CaseID: " + $x.item(0).text;	   
                write-output $caseID
                LogWrite -Message " [Update_SNIncident]: Updated the status [$IncidentStatus] for [$IncidenttID]."
            }	
        }
        else {
            write-output $oXmlHTTP.status
            write-output $oXmlHTTP.ResponseText
            LogWrite -Level ERROR "[Update_SNIncident]: Updating SN status error: $($oXmlHTTP.status) - [$IncidenttID]"
        }
        
    }
    catch {        
        LogWrite -Level ERROR "[Update_SNIncident]: An error occured while updating SN ticket: $_"                    
    }
}

Function SendEmailConfirmation {
    param([Parameter(Mandatory = $true)] $Request)
    $header = @"
<style>    
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: #1F4E79;
        font-size: 12pt;

    }   
    
    table {
		font-size: 12pt;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
        width:100%
	}    
    td {
		padding: 6px;
		margin: 0px;
		border: 0;
	}
   span.labelHeader{
        font-weight:bold;
        color: #1F4E79;
   }

</style>
"@

    $subject = "[NIH IC Admin Portal]"
    $reqType = $Request.RequestType # {Decomission or Provision} 
    $templateId = $Request.TemplateId # {Team ; M365 Group, SharePoint Site or Power BI Workspace}
    $reqTitle =  $Request.Title
    $reqStatus = $Request.RequestStatus    
    $reqURL = $Request.URL
    $reqObjectId = $Request.ObjectId
    $wsURL = "$($script:PowerBIUrl)/groups/$reqObjectId"

    switch ($reqType){
        "Provision" {
            switch ($templateId){
                "Team" {                   
                    $preContent = "<p>Your request for a new Microsoft Team is now completed. Please allow up to 24 hours for new changes to propagate across the services in the tenant before using the site.</p>"
                    $content = $Request | ConvertTo-Html -PreContent $preContent -Property @{Label="<span class='labelHeader'>ServiceNow Ticket</span>";Expression={$_.IncidentId}},@{Label="<span class='labelHeader'>Site Name</span>";Expression={$_.SiteName}},@{Label="<span class='labelHeader'>URL</span>";Expression={$_.URL}},@{Label="<span class='labelHeader'>Organization</span>";Expression={$_.SiteICName}},@{Label="<span class='labelHeader'>External Sharing Enabled (for SharePoint Content)</span>";Expression={$_.ExternalSharing}} -as List -Head $header #-Fragment
                }
                "PowerBIWorkspace" {
                    $preContent = "<p>Your request for a new Power BI Workspace is now completed.</p>"
                    $content = $Request | ConvertTo-Html -PreContent $preContent -Property @{Label="<span class='labelHeader'>ServiceNow Ticket</span>";Expression={$_.IncidentId}},@{Label="<span class='labelHeader'>Workspace Name</span>";Expression={"<a href='$wsURL'>$($_.SiteName)</a>"}},@{Label="<span class='labelHeader'>Organization</span>";Expression={$_.SiteICName}} -as List -Head $header #-Fragment
                }
                default {
                    $reqTitle = ([uri]$reqURL).AbsolutePath
                    $reqTitle = "- ..$reqTitle"
                    $preContent = "<p>Your request for a new site is now completed. Please find the details below:</p>" 
                    $content = $Request | ConvertTo-Html -PreContent $preContent -Property @{Label="<span class='labelHeader'>ServiceNow Ticket</span>";Expression={$_.IncidentId}},@{Label="<span class='labelHeader'>Site Name</span>";Expression={$_.SiteName}},@{Label="<span class='labelHeader'>URL</span>";Expression={$_.URL}},@{Label="<span class='labelHeader'>Organization</span>";Expression={$_.SiteICName}},@{Label="<span class='labelHeader'>External Sharing Enabled (for SharePoint Content)</span>";Expression={$_.ExternalSharing}} -as List -Head $header #-Fragment
                }
            }  
        }
        "Decomission" {
            $preContent = "<p>Your request for site decommission has been processed successfully. Please allow up to 24 hours for the changes to propagate across the services in the tenant.</p>"
        }
    
    }
    $subject = "$subject - $reqType - $reqStatus $reqTitle" 
    $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
             $content
             <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
    $body = [System.Web.HttpUtility]::HtmlDecode($body)
    $to = "$($Request.CreatedBy)"
    if ($Request.SiteOwner -ne "" -and $Request.SiteOwner -ne $Request.CreatedBy){
        $to+=";$($Request.SiteOwner)"
    }
    if ($Request.SecondaryOwnerEmail -ne "" -and $Request.SecondaryOwnerEmail -ne $Request.CreatedBy){
            $to+=";$($Request.SecondaryOwnerEmail)"
        }
    SendEmail -subject $subject -body $body -To $to
}

Function SendEmailToExistingOwner {
    param([Parameter(Mandatory = $true)] $Request)
    $header = @"
<style>    
    p {

        font-family: Arial, Helvetica, sans-serif;
        color: #1F4E79;
        font-size: 12pt;

    }   
    
    table {
		font-size: 12pt;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
        width:100%
	}    
    td {
		padding: 6px;
		margin: 0px;
		border: 0;
	}
   span.labelHeader{
        font-weight:bold;
        color: #1F4E79;
   }

</style>
"@

    $subject = "[NIH IC Admin Portal]"
    $groupId = $Request.GroupId.Trim()
    $ownerEmail = $Request.NewValue.Trim()

    $content = "Please inform that $Cc was added to the group ABC as an owner upon request through IC Admin Portal."
    $body = "<p><i>Note: This is an automated email. Please do not reply to this message.</i></p>
             $content
             <p><i>Thank you,</i> <br />NIH M365 Collaboration Support Team</p>"    
    $body = [System.Web.HttpUtility]::HtmlDecode($body)
    
    $to = "$($Request.CreatedBy)"
    if ($Request.SiteOwner -ne "" -and $Request.SiteOwner -ne $Request.CreatedBy){
        $to+=";$($Request.SiteOwner)"
    }
    

    SendEmail -subject $subject -body $body -To $to -EnabledCc -Cc $cc
}