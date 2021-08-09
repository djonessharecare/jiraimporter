

function Get-HttpBasicHeader() {
    
    $creds = Get-Credential
    $pair = "$($creds.UserName):$(ConvertFrom-SecureString -SecureString $creds.Password -AsPlainText)" 
    $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
    $basicAuthValue = "Basic $encodedCreds"
    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]";
    $headers.Add("Authorization", $basicAuthValue );
    $headers.Add("X-Atlassian-Token", "nocheck");
    $headers.Add("Content-Type", "application/json; charset=utf-8");
    $headers.Add("Accept", "*/*");
   

    return $headers
	
}

Function Get-CSVFileName() {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    # $OpenFileDialog.initialDirectory = [Environment]::GetFolderPath('') 
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.Title = "Choose CSV Data File to Import";
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.Title = "Choose CSV Data File to Import";
    return $OpenFileDialog.filename
}

Function Get-JSONFileName() {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    #$OpenFileDialog.initialDirectory = [Environment]::GetFolderPath('') 
    $OpenFileDialog.filter = "JSON (*.json)| *.json"
    $OpenFileDialog.Title = "Choose JSON Mapping Config File";
    $OpenFileDialog.ShowDialog() | Out-Null
    
    return $OpenFileDialog.filename
}

Function Get-CSVHeaders($filename) {
    # Use System.IO.StreamReader to get the first line. Much faster than Get-Content.
    $StreamReader = New-Object System.IO.StreamReader -Arg $filename;
    $headerRow = $StreamReader.ReadLine();
    $StreamReader.Close()

    [array]$headers = $headerRow -Split "," | % { $_.Trim() } | ? { $_ };

    #for each header column
    for ($i = 0; $i -lt $headers.count; $i++) {

        if ($i -eq 0) { continue } #skip first column.

        # if in any previous column, give it a generic header name
        if ($headers[0..($i - 1)] -contains $headers[$i]) {
            $headers[$i] = $headers[$i] + $i
        }	
    }

    return $headers
}

Function HandleError($ErrorObject, $jsonPayload, $Message) {
    Write-Host $ErrorObject
    # Dig into the exception to get the Response details.
    # Note that value__ is not a typo.
    Write-Host $Message
    Write-Host "StatusCode:" $ErrorObject.Exception.Response.StatusCode.value__ 
    Write-Host "StatusDescription:" $ErrorObject.Exception.Response.StatusDescription
    if ($ErrorObject.ErrorDetails.Message) {
        Write-Host "Inner Error: " $ErrorObject.ErrorDetails.Message
    }
}

Function HTMLtoMarkUp($html) {
    $retVal = $html -replace "<img\s[^\>]*src\s*=\s*[""']([^/<]*/)*([^""']+)[^\>]*\s*/*>", '';
    $retVal = $retVal -replace '{html}', '';
    $TempFile = New-TemporaryFile;
    $retVal | Out-File -FilePath $TempFile.FullName -Encoding UTF8;
    $pandocArgs = @(
        "--from=html",
        "--to=markdown",
        "--quiet",
        "--wrap=none",
        "--strip-comments"
        $TempFile.FullName);
    $retVal = & $pandocExePath $pandocArgs  | Out-String;
    $retVal = $retVal -replace ('┬á', '')
    return $retVal;
}

Function GetAtlassianDocFormat($textValue) {
    $returnObject = [ordered]@{
        type    = "doc"
        version = 1
        content = @(
            @{
                type    = "paragraph"
                content = @(
                    @{
                        type = "text"
                        text = $textValue
                    }
                )
            })
    }

    return $returnObject 

}



Function ProcessField($parentChildren, $metaFldSchemaType, $metaFldItemType, $metaFldName, $ColVal, $IssueType) {

    $returnObject = $null;
    if ($metaFldName -eq "environment") {

        $returnObject = GetAtlassianDocFormat($ColVal.Value );
    }
    else {
        #projects/issuetypes/fields/[fieldname]/schema
        switch ($metaFldSchemaType) {
            "project" {
                $hash = @{ key = $ColVal.Value }
                $returnObject = [PSCustomObject]$hash
                break;
            }
            "issuetype" {
                $val = $ColVal.Value;
                if ($val -eq "Sub-task") { 
                    $val = "Sub-Task";
                }           
                $hash = @{ name = $val };
                $returnObject = [PSCustomObject]$hash;
                break;
            }
            "issuelink" {
                $found = $parentChildren | where IssueId -eq $pIssueId
                $hash = @{ id = $found.id }
                $returnObject = $hash 
                break;
            }
            "issuetype" { 
                $returnObject = @{name = $ColVal.Value }
                break;
            }
            "priority" { 
                $returnObject = @{name = $ColVal.Value }
                break;
            }
            "string" {
                $convertedValue = "";
                if ($ColVal.Value -match '{html}') {
                    $convertedValue = HTMLtoMarkUp($ColVal.Value); 
                    $convertedValue = $convertedValue.substring(0, [System.Math]::Min(20000, $convertedValue.Length) );
                }
                else {
                    $convertedValue = $ColVal.Value.substring(0, [System.Math]::Min(20000, $ColVal.Value.Length) ); 
      
                }
                switch ($metaFldName) {

                    "description" { 
                        $returnObject = GetAtlassianDocFormat($convertedValue);
                        break;
                    }
                    default {
                        $returnObject = $convertedValue;
                    }
                }
                break;
            }
            "comments-page" {
           
                #convert comment to markup
                $convertedValue = HTMLtoMarkUp($ColVal.Value);
                $convertedValue = $convertedValue.substring(0, [System.Math]::Min(30000, $convertedValue.Length) );
                $hash = @{ text = HTMLtoMarkUp($ColVal.Value) };
                $returnObject = [PSCustomObject]$hash
                break;
            }
            "number" {
                $returnObject = [int]$ColVal.Value;
                break;
            }
            "datetime" {
                $date = [DateTime]$ColVal.Value;
                $returnObject = '{0:yyyy-MM-dd}' -f $date
                break;
            }
            "date" {
                $date = [DateTime]$ColVal.Value;
                $returnObject = '{0:yyyy-MM-dd}' -f $date
                break;
            }
            "array" {
                switch ($metaFldItemType) {
                    "string" {
                        $returnObject = $ColVal.Value;
                        break;
                    } 
                    "user" {
                        $returnObject = GetUserAccount  $ColVal.Value.Trim()
                        break;
                    } 
                    "version" {
                        #check if version exists and if not create it.
                        $exists = $fldMetaVersions | Where-Object { $_."name" -eq $ColVal.Value -and $_."issuetype" -eq $IssueType }
                        if (!$exists) {
                            #create it
                            $body = @{
                                name      = $ColVal.Value
                                projectId = $fldMetaProjectId
                            }
                            try {
                                $utfbody = [System.Text.Encoding]::UTF8.GetBytes( ($body | ConvertTo-Json -depth 10 -compress)) 
                                $newVersionResponse = Invoke-RestMethod -uri $restapiCreateVersion  -Headers $httpheaders -Method POST  -body $utfbody
                                $verHash = @{
                                    issuetype = $IssueType
                                    name      = $ColVal.Value
                                };
                                $fldMetaVersions.Add($verHash);
                            }
                            catch {
                                Write-Host "Unable to Create Version: {0}" -f  $ColVal.Value
                            }
                            
                        }
                        $hash = @{name = $ColVal.Value } 
                        $returnObject = [PSCustomObject]$hash


                    } 
                    # "component" {} 
                    # "attachment" {} 
                    # "json" {}
                    default {
                        $hash = @{name = $ColVal.Value } 
                        $returnObject = [PSCustomObject]$hash
                    }
                }
            

                break;
            }
            "user" {
                #try to look up the user if not found then set to default.
                $returnObject = GetUserAccount  $ColVal.Value.Trim()
                break;
            } 
            "option" {
                #try to look up the option 
                break;

            }
            "any" { 
                $returnObject = $ColVal.Value;
                break;
            }
            default {
                $returnObject = $ColVal.Value;
                break;
            }
        }
    }
    return $returnObject;
    
}

Function GetUserAccount($userSearchVal) {
    $assignableUserInfo = $null;
    #Get userids for project
    $retVal = $null;
    try {
        $fullURI = $restapiUserLookupByUsernameuri + "?query=" + $userSearchVal;
        $UserInfo = Invoke-RestMethod -uri $fullURI  -Headers $httpheaders -Method GET 
        if ($UserInfo.Length -gt 0) {
            $fullURI = "{0}?project={1}&accountId={2}" -f $restapiAssignableLookup, $projectKey, $UserInfo.accountId  
            $AssignableUserInfo = Invoke-RestMethod -uri $fullURI   -Headers $httpheaders -Method GET 
        }   
    }
    catch {
        handleerror($_, $null, "Failed to retrieve assignable users for user. Check auth creds"); 
    }
    if ($assignableUserInfo) {
        $hash = @{accountId = $assignableUserInfo.accountId } 
        $retVal = $hash ;
    }
    return $retVal;
}

Function GetJiraFieldFromMappingConfig($csvColName) {

    $mappingConfigFldObj = $null;
    $mappingConfigJiraField = $null;
    $mappingConfigCustomField = $null;
    $csvColName = $csvColName -replace '"', ''
    $csvColName = $csvColName -replace '\d+$', ''

    $property = $MappingConfigProps | where Name -like $csvColName 
    
    if ($property) {
        $mappingConfigFldObj = [PSCustomObject]$property.Value
        $mappingConfigJiraField = $mappingConfigFldObj."jira.field";
        $mappingConfigCustomField = $mappingConfigFldObj."existing.custom.field";  
    }
    #not a standard jira field is it a custom field
    if (!$mappingConfigJiraField -and $mappingConfigCustomField) {
        #if we cannot find it in the config pass back the csv column name
        $mappingConfigJiraField = "customfield_{0}" -f $mappingConfigCustomField;
    }
    if ($csvColName -eq "Parent id") { 
        $mappingConfigJiraField = "Parent";
    }
    return $mappingConfigJiraField;

    
}

Function  GetMetaDataForField($JiraField, $fldMetaIssueTypes, $csvRowIssueType) {
    #Get Metadata for Fields
    
    $returnObject = $null
    $metaFldKey = $null;
    $metaFldSchemaType = $null;

    $fmit = $fldMetaIssueTypes | where name -eq $csvRowIssueType 
    #$property.Value.name
    #Check the freshly gotten metadata to ensure config has not changed
    $fldMeta = $fmit.fields.$JiraField;
    if (!$fldMeta) {
        $searchVal = $JiraField #-replace ' ', '';
        $wildcardSearchVal = $searchVal.TrimEnd("s");
        $wildcardSearchVal = $searchVal.TrimEnd("\'");
        $wildcardSearchVal = $wildcardSearchVal + "*"
        $fldMeta = $fmit.fields.PsObject.Properties | where Value.key -eq $searchVal
        $fldMeta = $fmit.fields.PsObject.Properties | where Value.key -like $wildcardSearchVal
    }   
    else {
        $metaFldKey = $fldMeta.key;
        if ($fldMeta.schema) {
            $metaFldSchemaType = $fldMeta.schema.type.Length -gt 0?$fldMeta.schema.type:"string";
            $metaItemsType = $fldMeta.schema.items;
        }

        $hash = @{
            SchemaType = $metaFldSchemaType
            FieldKey   = $metaFldKey
            ItemsType  = $metaItemsType
        } 
        $returnObject = [PSCustomObject]$hash
    }
    return $returnObject
     
}

Function ProcessRow($csvRow) {
    #process csv row
    $csvRowIssueType = $_."Issue Type";

    $issueTypeMetaData = $JiraFieldInfo.projects.issuetypes | Where-Object name -eq $csvRowIssueType
    $verList = $issueTypeMetaData.fields.versions.allowedvalues |  Select-Object name;
    foreach ($versionName in $verList | Select-Object name ) {
        $verHash = @{
            issuetype = $csvRowIssueType
            name      = $versionName.name
        }
        $exists = $fldMetaVersions | Where-Object { $_."name" -eq $versionName.name -and $_."issuetype" -eq $csvRowIssueType }
        if (!$exists) {
            $fldMetaVersions.Add($verHash);
        }
    }

    #build the request body
    $returnObject = [ordered]@{
        update = @{}
        fields = @{
            project      = @{
                key = $projectKey
            }
            timetracking = @{
                remainingEstimate = $_."Remaining Estimate";
                originalEstimate  = $_."Original Estimate";
            }
        }
    }
 
    #process csv cols
    foreach ($csvCol in $csvRow.PsObject.Properties  | where Value | where Name -notmatch "comment") {
        #Take the import config mapping information and match fields with the CSV import
        # $object_properties.Name];
        #skip custom fields and comments so they can be last.

        try {


            $csvColName = $csvCol.Name -replace '"', ''
            $jiraField = GetJiraFieldFromMappingConfig($csvColName);
            if (!$jiraField) { continue; }
            $metaObject = GetMetaDataForField $jiraField $fldMetaIssueTypes $csvRowIssueType
            if (!$metaObject) { continue; }
            #ProcessField($parentChildren, $metaFldSchemaType, $metaFldName, $ColVal)
            $retFld = ProcessField  $parentChildList  $metaObject.SchemaType $metaObject.ItemsType $metaObject.FieldKey $csvCol  $csvRowIssueType
            #Check if the field already exists.  If so it is an array and just add it. 
            $currentField = $returnObject.fields.[string]$metaObject.FieldKey;
            if ($metaObject.SchemaType -eq "array") {
                if (!$currentField) {
                    $currentField = [System.Collections.ArrayList]::new();
                }
                if ($retFld.GetType().BaseType.Name -eq "Array") {
                    $retFld = $retFld[($retFld.Length - 1)]; 
                }
                $currentField.Add($retFld); 

            }  
            else {
                $currentField = $retFld;
            }
            $returnObject.fields.[string]$metaObject.FieldKey = $currentField;

        }
        catch {
            handleerror($_, $null, "Failed to process column"); 

        }
    }
    return [pscustomobject]$returnObject;
}

Function CreateJiraItems($csvRow, $issueid, $pissueid, $jsonBodyObject) {
    #send the payload to create the issue in jira
    #$output = ($jsonBodyObject | ConvertTo-Json -depth 100  -compress) 
    #$output | Out-File "C:\Users\douglas.jones\pandoc\jsonresultsnotworking.json"
    if ($jsonBodyObject.GetType().BaseType.Name -eq "Array") {
        $jsonBodyObject = $jsonBodyObject[($jsonBodyObject.Length - 1)]; 
    }
    $body = [System.Text.Encoding]::UTF8.GetBytes( ($jsonBodyObject | ConvertTo-Json -depth 50 -compress)) 
    try {
        $result = invoke-restmethod -uri $restapiuri -headers $httpheaders -method post  -body $body
        #add issue id and parent issue id to the array for use later
        $result | add-member -membertype noteproperty -name 'issueid' -value $issueid
        $result | add-member -membertype noteproperty -name 'parentissueid' -value $pissueid
        $parentchildlist.add($result);
        #start-sleep -s .1

        AddComments $result.key $csvRow;
        DoTransitions $result.key $csvRow;
        

    
    
    }
    catch {
        handleerror($_, ($jsonBodyObject | convertto-json -depth 50 -compress), "Failed to create issue in Jira"); 
            
    }
}


Function AddComments($issueKey, $row) {
    #Add Comments

    $comments = [System.Collections.ArrayList]::new();
    foreach ($csvCol in $row.PsObject.Properties  | where Value | where Name -match "comment" | Select-Object -First 10 ) {
        #if a comment add comment to comments list to add after creation.
        #convert comment to markup
        $commentMarkup = HTMLtoMarkUp($csvCol.Value);
        $comments.Add($commentMarkup);
    }
    try {
        #issue has to be created before adding comments.  
        $commenturi = $restapicommenturi + $issueKey + "/comment"
        #write-host "commenturi:" $commenturi;
                
        $comments | foreach-object {
            $commentbody = @{ body = $_ }
            $commentbody = [System.Text.Encoding]::UTF8.GetBytes( ($commentbody | ConvertTo-Json -depth 10 -compress)) 
            #$commentbody = ([pscustomobject]$commentbody | convertto-json -depth 10 -compress)
            invoke-restmethod -uri $commenturi  -headers $httpheaders -method post  -body $commentbody
        }
    }
    catch {
        handleerror($_, ([pscustomobject]$commentbody | convertto-json -depth 10 -compress), "Failed to create comment in Jira"); 
    }
}

Function DoTransitions($issueKey, $row) {
    try {
        #There is a separate workflow for Epic and other types
        $issueType = $_."Issue Type";
        $availableIssueTransitions = $null;
        #get available transition

        $availableIssueTransitions = GetIssueTransitions $issueKey
        switch ($row."status") {
            { ($_ -like "Done*") -or ($_ -like "Closed*") -or ($_ -like "Resolved*") } {
                #go to resolved first
                $trans = $availableIssueTransitions.Transitions | Where name -like "Resolved*" | Select-Object   
                $availableIssueTransitions = DoTransition $issueKey $trans.id "true"
                if (($_ -like "Done*") -or ($_ -eq "Closed*")) {
                    $trans = $availableIssueTransitions.Transitions | Where name -like "Closed*" | Select-Object  
                    DoTransition $issueKey $trans.id
                }
                break;
            } 
            { ($_ -like "In Progress*") -or ($_ -like "In Development*") -or ($_ -like "Stakeholder Approval*") -or ($_ -like "Ready for Testing*") } {
                $trans = $availableIssueTransitions.Transitions | Where name -like "Stakeholder Verification*" | Select-Object 
                $availableIssueTransitions = DoTransition $issueKey $trans.id "true"
                if (($_ -like "In Progress*") -or ($_ -like "In Development*")) {
                    $trans = $availableIssueTransitions.Transitions | Where name -like "Start Progress*" | Select-Object  
                    DoTransition $issueKey $trans.id
                }
                if (($_ -like "Ready for Testing*") ) {
                    $trans = $availableIssueTransitions.Transitions | Where name -like "QA Testing*" | Select-Object 
                    $availableIssueTransitions = DoTransition $issueKey $trans.id "true"

                }
                break;
            }
            default {
            }          
        }

        if ($issueType -eq "Epic") {
            #Create Issue goes to Open  
            #Open goes to In Progress
            #Open goest to Stakeholder Review
            #Open goes to resolved
            #resolved goes to closed
            #stakeholder review goes to Deployment Verification
            #Deployment Verification goes to Resolved
            #Deployment Verification goes to In Progress
            #stakeholder review goes to ready for stage
            #Ready for Stage goes to Deployed to Stage

        }

        #sharecare statuses
        #Non Epic Statuses - Open, In Progress, Stakeholder Review, QA Testing, Closed, Resolved, Deployment Verification
        #Epic Statuses - Open, Deployed to Stage, Ready for Stage, In Progress, Stakeholder Review, Closed, Resolved, Deployment Verification      
    }
    catch {
        handleerror($_, ([pscustomobject]$transition | convertto-json -depth 10 -compress), "Failed to execute transition in Jira"); 
    }
}

Function GetIssueTransitions($issueKey) {
    try {
        $transuri = $restapitransitionuri + $issueKey + "/transitions"
        return invoke-restmethod -uri $transuri -headers $httpheaders -method GET
    }
    catch {
        handleerror($_, $null, "Failed to execute transition in Jira"); 
    }


}

Function DoTransition($issueKey, $transitionid, $getTransitions ) {
    try {

        $transition = @{
            transition = @{
                id = $transitionid 
            }
        }

        $transuri = $restapitransitionuri + $issueKey + "/transitions?notifyUsers=false"
        $transbody = ([pscustomobject]$transition | convertto-json -depth 10 -compress)
        invoke-restmethod -uri $transuri -headers $httpheaders -method post  -body $transbody
        if ($getTransitions -eq "true") {
            return GetIssueTransitions $issueKey ;
        }
    }
    catch {
        handleerror($_, ([pscustomobject]$transition | convertto-json -depth 10 -compress), "Failed to execute transition in Jira"); 
    }

}




#Get Jira Configuration Files
$MappingConfigFileName = Get-JSONFileName;
$MappingConfigJSON = Get-Content -Raw -Path $MappingConfigFileName | ConvertFrom-Json
$project = [PSCustomObject]$MappingConfigJSON."config.project";
$projectKey = [PSCustomObject]$project."project.key";
$dateFormat = $MappingConfigJSON."config.date.format";
$MappingConfigProps = $MappingConfigJSON."config.field.mappings".psobject.properties

$restapiBaseUri = "[Put your base uri here]"
$restapiuri = "{0}/rest/api/3/issue?updateHistory=false&notifyUsers=false" -f $restapiBaseUri;
$restapicommenturi = "{0}/rest/api/2/issue/" -f $restapiBaseUri;
$restapitransitionuri = "{0}/rest/api/3/issue/" -f $restapiBaseUri;
$restapifieldsuri = "{0}/rest/api/3/issue/createmeta?projectKeys={1}&expand=projects.issuetypes.fields" -f $restapiBaseUri, $projectKey;
$restapiUserLookupByUsernameuri = "{0}/rest/api/3/user/search" -f $restapiBaseUri;
$restapiAssignableLookup = "{0}/rest/api/3/user/assignable/search" -f $restapiBaseUri;
$restapiCreateVersion = "{0}//rest/api/3/version" -f $restapiBaseUri;
$httpheaders = Get-HttpBasicHeader
$parentChildList = [System.Collections.ArrayList]::new()
$pandocExePath = 'C:\ProgramData\chocolatey\bin\pandoc.exe'


#Field Information from the destination jira system
$JiraFieldInfo = $null;
$fldMetaProjects = $null;
$fldMetaIssueTypes = $null;
$fldMetaProjectId = $null;
$fldMetaVersions = [System.Collections.ArrayList]::new(); ;
#Get Fields Meta Data For Project
try {
    $JiraFieldInfo = Invoke-RestMethod -uri $restapifieldsuri  -Headers $httpheaders -Method GET 
    $fldMetaProjects = $JiraFieldInfo.projects;
    
    if ($fldMetaProjects.Length -eq 0) {
        Write-Host "Failed to retrieve metadata for project. Check auth creds.  Exiting Program";
        exit;
    }
    $fldMetaIssueTypes = $fldMetaProjects.issuetypes;
    if ($fldMetaIssueTypes.Length -eq 0) {
        Write-Host "Retrieved project metadata but there were no IssueTypes returned.  Exiting Program";
        exit;
    }
    $fldMetaProjectId = $fldMetaProjects.id;

    

}
catch {
    handleerror($_, $null, "Failed to retrieve metadata for project. Check auth creds.  Exiting Program" ); 
    exit;
}


#Get CSV Data File
$fname = Get-CSVFileName;
$headers = Get-CSVHeaders($fname);
$csvFileData = $null;
try {
    $csvFileData = Import-Csv $fname -Header $headers  |  Select-Object -skip 1 
}
catch {
    Write-Host "Unable to load the CSV File.  Is the file open?  Exiting Program";
    exit;
}

if (!$csvFileData) {
    Write-Host "CSV FileData is null.  Exiting Program";
    exit;
}

#Make a request for each row to jira
$epics = $csvFileData | Where-Object { $_."Issue Type" -eq "Epic" } | Select-Object;
$issues = $csvFileData | Where-Object { $_."Issue Type" -ne "Epic" -and $_."Issue Type" -ne "Sub-task" } | Select-Object;
$subtasks = $csvFileData | Where-Object { $_."Issue Type" -eq "Sub-task" }  | Select-Object;
$epics | ForEach-Object {
    $issueId = $_."Issue id";
    $pIssueId = $_."Parent id";
    CreateJiraItems $_ $issueid  $pissueid (ProcessRow $_)  
}

$issues | ForEach-Object {
    $issueId = $_."Issue id";
    $pIssueId = $_."Parent id";
    CreateJiraItems $_ $issueid  $pissueid (ProcessRow $_)  
}

$subtasks | ForEach-Object {
    $issueId = $_."Issue id";
    $pIssueId = $_."Parent id";
    CreateJiraItems $_ $issueid  $pissueid (ProcessRow $_)  
}




	
	
	
	











