####
# Microsoft 365 CLI script to get obsolute Teams
# Created by Barry Bokdam
#####


#Variables
$WarningDate = (Get-Date).AddDays(-90)
$Today = (Get-Date)
$ObsoleteGroups = 0

#Connect to Microsoft 365
m365 login

#Get all Microsoft Teams 
Write-Host "Getting all Microsoft Teams"
$teams = m365 teams team list --output json
$teams = $teams | ConvertFrom-Json

#Getting SharePoint audit log
Write-Host "Getting Auditlog for SharePoint"
$AuditRecs = m365 tenant auditlog report --contentType "SharePoint" --output json
$AuditRecs = $AuditRecs | ConvertFrom-Json


#Get the information
foreach ($team in $teams) {
    #Get the SharePoint URL
    Write-Host "Checking site " $team.displayName
    #Check of the SharePoint sites exist
    $t = m365 teams team get --id $team.id --includeSiteUrl --output json
    $t = $t | ConvertFrom-Json
    $SPUrl = $t.siteUrl
    if ($null -eq $SPUrl)
        {
        Write-Host "SharePoint has never been used for the group" $Team.DisplayName -ForegroundColor blue
        $ObsoleteGroups++   
        }
    #Check auditlog for SharePoint Activity 
    $SPUrl = $SPUrl + "/*"
    $AuditRec = $AuditRecs | Where-Object {$_.ObjectId -like $SPUrl}
    If ($null -eq $AuditRec) 
    {
    Write-Host "No audit records found for" $team.displayName "-> It is potentially obsolete!" -ForegroundColor Green
    $ObsoleteGroups++   
    }
    Else 
        {
        Write-Host $AuditRec.Count "audit records found for " $team.displayname "the last is dated" $AuditRec.CreationTime[0] -ForegroundColor Yellow
    }

}

 
