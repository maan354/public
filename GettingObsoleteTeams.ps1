#################
# Microsoft 365 CLI script to get obsolute Teams
# Created by Barry Bokdam
###################

[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]
    $outputfile
)

#Variables
$WarningDate = (Get-Date).AddDays(-90)
$Today = (Get-Date)
$ObsoleteGroups = 0

#Connect to Microsoft 365
$m365Status = m365 status

if ($m365Status -eq "Logged Out") {
  # Connection to Microsoft 365
  m365 login
}
#Get current user
$m365 = m365 status --ouput json | ConvertFrom-Json
$user = $m365.ConnectedAs

#Get all Microsoft Teams 
Write-Host "Getting all Microsoft Teams"
$teams = m365 teams team list --output json | ConvertFrom-Json

#Getting SharePoint audit log
Write-Host "Getting Auditlog for SharePoint"
$AuditRecs = m365 tenant auditlog report --contentType "SharePoint" --output json | ConvertFrom-Json

#Get the information
foreach ($team in $teams) {
    #Get the SharePoint URL
    Write-Host "Checking site " $team.displayName
        
    #Add user as owner
    #m365 teams user add --teamId $team.id --userName $user --role Owner
    
    #Get info
    $t = m365 teams team get --id $team.id --includeSiteUrl --output json | ConvertFrom-Json
    Write-Host "Teamnaam is " $t.displayName -ForegroundColor Yellow
    $SPUrl = $t.siteUrl
    Write-Host "SiteUrl is " $SPUrl
    if ($null -eq $SPUrl)
        {
        Write-Host "SharePoint has never been used for the group" $Team.DisplayName -ForegroundColor blue
        $ObsoleteGroups++   
        }
    #Check auditlog for SharePoint Activity 
    $SPUrl = $SPUrl + "/"
    $AuditRec = $AuditRecs | Where-Object {$_.siteUrl -like $SPUrl}
    If ($null -eq $AuditRec) 
    {
    Write-Host "No audit records found for" $team.displayName "-> It is potentially obsolete!" -ForegroundColor Green
    $output =  "No audit records found for" +  $team.displayName +  "-> It is potentially obsolete!"
    $ObsoleteGroups++   
    }
    Else 
        {
        Write-Host $AuditRec.Count "audit records found for " $team.displayname "the last is dated" $AuditRec.CreationTime[0] -ForegroundColor Yellow
        $output = $AuditRec.Count + "audit records found for " + $team.displayname + "the last is dated" + $AuditRec.CreationTime[0]
    }
    #Remove user from Microsoft Teams
    #m365 teams user remove --teamId $team.id -userName $user --confirm
    $output | Out-File -FilePath $outputfile  -Append
}

 
