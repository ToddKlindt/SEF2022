break
# Things to install
# Use PowerShell Admin Console and -Force –Scope AllUsers
Install-Module PnP.PowerShell -Force -Scope AllUsers
Install-Module Microsoft.Online.SharePoint.PowerShell -Force -Scope AllUsers
Install-Module MicrosoftTeams -Force -Scope AllUsers

# Get modules installed in the current console
Get-InstalledModule

# Get all of the modules installed on the system
Get-Module -ListAvailable

# Update a module
Update-Module Microsoft.Online.SharePoint.PowerShell -Force

# Add a stored credential
Add-PnPStoredCredential -Name "https://contoso.sharepoint.com" -Username yourname@contoso.onmicrosoft.com
Connect-PnPOnline -Url "https://contoso.sharepoint.com”
Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/hr”

# Use a stored credential outside of the PnP
$Credentials = Get-PnPStoredCredential -Name "https://contoso.sharepoint.com" 
Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credentials $Credentials

# Find the Provisioning cmdlets
Get-Command -Module pnp.powershell -Noun "*template*"

# Multiple versions of the same module
Find-Module -Name pnp.powershell -AllVersions -AllowPrerelease
Install-Module -RequiredVersion
Install-Module -AllowClobber
Install-Module -Name pnp.powershell -Scope AllUsers -AllowClobber -Force
Install-Module -Name pnp.powershell -Scope AllUsers -AllowClobber -Force -RequiredVersion 1.9.0

# Example
Import-Module -Name PnP.PowerShell -RequiredVersion 1.9.0 -Prefix Old

# Upload a file
$web = https://tenant.sharepoint.com/sites/hr
$folder = "Shared Documents"
Connect-PnPOnline -Url $web
Add-PnPFile -Path '.\Boot fairs with Graphic design.docx' -Folder $folder

# Add a folder
Add-PnPFolder -Name "Folder 1" -Folder $folder
Add-PnPFile -Path '.\Building materials licences to budget for Storytelling.docx'  -Folder '$folder\Folder 1'

# Get all the files in all the document libraries
$docliblist = Get-PnPList -Includes DefaultViewUrl,IsSystemList | Where-Object -Property IsSystemList -EQ -Value $false | Where-Object -Property BaseType -EQ -Value "DocumentLibrary"
    Foreach ($doclib in $docliblist) 
        {
        $doclist = Get-PnPListItem -List $DocLib
        foreach ($doc in $doclist) {
            if ($null -ne ($doc.FieldValues).SharedWithUsers) {
                foreach ($user in (($doc.FieldValues).SharedWithUsers))  {
                    Write-Output "$(($doc.FieldValues).FileRef) - $($user.email)"
                    }
                }
             }
         } 

# Get extended object properties
$docliblist = Get-PnPList -Includes DefaultViewUrl,IsSystemList | Where-Object -Property IsSystemList -EQ -Value $false | Where-Object -Property BaseType -EQ -Value "DocumentLibrary“

    Foreach ($doclib in $docliblist) 
        {
            $doclibTitle = $doclib.Title
            $doclist = Get-PnPListItem -List $DocLib
            $doclist | ForEach-Object {  Get-PnPProperty -ClientObject $_ -Property File, ContentType, ComplianceInfo}
            foreach ($doc in $doclist) {
                [pscustomobject]@{
                    PSTypeName = 'TKPnPFile'
                    Library= $doclibTitle
                    Filename = ($doc.File).Name
                    ContentType = ($doc.ContentType).Name
                    Label = ($doc.ComplianceInfo).ComplianceTag
                }  
            }
        }

# Bulk undelete a lot of files
Connect-PnPOnline -Url https://sadtenant.sharepoint.com/ -Credentials SadTenantAdmin

# Get all of the deleted files matching our criteria
$bin = Get-PnPRecycleBinItem | Where-Object -Property Leafname -Like -Value "*.jpg" | Where-Object -Property Dirname -Like -Value “Important Photos/Shared Documents/*" | Where-Object -Property DeletedByEmail -EQ -Value baduser@sadtenant.phooey

# See the bad news
$bin.count

# Restore them all, writing out their name and the time at the end
$bin | ForEach-Object -begin { $a = 0} -Process {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }

# Do it in batches of 10,000
($bin[20001..30000]) | ForEach-Object -begin { $a = 0} -Process {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }

# Create a new site
# SPO
New-SPOSite # No group, can be groupified later

# PnP
New-PnPSite -Type TeamSite -Title "Modern Team Site" -Alias ModernTeamSite -IsPublic # Group, no Team. Can be Teamified

# Teams module
Connect-MicrosoftTeams -Credential (Get-PnPStoredCredential -Name "https://contoso-admin.sharepoint.com")
New-Team -DisplayName "Fancy Group" -Description "Fancy Group made by PowerShell?" -Alias FancyGroup -AccessType Public # All things

# Save site as template
Get-PnPSiteTemplate -Out customer.xml
Add-PnPListFoldersToSiteTemplate -Path customer.xml -List 'Data Storage' -Recursive
Invoke-PnPSiteTemplate -Path customer.xml -Handlers Lists, SiteSecurity

Get-Command -Module PnP.PowerShell -Name *temp*

# Get all the Flows
# Requires MSOL module and connection
Add-PowerAppsAccount
Get-AdminFlow | ForEach-Object { $ownername = (Get-MsolUser -ObjectId $_.CreatedBy.userId).DisplayName ; $owneremail = (Get-MsolUser -ObjectId $_.CreatedBy.userId).UserPrincipalName ; Write-Host $_.DisplayName, $ownername, $owneremail }

# Get your Flows
Get-Flow

# Get PowerApps
Get-PowerApp

# Disable Flow Button
# SPO Method
Connect-SPOService -Url https://flowhater-admin.sharepoint.com
$val = [Microsoft.Online.SharePoint.TenantAdministration.FlowsPolicy]::Disabled 
Set-SPOSite -Identity https://flowhater.sharepoint.com/sites/SadSite -DisableFlows $val

# PnP Method
Connect-PnPOnline -Url https://flowhater.sharepoint.com/sites/SadSite -Credentials 'MeanSiteAdmin'
Set-PnPSite -DisableFlows:$true

# From https://github.com/ToddKlindt/PowerShell/blob/master/TKM365Commands.psm1
# Hit the graph
Get-TKPnPGraphURI -uri https://graph.microsoft.com/v1.0/me/
Get-TKPnPGraphURI -uri https://graph.microsoft.com/v1.0/users | select displayName,userPrincipalName,id

# Get Current user
Get-TKPnPCurrentUser
Get-TKPnPCurrentUser -UseGraph

# Test PnP Connection
while (! $a) {
    
    try {
        $a = Connect-PnPOnline -Url https://tenant.sharepoint.com/sites/foo -erroraction stop
    }
    catch {
        Write-Warning "No connection"
        Write-Warning "Waiting 30 seconds..."
        Start-Sleep -Seconds 30
       }   
    }

# As a PS1 file
# Try to connect to a site
param($Site, $Retries, $TimeBetween)
$Connected = $false

while (($Retries -gt 0) -and (-not $Connected)) {
    if (Get-PnPTenantSite -Identity $Site) { 
        $Connected = $true
    } else {
        $Retries = $Retries - 1
        Start-Sleep -Seconds $TimeBetween
    }
}    

# Check to see if the Module is installed
if (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue  ) {
    Write-Host "Found it" 
    } else {
        Write-Warning "Couldn't PnP PowerShell module"
        Write-Warning "Install it with:"
        Write-Warning "Install-Module PnP.PowerShell"
        break
    }

# use this to time things
$start = Get-Date
# Do stuff here
# Get end time
$end = Get-Date
$totaltime = $end - $start
Write-Host "`nTime elapsed: $($totaltime.tostring("hh\:mm\:ss"))"

# usage
Connect-PnPOnline -Url https://1kgvf-admin.sharepoint.com
$SiteList = Get-PnPTenantSite
# Get each site's storage usage
$SiteList = Get-PnPTenantSite # -IncludeOneDriveSites
[System.Collections.ArrayList]$out = @()
foreach ($Site in $SiteList) {
    $tempout = [PSCustomObject]@{
        PSTypeName = 'TKPnPSite'
        URL     = $Site.url
        StorageUsageCurrent = $Site.StorageUsageCurrent
        }
    $out = $out + $tempout
}
$out | Get-Member


foreach ($Site in $SiteList) {
    Connect-PnPOnline -Url $Site.Url -Credentials $Credentials
    $Listlist = (Get-PnPList -Includes IsSystemList).Where({$_.IsSystemList -EQ $false})

    foreach ($DocLib in $Listlist) {
        $AllFiles = Get-PnPFolderItem -FolderSiteRelativeUrl $(($ListList[0].RootFolder.ServerRelativeUrl -split "/")[-1]) -Recursive -ItemType File
        {$total = 0}

        foreach ($File in $AllFiles) {
            $total = $total + $_.length
            $AllFiles[0].ServerRelativeUrl
        }
        
        {Write-Host "Total: $total"} 
        
    }
}

# Open https://github.com/ToddKlindt/TKTools/blob/master/DeleteOldVersions.ps1
