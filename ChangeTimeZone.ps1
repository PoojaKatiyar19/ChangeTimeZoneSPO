Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"

Import-Module -Name ".\CheckSiteExists.ps1"
$xmlFile = ".\Configurations.xml"
$xmlConfig = [System.Xml.XmlDocument](Get-Content $xmlFile)

$RootPath = $xmlConfig.Settings.LogsSettings.RootPath
$date = (Get-Date).ToString('yyyy-MM-dd-HHmm')

$ErrorLogs = $RootPath + "ErrorLogs_$($date).txt"
$TZLogs = $RootPath + "TZLogs_$($date).txt.txt"
$CSVLocation = $RootPath + "Sites.csv"

$userId = $xmlConfig.Settings.ConfigurationSettings.UserID
$Password = $xmlConfig.Settings.ConfigurationSettings.Password
$pwd = $(ConvertTo-SecureString $Password -AsPlainText -Force)
$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)

$sites = Import-csv -header URL $csvlocation

Write-Host "Number of sites: " $($sites.URL.Count) -ForegroundColor DarkYellow

#region changing the time zone here
foreach($site in $sites)
{
    
    try
    {
        $URL = $site.URL        

        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($URL)
        $ctx.credentials = $creds

        $SiteExists = Check-SiteExists -SiteURL $URL -Credentials $creds

        if($SiteExists -eq $true)
        {
            $RegionalSettings = $ctx.Web.RegionalSettings #Get all the regional settings
            $AllTZ = $ctx.Web.RegionalSettings.TimeZones  #Get all the time zones
            $CurrentTZ = $ctx.Web.RegionalSettings.TimeZone
            $ctx.Load($CurrentTZ)
            $ctx.Load($AllTZ)
            $ctx.ExecuteQuery()
            $timezone = $AllTZ | where{$_.Description -eq "(UTC+01:00) Amsterdam, Berlin, Bern, Rome, Stockholm, Vienna"}
            
            #Check if current time zone of the site is not equal to the desired time zone
            if($CurrentTZ.Id -ne $timezone.Id)
            {
                Write-Host "Changing the time zone for: " $URL -ForegroundColor Cyan
                $RegionalSettings.TimeZone = $timezone
                $ctx.Web.Update()
                $ctx.ExecuteQuery()

                "Time zone updated to European time zone for: " + $URL | Out-File -FilePath $TZLogs -Append
            }
            else
            {
                "Site is already in the same time zone: " + $URL | Out-File -FilePath $TZLogs -Append
            }
            
        }
        else
        {
            $URL + "  -------> does not exist!" | Out-File -FilePath $ErrorLogs -Append -
        }

    }
    catch
    {
        $URL + "`r`n" + $_.Exception.Item, $_.Exception.Message | Out-File -FilePath $ErrorLogs -Append
    }

}

#endregion




