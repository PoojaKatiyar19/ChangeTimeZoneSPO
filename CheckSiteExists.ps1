Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"

Function Check-SiteExists($SiteURL, $Credentials)
{
    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
    $ctx.Credentials = $Credentials
    $Web = $Ctx.Web
    $ctx.Load($web)
     
    Try 
    {
            $ctx.ExecuteQuery()
            Return $True
    }
    Catch [Exception] 
    {
            #Write-host $_.Exception.Message -f Red
            Return $False
    }       
}