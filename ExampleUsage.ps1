# Must have the SharePoint 2013 Client dlls installed see:
# http://www.microsoft.com/en-us/download/details.aspx?id=35585

[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Sharepoint.Client")
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Taxonomy")

$siteUrl = "https://sp2013dev:1000/"	
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)

## Example for Office 365 SharePoint Online
## http://msdn.microsoft.com/EN-US/library/microsoft.sharepoint.client.sharepointonlinecredentials.aspx

#$siteURL = "https://mysite.sharepoint.com"
#$username = "admin@mysite.sharepoint.com"
#$password = Read-Host -Prompt "Enter Password" -AsSecureString
#$SharePointOnlineCredentials = New-Object Microsoft.Sharepoint.Client.SharePointOnlineCredentials($username, $password)
#$ctx.Credentials = $SharePointOnlineCredentials

.\TaxonomyBackup.ps1 -ClientContext $ctx -LiteralPath "C:\TermStoreBackup"

.\TaxonomyRestore.ps1 -ClientContext $ctx -LiteralPath "C:\TermStoreBackup\2012-03-14_15-27-59_mms_bak.xml"