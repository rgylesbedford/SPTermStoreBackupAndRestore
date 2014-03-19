# Must have the SharePoint Client dlls installed see:
# SharePoint 2013 - http://www.microsoft.com/en-us/download/details.aspx?id=35585
# SharePoint Online - http://www.microsoft.com/en-us/download/details.aspx?id=42038

$CSOMdir = "${env:CommonProgramFiles}\microsoft shared\Web Server Extensions\16\ISAPI"
$excludeDlls = "*.Portable.dll"   
if ((Test-Path $CSOMdir -pathType container) -ne $true)
{
    $CSOMdir = "${env:CommonProgramFiles}\microsoft shared\Web Server Extensions\15\ISAPI"
    if ((Test-Path $CSOMdir -pathType container) -ne $true)
    {
        Throw "Please install the SharePoint 2013[1] or SharePoint Online[2] Client Components`n `n[1] http://www.microsoft.com/en-us/download/details.aspx?id=35585`n[2] http://www.microsoft.com/en-us/download/details.aspx?id=42038`n `n "
    }
}
$CSOMdlls = Get-Item "$CSOMdir\*.dll" -exclude $excludeDlls 
ForEach ($dll in $CSOMdlls) {
    [System.Reflection.Assembly]::LoadFrom($dll.FullName) | Out-Null
}

$siteUrl = "https://tenant.sharepoint.com"
$username = "username@tenant.com"
$domain = "domain"
$password = Read-Host -Prompt "Enter Password for $username" -AsSecureString

$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
# do not set $ctx.Credentials if you want default credentials to be used
#$ctx.Credentials = New-Object System.Net.NetworkCredential($username, $password, $domain)
#$ctx.Credentials = New-Object Microsoft.Sharepoint.Client.SharePointOnlineCredentials($username, $password)


$myScriptPath = (Split-Path -Parent $MyInvocation.MyCommand.Path)

#. "$myScriptPath\TaxonomyBackup.ps1" -ClientContext $ctx -LiteralPath "C:\TermStoreBackup"
. "$myScriptPath\TaxonomyRestore.ps1" -ClientContext $ctx -LiteralPath "C:\TermStoreBackup\2014-3-19_13-26-58_mms_bak.xml"
