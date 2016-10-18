$programFiles = [environment]::getfolderpath("programfiles")

add-type -Path $programFiles'\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll'
Write-Host 'To enable SharePoint app sideLoading, enter Site Url, username and password'
  
$siteurl = Read-Host 'Site Url'
$username = Read-Host "User Name"
$password = Read-Host -AsSecureString 'Password'
  
 if ($siteurl -eq '')
 {
     $siteurl = 'https://apress365.sharepoint.com'
     $username = 'mark@apress365.onmicrosoft.com'
     $password = ConvertTo-SecureString -String 'Jaguar05' -AsPlainText -Force
 }
 $outfilepath = $siteurl -replace ':', '_' -replace '/', '_'
  
 try
 {
    [Microsoft.SharePoint.Client.ClientContext]$cc = New-Object Microsoft.SharePoint.Client.ClientContext($siteurl)
    [Microsoft.SharePoint.Client.SharePointOnlineCredentials]$spocreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
  
    $cc.Credentials = $spocreds
     
    #Write-Host -ForegroundColor Yellow 'SideLoading feature is not enabled on the site:' $siteurl
      
    $site = $cc.Site;

    # This guid identifies the Development Feature
    $developerFeature = new-object System.Guid "e374875e-06b6-11e0-b0fa-57f5dfd72085" 
    $site.Features.Add($developerFeature, $true, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None);
    $cc.ExecuteQuery();
      
    Write-Host -ForegroundColor Green 'Developer feature enabled on site' $siteurl
    #Activated the Developer Site feature
    
    $test = Read-Host 'Enter to continue'

 }
  
 catch
 { 
     Write-Host -ForegroundColor Red 'Error encountered when trying to enable Developer feature' $siteurl, ':' $Error[0].ToString();
     $test = Read-Host 'Enter to continue'
 }

