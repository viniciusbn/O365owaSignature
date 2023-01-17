#Install required modules (uncomment if needed)
#Install-Module -Name AzureAD -Scope AllUsers -Force
#Install-Module -Name ExchangeOnlineManagement -Scope AllUsers -Force

#Load required modules
Write-Output "Loading required modules"
Import-Module -Name ExchangeOnlineManagement -Global | Out-Null
Import-Module -Name AzureAD -UseWindowsPowerShell -Global | Out-Null

#Script enviroment varibles
$AppID = "PutYourAppID"
$CertThumb = "PutYourCertThumbPrint"
$TenantID = "PutYourTenantID"
$OrganizationDomain = "PutYourPrincipalDomainName.onmicrosoft.com"

#Local where is the script and html template file
$DefaulWorkDir = "c:\Signature"

#Connect to AzureAD and ExchangeOnline services
Write-Output "Connecting to remote services ExchangeOnline and AzureAD"
Connect-AzureAD -ApplicationId $AppID -CertificateThumbprint $CertThumb -TenantId $TenantID
Connect-ExchangeOnline -AppId $AppID -CertificateThumbprint $CertThumb -Organization $OrganizationDomain -ShowBanner:$false

#Gets all users' data from Azure AD and saves it to an array.
#Adjust Where-Object as you wish
$users = (Get-AzureADUser -All $true | Where-Object {$_.AccountEnabled -like "True" -and $_.Mail -like "*cercena.com.br" -and $_.UserType -like "Member" -and $_.UserPrincipalName -like "*cercena.com.br"} | Select-Object DisplayName, JobTitle, TelephoneNumber, Mobile, Mail)

#Saves HTML code of a signature from the signature generator to a variable. Change the path to the location of your file.
$HTMLsig = Get-Content "$DefaulWorkDir\Template_HTML.html" | out-string
$UserCount = 0

Write-Output "Starting the routine" > $DefaulWorkDir\lastrun.txt

foreach($user in $users){
	#Replacing AzureAD placeholders with personal user data
	$HTMLSigX = $HTMLsig.replace('%%DisplayName%%', $user.DisplayName). `
	replace('%%JobTitle%%', $user.JobTitle). `
	replace('%%TelephoneNumber%%', $user.TelephoneNumber). `
	replace('%%Mobile%%', $user.Mobile). `
	replace('%%Mail%%', $user.Mail)
	
	<#Saves the personalized email signature in mailbox settings. It should be available in Outlook on the web right away.
	The -AutoAddSignature parameter sets the signature as default for new messages and -AutoAddSignatureOnReply does the same for replies.#>
	Set-MailboxMessageConfiguration $user.Mail -SignatureHTML $HTMLSigX -AutoAddSignature $true -AutoAddSignatureOnReply $true
	$UserCount++	
	Write-Output "Setting signature for user $($UserCount): $($user.Mail)" >> $DefaulWorkDir\lastrun.txt
}
Write-Output "Number of users affected: $($UserCount)"
date >> $DefaulWorkDir\lastrun.txt
Write-Output "Disconnecting from ExchangeOnline"
Disconnect-ExchangeOnline -Confirm:$false
Write-Output "Disconnecting from AzureAD"
Disconnect-AzureAD -Confirm:$false