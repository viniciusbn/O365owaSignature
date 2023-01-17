# O365owaSignature
Powershell script to set a OWS email signature with ldap data and html template.

System requisites:
This script was created to run with any version of powershell, on this implementation i have used powershell 7.3.1

Powershell requisites/modules:
- AzureAD
- ExchangeOnlineManagement

How to run:

1st - Create a AppID account, to use transparent and script like authentication.

Microsoft documentation
https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps


2st - Get the certificate thumbprint from 1st step and hardcode it on the $CertThumb varible.

3st - Set others varibles as needed on your deployment.

- $AppID =              "App ID generated on 1st step"
- $CertThumb =          "Certificate thumbprint generated on 1st step"
- $TenantID =           "Get it from your Azure console"
- $OrganizationDomain = "Your primary domain .onmicrosoft.com"

4st - Customie the varibles as needed on Template_HTML.html

- %%DisplayName%% =     User complete name
- %%JobTitle%% =        User jot title
- %%TelephoneNumber%% = On our enviroment this is the branch line
- %%Mobile%% =          User mobile phone number
- %%Mail%% =            User email address

You can make your own HTML template as you needed, just make sure to adjust variables that you want to replace.

5st - Automatize script

You can set your script to run automatically with task job in Windows Task Manager..