# O365owaSignature
Powershell script to set and automatize Microsoft 365 OWA mail signature with ldap data and html template.

System requisites:
This script was created to run with any version of powershell, on this implementation we have tested with powershell 7.3.1

Powershell requisites/modules:
- AzureAD
- ExchangeOnlineManagement

Office/Microsoft 365 requirements:
- AppID access

**How to run:**

**1st -** Create a AppID account ad self signed certificate.

Microsoft documentation
https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps


**2st -** Get the certificate thumbprint from 1st step and hardcode it on the $CertThumb varible.
In powershell
```powershell
Get-ChildItem -Path cert:\CurrentUser\My
```

**3st -** Define the anothers variables as needed on your deployment..

- $AppID =              "App ID generated on 1st step"
- $CertThumb =          "Certificate thumbprint generated on 1st step"
- $TenantID =           "Get it from your Azure console"
- $OrganizationDomain = "Your primary domain .onmicrosoft.com"

**4st -** Customize the variables as needed on Template_HTML.html

- %%DisplayName%% =     User complete name
- %%JobTitle%% =        User jot title
- %%TelephoneNumber%% = On our enviroment this is the branch line
- %%Mobile%% =          User mobile phone number
- %%Mail%% =            User email address

*You can make your own HTML template as you wich, just make sure to adjust variables that you want to replace.*

**5st -** Automatize script

You can set your script to run automatically, with job task in Windows Task Manager..
