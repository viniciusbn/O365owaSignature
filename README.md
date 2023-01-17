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


2st - Get the cert thumbprint from 1st step and hardcode this on the script line

