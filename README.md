# PS-PasswordExpiryNotification
<p align="center">
    <img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/PS-PasswordExpiryNotification.png" />
</p>

## Description
This script installs powershell files and configures a scheduled task that runs daily at 8am and processes all Active Directory users that are enabled without password never expires checked in their account options. Any account that has a password that is going to expire within the configurable number of days (Default is 14) will be sent an email to their email address. 

The email is dynamically crafted based on options configured during installation that are saved in a JSON config file. The JSON config file can be recreated if it's missing by deleting the config file and either running the installation script OR the main script in an interactive powershell session after installation.

## Dependencies/Prerequisites
### Dependencies
The following PowerShell modules are required and will be installed during installation to ensure the scripts run without error:
*   ActiveDirectory
*   [CredentialManager](https://github.com/davotronic5000/PowerShell_Credential_Manager)
*   [Mailozaurr](https://github.com/EvotecIT/Mailozaurr/tree/v2-speedygonzales)

Mailozaurr is a replacment for [Send-MailMessage because it has been deprecated](https://github.com/dotnet/platform-compat/blob/master/docs/DE0005.md) by Microsoft and should no longer be used.

### Prerequisites
This script requires that the account running the script be configured with 'Log on as a batch job' security permission. If you receive this error when you run the script you will need to modify the local security policy to allow the user account to run the scheduled task. If this setting is configured in a GPO it will need to be modified in that GPO, for domain controllers this is most likely the Default Domain Controllers Policy.

## Installation
### Run directly from github
```powershell
iwr https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Install-PasswordExpiryNotification.ps1 | iex
```

### Download and execute the script
Download and execute Install-PasswordExpiryNotification.ps1. The main scripts are embedded within and installed during execution after the options are configured.

## Options
### Supported SMTP Send Methods
The following SMTP Send Methods are available for configuration:
*   SMTPGRAPH (Send using a registered app with the Graph API) [Recommended]
*   SMTPAUTH (Office 365, Gmail, Zoho, Outlook, iCloud, Other) [See details below]
*   SMTPNOAUTH (Office 365 Direct Send, Google Restricted SMTP) [See details below]
*   SMTPRELAY (Custom options configured during installation)

***

#### SMTPGRAPH
> This method will allow the script to send emails using the Graph API with application permissions. A guide to configuring the app registration to get the Tenant ID, Client/App ID, and Client Secret is linked below. 
> 
> ⚠️ Mail.ReadWrite permission is also required when the size of email is above 4MB.
> 
> https://www.starwindsoftware.com/blog/sending-an-email-from-azure-using-microsoft-graph-api/
> 
> ⚠️ SMTP.Send application permissions allow the script to send as any account in the tenant. If you'd like to limit the scope to just a single email address for added security you will need to configure an application access policy and apply it to the Entra ID application you created.
> 
> https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access

***

#### SMTPAUTH
> If you are planning to use one of the SMTPAUTH methods with a dedicated mailbox to send emails, you will need to generate an app password for the account and use that when the script asks you to enter the password. Additionally you will need to disable security defaults in the client's Azure AD tenant. OAuth2 support is not provided in this script.

***

#### SMTPNOAUTH
> If you are planning to use one of the SMTPNOAUTH methods you will need to ensure the client has a static IP address where the server is located and that IP address has been added to the client's SPF DNS record. Otherwise the emails may be flagged as SPAM.

***

#### SMTPRELAY
> All options are asked during configuration so you can tailor the connection to your setup. If you need help setting up a SMTP relay to Office 365 you can use the guide below.
> 
> https://microsoftgeek.com/?p=1520

***

### Credentials
If you select to use SMTPGRAPH, SMTPAUTH, or SMTPRELAY with user authorization required you will be prompted to enter the client secret or account password. These credentials are stored securely in the Windows Credential Manager using the CredentialManager powershell module and are only accessible by the user account that installs and runs the script.

### JSON Config
These are the options that will be configured and saved to dynamically craft the HTML body and send the email. You can preinstall a JSON file with these options configured in C:\Scripts\AUPasswordExpiry prior to installation or running the main script and it will be detected automatically.
#### Table of JSON Config Variable
| Variable | Type | Description |
| --- | --- | --- |
| ClientName | String | The client's name, it will be used in the dynamic HTML generated during installation |
| ClientURL | String | The client's URL, the clients logo will use this link |
| ClientLogo | String | The client's logo, the https web URL or file path to a supported image extension (jpg,jpeg,png,gif,svg) |
| ClientDomain | String/Boolean | The client's DNS root of their domain. This is filled in automatically by the script if the client has a domain, otherwise its false |
| ClientVPN | Boolean | Does the client have a VPN for employees to use? |
| ClientAzure | Boolean | Does the client have Azure P1 or P2 licenses and password write-back enabled? |
| ClientSSPR | Boolean | Does the client have self service password reset enabled? |
| ClientSSPRButton | Boolean | Does the client have the lockscreen SSPR reset button enabled? |
| ExpireDays | Integer | The number of days before the users passwords expire to start emailing them notifications daily |
| SMTPMethod | String | The method used to send the email notification (SMTPGRAPH, SMTPAUTH, SMTPNOAUTH, SMTPRELAY) |
| SMTPServer | String | The SMTP server used to send the email |
| SMTPPort | Integer | The SMTP port to use when sending the email |
| SMTPTLS | Boolean | Use TLS when sending the email |
| TenantID | String | Tenant ID of application when using SMTPGRAPH SMTPMethod |
| SenderEmail | String | The email address used to send the password expiration notification email |
| EmailCredential | Boolean | AUPasswordExpiry password credential is saved in Credential Manager |

***

#### Client Logo
##### Validation
If a URL is provided the image will be downloaded and validated to check if it is a supported image type. If a file path is provided the extension will be validated to check if it is a supported image type. 

##### Sizing
It can be any width but will be resized in height to be 130px. I personally recommend a square PNG with transparent background sized to 130x130px.

##### Image Hosting
Client Logos can be hosted anywhere on the internet, however those links are subject to change by the website owners at any time. I would recommend creating a github repository and uploading the logos to it, then use the raw links to the logo directly from github. This site can be used to retrieve raw links to files easily, just copy and paste the github link to the file.

https://git-rawify.vercel.app/

***

#### Sample JSON File
```json
{
    "ClientName":  "Acme Corporation",
    "ClientURL":  "https://www.acmecorporation.com",
    "ClientLogo":  "https://th.bing.com/th/id/OIP.xDPJweY9GNABbZVUcw4TcwHaHa?rs=1\u0026pid=ImgDetMain",
    "ClientDomain":  "ACME.LOCAL",
    "ClientVPN":  true,
    "ClientAzure":  true,
    "ClientSSPR":  false,
    "ClientSSPRLockScreen":  false,
    "ExpireDays":  14,
    "SMTPMethod":  "SMTPAUTH",
    "SMTPServer":  "smtp.office365.com",
    "SMTPPort":  587,
    "SMTPTLS":  true,
    "TenantID": null,
    "SenderEmail":  "noreply@acmecorporation.com",
    "EmailCredential":  true
}
```

## Source Files
This folder contains the source files of the embedded files contained within the installation script and the HTML files used to visualize the HTML code during development.
