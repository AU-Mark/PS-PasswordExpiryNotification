# PS-PasswordExpiryNotification
<p align="center">
    <img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/refs/heads/main/Source%20Files/PS-PasswordExpiryNotification.png" />
</p>

## Description
This script installs powershell files and configures a scheduled task that runs daily at 8am and processes all Active Directory users that are enabled without password never expires checked in their account options. Any account that has a password that is going to expire within the configurable number of days (Default is 14) will be sent an email to their email address in the email address field of their Active Directory account. This email is dynamically crafted based on options configured during installation that are saved in a JSON config file. The JSON config file can be recreated if it's missing by deleting the config file and either running the installation script OR the main script in an interactive powershell session after installation.

## Dependencies/Prerequisites
The following modules are required and will be installed during installation to ensure the scripts run without error:
*   ActiveDirectory
*   CredentialManager
*   [Send-MailKitMessage](https://www.powershellgallery.com/packages/send-mailkitmessage)

## Installation
### Run directly from github
```powershell
iwr https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/refs/heads/main/Install-PasswordExpiryNotification.ps1 | iex
```

### Download and execute the script
Download and execute Install-PasswordExpiryNotification.ps1. The main scripts are embedded within and installed during execution after the options are configured.

## Options
### Supported SMTP Send Methods
The following SMTP Send Methods are available for configuration:
*   SMTPAUTH (Office 365, Gmail, Zoho, Outlook, iCloud, Other)
*   SMTPNOAUTH (Office 365 Direct Send, Google Restricted SMTP)
*   SMTPRELAY (Custom options configured during installation)

### JSON Options
These are the options that will be able to be configured to dynamically craft the HTML body of body and send the email. You can preinstall a JSON file with these options configured in C:\Scripts\AUPasswordExpiry prior to installation or running the main script and it will be detected automatically.
#### Individual Options
| Option | Type | Description |
| --- | --- | --- |
| ClientName | String | The client's name, it will be used in the dynamic HTML generated during installation |
| ClientURL | String | The client's URL, the clients logo will use this link |
| ClientLogo | String | The client's logo, the https web URL or file path to a supported image extension (jpg,jpeg,png,gif,svg) |
| ClientVPN | Boolean | Does the client have a VPN for employees to use? |
| ClientAzure | Boolean | Does the client have Azure P1 or P2 licenses and password write-back enabled? |
| ClientSSPR | Boolean | Does the client have self service password reset enabled? |
| ClientSSPRButton | Boolean | Does the client have the lockscreen SSPR reset button enabled? |
| ExpireDays | Integer | The number of days before the users passwords expire to start emailing them notifications daily |
| SMTPMethod | String | The method used to send the email notification (SMTPAUTH, SMTPNOAUTH, SMTPRELAY) |
| SMTPServer | String | The SMTP server used to send the email |
| SMTPPort | Integer | The SMTP port to use when sending the email |
| SMTPTLS | Boolean | Use TLS when sending the email |
| SenderEmail | String | The email address used to send the password expiration notification email |
| EmailCredential | Boolean | AUPasswordExpiry password credential is saved in Credential Manager |

#### Sample JSON File
```json
{
    "SMTPTLS":  true,
    "SMTPServer":  "smtp.office365.com",
    "ClientSSPRLockScreen":  false,
    "ClientName":  "Acme Incorporated",
    "ClientDomain":  "PMI.LOCAL",
    "ClientVPN":  true,
    "SMTPMethod":  "SMTPAUTH",
    "EmailCredential":  true,
    "ClientSSPR":  false,
    "ClientURL":  "https://www.acmeincorporated.com",
    "ClientLogo":  "https://th.bing.com/th/id/OIP.xDPJweY9GNABbZVUcw4TcwHaHa?rs=1\u0026pid=ImgDetMain",
    "SMTPPort":  587,
    "ClientAzure":  true,
    "SenderEmail":  "noreply@acmecorporation.com",
    "ExpireDays":  14
}
```
## Source Files
This folder contains the source files of the embedded files contained within the installation script and the HTML files used to visualize the HTML code during development.
