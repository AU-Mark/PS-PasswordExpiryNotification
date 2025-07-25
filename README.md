# PowerShell Password Expiry Notification System

<p align="center">
    <img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/PS-PasswordExpiryNotification.png" alt="PS-PasswordExpiryNotification Logo" />
</p>

<p align="center">
    <strong>Automated Active Directory password expiration notifications with customizable HTML email templates</strong>
</p>

<p align="center">
    <a href="#quick-start">Quick Start</a> • 
    <a href="#prerequisites">Prerequisites</a> • 
    <a href="#installation">Installation</a> • 
    <a href="#configuration">Configuration</a> • 
    <a href="#troubleshooting">Troubleshooting</a>
</p>

---

## 📋 Table of Contents

- [Overview](#overview)
- [✨ Features](#features)
- [⚠️ Important Requirements](#important-requirements)
- [📋 Prerequisites](#prerequisites)
- [🚀 Quick Start](#quick-start)
- [📦 Installation](#installation)
- [⚙️ Configuration](#configuration)
- [📧 SMTP Configuration Methods](#smtp-configuration-methods)
- [🎨 Customization](#customization)
- [🔧 Usage](#usage)
- [🛠️ Troubleshooting](#troubleshooting)
- [🔒 Security Considerations](#security-considerations)
- [📁 File Structure](#file-structure)
- [🤝 Contributing](#contributing)

---

## Overview

This PowerShell solution automatically monitors Active Directory user accounts and sends personalized HTML email notifications to users whose passwords are approaching expiration. The system runs as a scheduled task and dynamically generates email content based on your organization's specific configuration and capabilities.

## ✨ Features

- **📅 Automated Daily Monitoring**: Scheduled task runs daily at 8:00 AM
- **🎯 Smart User Targeting**: Only notifies enabled users with expiring passwords
- **📧 Multiple SMTP Methods**: Graph API, SMTP AUTH, SMTP Relay, and Unauthenticated SMTP
- **🎨 Dynamic HTML Templates**: Customizable email design based on client capabilities
- **🔐 Secure Credential Storage**: Uses Windows Credential Manager for authentication
- **📱 Responsive Email Design**: Mobile-friendly HTML email templates
- **⚙️ Flexible Configuration**: JSON-based configuration with interactive setup
- **🔍 Comprehensive Logging**: Detailed logging for monitoring and troubleshooting

---

## ⚠️ Important Requirements

> **🛑 CRITICAL:** This script requires specific system permissions and environment setup to function properly.

### **Essential Prerequisites**
- **Windows Server** with Active Directory Domain Services
- **"Log on as a batch job"** user rights for the executing account
- **PowerShell 5.1** or later with appropriate execution policy
- **Administrator privileges** during installation
- **Email infrastructure** (SMTP server, Office 365, or Graph API access)

### **Account Requirements**
The user account running this script must have:
- Local **"Log on as a batch job"** rights
- **Read access** to Active Directory users
- **Network access** to configured SMTP servers
- **Write permissions** to `C:\Scripts\AUPasswordExpiry`

---

## 📋 Prerequisites

### System Requirements

| Component | Requirement | Notes |
|-----------|------------|-------|
| **Operating System** | Windows Server 2016+ | Domain Controller or domain-joined server |
| **PowerShell** | 5.1 or later | PowerShell 5.1 and 7.x supported |
| **Active Directory** | Domain Services installed | Script queries AD users |
| **Network Access** | SMTP/Graph API connectivity | For sending email notifications |
| **Disk Space** | 50MB minimum | For scripts, logs, and dependencies |

### PowerShell Modules

The following modules are **automatically installed** during setup:

```powershell
# Required modules (auto-installed)
Install-Module ActiveDirectory      # AD user queries
Install-Module CredentialManager    # Secure credential storage  
Install-Module Mailozaurr          # Modern email sending (replaces deprecated Send-MailMessage)
```

### User Rights Assignment

> **⚠️ CRITICAL STEP:** The executing user account **must** have "Log on as a batch job" rights.

#### Method 1: Local Security Policy (Standalone Server)
1. Run `secpol.msc` as Administrator
2. Navigate to: **Security Settings** → **Local Policies** → **User Rights Assignment**
3. Double-click **"Log on as a batch job"**
4. Add your user account
5. Apply changes and restart if prompted

#### Method 2: Group Policy (Domain Environment)
1. Open **Group Policy Management Console**
2. Edit the appropriate GPO (often **Default Domain Controllers Policy** for DCs)
3. Navigate to: **Computer Configuration** → **Windows Settings** → **Security Settings** → **Local Policies** → **User Rights Assignment**
4. Configure **"Log on as a batch job"** to include your service account
5. Run `gpupdate /force` on target servers

#### Method 3: Command Line Verification
```powershell
# Check current user rights (requires admin privileges)
secedit /export /cfg "$env:TEMP\secpol.cfg" | Out-Null
$policy = Get-Content "$env:TEMP\secpol.cfg"
$hasRights = $policy -match "SeBatchLogonRight.*$env:USERNAME"
Remove-Item "$env:TEMP\secpol.cfg"
Write-Host "Has batch logon rights: $($hasRights -ne $null)"
```

### PowerShell Execution Policy

Ensure scripts can execute:

```powershell
# Check current policy
Get-ExecutionPolicy

# Set appropriate policy (run as Administrator)
Set-ExecutionPolicy RemoteSigned -Scope LocalMachine -Force

# Alternative: CurrentUser scope (no admin required)
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

---

## 🚀 Quick Start

### Option 1: Direct Installation from GitHub

```powershell
# Run directly from GitHub (requires internet access)
iwr https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Install-PasswordExpiryNotification.ps1 | iex
```

### Option 2: Download and Run

1. Download `Install-PasswordExpiryNotification.ps1`
2. Run as Administrator:
   ```powershell
   PowerShell.exe -ExecutionPolicy Bypass -File "Install-PasswordExpiryNotification.ps1"
   ```

The installation wizard will guide you through:
- ✅ Dependency installation
- ✅ Configuration setup  
- ✅ Credential storage
- ✅ Scheduled task creation
- ✅ Testing and validation

---

## 📦 Installation

### Step-by-Step Installation Process

#### 1. Pre-Installation Verification

```powershell
# Verify PowerShell version
$PSVersionTable.PSVersion

# Check execution policy
Get-ExecutionPolicy

# Verify you're running as Administrator
([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

# Test Active Directory connectivity
Get-ADDomain
```

#### 2. Run Installation Script

The installer will:
- ✅ Install required PowerShell modules
- ✅ Create directory structure (`C:\Scripts\AUPasswordExpiry`)
- ✅ Generate embedded PowerShell scripts
- ✅ Configure interactive setup wizard
- ✅ Create scheduled task
- ✅ Validate configuration

#### 3. Interactive Configuration

During installation, you'll configure:

| Setting | Description | Example |
|---------|-------------|---------|
| **Client Name** | Organization name for emails | `Acme Corporation` |
| **Client URL** | Organization website | `https://www.acme.com` |
| **Client Logo** | Logo URL or file path | `https://example.com/logo.png` |
| **VPN Access** | Users have VPN connectivity | `Yes/No` |
| **Azure P1/P2** | Password writeback enabled | `Yes/No` |
| **SSPR Enabled** | Self-service password reset | `Yes/No` |
| **Notification Days** | Days before expiry to notify | `14` (default) |
| **SMTP Method** | Email delivery method | `Graph API` (recommended) |

#### 4. SMTP Configuration

Choose your preferred email delivery method:
- **🏆 Graph API** (Recommended): Most reliable, requires app registration
- **📧 SMTP AUTH**: Traditional SMTP with authentication  
- **🔗 SMTP Relay**: Custom SMTP relay configuration
- **📤 Unauthenticated**: Direct send methods (requires SPF records)

#### 5. Final Steps

```powershell
# Verify installation
Test-Path "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1"

# Check scheduled task
Get-ScheduledTask -TaskName "Password Expiry Email Notification"

# Test email generation (optional)
& "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1" -Test
```

---

## ⚙️ Configuration

### Configuration File Location

The system uses a JSON configuration file stored at:
```
C:\Scripts\AUPasswordExpiry\clientconf.json
```

### Configuration Parameters

| Parameter | Type | Description | Example |
|-----------|------|-------------|---------|
| `ClientName` | String | Organization name displayed in emails | `"Acme Corporation"` |
| `ClientURL` | String | Organization website (logo link destination) | `"https://www.acme.com"` |
| `ClientLogo` | String | Logo image URL or file path | `"https://example.com/logo.png"` |
| `ClientDomain` | String/Boolean | AD domain (auto-detected) | `"ACME.LOCAL"` or `false` |
| `ClientVPN` | Boolean | Organization provides VPN access | `true` |
| `ClientAzure` | Boolean | Azure P1/P2 with password writeback | `true` |
| `ClientSSPR` | Boolean | Self-service password reset enabled | `false` |
| `ClientSSPRLockScreen` | Boolean | SSPR lockscreen button available | `false` |
| `ExpireDays` | Integer | Days before expiry to start notifications | `14` |
| `SMTPMethod` | String | Email delivery method | `"SMTPGRAPH"` |
| `SMTPServer` | String | SMTP server hostname/IP | `"smtp.office365.com"` |
| `SMTPPort` | Integer | SMTP server port | `587` |
| `SMTPTLS` | String | TLS encryption setting | `"Auto"` |
| `TenantID` | String | Azure tenant ID (Graph API only) | `"12345678-1234-..."` |
| `SenderEmail` | String | From email address | `"noreply@acme.com"` |
| `EmailCredential` | Boolean | Credentials stored in Credential Manager | `true` |

### Sample Configuration

```json
{
    "ClientName": "Acme Corporation",
    "ClientURL": "https://www.acmecorporation.com",
    "ClientLogo": "https://example.com/acme-logo.png",
    "ClientDomain": "ACME.LOCAL",
    "ClientVPN": true,
    "ClientAzure": true,
    "ClientSSPR": false,
    "ClientSSPRLockScreen": false,
    "ExpireDays": 14,
    "SMTPMethod": "SMTPGRAPH",
    "SMTPServer": null,
    "SMTPPort": null,
    "SMTPTLS": "Auto",
    "TenantID": "12345678-1234-5678-9abc-123456789def",
    "SenderEmail": "noreply@acmecorporation.com",
    "EmailCredential": true
}
```

### Reconfiguring the System

To modify configuration after installation:

```powershell
# Method 1: Delete config file and run main script interactively
Remove-Item "C:\Scripts\AUPasswordExpiry\clientconf.json"
& "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1"

# Method 2: Re-run installer
& "Install-PasswordExpiryNotification.ps1"

# Method 3: Manually edit JSON file
notepad "C:\Scripts\AUPasswordExpiry\clientconf.json"
```

---

## 📧 SMTP Configuration Methods

### 🏆 Method 1: Microsoft Graph API (Recommended)

**Best for:** Office 365/Microsoft 365 environments

**Advantages:**
- ✅ Most reliable delivery
- ✅ Modern authentication (OAuth2)
- ✅ Detailed delivery reporting
- ✅ No SMTP port dependencies

**Requirements:**
1. **App Registration** in Azure AD/Entra ID
2. **Required Permissions:**
   - `Mail.Send` (Application permission)
   - `Mail.ReadWrite` (if emails exceed 4MB)
3. **Tenant ID, Client ID, and Client Secret**

**Setup Guide:**
Follow this comprehensive guide for Graph API configuration:
[📖 Microsoft Graph API Email Setup Guide](https://www.starwindsoftware.com/blog/sending-an-email-from-azure-using-microsoft-graph-api/)

**Security Note:**
> ⚠️ The `Mail.Send` permission allows sending as ANY user in the tenant. For enhanced security, configure Application Access Policies to restrict scope to specific mailboxes.
> 
> [📖 Limit Mail Access to Specific Mailboxes](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access)

---

### 📧 Method 2: SMTP Authentication

**Best for:** Mixed environments, non-Microsoft email providers

**Supported Providers:**
- Office 365 (`smtp.office365.com:587`)
- Gmail (`smtp.gmail.com:587`) 
- Zoho (`smtp.zoho.com:587`)
- Outlook.com (`smtp-mail.outlook.com:587`)
- iCloud (`smtp.mail.me.com:587`)
- Custom SMTP servers

**Requirements:**
- Dedicated email account with SMTP access
- App-specific passwords (recommended over account passwords)
- TLS/SSL encryption support

**Important Notes:**
> ⚠️ **For Office 365:** You may need to disable Security Defaults in Azure AD to allow basic authentication, or preferably use Modern Authentication where supported.
>
> 🔐 **Password Security:** Always use app-specific passwords instead of main account passwords.

---

### 🔗 Method 3: SMTP Relay

**Best for:** Organizations with existing SMTP infrastructure

**Configuration Options:**
- Custom SMTP server settings
- TLS encryption (optional)  
- Authentication (optional)
- Custom ports and security settings

**Use Cases:**
- Internal mail servers
- Third-party email services
- Hybrid email environments

---

### 📤 Method 4: Unauthenticated SMTP

**Best for:** Simple environments with proper SPF configuration

**Available Options:**
- **Office 365 Direct Send:** Requires static IP in SPF record
- **Gmail Restricted SMTP:** Limited functionality

**Requirements:**
> ⚠️ **Critical:** Your server's IP address MUST be included in your domain's SPF DNS record, or emails will be marked as spam.

**SPF Record Example:**
```dns
v=spf1 ip4:203.0.113.10 include:spf.protection.outlook.com ~all
```

---

## 🎨 Customization

### Logo Requirements

**Supported Formats:** JPG, JPEG, PNG, GIF, SVG

**Sizing Guidelines:**
- **Height:** Automatically resized to 130px
- **Width:** Proportionally scaled
- **Recommended:** 130x130px square PNG with transparent background

**Hosting Options:**
1. **GitHub Repository** (Recommended for stability)
   - Upload to GitHub repo
   - Use [Git Rawify](https://git-rawify.vercel.app/) to get direct links
2. **Web URL** (Any accessible image URL)
3. **Local File Path** (Embedded as attachment)

### Email Template Customization

The system generates dynamic HTML emails based on your configuration:

**Conditional Content Blocks:**
- **Domain Users:** Ctrl+Alt+Delete instructions
- **VPN Users:** Enhanced connectivity options  
- **Azure P1/P2:** Office 365 password change options
- **SSPR Enabled:** Self-service reset instructions
- **Lockscreen SSPR:** Lock screen reset button guidance

**Template Structure:**
```
📧 Email Components
├── 🎨 Header (Logo + Company Name)
├── 📝 Personalized Greeting  
├── ⚠️ Expiration Warning
├── 📋 Method-Specific Instructions
│   ├── 🏢 Domain/VPN Instructions
│   ├── 🌐 Office 365 Web Interface
│   ├── 🔐 Self-Service Password Reset
│   └── 📞 Help Desk Contact
└── 📄 Footer (Company Copyright)
```

### Advanced Customization

To modify email templates:

1. **Edit HTML Template:**
   ```powershell
   notepad "C:\Scripts\AUPasswordExpiry\PasswordExpiryHTML.ps1"
   ```

2. **Test Changes:**
   ```powershell
   & "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1" -Test
   ```

3. **Custom CSS Styling:**
   - Modify inline styles in the HTML template
   - Ensure email client compatibility
   - Test across different email clients

---

## 🔧 Usage

### Manual Execution

```powershell
# Run normally (processes all expiring users)
& "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1"

# Test mode (generates sample HTML, opens in browser)
& "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1" -Test
```

### Scheduled Task Details

**Task Name:** `Password Expiry Email Notification`
**Schedule:** Daily at 8:00 AM
**User Context:** Installing user account
**Execution Policy:** Bypass

### Monitoring and Logs

**Log Location:** `C:\Scripts\AUPasswordExpiry\Logs\`

**Log Files:**
- `Install.log` - Installation process log
- `Debug.log` - Runtime execution log

**Log Levels:**
- **INFO:** Normal operations
- **WARNING:** Non-critical issues
- **ERROR:** Critical failures
- **DEBUG:** Detailed troubleshooting info

**Sample Log Monitoring:**
```powershell
# View recent log entries
Get-Content "C:\Scripts\AUPasswordExpiry\Logs\Debug.log" -Wait -Tail 20

# Check for errors
Select-String "ERROR" "C:\Scripts\AUPasswordExpiry\Logs\*.log"

# Monitor email notifications
Select-String "Notification email sent" "C:\Scripts\AUPasswordExpiry\Logs\Debug.log"
```

---

## 🛠️ Troubleshooting

### Common Issues and Solutions

#### ❌ "Log on as a batch job" Rights Error

**Error Message:**  
`The current user does not have "Log on as a batch job" rights`

**Solution:**
1. Follow the [User Rights Assignment](#user-rights-assignment) steps above
2. Verify with command:
   ```powershell
   secedit /export /cfg "$env:TEMP\secpol.cfg"
   Select-String "SeBatchLogonRight" "$env:TEMP\secpol.cfg"
   ```
3. Restart server if GPO changes were made
4. Re-run installation script

#### ❌ PowerShell Module Installation Failures

**Symptoms:** Module import errors, missing cmdlets

**Solutions:**
```powershell
# Check PowerShell version (5.1+ required)
$PSVersionTable.PSVersion

# Update PowerShellGet and PackageManagement
Install-Module PowerShellGet -Force -AllowClobber
Install-Module PackageManagement -Force -AllowClobber

# Manually install required modules
Install-Module ActiveDirectory -Force
Install-Module CredentialManager -Force  
Install-Module Mailozaurr -Force

# Verify module availability
Get-Module -ListAvailable ActiveDirectory, CredentialManager, Mailozaurr
```

#### ❌ Active Directory Access Issues

**Error Message:**  
`Unable to contact the server` or `Access denied`

**Solutions:**
```powershell
# Test AD connectivity
Test-NetConnection -ComputerName $env:LOGONSERVER.TrimStart('\\') -Port 389

# Verify AD module is available
Import-Module ActiveDirectory
Get-ADDomain

# Check user permissions
Get-ADUser $env:USERNAME -Properties MemberOf
```

#### ❌ Email Delivery Failures

**For Graph API Issues:**
1. Verify tenant ID, client ID, and client secret
2. Check application permissions in Azure AD
3. Ensure consent has been granted
4. Test Graph API connectivity:
   ```powershell
   # Test Graph API authentication
   $credential = Get-StoredCredential -Target AUPasswordExpiry
   Test-GraphConnection -ClientId $credential.UserName -ClientSecret $credential.GetNetworkCredential().Password
   ```

**For SMTP Issues:**
1. Verify SMTP server and port settings
2. Test network connectivity:
   ```powershell
   Test-NetConnection -ComputerName "smtp.office365.com" -Port 587
   ```
3. Check credential storage:
   ```powershell
   Get-StoredCredential -Target AUPasswordExpiry
   ```
4. Validate email account authentication

#### ❌ Scheduled Task Not Running

**Diagnostic Steps:**
```powershell
# Check task status
Get-ScheduledTask -TaskName "Password Expiry Email Notification"

# View task history
Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-TaskScheduler/Operational'; ID=201}

# Run task manually
Start-ScheduledTask -TaskName "Password Expiry Email Notification"

# Check task definition
Export-ScheduledTask -TaskName "Password Expiry Email Notification"
```

#### ❌ JSON Configuration Corruption

**Error Message:**  
`JSON syntax errors or corrupted config file`

**Solutions:**
```powershell
# Validate JSON syntax
Get-Content "C:\Scripts\AUPasswordExpiry\clientconf.json" | ConvertFrom-Json

# Recreate configuration
Remove-Item "C:\Scripts\AUPasswordExpiry\clientconf.json" -Force
& "C:\Scripts\AUPasswordExpiry\PasswordExpiryNotification.ps1"
```

### Debug Mode

Enable detailed logging:

```powershell
# Edit the main script to enable debug mode
$Debug = $True  # Change from $False to $True in PasswordExpiryNotification.ps1
```

### Getting Help

1. **Check Logs:** Always review log files first
2. **Run Test Mode:** Use `-Test` parameter to validate configuration
3. **Verify Prerequisites:** Ensure all requirements are met
4. **Manual Testing:** Run components individually to isolate issues

---

## 🔒 Security Considerations

### Credential Storage

- **Windows Credential Manager:** Credentials are encrypted and only accessible by the installing user
- **Account Isolation:** Use dedicated service accounts where possible
- **Least Privilege:** Grant minimal required permissions

### Email Security

- **TLS Encryption:** All SMTP connections use TLS/SSL encryption
- **Authentication:** Strong authentication methods (OAuth2 preferred)
- **SPF Records:** Configure proper SPF records for unauthenticated methods

### Access Control

- **File Permissions:** Restrict access to script directory
- **Log Security:** Protect log files from unauthorized access
- **Network Security:** Secure SMTP communications

### Best Practices

```powershell
# Example: Secure file permissions
$acl = Get-Acl "C:\Scripts\AUPasswordExpiry"
$acl.SetAccessRuleProtection($true,$false)
$adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Administrators","FullControl","ContainerInherit,ObjectInherit","None","Allow")
$systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM","FullControl","ContainerInherit,ObjectInherit","None","Allow")
$acl.RemoveAccessRuleAll($acl.AccessRuleProtection)
$acl.AddAccessRule($adminRule)
$acl.AddAccessRule($systemRule)
Set-Acl "C:\Scripts\AUPasswordExpiry" $acl
```

---

## 📁 File Structure

```
C:\Scripts\AUPasswordExpiry\
├── 📄 PasswordExpiryNotification.ps1    # Main execution script
├── 📄 PasswordExpiryHTML.ps1           # HTML email template generator
├── 📄 clientconf.json                  # Configuration file
└── 📁 Logs/                           # Log file directory
    ├── 📄 Install.log                  # Installation log
    └── 📄 Debug.log                    # Runtime log
```

### Script Dependencies

**Main Script (`PasswordExpiryNotification.ps1`):**
- Queries Active Directory for expiring passwords
- Loads configuration from JSON file
- Generates personalized emails using HTML template
- Sends notifications via configured SMTP method
- Logs all activities

**HTML Template (`PasswordExpiryHTML.ps1`):**
- Dynamically generates email content
- Handles conditional content blocks
- Applies responsive design for mobile compatibility
- Embeds logos and styling

**Configuration (`clientconf.json`):**
- Stores all customization settings
- Enables dynamic email generation
- Supports multiple deployment scenarios

---

## 🤝 Contributing  

We welcome contributions to improve this password notification system!

### How to Contribute

1. **Fork the repository**
2. **Create a feature branch:** `git checkout -b feature-name`
3. **Make your changes** with proper testing
4. **Update documentation** as needed
5. **Submit a pull request** with detailed description

### Areas for Contribution

- 🌐 Additional SMTP provider support
- 🎨 Email template improvements
- 🔧 PowerShell Core compatibility
- 📊 Enhanced reporting features
- 🌍 Internationalization/localization
- 🧪 Additional testing scenarios

### Reporting Issues

When reporting issues, please include:
- PowerShell version (`$PSVersionTable`)
- Operating system details
- Error messages and log excerpts  
- Configuration (sanitized)
- Steps to reproduce

---

<p align="center">
    <strong>🚀 Ready to get started?</strong><br>
    <a href="#quick-start">Jump to Quick Start Guide</a>
</p>
