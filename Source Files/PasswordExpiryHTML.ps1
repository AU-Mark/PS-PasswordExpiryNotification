$HTMLBegin = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <title>Password Expiry Notification</title>
    <style>
        body { margin: 0; padding: 0; width: 100%; height: 100%; background-color: #f0f0f0; color: #000; font-family: Arial, sans-serif; }
        table { border-spacing: 0; }
        .container { width: 100%; max-width: 1000px; margin: 0 auto; background-color: #fff; border-radius: 8px; overflow: hidden; }
        .header, .footer { padding: 20px; text-align: center; background-color: #808080; color: #fff; }
        .header .company-name { font-size: 28px; font-weight: 700; margin-top: 10px; margin-bottom: 0; }
        .content { padding: 20px; }
        .content h2 { text-align: center; }
        .content p { margin: 0 0 10px; }
        .instructions { margin: 20px 0; }
        .instructions p { margin: 10px 0; display: flex; align-items: center; }
        .instructions img { margin-right: 10px; vertical-align: middle; }
        .tab { margin-left: 20px; }
        .tab2 { margin-left: 40px; }
        hr {border: 1px solid #808080;}
        @media (max-width: 600px) { .container { width: 100%; } }
    </style>
</head>
<body>
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0">
        <tr>
            <td align="center" style="padding: 10px 0;">
                <table role="presentation" class="container" cellspacing="0" cellpadding="0" border="0">
                    <tr>
                        <td class="header" style="border-top-left-radius: 8px; border-top-right-radius: 8px;">
                            <a href="$($clientConfig["ClientURL"])" target="_blank">
                                <img src="
"@

$BodyBegin = @"
" alt="Company Logo" style="height: 130px; width: auto; display: block; margin: 0 auto;">
                            </a>
                            <div class="company-name">$($clientConfig["ClientName"])</div>
                        </td>
                    </tr>
                    <tr>
                        <td class="content">
                            <h2>Password Expiry Notification</h2>
                            <p>Hello&nbsp;<strong>$FirstName</strong>,</p>
                            <p>Your&nbsp;<strong>$($clientConfig["ClientName"])</strong>&nbsp;password will expire&nbsp;<strong>$ExpiryMsg</strong>. Please change your password using one of the methods below so you don't get locked out of your account. If you have any questions call Aunalytics for assistance at 1-855-799-DATA (1-855-799-3282).</p>
                            <div class="instructions">
"@

$InstructDomain = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Key.png" alt="Key Icon">&nbsp;If you work at a&nbsp;<strong>$($clientConfig["ClientName"]) Office</strong>:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;Press&nbsp;<strong>Ctrl + Alt + Delete</strong>&nbsp;on the keyboard.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Select&nbsp;<strong>Change a password</strong>&nbsp;from the menu.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Enter your old password, and your new password twice and press the enter key or the arrow button.</p>
                                <hr>
"@ 

$InstructDomainVPN = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Key.png" alt="Key Icon">&nbsp;If you work at a&nbsp;<strong>$($clientConfig["ClientName"]) Office or are connected to the $($clientConfig["ClientName"]) VPN</strong>:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;Press&nbsp;<strong>Ctrl + Alt + Delete</strong>&nbsp;on the keyboard.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Select&nbsp;<strong>Change a password</strong>&nbsp;from the menu.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Enter your old password, and your new password twice and press the enter key or the arrow button.</p>
                                <hr>
"@ 

$InstructEntra = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Key.png" alt="Key Icon">&nbsp;Reset your password with&nbsp;<strong>Ctrl + Alt + Delete</strong><p>
                                <p class="tab"><strong>1.</strong>&nbsp;Press&nbsp;<strong>Ctrl + Alt + Delete</strong>&nbsp;on the keyboard.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Select&nbsp;<strong>Change a password</strong>&nbsp;from the menu.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Enter your old password, and your new password twice and press the enter key or the arrow button.</p>
                                <hr>
"@ 

$Instruct365 = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Microsoft365.png" alt="Microsoft 365 Icon">&nbsp;Reset your password in a web browser with&nbsp;<strong>Microsoft 365</strong>:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;Open a web browser on your workstation.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Go to&nbsp<a href="https://www.microsoft365.com" target="_blank">https://www.microsoft365.com</a>&nbspand if needed sign-in with your account.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Click on your avatar icon in the upper right corner of the screen.</p>
                                <p class="tab"><strong>4.</strong>&nbsp;Select&nbsp;<strong>Account Info</strong>&nbsp;from the dropdown menu.</p>
                                <p class="tab"><strong>5.</strong>&nbsp;Navigate to the&nbsp;<strong>Security</strong>&nbsp;section.</p>
                                <p class="tab"><strong>6.</strong>&nbsp;Click on&nbsp;<strong>Change Password</strong>.</p>
                                <p class="tab"><strong>7.</strong>&nbsp;Follow the prompts to enter your current password and your new password.</p>
                                <p class="tab"><strong>8.</strong>&nbsp;Confirm your new password and click&nbsp;<strong>Submit</strong>&nbsp;to complete the process.</p>
                                <hr>
"@

$InstructSSPR = @"
								<p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/UserShield.png" alt="User Shield Icon">&nbsp;Reset your password in a web browser using&nbsp;<strong>Microsoft Self-Service Password Reset</strong>:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;Go to&nbsp;<a href="https://passwordreset.microsoftonline.com" target="_blank">https://passwordreset.microsoftonline.com</a></p>
                                <p class="tab"><strong>2.</strong>&nbsp;Enter your username and follow the prompts to verify your identity.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Once verified, enter your new password and confirm it.</p>
                                <p class="tab"><strong>4.</strong>&nbsp;Click&nbsp;<strong>Submit</strong>&nbsp;to complete the password reset process.</p>
                                <hr>
"@ 

$InstructSSPRButton = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Desktop.png" alt="icon-desktop-25" border="0">&nbsp;Reset your password using the&nbsp;<strong>Password Reset button on the lock screen</strong>:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;On the lock screen, click the&nbsp;<strong>Password Reset</strong>&nbsp;button.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Enter your username and follow the prompts to verify your identity.</p>
                                <p class="tab"><strong>3.</strong>&nbsp;Once verified, enter your new password and confirm it.</p>
                                <p class="tab"><strong>4.</strong>&nbsp;Click&nbsp;<strong>Submit</strong>&nbsp;to complete the password reset process.</p>
                                <hr>
"@ 

$InstructPhone = @"
                                <p><img src="https://raw.githubusercontent.com/AU-Mark/PS-PasswordExpiryNotification/main/Source%20Files/Icons/Phone.png" alt="Phone Icon">&nbsp;<strong>Call Aunalytics</strong>&nbsp;and request a password reset:</p>
                                <p class="tab"><strong>1.</strong>&nbsp;Call Aunalytics at phone number&nbsp;<strong>1-855-799-DATA (1-855-799-3282)</strong>.</p>
                                <p class="tab"><strong>2.</strong>&nbsp;Request a password reset (Aunalytics will verify your identity before resetting your password)</p>
"@ 

$HTMLEnd = @"
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td class="footer" style="border-bottom-left-radius: 8px; border-bottom-right-radius: 8px;">
                            &copy; 2025 $($clientConfig["ClientName"]). All rights reserved.
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@

# Customize the HTML based on the clients configuration
$EmailBody = $HTMLBegin

If ($clientConfig["ClientLogo"] -like "http*") {
    $EmailBody = $EmailBody + $clientConfig["ClientLogo"] + $BodyBegin
} Else {
    If ($Test) {
        $CompanyLogoFile = "file:///$($clientConfig['ClientLogo'])"
        $EmailBody = $EmailBody + "$CompanyLogoFile" + $BodyBegin
    } Else {
        $CompanyLogoFile = Split-Path -Path $clientConfig['ClientLogo'] -Leaf
        $EmailBody = $EmailBody + "cid:$CompanyLogoFile" + $BodyBegin
    }
}

# Does the client have a AD domain?
If ($clientConfig["ClientDomain"]) {
    # Does the client have a VPN?
    If ($clientConfig["ClientVPN"] -eq $False) {
        $EmailBody = $EmailBody + $InstructDomain
    } Else {
        $EmailBody = $EmailBody + $InstructDomainVPN
    }
} Else {
    $EmailBody = $EmailBody + $InstructEntra
}

If ($clientConfig["ClientSSPR"] -eq $True) {
    $EmailBody = $EmailBody + $InstructSSPR
}

If ($clientConfig["ClientSSPRLockScreen"] -eq $True) {
    $EmailBody = $EmailBody + $InstructSSPRButton
}

If ($clientConfig["ClientAzure"] -eq $True) {
    $EmailBody = $EmailBody + $Instruct365
}

$EmailBody = $EmailBody + $InstructPhone

$EmailBody = $EmailBody + $HTMLEnd