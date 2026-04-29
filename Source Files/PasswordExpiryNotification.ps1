<#
.SYNOPSIS
  Checks the password expiration time of AD user accounts and sends out an email to notify them when their password will expire
.DESCRIPTION
  Checks if a users password is going to expire within $ExpireDays and emails the user to notify them to change their password
.INPUTS
  None
.OUTPUTS
  None
.NOTES
  Version:        1.0
  Author:         Mark Newton
  Email:          mark.newton@aunalytics.com
  Creation Date:  03/08/2022
.EXAMPLE
  PowerShell.exe -ExecutionPolicy Bypass -File PasswordChangeNotification.ps1
#>

param (
    [switch] $Test
)

################################################################################################################################
#                                                             Globals                                                          #
################################################################################################################################
# Debug mode will log users whos password are not expiring
$Debug = $False

# This is a special character used with write-color and read-host lines
$zeroWidthSpace = [char]0x200B

################################################################################################################################
#                                                            Functions                                                         #
################################################################################################################################
Function Check-ModuleStatus {
    <#
    .DESCRIPTION
    Checks whether the supplied module name is installed. If not it force installs and imports it, otherwise it just imports it.

    .EXAMPLE
    Check-ModuleStatus -Service "MSOnline"
    OR
    Check-ModuleStatus -Service "MSOnline" -Silent $True
    
    .PARAMETERS
    [String]$Module - Name of module to check installation and import status
    [Boolean]$Silent - Flag to display warning if required modules are not installed

    .RETURNS
    [Boolean] - $True if the module is installed and imported or $False 
    #>

    Param(
        [Parameter(Mandatory=$True)][String]$Name,
        [Parameter(Mandatory=$False)][Boolean]$Silent
    )

    if ((Get-PackageProvider).Name -notcontains 'NuGet') {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
        Import-PackageProvider -Name NuGet -Force
    } 

    If (Get-Module -ListAvailable -Name $Name) {
        Import-Module $Name
        Write-Color "Imported module $Name" -Color Green -L
        Return $True
    } Else {
        If ($Silent -eq $True) {
            Install-Module -Name $Name -Force 
            Import-Module $Name
            Write-Color "Installed and imported module $Name" -Color Green -L
            Return $True
        } Else {
            Write-Host "WARNING: $Name module is not installed. It will need to be installed with an admin PowerShell session before continuing"
            Return $False
        }
    }
    Return $False
}

Function Write-Color {
    <#
    .SYNOPSIS
    Write-Color is a wrapper around Write-Host delivering a lot of additional features for easier color options.

    .DESCRIPTION
    Write-Color is a wrapper around Write-Host delivering a lot of additional features for easier color options.

    It provides:
    - Easy manipulation of colors,
    - Logging output to file (log)
    - Nice formatting options out of the box.
    - Ability to use aliases for parameters

    .PARAMETER Text
    Text to display on screen and write to log file if specified.
    Accepts an array of strings.

    .PARAMETER Color
    Color of the text. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.
    Available colors are: Black, Blue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White

    .PARAMETER BackGroundColor
    Color of the background. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.
    Available colors are: Black, Blue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White

    .PARAMETER HorizontalCenter
    Calculates the window width and inserts spaces to make the text center according to the present width of the powershell window. Default is false.

    .PARAMETER VerticalCenter
    Calculates the window height and inserts newlines to make the text center according to the present height of the powershell window. Default is false.

    .PARAMETER StartTab
    Number of tabs to add before text. Default is 0.

    .PARAMETER LinesBefore
    Number of empty lines before text. Default is 0.

    .PARAMETER LinesAfter
    Number of empty lines after text. Default is 0.

    .PARAMETER StartSpaces
    Number of spaces to add before text. Default is 0.

    .PARAMETER LogFile
    Path to log file. If not specified no log file will be created.

    .PARAMETER DateTimeFormat
    Custom date and time format string. Default is yyyy-MM-dd HH:mm:ss

    .PARAMETER LogTime
    If set to $true it will add time to log file. Default is $true.

    .PARAMETER LogRetry
    Number of retries to write to log file, in case it can't write to it for some reason, before skipping. Default is 2.

    .PARAMETER Encoding
    Encoding of the log file. Default is Unicode.

    .PARAMETER ShowTime
    Switch to add time to console output. Default is not set.

    .PARAMETER NoNewLine
    Switch to not add new line at the end of the output. Default is not set.

    .PARAMETER NoConsoleOutput
    Switch to not output to console. Default all output goes to console.

    .EXAMPLE
    Write-Color -Text "Red ", "Green ", "Yellow " -Color Red,Green,Yellow

    .EXAMPLE
    Write-Color -Text "This is text in Green ",
                      "followed by red ",
                      "and then we have Magenta... ",
                      "isn't it fun? ",
                      "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan

    .EXAMPLE
    Write-Color -Text "This is text in Green ",
                      "followed by red ",
                      "and then we have Magenta... ",
                      "isn't it fun? ",
                      "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan -StartTab 3 -LinesBefore 1 -LinesAfter 1

    .EXAMPLE
    Write-Color "1. ", "Option 1" -Color Yellow, Green
    Write-Color "2. ", "Option 2" -Color Yellow, Green
    Write-Color "3. ", "Option 3" -Color Yellow, Green
    Write-Color "4. ", "Option 4" -Color Yellow, Green
    Write-Color "9. ", "Press 9 to exit" -Color Yellow, Gray -LinesBefore 1

    .EXAMPLE
    Write-Color -LinesBefore 2 -Text "This little ","message is ", "written to log ", "file as well." `
                -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt" -TimeFormat "yyyy-MM-dd HH:mm:ss"
    Write-Color -Text "This can get ","handy if ", "want to display things, and log actions to file ", "at the same time." `
                -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt"

    .EXAMPLE
    Write-Color -T "My text", " is ", "all colorful" -C Yellow, Red, Green -B Green, Green, Yellow
    Write-Color -t "my text" -c yellow -b green
    Write-Color -text "my text" -c red

    .EXAMPLE
    Write-Color -Text "Testuj?? czy si?? ??adnie zapisze, czy b??d?? problemy" -Encoding unicode -LogFile 'C:\temp\testinggg.txt' -Color Red -NoConsoleOutput

    .NOTES
    Understanding Custom date and time format strings: https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
    Project support: https://github.com/EvotecIT/PSWriteColor
    Original idea: Josh (https://stackoverflow.com/users/81769/josh)

    #>
    [alias('Write-Colour')]
    [CmdletBinding()]
    param (
        [alias ('T')] [string[]]$Text,
        [alias ('C', 'ForegroundColor', 'FGC')][ConsoleColor[]]$Color = [ConsoleColor]::White,
        [alias ('B', 'BGC')][ConsoleColor[]]$BackGroundColor = $null,
        [bool] $VerticalCenter = $False,
        [bool] $HorizontalCenter = $False,
        [alias ('Indent')][int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [int] $StartSpaces = 0,
        [alias ('Logging', 'L')][switch] $Log,
        [alias ('LN')][string] $LogName = "Debug",
        [alias ('LF')][string] $LogFile = "$PSScriptRoot\Logs\$LogName.log",
        [alias ('LL', 'LogLvl')][string] $LogLevel = "INFO",
        [alias ('NLT')][bool] $NoLogTime = $False,
        [Alias('DateFormat', 'TimeFormat', 'Timestamp', 'TS')][string] $DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
        [int] $LogRetry = 2,
        [ValidateSet('unknown', 'string', 'unicode', 'bigendianunicode', 'utf8', 'utf7', 'utf32', 'ascii', 'default', 'oem')][string]$Encoding = 'Unicode',
        [switch] $ShowTime,
        [switch] $NoNewLine,
        [alias('HideConsole', 'NoConsole', 'LogOnly', 'LO')][switch] $NoConsoleOutput
    )
    if (-not $NoConsoleOutput) {
        $DefaultColor = $Color[0]
        if ($null -ne $BackGroundColor -and $BackGroundColor.Count -ne $Color.Count) {
            Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated."
            return
        }
        If ($VerticalCenter) {
            for ($i = 0; $i -lt ([Math]::Max(0, $Host.UI.RawUI.BufferSize.Height / 2) - 1); $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        } # Center the output vertically according to the powershell window size
        if ($LinesBefore -ne 0) {
            for ($i = 0; $i -lt $LinesBefore; $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        } # Add empty line before
        If ($HorizontalCenter) {
            $MessageLength = 0
            ForEach ($Value in $Text) {
                $MessageLength += $Value.Length
            }
            Write-Host ("{0}" -f (' ' * ([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($MessageLength / 2)))) -NoNewline 
        } # Center the line horizontally according to the powershell window size
        if ($StartTab -ne 0) {
            for ($i = 0; $i -lt $StartTab; $i++) {
                Write-Host -Object "`t" -NoNewline 
            } 
        }  # Add TABS before text
        
        if ($StartSpaces -ne 0) {
            for ($i = 0; $i -lt $StartSpaces; $i++) {
                Write-Host -Object ' ' -NoNewline 
            } 
        }  # Add SPACES before text
        if ($ShowTime) {
            Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline -ForegroundColor DarkGray
        } # Add Time before output
        if ($Text.Count -ne 0) {
            if ($Color.Count -ge $Text.Count) {
                # the real deal coloring
                if ($null -eq $BackGroundColor) {
                    for ($i = 0; $i -lt $Text.Length; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewline 
                    }
                } else {
                    for ($i = 0; $i -lt $Text.Length; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewline 
                        
                    }
                }
            } else {
                if ($null -eq $BackGroundColor) {
                    for ($i = 0; $i -lt $Color.Length ; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewline 
                        
                    }
                    for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -NoNewline 
                        
                    }
                }
                else {
                    for ($i = 0; $i -lt $Color.Length ; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewline 
                        
                    }
                    for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackGroundColor[0] -NoNewline 
                        
                    }
                }
            }
        }
        if ($NoNewLine -eq $true) {
            Write-Host -NoNewline 
        }
        else {
            Write-Host 
        } # Support for no new line
        if ($LinesAfter -ne 0) {
            for ($i = 0; $i -lt $LinesAfter; $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        }  # Add empty line after
    }
    if ($Text.Count -and $Log) {
        if (!(Test-Path -Path "$PSScriptRoot\Logs")) {
            New-Item -ItemType "Directory" -Path "$PSScriptRoot\Logs"
        }
    
        if (!(Test-Path -Path "$PSScriptRoot\Logs\$LogName.log")) {
            Write-Output "[$([datetime]::Now.ToString($DateTimeFormat))][INFO] Logging started" | Out-File -FilePath "$PSScriptRoot\Logs\$LogName.log" -Append
        }

        # Save to file
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        $Saved = $false
        $Retry = 0
        Do {
            $Retry++
            try {
                if ($NoLogTime) {
                    "$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$false
                } else {
                    "[$([datetime]::Now.ToString($DateTimeFormat))][$LogLevel] $TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$false
                }
                $Saved = $true
            }
            catch {
                if ($Saved -eq $false -and $Retry -eq $LogRetry) {
                    Write-Warning "Write-Color - Couldn't write to log file $($_.Exception.Message). Tried ($Retry/$LogRetry))"
                }
                else {
                    Write-Warning "Write-Color - Couldn't write to log file $($_.Exception.Message). Retrying... ($Retry/$LogRetry)"
                }
            }
        } Until ($Saved -eq $true -or $Retry -ge $LogRetry)
    }
}

function Validate-URI {
    param (
        [string]$URI
    )

    If ($Null -ne $URI -and $URI -ne "") {
        # Define regex patterns for web URI and file URI
        $webUriPattern = '^https?://'
        $fileUriPattern = '^[a-zA-Z]:\\'

        # Define supported image file extensions
        $supportedImageExtensions = @('.jpg', '.jpeg', '.png', '.gif', '.svg')

        # Check if the input matches a web URI
        if ($URI -match $webUriPattern) {
            try {
                # Create a WebRequest to get the file extension
                $webRequest = [System.Net.WebRequest]::Create($URI)
                $webResponse = $webRequest.GetResponse()
                $contentType = $webResponse.ContentType
                $webResponse.Close()

                # Check if the content type is an image
                if ($contentType -match 'image/(jpeg|png|gif|svg)') {
                    Return $True
                } else {
                    Write-Color "The web URI you entered does not link to a supported image file. Only JPEG, JPG, PNG, GIF, and SVG files are supported. Please use a different link or a file path." -Color Red -LinesAfter 1
                    Return $False
                }
            } catch {
                Write-Color "The web URI you entered is not accessible or does not link to a supported image file. Please use a different link or a file path." -Color Red -LinesAfter 1
                Return $False
            }
        }
        # Check if the input matches a file URI
        elseif ($URI -match $fileUriPattern) {

            # Test if the file exists
            if (Test-Path -Path $URI) {
                # Check if the file extension is supported
                $fileExtension = [System.IO.Path]::GetExtension($URI).ToLower()
                if ($supportedImageExtensions -contains $fileExtension) {
                    Return $True
                } else {
                    Write-Color "The file path you entered does not link to a supported image file. Only JPEG, JPG, PNG, GIF, and SVG files are supported. Please use a different link or a file path." -Color Red -LinesAfter 1
                    Return $False
                }
            } else {
                Write-Color "The file path you entered is not accessible. Please use a different link or a file path." -Color Red -LinesAfter 1
                Return $False
            }
        }
        else {
            Write-Color "You did not enter a valid web URI or a valid file path. Please use a different link or a file path." -Color Red -LinesAfter 1
            Return $False
        }
    }
}

function Validate-Server {
    param (
        [string]$Server
    )

    # Define regex patterns for FQDN, IPv4, and IPv6
    $fqdnPattern = '^(?=.{1,253}$)(?:(?!\d+\.)[a-zA-Z0-9-_]{1,63}\.?)+(?:[a-zA-Z]{2,})$'
    $ipv4Pattern = '^(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    $ipv6Pattern = '^(?:[a-fA-F0-9]{1,4}:){7}[a-fA-F0-9]{1,4}$|^(?:[a-fA-F0-9]{1,4}:){1,7}:$|^(?:[a-fA-F0-9]{1,4}:){1,6}:[a-fA-F0-9]{1,4}$|^(?:[a-fA-F0-9]{1,4}:){1,5}(?::[a-fA-F0-9]{1,4}){1,2}$|^(?:[a-fA-F0-9]{1,4}:){1,4}(?::[a-fA-F0-9]{1,4}){1,3}$|^(?:[a-fA-F0-9]{1,4}:){1,3}(?::[a-fA-F0-9]{1,4}){1,4}$|^(?:[a-fA-F0-9]{1,4}:){1,2}(?::[a-fA-F0-9]{1,4}){1,5}$|^(?:[a-fA-F0-9]{1,4}:){1,6}:(?:[a-fA-F0-9]{1,4}){1,6}$|^:(?::[a-fA-F0-9]{1,4}){1,7}$|^(?:[a-fA-F0-9]{1,4}:){1,7}:$'

    # Validate the input string
    if ($Server -match $fqdnPattern -or $Server -match $ipv4Pattern -or $Server -match $ipv6Pattern) {
        return $true
    } else {
        Write-Color "$prompt is not a valid FQDN or IP Address. Please try again." -Color Red -LinesAfter 1
        return $false
    }
}

function Validate-Email {
    param (
        [string]$emailAddress
    )

    return (Test-EmailAddress $emailAddress).IsValid
}

function Validate-MXRecord {
param (
    [string]$SenderDomain
)

    # Define regex patterns for FQDN, IPv4, and IPv6
    $fqdnPattern = '^(?=.{1,253}$)(?:(?!\d+\.)[a-zA-Z0-9-_]{1,63}\.?)+(?:[a-zA-Z]{2,})$'

    # Validate the domain
    if ($SenderDomain -match $fqdnPattern) {
        return $true
    } else {
        Write-Color "$prompt is not a valid domain name. Please check the domain name you entered and try again." -Color Red -LinesAfter 1 -L -LogLvl "ERROR"
        return $false
    }

    # Validate the MX record
    Try {
        (Find-MxRecord -DomainName $SenderDomain -DNSProvider Google).MX
        Return $True
    } Catch {
        Write-Color "$SenderDomain MX record lookup did not find a DNS record. Please check the domain name you entered and try again" -Color Red -LinesAfter 1 -L -LogLvl "ERROR"
        Return $False
    }
}

function Prompt-Question {
    param (
        [string]$question,
        [System.Collections.Specialized.OrderedDictionary]$answers
    )

    # Output the question
    Write-Color -Text "$question" -Color White -LinesBefore 1

    # Output each possible answer with a number
    $i = 1
    $answerKeys = @()
    foreach ($key in $answers.Keys) {
        Write-Color -Text "$i. ","$key" -Color Yellow,White
        $answerKeys += $key
        $i++
    }

    # Prompt the user to enter a number
    Write-Color -Text "Enter the number of your choice" -Color White -NoNewline -LinesBefore 1; $selection = Read-Host "$zeroWidthSpace"

    # Validate the user input
    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $answers.Count) {
        # Valid selection, return the corresponding answer
        return $answers[$answerKeys[$selection - 1]]
    } else {
        # Invalid selection, prompt again
        Write-Color "Invalid selection. Please try again." -Color Yellow
        Prompt-Question -question $question -answers $answers
    }
}

function Prompt-Input {
    param (
        [string]$PromptMessage,
        [string]$DefaultValue = "",
        [switch]$Required,
        [switch]$ValidateServer,
        [switch]$ValidateURI,
        [switch]$ValidateEmail,
        [switch]$ValidateMX,
        [switch]$Password
    )

    If ($Password) {
        Write-Color -Text "$PromptMessage" -Color White -NoNewline; $prompt = Read-Host "$zeroWidthSpace" -AsSecureString
    } Else {
        Write-Color -Text "$PromptMessage" -Color White -NoNewline; $prompt = Read-Host "$zeroWidthSpace"
    }

    if ([string]::IsNullOrWhiteSpace($prompt)) {
        If ($Required) {
            Write-Color "This is a required input for this script to function. If you need to gather this information you can run this script again later." -Color Yellow
            Prompt-Input -PromptMessage $PromptMessage
        } Else {
            return $DefaultValue
        }
    }

    If ($Password) {
        Write-Color -Text "$PromptMessage second time to confirm" -Color White -NoNewline; $prompt2 = Read-Host "$zeroWidthSpace" -AsSecureString

        # Convert secure strings to plain text
        $plainText1 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($prompt))
        $plainText2 = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($prompt2))
        
        If ($plainText1 -eq $plainText2) {
            Return $plainText1
        } Else {
            Write-Color "The passwords entered did not match! Please try again..." -Color Yellow -LinesBefore 1 -LinesAfter 1
            Prompt-Input -PromptMessage $PromptMessage -Password -Required
        }
    } Else {
        Write-Color "You entered:"," $prompt" -Color White,Green -LinesAfter 1 -LinesBefore 1

        Write-Color -Text "Is this correct"," (","Y","/N)" -Color White,Yellow,Green,Yellow -NoNewline; $verifyprompt = Read-Host "$zeroWidthSpace"
        Write-Color ' '
        switch ($verifyprompt.ToLower()) {
            "y" { 
                If ($ValidateServer) {
                    If (Validate-Server -Server $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateURI) {
                    If (Validate-URI -URI $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateEmail) {
                    If (Validate-Email -emailAddress $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateMX) {
                    If (Validate-MXRecord -SenderDomain $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } Else {
                    return $prompt 
                }
            }
            "n" { 
                Prompt-Input -PromptMessage $PromptMessage 
            }
            default {
                If ($ValidateServer) {
                    If (Validate-Server -Server $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateURI) {
                    If (Validate-URI -URI $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateEmail) {
                    If (Validate-Email -emailAddress $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } ElseIf ($ValidateMX) {
                    If (Validate-MXRecord -SenderDomain $prompt) {
                        return $prompt 
                    } Else {
                        Prompt-Input -PromptMessage $PromptMessage
                    }
                } Else {
                    return $prompt 
                }
            }
        }
    }
    Write-Color " "
}

function Prompt-Integer {
    param (
        [string]$PromptMessage,
        [int]$DefaultValue,
        [switch]$Required
    )

    while ($true) {
        Write-Color -Text "$PromptMessage" -Color White -NoNewline; $prompt = Read-Host "$zeroWidthSpace"

        if ([string]::IsNullOrWhiteSpace($prompt)) {
                Write-Color "No integer was defined. Default value of $DefaultValue will be used." -Color Yellow
                return $DefaultValue
        } elseif ($prompt -match '^\d+$') {
            $integer = [int]$prompt
            Write-Color "You entered:"," $integer" -Color White,Green -LinesAfter 1 -LinesBefore 1
            Write-Color -Text "Is this correct"," (","Y","/N)" -Color White,Yellow,Green,Yellow -NoNewline; $verifyprompt = Read-Host "$zeroWidthSpace"
            switch ($verifyprompt.ToLower()) {
                "y" { return $integer }
                "n" { Prompt-Input -PromptMessage $PromptMessage }
                default { return $integer }
            }
        } else {
            Write-Host "Invalid input. Please enter a valid integer." -ForegroundColor Red
            Prompt-Input -PromptMessage $PromptMessage
        }
    }
}

function Prompt-Bool {
    param (
        [string]$PromptMessage,
        [switch]$DefaultYes
    )

    If ($DefaultYes) {
        Write-Color -Text "$PromptMessage"," (","Y","/N)" -Color White,Yellow,Green,Yellow -NoNewline; $prompt = Read-Host "$zeroWidthSpace"
        switch ($prompt.ToLower()) {
            "y" { return $true }
            "n" { return $false }
            default { return $true }
        }
        Write-Color " "
    } Else {
        Write-Color -Text "$PromptMessage"," (Y/","N",")" -Color White,Yellow,Green,Yellow -NoNewline; $prompt = Read-Host "$zeroWidthSpace"
        switch ($prompt.ToLower()) {
            "y" { return $true }
            "n" { return $false }
            default { return $false }
        }
        Write-Color " "
    }
}

function Get-SMTPService {
    param(
        [string] $SMTPMethod
    )

    Switch ($SMTPMethod) {
        "SMTPAUTH" {
            $question = "What email provider will be used to send the email notifications?"
            $answers = [ordered]@{
                "Office 365" = "Office 365"
                "Gmail" = "Gmail"
                "Zoho" = "Zoho"
                "Outlook" = "Outlook"
                "iCloud" = "iCloud"
                "Other" = "Other"
            }
        
            $SMTPService = Prompt-Question -question $question -answers $answers
        
            # Initialize variables
            $SMTPServer = ""
            $SMTPPort = 0
        
            # Determine SMTP settings based on the service
            switch ($SMTPService) {
                "Gmail" {
                    $SMTPServer = "smtp.gmail.com"
                    $SMTPPort = 587
                }
                "Office 365" {
                    $SMTPServer = "smtp.office365.com"
                    $SMTPPort = 587
                }
                "Zoho" {
                    $SMTPServer = "smtp.zoho.com"
                    $SMTPPort = 587
                }
                "Outlook" {
                    $SMTPServer = "smtp-mail.outlook.com"
                    $SMTPPort = 587
                }
                "iCloud" {
                    $SMTPServer = "smtp.mail.me.com"
                    $SMTPPort = 587
                }
                "Other" {
                    $SMTPServer = Prompt-Input -PromptMessage "Enter the SMTP servers FQDN or IP address" -ValidateServer
                    $SMTPPort = Prompt-Integer -PromptMessage "Enter port used by this SMTP server (Default is 25)" -DefaultValue 25
                }
            }
        }
        "SMTPRELAY" {
            $SMTPServer = Prompt-Input -PromptMessage "Enter the SMTP servers FQDN or IP address" -ValidateServer
            $SMTPPort = Prompt-Integer -PromptMessage "Enter port used by this SMTP server (Default is 25)" -DefaultValue 25
        }
        "SMTPNOAUTH" {
            $question = "Which unauthenticated SMTP service will the email be sent with?"
            $answers = [ordered]@{
                "Office 365 Direct Send" = "Office 365"
                "Gmail Restricted SMTP" = "Gmail"
            }

            $SMTPService = Prompt-Question -question $question -answers $answers

            switch ($SMTPService) {
                "Office 365" {
                    Write-Color "The emails will be sent with the Direct Send method of O365 which requires the client to have static IPs and they have to be added to their DNS SPF record otherwise they will be flagged as SPAM" -Color Yellow -LinesBefore 1 -LinesAfter 1 -L 
                    $SenderDomain = Prompt-Input -PromptMessage "Enter the client's email domain that will send the email (e.g. aunalytics.com)" -ValidateMX
                    $SMTPServer = (Resolve-DnsName -Name $SenderDomain -Type MX -Server 8.8.8.8).NameExchange
                    $SMTPPort = 25
                }
                "Gmail" {
                    $SMTPServer = "aspmx.l.google.com"
                    $SMTPPort = 25
                }
            }
        }
    }

    return $SMTPServer, $SMTPPort
}

Function Create-NewCredential {
    param(
        [switch]$Graph
    )
    
    If ($Graph) {
        $ClientID = Prompt-Input -PromptMessage "Enter the Client/App ID" -Required
        $ClientSecret = Prompt-Input -PromptMessage "Enter the Client Secret" -Password -Required

        # Store the credential in Credential Manager
        New-StoredCredential -Target AUPasswordExpiry -Username $ClientID -Password $ClientSecret -Persist LocalMachine | Out-Null

        Write-Color "The AUPasswordExpiry credential was saved in Credential Manager under account $(whoami.exe)" -Color Green -L -LinesBefore 1

        return $ClientID
    } Else {
        $SenderEmail = Prompt-Input -PromptMessage "Enter the email address for the account we will authenticate with" -ValidateEmail -Required
        $SenderPassword = Prompt-Input -PromptMessage "Enter the password for the account" -Password -Required

        # Store the credential in Credential Manager
        New-StoredCredential -Target AUPasswordExpiry -Username $SenderEmail -Password $SenderPassword -Persist LocalMachine | Out-Null

        Write-Color "The AUPasswordExpiry credential was saved in Credential Manager under account $(whoami.exe)" -Color Green -L -LinesBefore 1
    }
}

Function Get-ClientCredential {
    param(
        [string]$SMTPMethod
    )

    # Are we in a noninteractive session?
    $noninteractive = ([Environment]::GetCommandLineArgs() -contains '-NonInteractive')

    If ($SMTPMethod -eq "SMTPGRAPH") {
        $ClientCredential = Get-StoredCredential -Target AUPasswordExpiry

        If ($Null -eq $ClientCredential) {
            Write-Color "AUPasswordExpiry credential not found in Credential Manager. This credential must exist to send the email notification with the Graph API." -Color Yellow -L -LogLvl "WARNING" -LinesBefore 1
            
            If ($noninteractive) {
                Write-Color "AUPasswordExpiry credential must exist to send the email notification with SMTPGRAPH. Run this script again in an interactive powershell session to recreate the credential. Exiting..." -Color Red -L -LogLvl "ERROR" -LinesBefore 1
                Exit 1
            } Else {
                $NewCredential = Prompt-Bool -PromptMessage "Would you like to create a new Graph API credential now?"
                If ($NewCredential) {
                    Create-NewCredential -Graph
                    $ClientCredential = Get-StoredCredential -Target AUPasswordExpiry
                } Else {
                    Write-Color "AUPasswordExpiry credential must exist to send the email notification with SMTPGRAPH. Run this script again in an interactive powershell session to recreate the credential. Exiting..." -Color Red -L -LogLvl "ERROR" -LinesBefore 1
                    Exit 1
                }
            }
        }

        $ClientSecret = $ClientCredential.GetNetworkCredential().Password
        $Credential = ConvertTo-GraphCredential -ClientID $ClientCredential.UserName -ClientSecret $ClientSecret -DirectoryID $clientConfig["TenantID"]
    } Else {
        $Credential = Get-StoredCredential -Target AUPasswordExpiry

        If ($Null -eq $Credential) {
            Write-Color "AUPasswordExpiry credential not found in Credential Manager. This credential must exist to send the email notification with a dedicated email account." -Color Yellow -L -LogLvl "WARNING" -LinesBefore 1
            
            If ($noninteractive) {
                Write-Color "AUPasswordExpiry credential must exist to send the email notification with SMTPGRAPH. Run this script again in an interactive powershell session to recreate the credential. Exiting..." -Color Red -L -LogLvl "ERROR" -LinesBefore 1
                Exit 1
            } Else {
                $NewCredential = Prompt-Bool -PromptMessage "Would you like to create a new SMTP AUTH email account credential now?"
                If ($NewCredential) {
                    Create-NewCredential
                    $Credential = Get-StoredCredential -Target AUPasswordExpiry
                } Else {
                    Write-Color "AUPasswordExpiry credential must exist to send the email notification with SMTPAUTH. Run this script again in an interactive powershell session to recreate the credential. Exiting..." -Color Red -L -LogLvl "ERROR" -LinesBefore 1
                    Exit 1
                }
            }
        }
    }

    Return $Credential
}

Function Add-ClientConfig {
    # Initialize all variables to nulls
    $ClientName = $Null
    $ClientURL = $Null
    $ClientLogo = $Null
    $ClientDomain = $Null
    $ClientVPN = $Null
    $ClientAzure = $Null
    $ClientSSPR = $Null
    $ClientSSPRLockScreen = $Null
    $ExpireDays = $Null
    $SMTPMethod = $Null
    $SMTPServer = $Null
    $SMTPPort = $Null
    $SMTPTLS = $Null
    $TenantID = $Null
    $SenderEmail = $Null
    $EmailCredential = $Null

    # Prompt the user for client input
    $ClientName = Prompt-Input -PromptMessage "Enter the client's company name (This will be displayed in the email body)" -Required
    $ClientURL = Prompt-Input -PromptMessage "Enter the client's website URL" 
    $ClientLogo = Prompt-Input -PromptMessage "Enter a URL or the file path to a logo image for the client (This logo will be featured at the top of the email body)" -ValidateURI
    $ClientVPN = Prompt-Bool -PromptMessage "Does this client have a VPN for end users?"
    $ClientAzure = Prompt-Bool -PromptMessage "Does this client have Azure P1 or P2 licenses and have Password Writeback enabled in AAD/Entra Connect?"
    $ClientSSPR = Prompt-Bool -PromptMessage "Does this client have Microsoft SSPR (Self Sevrice Password Reset) enabled?"
    $ClientSSPRLockScreen = Prompt-Bool -PromptMessage "Does this client have the Reset Password SSPR lock screen button enabled?"
    $ExpireDays = Prompt-Integer -PromptMessage "How many days before users password expire should they start being notified? (Default is 14 days)" -DefaultValue 14

    # Retrieve the root domain. If it errors then no domain was found.
    Try {
        $ClientDomain = (Get-ADDomain).DNSROOT
    } Catch {
        $ClientDomain = $False
    }

    # Define the questions and answers for the SMTP Send Method
    $question = "What SMTP send method will be used?"
    $answers = [ordered]@{
        "Microsoft 365 Graph API: Requires App Registration setup in the client's tenant (Recommended)" = "SMTPGRAPH"
        "SMTP AUTH: Office 365, Gmail, Zoho, Outlook, iCloud, Other" = "SMTPAUTH"
        "SMTP Relay: Manual Setup" = "SMTPRELAY"
        "Unauthenticated SMTP: Office 365 Direct Send, Gmail Restricted SMTP" = "SMTPNOAUTH"
    }

    # Prompt the question to the user
    $SMTPChoice = Prompt-Question -question $question -answers $answers

    # Configure SMTP Options based on the users choice
    switch ($SMTPChoice) {
        "SMTPAUTH" {
            $SMTPMethod = $SMTPChoice
            $SMTPServer, $SMTPPort = Get-SMTPService -SMTPMethod $SMTPMethod
            $SMTPTLS = "Auto"
            Create-NewCredential
            $SenderEmail = Prompt-Input -PromptMessage "Enter the email address that will send the email" -ValidateEmail
            $EmailCredential = $True
        }
        "SMTPRELAY" {
            $SMTPMethod = $SMTPChoice
            $SMTPServer, $SMTPPort = Get-SMTPService -SMTPMethod $SMTPMethod
            $TLS = Prompt-Bool -PromptMessage "Does this SMTP Relay require TLS?"
            If ($TLS) {
                $SMTPTLS = "Auto"
            } Else {
                $SMTPTLS = "None"
            }
            $DedicatedEmail = Prompt-Bool -PromptMessage "Does this SMTP Relay required user authentication?"
            If ($DedicatedEmail) {
                Create-NewCredential
                $SenderEmail = Prompt-Input -PromptMessage "Enter the email address that will send the email" -ValidateEmail
                $EmailCredential = $True
            } Else {
                $SenderEmail = Prompt-Input -PromptMessage "Enter the email address that will send the email" -ValidateEmail
            }
        }
        "SMTPNOAUTH" {
            $SMTPMethod = $SMTPChoice
            $SMTPServer, $SMTPPort = Get-SMTPService -SMTPMethod $SMTPMethod
            $SMTPTLS = "Auto"
            $SenderEmail = Prompt-Input -PromptMessage "Enter the email address that will send the email" -ValidateEmail
            $EmailCredential = $False
        }
        "SMTPGRAPH" {
            $SMTPMethod = $SMTPChoice
            $TenantID = Prompt-Input -PromptMessage "Enter the client's Tenant ID" -Required
            Create-NewCredential -Graph
            $SMTPTLS = "Auto"
            $SenderEmail = Prompt-Input -PromptMessage "Enter the email address that will send the email" -ValidateEmail
            $EmailCredential = $True
        }
    }

    # Create a hashtable to store the parameters
    $clientConfig = @{
        ClientName = $ClientName
        ClientURL = $ClientURL
        ClientLogo = $ClientLogo
        ClientDomain = $ClientDomain
        ClientVPN = $ClientVPN
        ClientAzure = $ClientAzure
        ClientSSPR = $ClientSSPR
        ClientSSPRLockScreen = $ClientSSPRLockScreen
        ExpireDays = $ExpireDays
        SMTPMethod = $SMTPMethod
        SMTPServer = $SMTPServer
        SMTPPort = $SMTPPort
        SMTPTLS = $SMTPTLS
        TenantID = $TenantID
        SenderEmail = $SenderEmail
        EmailCredential = $EmailCredential
    }

    # Convert the hashtable to JSON
    $json = $clientConfig | ConvertTo-Json -Depth 3

    # Save the JSON to a file
    $json | Out-File -FilePath "$PSScriptRoot\clientconf.json" -Encoding utf8 -Force

    # Return the hashtable content
    Return $clientConfig
}

Function Get-ClientConfig {
    # Define the path to the JSON file
    $jsonFilePath = "$PSScriptRoot\clientconf.json"

    # Check if the file exists
    if (Test-Path -Path $jsonFilePath) {
        Try {
            # Read the JSON file content
            $jsonContent = Get-Content -Path $jsonFilePath -Raw

            # Convert the JSON content to a hashtable
            $clientConfigJson = ConvertFrom-Json -InputObject $jsonContent
            $clientConfig = @{}
            foreach ($property in $clientConfigJson.PSObject.Properties) {
                $clientConfig[$property.Name] = $property.Value
            }

            # Client config loaded successfully
            Write-Color "Client configuration loaded successfully."

            # Validate credentials exist early (fail-fast) - the actual credential will be
            # fetched again in Send-PasswordExpiry when needed
            If ($clientConfig["EmailCredential"]) {
                $null = Get-ClientCredential -SMTPMethod $clientConfig["SMTPMethod"]
            }

            # Return the hashtable content
            Return $clientConfig
        } Catch {
            # Determine if we are running in an non-ineractive session
            $noninteractive = ([Environment]::GetCommandLineArgs() -contains '-NonInteractive')

            If ($noninteractive) {
                Write-Color "Your JSON config file at $jsonFilePath has syntax errors or is corrupted. Check the config file or overwrite by running this script again again and creating a new config in an interactive powershell session." -Color Red -L -LogLvl "ERROR" -NoConsoleOutput
                Exit
            } Else {
                Write-Color "Your JSON config file at $jsonFilePath has syntax errors or is corrupted. Check the config file or overwrite it by running this script again and creating a new config." -Color Red -L -LogLvl "ERROR"
                Exit
            }
        }
    } Else {
        # Determine if we are running in an non-ineractive session
        $noninteractive = ([Environment]::GetCommandLineArgs() -contains '-NonInteractive')

        If ($noninteractive) {
            Write-Color "Cannot create client config json file while running in a non-interactive session. Please launch $PSScriptRoot\PasswordExpiryEmail.ps1 in an interactive powershell session to create a new client configuration" -Color Red -L -LogLvl "ERROR" -NoConsoleOutput
            Exit
        } Else {
            # No client config was found, create a new one
            Write-Color "Client configuration was not found. Creating a new one." -Color Yellow -L -LogLvl "WARNING" -NoConsoleOutput

            # Add client config
            $clientConfig = Add-ClientConfig

            # Return the hashtable content
            Return $clientConfig
        }
    }
}

function Convert-HTMLToPlainText {
    param (
        [string]$htmlContent
    )

    # Remove CSS
    $htmlContent = [regex]::Replace($htmlContent, '<style[^>]*>.*?</style>', '', 'Singleline')

    # Replace &nbsp; with spaces
    $htmlContent = $htmlContent -replace '&nbsp;', ' '

    # Remove text between <title> tags
    $htmlContent = [regex]::Replace($htmlContent, '<title[^>]*>.*?</title>', '', 'Singleline')

    # Replace &copy; with the word "Copyright"
    $htmlContent = $htmlContent -replace '&copy;', 'Copyright'

    # Use regex to remove HTML tags
    $plainText = [regex]::Replace($htmlContent, '<[^>]+>', '')

    # Remove multiple blank lines, allowing only one blank line between paragraphs
    $plainText = [regex]::Replace($plainText, "(\r?\n\s*){2,}", "`n`n")

    return $plainText
}

Function Send-PasswordExpiry {
    <#
    .SYNOPSIS
        Sends a password expiry notification email to a user.
    .DESCRIPTION
        Sends an email notification using either Microsoft Graph API or SMTP depending on configuration.
        Returns $True on success, $False on failure.
    .OUTPUTS
        [Boolean] - $True if email was sent successfully, $False otherwise
    #>
    param(
        [Parameter(Mandatory = $True)] [hashtable]$clientConfig,
        [Parameter(Mandatory = $True)] [string]$EmailRecipient,
        [Parameter(Mandatory = $True)] [string]$EmailSubject,
        [Parameter(Mandatory = $True)] [string]$EmailBody
    )

    If ($clientConfig["EmailCredential"]) {
        $Credential = Get-ClientCredential -SMTPMethod $clientConfig["SMTPMethod"]
    }

    $SMTPTLS = $clientConfig["SMTPTLS"]

    #SMTP server ([string], required)
    $SMTPServer = $clientConfig["SMTPServer"]

    #port ([int], required)
    $Port = $clientConfig["SMTPPort"]

    #sender ([string], required)
    $From = $clientConfig["SenderEmail"]

    #recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
    $RecipientList = @($EmailRecipient)

    #cc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$CCList = [MimeKit.InternetAddressList]::new()
    #$CCList.Add([MimeKit.InternetAddress]"CCRecipient1EmailAddress")

    #bcc list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$BCCList = [MimeKit.InternetAddressList]::new()
    #$BCCList.Add([MimeKit.InternetAddress]"BCCRecipient1EmailAddress")

    #subject ([string], optional)
    $Subject = [string]$EmailSubject

    #text body ([string], optional)
    $TextBody = Convert-HTMLToPlainText -htmlContent $EmailBody

    #HTML body ([string], optional)
    $HTMLBody = $EmailBody

    #attachment list ([System.Collections.Generic.List[string]], optional)
    $AttachmentList = [System.Collections.Generic.List[string]]::new()
    If ($clientConfig["ClientLogo"] -notlike "http*") {
        $AttachmentList.Add($clientConfig["ClientLogo"])
    }

    # Determine send method and build parameters
    $SendMethod = $clientConfig["SMTPMethod"]

    If ($clientConfig["EmailCredential"]) {
        If ($SendMethod -eq "SMTPGRAPH") {
            Write-Color "Sending email via Microsoft Graph API to $EmailRecipient" -Color Cyan -L -LogLvl "INFO"
            $Parameters = @{
                "Credential" = $Credential
                "From" = $From
                "To" = $RecipientList
                "Subject" = $Subject
                "Text" = $TextBody
                "HTML" = $HTMLBody
                "Attachment" = $AttachmentList
                "Priority" = "High"
                "Graph" = $True
            }
        } Else {
            Write-Color "Sending email via SMTP ($SMTPServer`:$Port) to $EmailRecipient" -Color Cyan -L -LogLvl "INFO"
            $Parameters = @{
                "SecureSocketOptions" = $SMTPTLS
                "Credential" = $Credential
                "Server" = $SMTPServer
                "Port" = $Port
                "From" = $From
                "To" = $RecipientList
                "Subject" = $Subject
                "Text" = $TextBody
                "HTML" = $HTMLBody
                "Attachment" = $AttachmentList
                "Priority" = "High"
            }
        }
    } Else {
        Write-Color "Sending email via unauthenticated SMTP ($SMTPServer`:$Port) to $EmailRecipient" -Color Cyan -L -LogLvl "INFO"
        $Parameters = @{
            "SecureSocketOptions" = $SMTPTLS
            "Server" = $SMTPServer
            "Port" = $Port
            "From" = $From
            "To" = $RecipientList
            "Subject" = $Subject
            "Text" = $TextBody
            "HTML" = $HTMLBody
            "Attachment" = $AttachmentList
            "Priority" = "High"
        }
    }

    # Send message and return success/failure status
    Try {
        $Result = Send-EmailMessage @Parameters -ErrorAction Stop

        # Mailozaurr returns a result object with Status/Error properties
        # Graph API failures (e.g. missing Mail.Send permission) do NOT throw exceptions
        If ($Result.Status) {
            Write-Color "Email successfully sent to $EmailRecipient" -Color Green -L -LogLvl "INFO"
            Return $True
        } Else {
            Write-Color -Text "FAILED to send email to $EmailRecipient - ","Error: ","$($Result.Error)" -Color Red,Yellow,Magenta -L -LogLvl "ERROR"
            Return $False
        }
    } Catch {
        Write-Color -Text "FAILED to send email to $EmailRecipient - ","Err Line: ","$($_.InvocationInfo.ScriptLineNumber) ","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Yellow,Magenta,Yellow,Magenta,Yellow,Magenta -L -LogLvl "ERROR"
        Return $False
    }
}

Function Check-ProgramInstalled {
    # Function to check if a program is installed in both x86 and x64 registry paths
    param (
        [string]$ProgramName
    )
    $x86Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $x64Path = "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"

    $installedX86 = Get-ItemProperty -Path $x86Path | Where-Object { $_.DisplayName -like "*$ProgramName*" }
    $installedX64 = Get-ItemProperty -Path $x64Path | Where-Object { $_.DisplayName -like "*$ProgramName*" }

    return ($installedX86 -ne $null -or $installedX64 -ne $null)
}

function Check-LogonAsBatchJobRights {
    # Function to check if the current user has "Log on as a batch job" rights# Function to check if the current user has "Allow log on as a batch job" rights

    # Export the local security policy
    secedit /export /cfg "$env:TEMP\secpol.cfg" | Out-Null

    # Read the policy file
    $policy = Get-Content "$env:TEMP\secpol.cfg"

    # Check if the user is listed in the "SeBatchLogonRight" policy
    $hasRights = $policy -match "SeBatchLogonRight = .*$env:USERNAME"

    # Clean up the exported policy file
    Remove-Item "$env:TEMP\secpol.cfg"

    return $hasRights
}

################################################################################################################################
#                                                               Main                                                           #
################################################################################################################################
Try {
    If ((Check-ModuleStatus -Name "CredentialManager" -Silent $True) -eq $False) {
        Write-Color "CredentialManager module was not found. Please install this module for this script to run properly. Exiting script..." -Color Red -L -LinesBefore 1
        Exit
    } 

    If ((Check-ModuleStatus -Name "ActiveDirectory" -Silent $True) -eq $False) {
        Write-Color "ActiveDirectory module was not found. Please install this module for this script to run properly. Exiting script..." -Color Red -L -LinesBefore 1
        Exit
    }

    If ((Check-ModuleStatus -Name "Mailozaurr" -Silent $True) -eq $False) {
        Write-Color "Mailozaurr module was not found. Please install this module for this script to run properly. Exiting script..." -Color Red -L -LinesBefore 1
        Exit
    }

    If (-not (Check-LogonAsBatchJobRights)) {
        Write-Color "The current user $ENV:Username does not have `"Log on as a batch job`" rights which are required for the scheduled task to run as the current logged on user. This may be configured in the local security policy or via GPO. If on a domain controller its typically configured in the Default Domain Controllers Policy GPO. Please ensure these rights are granted and then run this script again." -Color Red -L -LogLvl "ERROR" -LinesBefore 1
        Exit 1
    }

    Write-Color "Checking for client config..." -Color White -L -LinesBefore 1
    $clientConfig = Get-ClientConfig

    $Users = Get-ADUser -Filter * -Properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress, sAMAccountName | Where-Object {$_.Enabled -eq "True"} | Where-Object { $_.PasswordNeverExpires -eq $false } | Where-Object { $_.passwordexpired -eq $false }
    $DomainMaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    # Process Each User for Password Expiry
    ForEach ($User in $Users) {
        # Get the users name and correct the last, first naming convention from Trillum's AD
        $LastName = $User.Surname
        $FirstName = $User.GivenName
        $FullName = "$FirstName $LastName"

        $UserEmail = $User.EmailAddress
        $UserPasswordLastSet = $User.PasswordLastSet
        $PasswordPolicy = (Get-ADUserResultantPasswordPolicy $User)
        $Username = $User.sAMAccountName

        # Check for User Password Policy
        If ($null -ne $PasswordPolicy){
            $MaxPasswordAge = ($PasswordPolicy).MaxPasswordAge
        } Else {
            # No User Password Policy, set to Domain Default
            $MaxPasswordAge = $DomainMaxPasswordAge
        }

        # Calculate the expiration date based on when the user last reset their password and the max password age before expiration
        $ExpiresOn = $UserPasswordLastSet + $MaxPasswordAge
        $Today = (Get-Date)
        $DaysToExpire = (New-TimeSpan -Start $Today -End $ExpiresOn).Days

        If ($DaysToExpire -gt 1) {
            $ExpiryMsg = "in $daystoexpire days"
            $EmailSubject="Your $($clientConfig["ClientName"]) password will expire in $ExpiryMsg."
        } Else {
            $ExpiryMsg = "today"
            $EmailSubject="Your $($clientConfig["ClientName"]) password will expire $ExpiryMsg."
        }

        # If a user has no email address listed
        If ($null -eq $UserEmail) {
            Write-Color "No email address was found for $Username" -Color Yellow -L -LogLvl 'WARNING'
            Continue 
        }

        # If Test switch, output HTML from first user run and then exit script
        if ($Test) {
            # Dot source the PowerShell HTML file
            . ("$PSScriptRoot\PasswordExpiryHTML.ps1")

            # Create temp HTML file
            $tempFilePath = [System.IO.Path]::GetTempFileName()
            $tempFilePath = "$tempFilePath.html"
            # Dump the EmailBody into the html temp file
            Set-Content -Path $tempFilePath -Value $EmailBody

            # Open the HTML file with Chrome or Edge browser for viewing
            $edgeInstalled = Check-ProgramInstalled -ProgramName "Microsoft Edge"
            $chromeInstalled = Check-ProgramInstalled -ProgramName "Google Chrome"
            if ($edgeInstalled) {
                Start-Process "msedge.exe" -ArgumentList $tempFilePath
            } elseif ($chromeInstalled) {
                Start-Process "chrome.exe" -ArgumentList $tempFilePath
            }

            Exit 0
        } 

        # Send Email Message
        If (($DaysToExpire -ge 0) -and ($DaysToExpire -le $clientConfig["ExpireDays"])) {
            # Dot source the PowerShell HTML file
            . ("$PSScriptRoot\PasswordExpiryHTML.ps1")

            # Send the email and check if it was successful
            $EmailSent = Send-PasswordExpiry -clientConfig $clientConfig -EmailRecipient $UserEmail -EmailSubject $EmailSubject -EmailBody $EmailBody

            If ($EmailSent) {
                Write-Color "Password for $FullName will expire $ExpiryMsg. Notification email sent to user at $UserEmail" -Color Green -L -NoConsoleOutput
            } Else {
                Write-Color "Password for $FullName will expire $ExpiryMsg. FAILED to send notification email to $UserEmail" -Color Red -L -LogLvl "ERROR" -NoConsoleOutput
            }
        } Else {
            # Log Non Expiring Password
            If ($Debug -eq $True) {
                #Write-Host "No password expiration was found for $Username"
                Write-Color "No password expiration was found for $Username. There password will expire $ExpiryMsg" -Color White -L -LogLvl 'DEBUG'
            }
            Continue
        }
    } 
} Catch {
    Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber) ","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -L -LogLvl "ERROR"
    Exit 1
}

