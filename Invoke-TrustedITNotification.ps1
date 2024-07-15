<#
    .SYNOPSIS
        Sends a Trusted IT-branded toast notification.
#>



[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Reboot", "ServerProblem")]
    [string]
    $Preset = "Reboot",

    [Parameter(Mandatory = $false)]
    [string]
    $LogFile = "C:\ProgramData\Trusted IT\Logs\Toast.log"
)



# ===[ Start Logging ]===
if (-not (Test-Path -Path $LogFile)) {
    New-Item -Path $LogFile -ItemType File -Force
}
Start-Transcript -Path $LogFile -Append -UseMinimalHeader
Write-Output "===[ NEW RUN - $(Get-Date -Format dd-MM-yyyy_hh-mm-ss) ]==="



# ===[ Verify Technician Input ]===
Write-Output "Checking technician input."
$AcceptedPresets = @("Reboot", "ServerProblem")
if ("$Preset" -notin $AcceptedPresets) {
    Write-Output "Not an accepted preset. Try one of: $AcceptedPresets."
    Stop-Transcript
    exit 1
}
else {
    Write-Output "The preset '$Preset' is valid. Continuing..."
}



# ===[ Create a Protocol ]===
Write-Output "Checking protocol handlers."

# Restarts immediately when the user clicks "Reboot Now" on the "Reboot" preset
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
$Handler = Get-Item "HKCR:\TrustedITReboot" -ErrorAction SilentlyContinue
if (-not $Handler) {
    Write-Output "TrustedITReboot handler not found. Creating..."
    New-Item -Path "HKCR:\TrustedITReboot" -Force
    Set-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "(DEFAULT)" -Value "TrustedITReboot" -Force
    Set-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "URL Protocol" -Value "" -Force
    New-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "EditFlags" -PropertyType DWORD -Value 2162688
    
    New-Item -Path "HKCR:\TrustedITReboot\Shell\Open\Command" -Force
    Set-ItemProperty "HKCR:\TrustedITReboot\Shell\Open\Command" -Name "(DEFAULT)" -Value "C:\Windows\System32\shutdown.exe -r -t 00" -Force
    Write-Output "TrustedITReboot handler has been created. Continuing..."
}
else {
    Write-Output "TrustedITReboot handler already exists. Continuing..."
}



# ===[ Install Modules ]===
Write-Output "Checking for modules."
if (-not(Get-Module -Name BurntToast -ListAvailable)) {
    try {
        Write-Output "BurntToast module is not installed. Installing..."
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name BurntToast -Scope CurrentUser | Out-Null
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted
        Write-Output "BurntToast module has been installed. Continuing..."
    }
    catch {
        Write-Output "Could not install module: BurntToast: $_"
        Stop-Transcript
        exit 2
    }
}
Import-Module -Name BurntToast



# ===[ Download Images ]===
Write-Output "Checking images."
try {
    if (-not (Test-Path -Path "C:\ProgramData\Trusted IT")) {
        Write-Output "Creating working folder..."
        New-Item -Path "C:\ProgramData\Trusted IT" -ItemType Directory
    }
    $WorkingFolder = "C:\ProgramData\Trusted IT"
}
catch {
    Write-Output "Could not create working folder: $_"
    Stop-Transcript
    exit 3
}


try {
    if (-not (Test-Path -Path "$WorkingFolder\TrustedITLogo_Hero.jpg")) {
        Write-Output "Downloading hero image..."
        $HeroURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITLogo_Hero.jpg"
        Invoke-WebRequest -Uri $HeroURL -OutFile "$WorkingFolder\TrustedITLogo_Hero.jpg"
        Write-Output "Downloaded hero image successfully."
    }
    $HeroPath = "$WorkingFolder\TrustedITLogo_Hero.jpg"
}
catch {
    Write-Output "Could not download hero image: $_"
    Stop-Transcript
    exit 4
}


try {
    if (-not (Test-Path -Path "$WorkingFolder\TrustedITMSP_Icon.png")) {
        Write-Output "Downloading MSP image..."
        $IconURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITMSP_Icon.png"
        Invoke-WebRequest -Uri $IconURL -OutFile "$WorkingFolder\TrustedITMSP_Icon.png"
        Write-Output "Downloaded MSP image successfully."
    }
    $IconPath   = "$WorkingFolder\TrustedITMSP_Icon.png"
}
catch {
    Write-Output "Could not download app logo image: $_"
    Stop-Transcript
    exit 5
}


try {
    if (-not (Test-Path -Path "$WorkingFolder\TrustedITIcon.ico")) {
        Write-Output "Downloading app icon image..."
        $IconURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITIcon.ico"
        Invoke-WebRequest -Uri $IconURL -OutFile "$WorkingFolder\TrustedITIcon.ico"
        Write-Output "Downloaded app icon image successfully."
    }
}
catch {
    Write-Output "Could not download notification handler image: $_"
    Stop-Transcript
    exit 6
}



# ===[ Set Custom Notification App ]===
Write-Host "Checking notification handler app."
if (-not (Test-Path -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\TrustedIT.Notifications")) {
    try {
        Write-Output "Creating notification handler app 'TrustedIT.Notifications'..."
        New-BTAppId -AppId "TrustedIT.Notifications"
        
        $HKCR = Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue
        if (-not ($HKCR)) { New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -Scope Script | Out-Null }

        $RegPath    = "HKCR:\AppUserModelId"
        $AppIdPath  = "$RegPath\TrustedIT.Notifications"
        if (-not (Test-Path $AppIdPath)) { New-Item -Path $RegPath -Name "TrustedIT.Notifications" -Force | Out-Null }

        $DisplayName = (Get-ItemProperty -Path $AppIdPath -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
        if ($DisplayName -ne "Trusted IT") {
            New-ItemProperty -Path $AppIdPath -Name DisplayName -Value "Trusted IT" -PropertyType String -Force | Out-Null
        }

        $AppIconPath  = (Get-ItemProperty -Path $AppIdPath -Name IconUri -ErrorAction SilentlyContinue).IconUri
        if ($AppIconPath -ne "$WorkingFolder\TrustedITIcon.ico") {
            New-ItemProperty -Path $AppIdPath -Name IconUri -Value "$WorkingFolder\TrustedITIcon.ico" -PropertyType String -Force | Out-Null
        }
        Write-Output "Created notification handler app successfully."
    }
    catch {
        Write-Output "Could not set custom notification app: $_"
        Stop-Transcript
        exit 7
    }
}
else {
    Write-Output "Notification app 'TrustedIT.Notifications' already exists."
}



# ===[ Send Notification ]===
switch ($Preset) {
    "Reboot" {
        try {
            Write-Output "'Reboot' preset selected."
            $5Min = New-BTSelectionBoxItem -Id 5 -Content "5 Minutes"
            $10Min = New-BTSelectionBoxItem -Id 10 -Content "10 Minutes"
            $1Hour = New-BTSelectionBoxItem -Id 60 -Content "1 Hour"
            $4Hour = New-BTSelectionBoxItem -Id 240 -Content "4 Hours"
            $1Day = New-BTSelectionBoxItem -Id 1440 -Content "1 Day"
            $Items = $5Min, $10Min, $1Hour, $4Hour, $1Day
            $SelectionBox = New-BTInput -Id "SnoozeTime" -DefaultSelectionBoxItemId 10 -Items $Items

            $Uptime = ((Get-Date) - (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime).Days
            $BTText1 = New-BTText -Text "Your computer has been powered on for $Uptime days."
            $BTText2 = New-BTText -Text "Please restart your computer when convenient."
            $IconImage = New-BTImage -Source $IconPath -AppLogoOverride -Crop Circle
            $HeroImage = New-BTImage -Source $HeroPath -HeroImage
            $Audio = New-BTAudio -Source ms-winsoundevent:Notification.Looping.Call2
            $ActionButton = New-BTButton -Content "Restart now" -Arguments "TrustedITReboot:" -ActivationType Protocol
            $DismissButton = New-BTButton -Content "Remind me later" -Snooze -Id "SnoozeTime"
            $ButtonHolder = New-BTAction -Buttons $DismissButton, $ActionButton -Inputs $SelectionBox

            $Binding = New-BTBinding -Children $BTText1, $BTText2 -AppLogoOverride $IconImage -HeroImage $HeroImage
            $Visual = New-BTVisual -BindingGeneric $Binding
            $Content = New-BTContent -Visual $Visual -Audio $Audio -Duration Long -Scenario IncomingCall -Actions $ButtonHolder
            Write-Output "Preset options loaded."
        }
        catch {
            Write-Output "Error when setting notification parameters: $_"
            Stop-Transcript
            exit 8
        }


        try {
            Write-Output "Attempting to send notification."
            Submit-BTNotification -Content $Content -AppId "TrustedIT.Notifications"
            Write-Output "Notification submimtted."
        }
        catch {
            Write-Output "Error sending notification to user: $_"
            Stop-Transcript
            exit 9
        }


        Write-Output "Script completed!"
        Stop-Transcript
        exit 0
    }

    "ServerProblem" {
        try {
            Write-Output "'ServerProblem' preset selected."
            $BTText = New-BTText -Text "We're aware of a problem with the server and we're working on it. We'll update you as soon as possible."
            $IconImage = New-BTImage -Source $IconPath -AppLogoOverride -Crop Circle
            $HeroImage = New-BTImage -Source $HeroPath -HeroImage
            $Audio = New-BTAudio -Source ms-winsoundevent:Notification.Looping.Call2
            $ActionButton = New-BTButton -Content "Please keep me updated." -Arguments "TODO"
            $DismissButton = New-BTButton -Content "I understand, thank you." -Dismiss
            $ButtonHolder = New-BTAction -Buttons $ActionButton, $DismissButton

            $Binding = New-BTBinding -Children $BTText -AppLogoOverride $IconImage -HeroImage $HeroImage
            $Visual = New-BTVisual -BindingGeneric $Binding
            $Content = New-BTContent -Visual $Visual -Audio $Audio -Duration Long -Scenario IncomingCall -Actions $ButtonHolder
            Write-Output "Preset options loaded."
        }
        catch {
            Write-Output "Error when setting notification parameters: $_"
            Stop-Transcript
            exit 8
        }


        try {
            Write-Output "Attempting to send notification."
            Submit-BTNotification -Content $Content -AppId "TrustedIT.Notifications"
            Write-Output "Notification submimtted."
        }
        catch {
            Write-Output "Error sending notification to user: $_"
            Stop-Transcript
            exit 9
        }


        Write-Output "Script completed!"
        Stop-Transcript
        exit 0
    }
}
