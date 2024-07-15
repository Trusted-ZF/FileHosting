<#
    .SYNOPSIS
        Sends a Trusted IT-branded toast notification.
#>



[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidateSet("Reboot", "ServerProblem")]
    [string]
    $Preset = "Reboot"
)



# ===[ Verify Technician Input ]===
$AcceptedPresets = @("Reboot", "ServerProblem")
if ("$Preset" -notin $AcceptedPresets) {
    Write-Output "Not an accepted preset. Try one of: $AcceptedPresets."
}



# ===[ Create a Protocol ]===
# Restarts immediately when the user clicks "Reboot Now" on the "Reboot" preset
New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
$Handler = Get-Item "HKCR:\TrustedITReboot" -ErrorAction SilentlyContinue
if (-not $Handler) {
    New-Item -Path "HKCR:\TrustedITReboot" -Force
    Set-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "(DEFAULT)" -Value "TrustedITReboot" -Force
    Set-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "URL Protocol" -Value "" -Force
    New-ItemProperty -Path "HKCR:\TrustedITReboot" -Name "EditFlags" -PropertyType DWORD -Value 2162688
    
    New-Item -Path "HKCR:\TrustedITReboot\Shell\Open\Command" -Force
    Set-ItemProperty "HKCR:\TrustedITReboot\Shell\Open\Command" -Name "(DEFAULT)" -Value "C:\Windows\System32\shutdown.exe -r -t 00" -Force
}



# ===[ Install Modules ]===
if (-not(Get-Module -Name BurntToast -ListAvailable)) {
    try {
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name BurntToast -Scope CurrentUser | Out-Null
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted
    }
    catch {
        Write-Output "Could not install module: BurntToast: $_"
    }
}
Import-Module -Name BurntToast



# ===[ Download Images ]===
try {
    if (-not (Test-Path -Path "C:\ProgramData\Trusted IT")) {
        New-Item -Path "C:\ProgramData\Trusted IT" -ItemType Directory
    }
    $WorkingFolder = "C:\ProgramData\Trusted IT"
}
catch {
    Write-Output "Could not create working folder: $_"
}


try {
    if (-not (Test-Path -Path "$WorkingFolder\TrustedITLogo_Hero.jpg")) {
        $HeroURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITLogo_Hero.jpg"
        Invoke-WebRequest -Uri $HeroURL -OutFile "$WorkingFolder\TrustedITLogo_Hero.jpg"
    }
    $HeroPath = "$WorkingFolder\TrustedITLogo_Hero.jpg"
}
catch {
    Write-Output "Could not download hero image: $_"
}


try {
    if (-not (Test-Path -Path "$WorkingFolder\TrustedITMSP_Icon.png")) {
        $IconURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITMSP_Icon.png"
        Invoke-WebRequest -Uri $IconURL -OutFile "$WorkingFolder\TrustedITMSP_Icon.png"
    }
    $IconPath   = "$WorkingFolder\TrustedITMSP_Icon.png"
}
catch {
    Write-Output "Could not download app logo image: $_"
}



# ===[ Set Custom Notification App ]===
try {
    New-BTAppId -AppId "TrustedIT.Notifications"
    
    $RegPath    = "HKCU:\Software\Classes\AppUserModelId"
    $AppIdPath  = "$RegPath\TrustedIT.Notifications"
    if (-not (Test-Path $AppIdPath)) { New-Item -Path $RegPath -Name "TrustedIT.Notifications" -Force | Out-Null }

    $DisplayName = (Get-ItemProperty -Path $AppIdPath -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
    if ($DisplayName -ne "Trusted IT") {
        New-ItemProperty -Path $AppIdPath -Name DisplayName -Value "Trusted IT" -PropertyType String -Force | Out-Null
    }

    $IconPath = (Get-ItemProperty -Path $AppIdPath -Name IconUri -ErrorAction SilentlyContinue).IconUri
    if ($IconPath -ne "C:\ProgramData\Trusted IT\TrustedITIcon.ico") {
        New-ItemProperty -Path $AppIdPath -Name IconUri -Value "C:\ProgramData\Trusted IT\TrustedITIcon.ico" -PropertyType String -Force | Out-Null
    }
}
catch {
    Write-Output "Could not set custom notification app: $_"
}



# ===[ Send Notification ]===
switch ($Preset) {
    "Reboot" {
        try {
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
        }
        catch {
            Write-Output "Error when setting notification parameters: $_"
        }


        try {
            Submit-BTNotification -Content $Content -AppId "TrustedIT.Notifications"
        }
        catch {
            Write-Output "Error sending notification to user: $_"
        }


        Write-Output "Script completed!"
        exit 0
    }

    "ServerProblem" {
        try {
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
        }
        catch {
            Write-Output "Error when setting notification parameters: $_"
        }


        try {
            Submit-BTNotification -Content $Content -AppId "TrustedIT.Notifications"
        }
        catch {
            Write-Output "Error sending notification to user: $_"
        }


        Write-Output "Script completed!"
        exit 0
    }
}
