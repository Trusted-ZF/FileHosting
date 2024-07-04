<#
    .SYNOPSIS
        Sends a Trusted IT-branded toast notification.
#>


# ===[ Verify Technician Input ]===
$AcceptedPresets = @("Reboot")
if ("@Preset@" -notin $AcceptedPresets) {
    Write-Output "Not an accepted preset. Try one of: $AcceptedPresets."
}


# ===[ Variables ]===
$Preset = "@Preset@"


# ===[ Install Modules ]===
if (-not(Get-Module -Name BurntToast -ListAvailable)) {
    try {
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name BurntToast -Scope AllUsers -AcceptLicense:$true | Out-Null
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted
    }
    catch {
        Write-Output "Could not install module: BurntToast: $_"
    }
}
Import-Module -Name BurntToast


# ===[ Download Images ]===
try {
    $HeroURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITLogo_Hero.jpg"
    Invoke-WebRequest -Uri $HeroURL -OutFile "$env:TEMP\TrustedITLogo_Hero.jpg"
    $HeroPath = "$env:TEMP\TrustedITLogo_Hero.jpg"
}
catch {
    Write-Output "Could not download hero image: $_"
}

try {
    $IconURL    = "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITMSP_Icon.png"
    Invoke-WebRequest -Uri $IconURL -OutFile "$env:TEMP\TrustedITMSP_Icon.png"
    $IconPath   = "$env:TEMP\TrustedITMSP_Icon.png"
}
catch {
    Write-Output "Could not download app logo image: $_"
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
            $ActionButton = New-BTButton -Content "Restart now" -Arguments "TODO"
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
            Submit-BTNotification -Content $Content
        }
        catch {
            Write-Output "Error sending notification to user: $_"
        }

        Write-Output "Script completed! $_"
        exit 0
    }
}
