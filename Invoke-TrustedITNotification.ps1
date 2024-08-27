<#
    .SYNOPSIS
    Sends a Trusted IT-branded balloon notification.

    .PARAMETER Preset
    Uses a preset notification template.
#>

[CmdletBinding()]
param (
    # Uses a preset notification template.
    [ValidateSet("Reboot", "ServerProblem")]
    [string]$Preset
)



# Validate parameters
$Preset = $Preset.ToLower()



# Create working folder and start logging
$WorkingFolder = "C:\ProgramData\Trusted IT"
if (-not (Test-Path -Path $WorkingFolder)) {
    New-Item -Path $WorkingFolder -ItemType Directory
}

$LogFolder = "C:\ProgramData\Trusted IT\Logs\Notifications"
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory
}
$LogFile = New-Item -Path "$LogFolder\User-$(Get-Date -Format yyyy-MM-dd_hh-mm-ss).log"
Start-Transcript -Path $LogFile
Write-Output "New notification run started."



# Install required BurntToast module
if (-not (Get-Module -Name "BurntToast" -ListAvailable)) {
    try {
        Write-Output "Module BurntToast is not available. Installing..."

        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name "BurntToast" -Scope CurrentUser | Out-Null
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted

        Write-Output "Done."
    }
    catch {
        Write-Output "[ERROR]: Could not install BurntToast module: $($_.Exception.Message)"
    }
}



# Set image locations
$HeroImage = "$WorkingFolder\Hero-TrustedITLogo.jpg"
$IconImage = "$WorkingFolder\Icon-TrustedITMSP.png"



# Handle templates
switch ($Preset) {
    "reboot"		{
        try {
            $5Min		= New-BTSelectionBoxItem -Id 5 -Content "5 Minutes"
            $10Min		= New-BTSelectionBoxItem -Id 10 -Content "10 Minutes"
            $1Hour		= New-BTSelectionBoxItem -Id 60 -Content "1 Hour"
            $4Hour		= New-BTSelectionBoxItem -Id 240 -Content "4 Hours"
            $1Day		= New-BTSelectionBoxItem -Id 1440 -Content "1 Day"
            $Items		= $5Min, $10Min, $1Hour, $4Hour, $1Day
            $Selection	= New-BTInput -Id "SnoozeTime" -DefaultSelectionBoxItemId 10 -Items $Items

            $Uptime		= ((Get-Date) - (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime).Days
            $BTText1	= New-BTText -Text "Your computer has been powered on for $Uptime days."
            $BTText2	= New-BTText -Text "Please restart your computer when convenient."
            $BTHero		= New-BTImage -Source $HeroImage -HeroImage
            $BTIcon		= New-BTImage -Source $IconImage -AppLogoOverride -Crop Circle
            $BTAudio	= New-BTAudio -Source ms-winsoundevent:Notification.Looping.Call2

            $Action		= New-BTButton -Content "Restart now" -Arguments "TrustedIT.Reboot:" -ActivationType Protocol
            $Dismiss	= New-BTButton -Content "Snooze" -Snooze -Id "SnoozeTime"
            $BTHolder	= New-BTAction -Buttons $Dismiss, $Action -Inputs $Selection

            $BTBinding	= New-BTBinding -Children $BTText1, $BTText2 -HeroImage $BTHero -AppLogoOverride $BTIcon
            $BTVisual	= New-BTVisual -BindingGeneric $BTBinding
            $BTContent	= New-BTContent -Visual $BTVisual -Audio $BTAudio -Duration Long -Scenario IncomingCall -Actions $BTHolder
            Submit-BTNotification -Content $BTContent -AppId "TrustedIT.Notifications"
            Write-Output "Notification submitted successfully."
        }
        catch {
            Write-Output "[ERROR]: Could not submit notification: $($_.Exception.Message)"
        }
    }
    "serverproblem" {
        try {
            $BTText     = New-BTText -Text "We're aware of a problem with the server and we're working on it. We'll update you as soon as possible."
            $BTHero     = New-BTImage -Source $HeroImage -HeroImage
            $BTIcon     = New-BTImage -Source $IconImage -AppLogoOverride -Crop Circle
            $BTAudio    = New-BTAudio -Source ms-winsoundevent:Notification.Looping.Call2

            $Action     = New-BTButton -Content "Please keep me updated." -Arguments "none"
            $Dismiss    = New-BTButton -Content "I understand, thank you." -Dismiss
            $BTHolder   = New-BTAction -Buttons $Dismiss, $Action

            $BTBinding  = New-BTBinding -Children $BTText -HeroImage $BTHero -AppLogoOverride $BTIcon
            $BTVisual   = New-BTVisual -BindingGeneric $BTBinding
            $BTContent  = New-BTContent -Visual $BTVisual -Audio $BTAudio -Duration Long -Scenario IncomingCall -Actions $BTHolder
            Submit-BTNotification -Content $BTContent -AppId "TrustedIT.Notifications"
            Write-Output "Notification submitted successfully."
        }
        catch {
            Write-Output "[ERROR]: Could not submit notification: $($_.Exception.Message)"
        }
    }
}

Stop-Transcript
