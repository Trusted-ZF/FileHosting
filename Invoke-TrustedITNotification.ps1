<#
    .SYNOPSIS
    Sends a Trusted IT-branded balloon notification.

    .PARAMETER Preset
    Uses a preset notification template.
#>

[CmdletBinding()]
param (
    # Uses a preset notification template.
    [ValidateSet("ADSyncAlert", "Reboot", "ServerProblem")]
    [string]$Preset
)



# Install required BurntToast module
if (-not (Get-Module -Name "BurntToast" -ListAvailable)) {
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module -Name "BurntToast" -Scope CurrentUser -AcceptLicense:$true | Out-Null
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Untrusted
}



# Create working folder
$WorkingFolder = "C:\ProgramData\Trusted IT"
if (-not (Test-Path -Path $WorkingFolder)) {
    New-Item -Path $WorkingFolder -ItemType Directory
}



# Hero image
$HeroImage = "$WorkingFolder\Hero-TrustedITLogo.jpg"
if (-not (Test-Path -Path $HeroImage)) {
    $HeroURL = "https://raw.githubusercontent.com/Trusted-ZF/FileHosting/main/TrustedITLogo_Hero.jpg"
    Invoke-WebRequest -Uri $HeroURL -OutFile $HeroImage
}



# Icon image
$IconImage = "$WorkingFolder\Icon-TrustedITMSP.png"
if (-not (Test-Path -Path $IconImage)) {
    $IconURL = "https://raw.githubusercontent.com/Trusted-ZF/FileHosting/main/TrustedITMSP_Icon.png"
    Invoke-WebRequest -Uri $IconURL -OutFile $IconImage
}



# Notification App image
$NotifyIcon = "$WorkingFolder\AppIcon-TrustedIT.ico"
if (-not (Test-Path -Path $NotifyIcon)) {
    $AppURL	= "https://raw.githubusercontent.com/Trusted-ZF/ToastNotifications/main/TrustedITIcon.ico"
    Invoke-WebRequest -Uri $AppURL -OutFile $NotifyIcon
}



# Clear old reboot handlers
New-PSDrive -Name "HKCR" -PSProvider Registry -Root HKEY_CLASSES_ROOT -ErrorAction SilentlyContinue | Out-Null
$HandlerPaths = @("HKCR:\TrustedIT-Reboot", "HKCR:\TrustedITReboot")
$HandlerPaths | ForEach-Object {
    Remove-Item -Path $_ -Recurse -Force -ErrorAction SilentlyContinue
}



# Create reboot handler
$HandlerPath = "HKCR:\TrustedIT.Reboot"
$Handler = Get-Item $HandlerPath -ErrorAction SilentlyContinue
if (-not $Handler) {
    New-Item -Path $HandlerPath -Force
    Set-ItemProperty -Path $HandlerPath -Name "(DEFAULT)" -Value "url:TrustedIT.Reboot" -Force
    Set-ItemProperty -Path $HandlerPath -Name "URL Protocol" -Value "" -Force
    New-ItemProperty -Path $HandlerPath -Name "EditFlags" -PropertyType DWORD -Value 2162688
    New-Item -Path "$HandlerPath\Shell\Open\Command" -Force
    Set-ItemProperty "$HandlerPath\Shell\Open\Command" -Name "(DEFAULT)" -Value "C:\Windows\System32\shutdown.exe -r -t 00" -Force
}



# Create notification handler app
New-BTAppId -AppId "TrustedIT.Notifications"
$RegPath = "HKCR:\AppUserModelId"
$AppIdPath = "$RegPath\TrustedIT.Notifications"
if (-not (Test-Path -Path $AppIdPath)) {
    New-Item -Path $RegPath -Name "TrustedIT.Notifications" -Force | Out-Null
}

$AppDisplayName = (Get-ItemProperty -Path $AppIdPath -Name DisplayName -ErrorAction SilentlyContinue).DisplayName
if ($AppDisplayName -ne "Trusted IT") {
    New-ItemProperty -Path $AppIdPath -Name DisplayName -Value "Trusted IT" -PropertyType String -Force | Out-Null
}

$AppIconPath = (Get-ItemProperty -Path $AppIdPath -Name IconUri -ErrorAction SilentlyContinue).IconUri
if ($AppIconPath -ne "$NotifyIcon") {
    New-ItemProperty -Path $AppIdPath -Name IconUri -Value $NotifyIcon -PropertyType String -Force | Out-Null
}



# Handle templates
switch ($Preset) {
    "ADSyncAlert"	{
        $BTHero		= New-BTImage -Source $HeroImage
        $BTIcon		= New-BTImage -Source $IconImage
        $TenantName	= (Get-MgOrganization).DisplayName
        New-BurntToastNotification -Text "ADSync is not running for $TenantName!" -HeroImage $BTHero -AppLogo $BTIcon -AppId "TrustedIT.Notifications"
    }
    "Reboot"		{
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
    }
    "ServerProblem" {
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
    }
}
