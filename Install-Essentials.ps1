<#
.SYNOPSIS
    Installs essential apps via winget with silent mode and creates desktop shortcuts on the user's Desktop.
#>

# Welcome message
Write-Host @"
============================================
  FreshStart - Essential Windows App Installer
============================================
"@ -ForegroundColor Cyan
Write-Host "This script helps you install popular apps and puts shortcuts on your desktop!" -ForegroundColor Yellow
Write-Host "\nIf you pasted a one-liner, just follow the menu. If prompted, allow PowerShell to run as administrator.\n" -ForegroundColor Green

# Ensure running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator. Right-click PowerShell and select 'Run as administrator', or approve the UAC prompt."
    Start-Sleep -Seconds 5
    exit
}

# Check for winget
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "\nERROR: 'winget' (Windows Package Manager) is not installed on your system." -ForegroundColor Red
    Write-Host "Please update Windows or install winget from the Microsoft Store, then re-run this script." -ForegroundColor Red
    exit
}


function Create-DesktopShortcut {
    param(
        [string]$Name,
        [string]$TargetPath
    )
    $desktop = [Environment]::GetFolderPath('Desktop')
    $shell   = New-Object -ComObject WScript.Shell
    $link    = $shell.CreateShortcut("$desktop\$Name.lnk")
    $link.TargetPath = $TargetPath
    $link.Save()
}

# Define categories and apps
$categories = @{
    "Browsers" = @(
        @{Name='Google Chrome';        Id='Google.Chrome';          Exe='chrome.exe'},
        @{Name='Microsoft Edge';       Id='Microsoft.Edge';         Exe='msedge.exe'},
        @{Name='Mozilla Firefox';      Id='Mozilla.Firefox';        Exe='firefox.exe'},
        @{Name='Brave';                Id='Brave.Brave';            Exe='brave.exe'},
        @{Name='Opera GX';             Id='Opera.OperaGX';          Exe='launcher.exe'}
    )
    "Communication" = @(
        @{Name='Discord';              Id='Discord.Discord';        Exe='Discord.exe'},
        @{Name='Slack';                Id='SlackTechnologies.Slack'; Exe='slack.exe'},
        @{Name='Microsoft Teams';      Id='Microsoft.Teams';        Exe='Teams.exe'},
        @{Name='Zoom';                 Id='Zoom.Zoom';              Exe='Zoom.exe'},
        @{Name='Skype';                Id='Microsoft.Skype';        Exe='Skype.exe'}
    )
    "Compression" = @(
        @{Name='7-Zip';                Id='7zip.7zip';              Exe='7zFM.exe'},
        @{Name='WinRAR';               Id='RarLab.WinRAR';          Exe='WinRAR.exe'},
        @{Name='PeaZip';               Id='PeaZip.PeaZip';          Exe='PeaZip.exe'},
        @{Name='Bandizip';             Id='Bandisoft.Bandizip';     Exe='Bandizip.exe'}
    )
    "Development" = @(
        @{Name='Visual Studio Code';   Id='Microsoft.VisualStudioCode'; Exe='Code.exe'},
        @{Name='Git';                  Id='Git.Git';                Exe='git.exe'},
        @{Name='Node.js';              Id='OpenJS.NodeJS';          Exe='node.exe'},
        @{Name='Python 3';             Id='Python.Python.3';        Exe='python.exe'},
        @{Name='Docker Desktop';       Id='Docker.DockerDesktop';   Exe='Docker Desktop.exe'},
        @{Name='Postman';              Id='Postman.Postman';        Exe='postman.exe'}
    )
    "Media" = @(
        @{Name='VLC Media Player';     Id='VideoLAN.VLC';           Exe='vlc.exe'},
        @{Name='Spotify';              Id='Spotify.Spotify';        Exe='Spotify.exe'},
        @{Name='iTunes';               Id='Apple.iTunes';           Exe='iTunes.exe'},
        @{Name='OBS Studio';           Id='OBSProject.OBSStudio';   Exe='obs64.exe'},
        @{Name='GIMP';                 Id='GIMP.GIMP';              Exe='gimp-2.10.exe'}
    )
}

# Display menu and read selection
Write-Host "Select categories to install (comma-separated numbers):`n"
$menu = @()
$idx = 1
foreach ($cat in $categories.Keys) {
    Write-Host "[$idx] $cat"
    $menu += $cat
    $idx++
}
$sel = Read-Host "Enter choices"
$nums = $sel -split ',' | % { $_.Trim() } | Where-Object { $_ -match '^[0-9]+$' }
$chosenCats = $nums | % { $menu[$_ - 1] } | Where-Object { $_ }

# Aggregate apps
$toInstall = foreach ($c in $chosenCats) { $categories[$c] }

# Install and create shortcuts
foreach ($app in $toInstall) {
    Write-Host "Installing $($app.Name)..."
    winget install --id $app.Id --silent --accept-source-agreements --accept-package-agreements

    $exePath = (Get-Command $app.Exe -ErrorAction SilentlyContinue).Source
    if ($exePath) {
        Create-DesktopShortcut -Name $app.Name -TargetPath $exePath
        Write-Host "Shortcut for $($app.Name) created on Desktop.`n"
    } else {
        Write-Warning "Could not locate $($app.Name) executable; shortcut skipped.`n"
    }
}

Write-Host "All done! Check your desktop for shortcuts."
