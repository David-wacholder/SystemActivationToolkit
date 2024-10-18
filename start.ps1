# Function to display error popup using VBS
function Show-VBSError {
    param (
        [string]$message
    )
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup($message, 0, "Error", 16)
}

# Function to display success popup using VBS
function Show-VBSSuccess {
    param (
        [string]$message
    )
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup($message, 0, "Success", 64)
}

# Function to check if the file exists and display the full path in English
function Check-File {
    param (
        [string]$filePath
    )
    $fullPath = Join-Path -Path $PSScriptRoot -ChildPath $filePath
    if (-not (Test-Path -Path $fullPath)) {
        Show-VBSError "File not found: $filePath"
        Exit
    } else {
        Write-Output "File found: $fullPath"
    }
}

# Function to check for internet connection
function Check-InternetConnection {
    try {
        $connection = Test-Connection -ComputerName "www.google.com" -Count 1 -Quiet
        if (-not $connection) {
            Show-VBSError "No internet connection. Operation aborted."
            Exit
        }
    }
    catch {
        Show-VBSError "No internet connection. Operation aborted."
        Exit
    }
}

# Getting parameters from user input
$params = $args -join ' '

# Checking for administrator privileges
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    # If no admin rights, re-run the script as admin with all parameters
    Start-Process powershell -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $params" -Verb RunAs
    Exit
}

# Changing directory to the script folder
Set-Location -Path $PSScriptRoot

# Checking internet connection
Check-InternetConnection

# File names
$windowsfileName = "win_active.cmd"
$officefileName = "office.cmd"
$timefileName = "time.bat"
$OEMfileName = "OEM_Configurator.bat"
$browserUpdateFile = "BrowserUpdate.ps1"

# Checking if files exist
Check-File $windowsfileName
Check-File $officefileName
Check-File $timefileName
Check-File $OEMfileName
Check-File $browserUpdateFile

# System actions
powershell -command "Enable-ComputerRestore -Drive C:"
vssadmin resize shadowstorage /for=C: /on=C: /maxsize=13%

# Choosing parts to execute
Write-Output "Please choose the part you want to execute:"
Write-Output "1. All parts - Win Active, Win Office, Time, OEM_Configurator, BrowserUpdate"
Write-Output "2. All parts without OEM_Configurator"
Write-Output "3. Run Win Active"
Write-Output "4. Run Win Office"
Write-Output "5. Run Time"
Write-Output "6. Run OEM_Configurator"
Write-Output "7. Run BrowserUpdate.ps1"

$userChoice = Read-Host "Enter the number of the part you want to execute"

# Default to part 1 if no input is provided
if ([string]::IsNullOrWhiteSpace($userChoice)) {
    $userChoice = "1"
}

switch ($userChoice) {
    "1" {
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/E"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $windowsfileName"
        Write-Output "Windows Active executed successfully"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $officefileName"
        Write-Output "Office executed successfully"
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/D /P"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $timefileName"
        Write-Output "Time executed successfully"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $OEMfileName"
        Write-Output "OEM Configurator executed successfully"
        Start-Process -Wait -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File $browserUpdateFile"
        Write-Output "BrowserUpdate executed successfully"
        Write-Output "All parts executed successfully"
    }
    "2" {
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/E"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $windowsfileName"
        Write-Output "Windows Active executed successfully"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $officefileName"
        Write-Output "Office executed successfully"
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/D /P"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $timefileName"
        Write-Output "Time executed successfully"
        Start-Process -Wait -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File $browserUpdateFile"
        Write-Output "BrowserUpdate executed successfully"
    }
    "3" {
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/E"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $windowsfileName"
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/D /P"
        Write-Output "Windows Active executed successfully"
    }
    "4" {
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/E"
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $officefileName"
        Start-Process -Wait -FilePath "Wub.exe" -ArgumentList "/D /P"
        Write-Output "Office executed successfully"
    }
    "5" {
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $timefileName"
        Write-Output "Time executed successfully"
    }
    "6" {
        Start-Process -Wait -FilePath "cmd.exe" -ArgumentList "/c $OEMfileName"
        Write-Output "OEM Configurator executed successfully"
    }
    "7" {
        Start-Process -Wait -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File $browserUpdateFile"
        Write-Output "BrowserUpdate executed successfully"
    }
    Default {
        Write-Output "Invalid choice"
        Exit
    }
}

# Success message at the end of the script
Show-VBSSuccess "All operations completed successfully"

Write-Output "Thanks to YK"
Start-Sleep -Seconds 3
Exit
