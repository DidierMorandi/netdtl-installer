param(
    [string]$PythonVersion = "3.12.10"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Write-Log {
    param([string]$Message)
    Write-Host "[NetDTL] $Message"
}

function Assert-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Ce script doit être exécuté en administrateur."
    }
}

function Download-File {
    param(
        [string]$Url,
        [string]$Destination
    )

    Write-Log "Téléchargement : $Url"
    Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing

    if (-not (Test-Path $Destination)) {
        throw "Téléchargement échoué : $Url"
    }
}

function Command-Exists {
    param([string]$Command)
    return $null -ne (Get-Command $Command -ErrorAction SilentlyContinue)
}

function Install-Python {
    if (Command-Exists "python") {
        Write-Log "Python déjà installé."
        return
    }

    $installer = Join-Path $TempDir "python-installer.exe"

    Download-File `
        -Url "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe" `
        -Destination $installer

    Write-Log "Installation Python..."
    Start-Process $installer `
        -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1 Include_pip=1" `
        -Wait `
        -NoNewWindow
}

function Install-PythonDeps {
    Write-Log "Installation dépendances Python..."

    python -m pip install --upgrade pip
    python -m pip install pillow pandas openpyxl reportlab
}

function Install-Nmap {
    $nmap1 = "C:\Program Files\Nmap\nmap.exe"
    $nmap2 = "C:\Program Files (x86)\Nmap\nmap.exe"

    if ((Test-Path $nmap1) -or (Test-Path $nmap2)) {
        Write-Log "Nmap déjà installé."
        return
    }

    $installer = Join-Path $TempDir "nmap-setup.exe"

    Download-File `
        -Url "https://nmap.org/dist/nmap-7.97-setup.exe" `
        -Destination $installer

    Write-Log "Installation Nmap..."
    Start-Process $installer `
        -ArgumentList "/S" `
        -Wait `
        -NoNewWindow
}

function Install-XAMPP {
    $xampp = "C:\xampp\xampp-control.exe"

    if (Test-Path $xampp) {
        Write-Log "XAMPP déjà installé."
        return
    }

    $installer = Join-Path $TempDir "xampp-installer.exe"

    $url = "https://sourceforge.net/projects/xampp/files/XAMPP%20Windows/8.2.12/xampp-windows-x64-8.2.12-0-VS16-installer.exe/download"

    Download-File -Url $url -Destination $installer

    Write-Log "Installation XAMPP..."
    Start-Process $installer `
        -ArgumentList "--mode unattended" `
        -Wait `
        -NoNewWindow
}

# MAIN
Assert-Admin

$TempDir = Join-Path $env:TEMP "NetDTL_Installer"
New-Item -ItemType Directory -Force -Path $TempDir | Out-Null

Write-Log "Début installation dépendances..."

Install-Python
Install-PythonDeps
Install-Nmap
Install-XAMPP

Write-Log "Installation terminée."