param(
    [Parameter(Mandatory = $true)]
    [string]$AdminUser,

    [Parameter(Mandatory = $true)]
    [string]$AdminPass,

    [Parameter(Mandatory = $true)]
    [string]$CIDR
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$NetDTLPath = "C:\xampp\htdocs\netdtl"
$DbName = "netdtl"

function Write-Log {
    param([string]$Message)
    Write-Host "[NetDTL] $Message"
}

function Find-Nmap {
    $paths = @(
        "C:\Program Files\Nmap\nmap.exe",
        "C:\Program Files (x86)\Nmap\nmap.exe"
    )

    foreach ($path in $paths) {
        if (Test-Path $path) {
            return $path
        }
    }

    throw "Nmap introuvable."
}

function Find-MySQL {
    $mysql = "C:\xampp\mysql\bin\mysql.exe"

    if (-not (Test-Path $mysql)) {
        throw "Client MariaDB introuvable."
    }

    return $mysql
}

function Wait-MariaDB {
    param([string]$MySQL)

    Write-Log "Attente MariaDB..."

    for ($i = 0; $i -lt 30; $i++) {
        try {
            & $MySQL -u root -e "SELECT 1;" *> $null
            Write-Log "MariaDB OK."
            return
        }
        catch {
            Start-Sleep -Seconds 2
        }
    }

    throw "MariaDB indisponible."
}

function Create-Database {
    param([string]$MySQL)

    Write-Log "Création base netdtl..."

    & $MySQL -u root -e `
        "CREATE DATABASE IF NOT EXISTS netdtl CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;"
}

function Import-SQL {
    param(
        [string]$MySQL,
        [string]$SqlFile
    )

    if (-not (Test-Path $SqlFile)) {
        throw "netdtl.sql introuvable."
    }

    Write-Log "Import SQL..."

    cmd.exe /c "`"$MySQL`" -u root netdtl < `"$SqlFile`""
}

function Generate-Config {
    param([string]$NmapPath)

    $template = Join-Path $NetDTLPath "db.example.php"
    $target = Join-Path $NetDTLPath "db.php"

    if (-not (Test-Path $template)) {
        throw "db.example.php introuvable."
    }

    Write-Log "Génération db.php..."

    $content = Get-Content $template -Raw

    $content = $content -replace "changeme", $AdminPass
    $content = $content -replace "admin", $AdminUser
    $content = $content -replace "192.168.1.0/24", $CIDR
    $content = $content -replace "C:\\\\Program Files\\\\Nmap\\\\nmap.exe", ($NmapPath -replace "\\","\\\\")

    Set-Content `
        -Path $target `
        -Value $content `
        -Encoding UTF8
}

function Validate-Web {
    Write-Log "Validation NetDTL..."

    $pair = "$AdminUser`:$AdminPass"
    $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
    $token = [Convert]::ToBase64String($bytes)

    $headers = @{
        Authorization = "Basic $token"
    }

    $resp = Invoke-WebRequest `
        -Uri "http://localhost/netdtl/" `
        -Headers $headers `
        -UseBasicParsing `
        -TimeoutSec 20

    if ($resp.StatusCode -ne 200) {
        throw "NetDTL inaccessible."
    }

    Write-Log "Validation OK."
}

# MAIN
Write-Log "Configuration NetDTL"

if (-not (Test-Path $NetDTLPath)) {
    throw "Répertoire NetDTL introuvable."
}

$nmap = Find-Nmap
$mysql = Find-MySQL

Wait-MariaDB -MySQL $mysql
Create-Database -MySQL $mysql

$sql = Join-Path $NetDTLPath "netdtl.sql"

Import-SQL `
    -MySQL $mysql `
    -SqlFile $sql

Generate-Config -NmapPath $nmap
Validate-Web

Write-Log "Configuration terminée."