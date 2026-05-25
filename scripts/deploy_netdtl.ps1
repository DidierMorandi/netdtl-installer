param(
    [string]$RepoZip = "https://github.com/DidierMorandi/netdtl/archive/refs/tags/v1.0.0.zip",
    [string]$TargetPath = "C:\xampp\htdocs\netdtl"
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

function Write-Log {
    param([string]$Message)
    Write-Host "[NetDTL] $Message"
}

function Assert-XAMPP {
    if (-not (Test-Path "C:\xampp")) {
        throw "XAMPP introuvable. L'installation des dépendances a probablement échoué."
    }
}

function Download-Repo {
    param(
        [string]$Url,
        [string]$Destination
    )

    Write-Log "Téléchargement NetDTL depuis GitHub..."

    Invoke-WebRequest `
        -Uri $Url `
        -OutFile $Destination `
        -UseBasicParsing

    if (-not (Test-Path $Destination)) {
        throw "Téléchargement du dépôt échoué."
    }
}

function Start-XAMPPServices {
    Write-Log "Démarrage Apache / MariaDB..."

    $apacheService = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match "apache"
    }

    $mysqlService = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match "mysql|mariadb"
    }

    if ($apacheService) {
        foreach ($svc in $apacheService) {
            if ($svc.Status -ne "Running") {
                Start-Service $svc.Name
            }
        }
    }
    else {
        $apacheExe = "C:\xampp\apache\bin\httpd.exe"

        if (Test-Path $apacheExe) {
            Start-Process $apacheExe -WindowStyle Hidden
        }
    }

    if ($mysqlService) {
        foreach ($svc in $mysqlService) {
            if ($svc.Status -ne "Running") {
                Start-Service $svc.Name
            }
        }
    }
    else {
        $mysqlExe = "C:\xampp\mysql\bin\mysqld.exe"

        if (Test-Path $mysqlExe) {
            Start-Process $mysqlExe -WindowStyle Hidden
        }
    }

    Start-Sleep -Seconds 10
}

function Deploy-NetDTL {
    $tempZip = Join-Path $env:TEMP "netdtl_v1.zip"
    $extractPath = Join-Path $env:TEMP "netdtl_extract"

    if (Test-Path $extractPath) {
        Remove-Item $extractPath -Recurse -Force
    }

    if (Test-Path $tempZip) {
        Remove-Item $tempZip -Force
    }

    Download-Repo `
        -Url $RepoZip `
        -Destination $tempZip

    Write-Log "Extraction archive..."

    Expand-Archive `
        -Path $tempZip `
        -DestinationPath $extractPath `
        -Force

    $source = Join-Path $extractPath "netdtl-1.0.0"

    if (-not (Test-Path $source)) {
        throw "Sources NetDTL introuvables après extraction."
    }

    if (Test-Path $TargetPath) {
        Write-Log "Suppression ancienne installation..."
        Remove-Item $TargetPath -Recurse -Force
    }

    Write-Log "Déploiement NetDTL..."

    Copy-Item `
        -Path $source `
        -Destination $TargetPath `
        -Recurse `
        -Force
}

function Validate-Install {
    $required = @(
        "index.php",
        "db.example.php",
        "netdtl.sql"
    )

    foreach ($file in $required) {
        $path = Join-Path $TargetPath $file

        if (-not (Test-Path $path)) {
            throw "Fichier manquant après déploiement : $file"
        }
    }

    Write-Log "Validation fichiers OK."
}

# MAIN
Write-Log "Déploiement NetDTL v1.0.0"

Assert-XAMPP
Deploy-NetDTL
Start-XAMPPServices
Validate-Install

Write-Log "Déploiement terminé."