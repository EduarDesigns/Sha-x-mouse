# --- Script de Configuración Automática v4.0 ---
# Este script instala todo el software necesario en una instancia EC2 de Windows Server
# Se ejecuta desde user-data que descarga este archivo desde GitHub

# Establece un archivo de registro para poder verificar el progreso.
$logFile = "C:\install_log.txt"
$ErrorActionPreference = "Continue"  # Continuar incluso si hay errores

# Verificar que el script se está ejecutando
$scriptStartTime = Get-Date

try {
    "================================================================================" | Out-File $logFile -Append
    "INICIO DEL SCRIPT COMPLETO DE INSTALACION" | Out-File $logFile -Append
    "Fecha/Hora: $(Get-Date)" | Out-File $logFile -Append
    "Usuario: $env:USERNAME" | Out-File $logFile -Append
    "Ruta del script: $PSCommandPath" | Out-File $logFile -Append
    "Directorio de trabajo: $(Get-Location)" | Out-File $logFile -Append
    "================================================================================" | Out-File $logFile -Append
    
    # 1. Instalar Chocolatey
    "Instalando Chocolatey..." | Out-File $logFile -Append
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        "Chocolatey instalado correctamente." | Out-File $logFile -Append
    }
    catch {
        "ERROR al instalar Chocolatey: $_" | Out-File $logFile -Append
        throw "No se pudo instalar Chocolatey, abortando instalación"
    }

# 2. Usar Chocolatey para instalar el software base.
"Instalando software base..." | Out-File $logFile -Append
choco install python -y
choco install git -y
choco install googlechrome -y
choco install sharex -y
choco install onedrive -y
choco install x-mouse-button-control -y

# 3. Instalar AWS VPN Client desde el instalador MSI oficial
"Iniciando la instalación de AWS VPN Client..." | Out-File $logFile -Append
try {
    $vpnUrl = "https://d20adtppz83p9s.cloudfront.net/WPF/latest/AWS_VPN_Client.msi"
    $vpnMsiPath = "$env:TEMP\AWS_VPN_Client.msi"
    
    "Descargando el instalador del cliente VPN..." | Out-File $logFile -Append
    Invoke-WebRequest -Uri $vpnUrl -OutFile $vpnMsiPath

    "Instalando el cliente VPN silenciosamente..." | Out-File $logFile -Append
    # Usamos msiexec para una instalación silenciosa (/qn) y sin reinicios (/norestart)
    Start-Process msiexec.exe -ArgumentList "/i `"$vpnMsiPath`" /qn /norestart" -Wait

    "Instalación de AWS VPN Client completada." | Out-File $logFile -Append
    Remove-Item $vpnMsiPath # Limpiamos el instalador
}
catch {
    "ERROR: Falló la instalación de AWS VPN Client." | Out-File $logFile -Append
}

# 4. Instalar las utilidades de Python (owocr) - MÉTODO ULTRA-ROBUSTO
"Iniciando la instalación de owocr (método a prueba de fallos)..." | Out-File $logFile -Append
try {
    # Paso 1: Esperar un poco para que Python y Git se instalen completamente
    Start-Sleep -Seconds 15
    
    # Paso 2: Intentar recargar variables de entorno (puede fallar, no es crítico)
    try {
        refreshenv
        "Variables de entorno recargadas." | Out-File $logFile -Append
    }
    catch {
        "ADVERTENCIA: No se pudo recargar variables de entorno (puede ser normal)." | Out-File $logFile -Append
    }
    
    # Paso 3: Buscar Python en ubicaciones comunes
    $pythonPath = $null
    $possiblePaths = @(
        "C:\Python*",
        "C:\Program Files\Python*",
        "C:\Program Files (x86)\Python*",
        "$env:LOCALAPPDATA\Programs\Python\Python*"
    )
    
    # Primero intentar con Get-Command
    try {
        $pythonPath = (Get-Command python.exe -ErrorAction Stop).Source
        "Python encontrado mediante Get-Command: $pythonPath" | Out-File $logFile -Append
    }
    catch {
        "Python no encontrado en PATH. Buscando en ubicaciones comunes..." | Out-File $logFile -Append
        
        # Buscar en ubicaciones comunes
        foreach ($basePath in $possiblePaths) {
            $pythonDirs = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue | Sort-Object Name -Descending
            foreach ($dir in $pythonDirs) {
                $testPath = Join-Path $dir.FullName "python.exe"
                if (Test-Path $testPath) {
                    $pythonPath = $testPath
                    "Python encontrado en: $pythonPath" | Out-File $logFile -Append
                    break
                }
            }
            if ($pythonPath) { break }
        }
    }
    
    if (-not $pythonPath -or -not (Test-Path $pythonPath)) {
        throw "No se pudo encontrar python.exe en ninguna ubicación conocida."
    }
    
    # Paso 4: Verificar que Python funciona
    $pythonDir = Split-Path -Parent $pythonPath
    $pythonScriptsDir = Join-Path -Path $pythonDir -ChildPath "Scripts"
    
    # Añadir rutas al PATH de la sesión actual
    $env:Path = "$pythonDir;$pythonScriptsDir;$($env:Path)"
    "Ruta de Python configurada: $pythonDir" | Out-File $logFile -Append
    
    # Verificar versión de Python
    $pythonVersion = & $pythonPath --version 2>&1
    "Versión de Python: $pythonVersion" | Out-File $logFile -Append
    
    # Paso 5: Verificar que git está instalado
    "Verificando instalacion de git..." | Out-File $logFile -Append
    try {
        refreshenv
        $gitPath = (Get-Command git.exe -ErrorAction Stop).Source
        "Git encontrado: $gitPath" | Out-File $logFile -Append
    }
    catch {
        "ADVERTENCIA: Git no encontrado en PATH. Intentando buscar..." | Out-File $logFile -Append
        $gitPossiblePaths = @(
            "C:\Program Files\Git\bin\git.exe",
            "C:\Program Files (x86)\Git\bin\git.exe"
        )
        $gitPath = $null
        foreach ($path in $gitPossiblePaths) {
            if (Test-Path $path) {
                $gitPath = $path
                $gitDir = Split-Path -Parent (Split-Path -Parent $path)
                $env:Path = "$gitDir\bin;$($env:Path)"
                "Git encontrado en: $gitPath" | Out-File $logFile -Append
                break
            }
        }
        if (-not $gitPath) {
            throw "Git no se pudo encontrar. Asegurate de que se instalo correctamente."
        }
    }
    
    # Paso 6: Actualizar pip
    "Actualizando pip..." | Out-File $logFile -Append
    $pipUpgradeOutput = & $pythonPath -m pip install --upgrade pip 2>&1
    $pipUpgradeOutput | Out-File $logFile -Append
    if ($LASTEXITCODE -ne 0) {
        "ADVERTENCIA: pip upgrade fallo con codigo $LASTEXITCODE" | Out-File $logFile -Append
    }
    
    # Paso 7: Instalar owocr directamente desde GitHub sin clonar localmente
    "Instalando owocr desde el repositorio personalizado en GitHub..." | Out-File $logFile -Append
    $owocrRepoUrl = "git+https://github.com/EduarDesigns/my-owocr.git"
    
    # Instalar directamente desde GitHub con extras [lens]
    # Pip clonará temporalmente, instalará y limpiará automáticamente
    "Instalando owocr con extras [lens] directamente desde GitHub..." | Out-File $logFile -Append
    $owocrInstallOutput = & $pythonPath -m pip install "$owocrRepoUrl[lens]" 2>&1
    $owocrInstallOutput | Out-File $logFile -Append
    
    if ($LASTEXITCODE -ne 0) {
        "ADVERTENCIA: Instalacion con extras fallo. Intentando sin extras primero..." | Out-File $logFile -Append
        # Intentar sin extras primero
        $owocrInstallOutput2 = & $pythonPath -m pip install $owocrRepoUrl 2>&1
        $owocrInstallOutput2 | Out-File $logFile -Append
        if ($LASTEXITCODE -eq 0) {
            "Instalacion base exitosa. Instalando extras [lens]..." | Out-File $logFile -Append
            # Instalar los extras después
            $owocrLensOutput = & $pythonPath -m pip install "$owocrRepoUrl[lens]" 2>&1
            $owocrLensOutput | Out-File $logFile -Append
            if ($LASTEXITCODE -ne 0) {
                "ADVERTENCIA: No se pudieron instalar los extras [lens]. Intentando desde PyPI..." | Out-File $logFile -Append
                # Intentar instalar solo los extras desde PyPI (asumiendo que owocr ya está instalado)
                & $pythonPath -m pip install "google-generativeai" 2>&1 | Out-File $logFile -Append
            }
        } else {
            throw "La instalacion de owocr desde el repositorio fallo con codigo $LASTEXITCODE"
        }
    } else {
        "Instalacion exitosa desde GitHub." | Out-File $logFile -Append
    }
    
    # Paso 8: Verificar instalacion
    $owocrCheck = & $pythonPath -m pip show owocr 2>&1
    if ($LASTEXITCODE -eq 0) {
        "Verificacion de instalacion:" | Out-File $logFile -Append
        $owocrCheck | Out-File $logFile -Append
        "Instalacion de owocr completada con exito." | Out-File $logFile -Append
    } else {
        throw "No se pudo verificar la instalacion de owocr"
    }
}
catch {
    $errorMessage = "ERROR: Ocurrió un problema durante la instalación de owocr: $_"
    $errorMessage | Out-File $logFile -Append
    "Detalles del error: $($_.Exception.Message)" | Out-File $logFile -Append
    if ($_.ScriptStackTrace) {
        "Stack trace: $($_.ScriptStackTrace)" | Out-File $logFile -Append
    }
}

# 5. Descargar e instalar Microsoft Office LTSC 2024
"Iniciando la preparación para instalar Office LTSC 2024..." | Out-File $logFile -Append
try {
    $officeDir = "C:\OfficeInstall"
    New-Item -ItemType Directory -Force -Path $officeDir
    $odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20278.exe"
    $odtExtractorPath = "$officeDir\odt_extractor.exe"
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtExtractorPath
    Start-Process -FilePath $odtExtractorPath -ArgumentList "/quiet /extract:$officeDir" -Wait
    $xmlContent = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume"><Language ID="en-US" /></Product>
  </Add>
  <AcceptEULA>TRUE</AcceptEULA>
  <Display Level="None" />
</Configuration>
"@
    $xmlPath = "$officeDir\configuration.xml"
    $xmlContent | Out-File -FilePath $xmlPath -Encoding utf8
    Start-Process -FilePath "$officeDir\setup.exe" -ArgumentList "/download `"$xmlPath`"" -Wait
    Start-Process -FilePath "$officeDir\setup.exe" -ArgumentList "/configure `"$xmlPath`"" -Wait
    "Instalación de Office LTSC 2024 completada." | Out-File $logFile -Append
    Remove-Item -Path $officeDir -Recurse -Force
    
    # Esperar un poco más para asegurar que Office esté completamente instalado
    Start-Sleep -Seconds 30
}
catch {
    "ERROR: Ocurrió un problema grave durante la instalación de Office." | Out-File $logFile -Append
}

# 5.5. Instalar plantillas y addin de Word desde archivos extraídos
"Iniciando la instalación de plantillas y addin de Word..." | Out-File $logFile -Append
try {
    # Los archivos ya fueron descargados y descomprimidos por el user-data
    # Solo necesitamos procesar los archivos extraídos
    $tempExtractPath = "$env:TEMP\att_extracted"
    
    # Esperar un poco por si la extracción aún está en progreso
    $maxWait = 60  # segundos
    $waited = 0
    while (-not (Test-Path $tempExtractPath) -and $waited -lt $maxWait) {
        Start-Sleep -Seconds 5
        $waited += 5
        "Esperando que los archivos esten extraidos... ($waited/$maxWait segundos)" | Out-File $logFile -Append
    }
    
    # Verificar que la extracción fue exitosa
    if (Test-Path $tempExtractPath) {
        "Extracción completada. Copiando archivos..." | Out-File $logFile -Append
        
        # 1. Copiar Normal.dotm a %appdata%\Microsoft\Templates\
        $templatesDir = "$env:APPDATA\Microsoft\Templates"
        New-Item -ItemType Directory -Force -Path $templatesDir | Out-Null
        $normalDotm = Join-Path -Path $tempExtractPath -ChildPath "Normal.dotm"
        if (Test-Path $normalDotm) {
            Copy-Item -Path $normalDotm -Destination $templatesDir -Force
            "Normal.dotm copiado a $templatesDir" | Out-File $logFile -Append
        } else {
            "ADVERTENCIA: No se encontró Normal.dotm en el archivo extraído." | Out-File $logFile -Append
        }
        
        # 2. Copiar ATTESTATION 127.dotm a Documents\Custom Office Templates\ y al escritorio
        $attestationDotm = Join-Path -Path $tempExtractPath -ChildPath "ATTESTATION 127.dotm"
        if (Test-Path $attestationDotm) {
            # Copiar a Documents\Custom Office Templates\ (crear carpeta si no existe)
            $customTemplatesDir = "$env:USERPROFILE\Documents\Custom Office Templates"
            New-Item -ItemType Directory -Force -Path $customTemplatesDir | Out-Null
            Copy-Item -Path $attestationDotm -Destination $customTemplatesDir -Force
            "ATTESTATION 127.dotm copiado a $customTemplatesDir" | Out-File $logFile -Append
            
            # Copiar también al escritorio
            $desktopPath = [Environment]::GetFolderPath("Desktop")
            Copy-Item -Path $attestationDotm -Destination $desktopPath -Force
            "ATTESTATION 127.dotm copiado al escritorio: $desktopPath" | Out-File $logFile -Append
            "NOTA: ATTESTATION 127.dotm estara disponible como plantilla personalizada en Word." | Out-File $logFile -Append
        } else {
            "ADVERTENCIA: No se encontro ATTESTATION 127.dotm en el archivo extraido." | Out-File $logFile -Append
        }
        
        # 3. Copiar carpeta Document Building Blocks a %appdata%\Microsoft\
        $buildingBlocksSource = Join-Path -Path $tempExtractPath -ChildPath "Document Building Blocks"
        $buildingBlocksDest = "$env:APPDATA\Microsoft\Document Building Blocks"
        if (Test-Path $buildingBlocksSource) {
            if (Test-Path $buildingBlocksDest) {
                "Carpeta Document Building Blocks ya existe. Reemplazando..." | Out-File $logFile -Append
                Remove-Item -Path $buildingBlocksDest -Recurse -Force
            }
            Copy-Item -Path $buildingBlocksSource -Destination $buildingBlocksDest -Recurse -Force
            "Carpeta Document Building Blocks copiada a $env:APPDATA\Microsoft\" | Out-File $logFile -Append
        } else {
            "ADVERTENCIA: No se encontró la carpeta Document Building Blocks en el archivo extraído." | Out-File $logFile -Append
        }
        
        # 4. Copiar carpeta wordin al escritorio para instalacion manual
        $wordinDir = Join-Path -Path $tempExtractPath -ChildPath "wordin"
        if (Test-Path $wordinDir) {
            "Copiando carpeta wordin al escritorio..." | Out-File $logFile -Append
            $desktopPath = [Environment]::GetFolderPath("Desktop")
            $wordinDest = Join-Path -Path $desktopPath -ChildPath "wordin"
            
            # Si ya existe, eliminar la anterior
            if (Test-Path $wordinDest) {
                "La carpeta wordin ya existe en el escritorio. Reemplazando..." | Out-File $logFile -Append
                Remove-Item -Path $wordinDest -Recurse -Force
            }
            
            Copy-Item -Path $wordinDir -Destination $wordinDest -Recurse -Force
            "Carpeta wordin copiada al escritorio: $wordinDest" | Out-File $logFile -Append
            "NOTA: Puedes instalar el addin manualmente ejecutando setup.exe desde la carpeta wordin en el escritorio." | Out-File $logFile -Append
        } else {
            "ADVERTENCIA: No se encontro la carpeta wordin en el archivo extraido." | Out-File $logFile -Append
        }
        
        # Limpiar archivos temporales extraídos (ya se copiaron al escritorio)
        "Limpiando archivos temporales extraidos..." | Out-File $logFile -Append
        Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        "Limpieza de archivos temporales completada." | Out-File $logFile -Append
        "Instalación de plantillas y addin de Word completada." | Out-File $logFile -Append
    } else {
        "ERROR: Los archivos extraidos no se encontraron en $tempExtractPath" | Out-File $logFile -Append
        "Es posible que la extraccion aun este en progreso o haya fallado." | Out-File $logFile -Append
    }
}
catch {
    "ERROR: Ocurrió un problema durante la instalación de plantillas y addin de Word: $_" | Out-File $logFile -Append
}

# 6. Configurar el backup de ShareX
"Iniciando la configuración del backup de ShareX..." | Out-File $logFile -Append
try {
    Stop-Process -Name ShareX -Force -ErrorAction SilentlyContinue
    $sharexBackupUrl = "https://github.com/EduarDesigns/Sha-x-mouse/releases/download/v1.0/ShareX_Backup.zip"
    $sharexSettingsDir = "C:\Users\Administrator\Documents\ShareX"
    $tempZipPath = "$env:TEMP\ShareX_Backup.zip"
    New-Item -ItemType Directory -Force -Path $sharexSettingsDir
    Invoke-WebRequest -Uri $sharexBackupUrl -OutFile $tempZipPath
    Expand-Archive -Path $tempZipPath -DestinationPath $sharexSettingsDir -Force
    Remove-Item -Path $tempZipPath
    "Configuración de ShareX restaurada con éxito." | Out-File $logFile -Append
}
catch {
    "ERROR: No se pudo restaurar el backup de ShareX." | Out-File $logFile -Append
}

# 7. Configurar X-Mouse Button Control
"Iniciando la configuración de X-Mouse Button Control..." | Out-File $logFile -Append
try {
    Stop-Process -Name XMouseButtonControl -Force -ErrorAction SilentlyContinue
    $xmbcSettingsUrl = "https://raw.githubusercontent.com/EduarDesigns/Sha-x-mouse/main/XMBCSettings.xml"
    $xmbcDir = "$env:APPDATA\Highresolution Enterprises\XMouseButtonControl"
    $xmbcFile = "$xmbcDir\XMBCSettings.xml"
    New-Item -ItemType Directory -Force -Path $xmbcDir
    Invoke-WebRequest -Uri $xmbcSettingsUrl -OutFile $xmbcFile
    "Configuración de X-Mouse Button Control restaurada con éxito." | Out-File $logFile -Append
}
catch {
    "ERROR: No se pudo restaurar la configuración de XMBC." | Out-File $logFile -Append
}

    "--- PROCESO DE CONFIGURACION FINALIZADO ---" | Out-File $logFile -Append
    "Finalizado a las $(Get-Date)" | Out-File $logFile -Append
}
catch {
    "================================================================================" | Out-File $logFile -Append
    "ERROR CRITICO EN EL SCRIPT DE INSTALACION" | Out-File $logFile -Append
    "================================================================================" | Out-File $logFile -Append
    "Error: $_" | Out-File $logFile -Append
    "Detalles: $($_.Exception.Message)" | Out-File $logFile -Append
    if ($_.ScriptStackTrace) {
        "Stack trace:" | Out-File $logFile -Append
        $_.ScriptStackTrace | Out-File $logFile -Append
    }
    "El script terminó con errores a las $(Get-Date)" | Out-File $logFile -Append
    exit 1
}

