# --- Script de Configuración Automática v4.0 ---
# Este script instala todo el software necesario en una instancia EC2 de Windows Server
# Se ejecuta desde user-data que descarga este archivo desde GitHub

# Establece un archivo de registro para poder verificar el progreso.
$logFile = "C:\install_log.txt"
"================================================================================" | Out-File $logFile
"Inicio del script de instalación a las $(Get-Date)" | Out-File $logFile -Append
"================================================================================" | Out-File $logFile -Append

# 1. Instalar Chocolatey
"Instalando Chocolatey..." | Out-File $logFile -Append
Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

# 2. Usar Chocolatey para instalar el software base.
"Instalando software base..." | Out-File $logFile -Append
choco install python -y
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
    # Paso 1: Esperar un poco para que Python se instale completamente
    Start-Sleep -Seconds 10
    
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
    
    # Paso 5: Actualizar pip
    "Actualizando pip..." | Out-File $logFile -Append
    $pipUpgradeOutput = & $pythonPath -m pip install --upgrade pip 2>&1
    $pipUpgradeOutput | Out-File $logFile -Append
    if ($LASTEXITCODE -ne 0) {
        "ADVERTENCIA: pip upgrade falló con código $LASTEXITCODE" | Out-File $logFile -Append
    }
    
    # Paso 6: Instalar owocr
    "Instalando owocr..." | Out-File $logFile -Append
    $owocrOutput = & $pythonPath -m pip install owocr 2>&1
    $owocrOutput | Out-File $logFile -Append
    if ($LASTEXITCODE -ne 0) {
        throw "La instalación de owocr falló con código $LASTEXITCODE"
    }
    
    # Paso 7: Instalar owocr con extras [lens]
    "Instalando owocr[lens]..." | Out-File $logFile -Append
    $owocrLensOutput = & $pythonPath -m pip install 'owocr[lens]' 2>&1
    $owocrLensOutput | Out-File $logFile -Append
    if ($LASTEXITCODE -ne 0) {
        "ADVERTENCIA: La instalación de owocr[lens] falló con código $LASTEXITCODE" | Out-File $logFile -Append
        "Esto puede ser normal si ya está instalado o si no se necesitan los extras." | Out-File $logFile -Append
    }
    
    # Paso 8: Verificar instalación
    $owocrCheck = & $pythonPath -m pip show owocr 2>&1
    if ($LASTEXITCODE -eq 0) {
        "Verificación de instalación:" | Out-File $logFile -Append
        $owocrCheck | Out-File $logFile -Append
        "Instalación de owocr completada con éxito." | Out-File $logFile -Append
    } else {
        throw "No se pudo verificar la instalación de owocr"
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
    # Variable para controlar la limpieza de archivos temporales
    $script:skipCleanup = $false
    
    # Los archivos ya fueron descargados y descomprimidos por el user-data
    # Solo necesitamos procesar los archivos extraídos
    $tempExtractPath = "$env:TEMP\att_extracted"
    
    # Esperar un poco por si la extracción aún está en progreso
    $maxWait = 60  # segundos
    $waited = 0
    while (-not (Test-Path $tempExtractPath) -and $waited -lt $maxWait) {
        Start-Sleep -Seconds 5
        $waited += 5
        "Esperando que los archivos estén extraídos... ($waited/$maxWait segundos)" | Out-File $logFile -Append
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
        
        # 2. Copiar ATTESTATION 127.dotm a Templates (será registrado como nueva plantilla)
        $attestationDotm = Join-Path -Path $tempExtractPath -ChildPath "ATTESTATION 127.dotm"
        if (Test-Path $attestationDotm) {
            Copy-Item -Path $attestationDotm -Destination $templatesDir -Force
            "ATTESTATION 127.dotm copiado a $templatesDir" | Out-File $logFile -Append
            "NOTA: ATTESTATION 127.dotm estará disponible como plantilla personalizada en Word." | Out-File $logFile -Append
        } else {
            "ADVERTENCIA: No se encontró ATTESTATION 127.dotm en el archivo extraído." | Out-File $logFile -Append
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
        
        # 4. Instalar el addin de Word desde la carpeta wordin
        $wordinDir = Join-Path -Path $tempExtractPath -ChildPath "wordin"
        $setupExe = Join-Path -Path $wordinDir -ChildPath "setup.exe"
        if (Test-Path $setupExe) {
            "Instalando addin de Word desde wordin..." | Out-File $logFile -Append
            "Ruta del instalador: $setupExe" | Out-File $logFile -Append
            
            # Asegurar que Word no esté ejecutándose
            Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            
            # Función para ejecutar con timeout
            function Start-ProcessWithTimeout {
                param(
                    [string]$FilePath,
                    [string]$Arguments,
                    [int]$TimeoutSeconds = 300
                )
                
                $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -NoNewWindow
                $processId = $process.Id
                "Proceso iniciado con PID: $processId" | Out-File $logFile -Append
                
                $completed = $false
                $elapsed = 0
                $checkInterval = 5
                
                while ($elapsed -lt $TimeoutSeconds -and -not $completed) {
                    Start-Sleep -Seconds $checkInterval
                    $elapsed += $checkInterval
                    
                    try {
                        $runningProcess = Get-Process -Id $processId -ErrorAction Stop
                        if ($runningProcess.HasExited) {
                            $completed = $true
                            $exitCode = $runningProcess.ExitCode
                            "Proceso terminó con código de salida: $exitCode" | Out-File $logFile -Append
                            return $exitCode
                        }
                    }
                    catch {
                        $completed = $true
                        "Proceso ya no existe (puede haber terminado o haber sido terminado)." | Out-File $logFile -Append
                        return $null
                    }
                }
                
                if (-not $completed) {
                    "TIMEOUT: El proceso no terminó en $TimeoutSeconds segundos. Terminando proceso..." | Out-File $logFile -Append
                    try {
                        Stop-Process -Id $processId -Force -ErrorAction Stop
                        "Proceso terminado forzadamente." | Out-File $logFile -Append
                    }
                    catch {
                        "No se pudo terminar el proceso (puede que ya haya terminado)." | Out-File $logFile -Append
                    }
                    return $null
                }
            }
            
            # Intentar instalación silenciosa con diferentes parámetros comunes
            $installSuccess = $false
            $installMethods = @(
                @{Args = "/S"; Name = "/S (silent)"},
                @{Args = "/quiet"; Name = "/quiet"},
                @{Args = "/SILENT"; Name = "/SILENT"},
                @{Args = "/VERYSILENT"; Name = "/VERYSILENT"},
                @{Args = "/qn"; Name = "/qn (MSI)"}
            )
            
            foreach ($method in $installMethods) {
                if ($installSuccess) { break }
                
                "Intentando instalación con parámetros: $($method.Name)..." | Out-File $logFile -Append
                try {
                    $exitCode = Start-ProcessWithTimeout -FilePath $setupExe -Arguments $method.Args -TimeoutSeconds 300
                    
                    if ($exitCode -eq 0 -or $null -eq $exitCode) {
                        "Instalación completada con método $($method.Name)." | Out-File $logFile -Append
                        $installSuccess = $true
                    }
                    else {
                        "El proceso terminó con código de error: $exitCode (método $($method.Name))" | Out-File $logFile -Append
                        "Intentando siguiente método..." | Out-File $logFile -Append
                    }
                }
                catch {
                    "ERROR con método $($method.Name): $_" | Out-File $logFile -Append
                    "Intentando siguiente método..." | Out-File $logFile -Append
                }
            }
            
            if (-not $installSuccess) {
                "ADVERTENCIA: No se pudo instalar el addin con ningún método de instalación silenciosa." | Out-File $logFile -Append
                "El addin puede requerir instalación manual ejecutando: $setupExe" | Out-File $logFile -Append
                "NOTA: Los archivos extraídos están en: $tempExtractPath" | Out-File $logFile -Append
                "NOTA: Los archivos NO se limpiarán para permitir instalación manual." | Out-File $logFile -Append
                $script:skipCleanup = $true
            }
        } else {
            "ADVERTENCIA: No se encontró setup.exe en la carpeta wordin." | Out-File $logFile -Append
            "Buscando en: $wordinDir" | Out-File $logFile -Append
            $filesInWordin = Get-ChildItem -Path $wordinDir -ErrorAction SilentlyContinue
            if ($filesInWordin) {
                "Archivos encontrados en wordin:" | Out-File $logFile -Append
                $filesInWordin | ForEach-Object { "  - $($_.Name)" } | Out-File $logFile -Append
            }
        }
        
        # Limpiar archivos temporales solo si la instalación fue exitosa
        if (-not $script:skipCleanup) {
            "Limpiando archivos temporales extraídos..." | Out-File $logFile -Append
            Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
            "Limpieza de archivos temporales completada." | Out-File $logFile -Append
        } else {
            "Los archivos temporales se mantienen para instalación manual en: $tempExtractPath" | Out-File $logFile -Append
        }
        "Instalación de plantillas y addin de Word completada." | Out-File $logFile -Append
    } else {
        "ERROR: Los archivos extraídos no se encontraron en $tempExtractPath" | Out-File $logFile -Append
        "Es posible que la extracción aún esté en progreso o haya fallado." | Out-File $logFile -Append
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

"--- PROCESO DE CONFIGURACIÓN FINALIZADO ---" | Out-File $logFile -Append
"Finalizado a las $(Get-Date)" | Out-File $logFile -Append

