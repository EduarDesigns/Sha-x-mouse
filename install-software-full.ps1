<powershell>
# --- Script mínimo de user-data que descarga y ejecuta el script completo desde GitHub ---
# Este script evita el límite de 16KB del user-data de EC2

$logFile = "C:\install_log.txt"
"Inicio del user-data a las $(Get-Date)" | Out-File $logFile
"Descargando script completo desde GitHub..." | Out-File $logFile -Append

try {
    # URL del script completo en GitHub (cambia esto por la URL real de tu repositorio)
    # Ejemplo: https://raw.githubusercontent.com/TU_USUARIO/TU_REPO/main/install-software-full.ps1
    $scriptUrl = "https://raw.githubusercontent.com/EduarDesigns/Sha-x-mouse/refs/heads/main/install-software-full.ps1"
    
    # Descargar el script completo
    $scriptPath = "$env:TEMP\install-software-full.ps1"
    Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath -UseBasicParsing
    
    "Script descargado exitosamente." | Out-File $logFile -Append
    
    # Descargar y descomprimir att.zip (con clave privada)
    "Descargando y descomprimiendo att.zip..." | Out-File $logFile -Append
    try {
        # Instalar 7-Zip si no está disponible
        $7zipPath = "C:\Program Files\7-Zip\7z.exe"
        if (-not (Test-Path $7zipPath)) {
            "Instalando 7-Zip..." | Out-File $logFile -Append
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            choco install 7zip -y
            refreshenv
            $7zipPath = "C:\Program Files\7-Zip\7z.exe"
        }
        
        # Descargar att.zip
        $attZipUrl = "https://github.com/EduarDesigns/Sha-x-mouse/releases/download/v1.0/att.zip"
        $tempAttZipPath = "$env:TEMP\att.zip"
        $tempExtractPath = "$env:TEMP\att_extracted"
        
        Invoke-WebRequest -Uri $attZipUrl -OutFile $tempAttZipPath -UseBasicParsing
        
        # Crear directorio temporal para extracción
        New-Item -ItemType Directory -Force -Path $tempExtractPath | Out-Null
        
        # Descomprimir con contraseña (clave privada - no se sube a GitHub)
        $password = "11234205"
        $extractArgs = "x `"$tempAttZipPath`" -o`"$tempExtractPath`" -p$password -y"
        Start-Process -FilePath $7zipPath -ArgumentList $extractArgs -Wait -NoNewWindow
        
        # Limpiar el ZIP después de extraer (mantener solo los archivos extraídos)
        Remove-Item -Path $tempAttZipPath -Force -ErrorAction SilentlyContinue
        
        "Archivo att.zip descargado y descomprimido exitosamente." | Out-File $logFile -Append
    }
    catch {
        "ADVERTENCIA: Error al descargar/descomprimir att.zip: $_" | Out-File $logFile -Append
        "El script continuará, pero las plantillas de Word pueden no estar disponibles." | Out-File $logFile -Append
    }
    
    "Ejecutando script completo..." | Out-File $logFile -Append
    "Ruta del script: $scriptPath" | Out-File $logFile -Append
    
    # Verificar que el script existe y tiene contenido
    if (Test-Path $scriptPath) {
        $scriptSize = (Get-Item $scriptPath).Length
        "Tamaño del script descargado: $scriptSize bytes" | Out-File $logFile -Append
    } else {
        throw "El script no se descargó correctamente o no existe en $scriptPath"
    }
    
    # Ejecutar el script descargado y capturar salida y errores
    try {
        $process = Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"" `
            -Wait `
            -NoNewWindow `
            -PassThru `
            -RedirectStandardOutput "$env:TEMP\script_output.txt" `
            -RedirectStandardError "$env:TEMP\script_errors.txt"
        
        $exitCode = $process.ExitCode
        "Script ejecutado. Código de salida: $exitCode" | Out-File $logFile -Append
        
        # Leer la salida del script si existe
        if (Test-Path "$env:TEMP\script_output.txt") {
            $output = Get-Content "$env:TEMP\script_output.txt" -ErrorAction SilentlyContinue
            if ($output) {
                "Salida del script:" | Out-File $logFile -Append
                $output | Out-File $logFile -Append
            }
        }
        
        # Leer los errores del script si existen
        if (Test-Path "$env:TEMP\script_errors.txt") {
            $errors = Get-Content "$env:TEMP\script_errors.txt" -ErrorAction SilentlyContinue
            if ($errors) {
                "Errores del script:" | Out-File $logFile -Append
                $errors | Out-File $logFile -Append
            }
        }
        
        if ($exitCode -ne 0) {
            "ADVERTENCIA: El script terminó con código de error $exitCode" | Out-File $logFile -Append
        }
    }
    catch {
        $errorMsg = "ERROR al ejecutar el script: $_"
        $errorMsg | Out-File $logFile -Append
        "Detalles: $($_.Exception.Message)" | Out-File $logFile -Append
        if ($_.ScriptStackTrace) {
            "Stack trace: $($_.ScriptStackTrace)" | Out-File $logFile -Append
        }
    }
    
    "Limpiando archivos temporales..." | Out-File $logFile -Append
    Remove-Item -Path $scriptPath -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:TEMP\script_output.txt" -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:TEMP\script_errors.txt" -Force -ErrorAction SilentlyContinue
    "Limpieza completada." | Out-File $logFile -Append
}
catch {
    $errorMsg = "ERROR al descargar o ejecutar el script: $_"
    $errorMsg | Out-File $logFile -Append
    "Detalles: $($_.Exception.Message)" | Out-File $logFile -Append
}
</powershell>
