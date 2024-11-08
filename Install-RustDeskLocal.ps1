# Local RustDesk Installation Script
$ErrorActionPreference = 'Stop'

function Write-Log {
    param($Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    if ($IsError) {
        Write-Host $logMessage -ForegroundColor Red
    } else {
        Write-Host $logMessage
    }
}

try {
    Write-Log "Starting local RustDesk installation..."
    
    # Setup variables
    $tempDir = "C:\Windows\Temp\RustDeskInstall"
    $msiPath = Join-Path $tempDir "coatsinkrustdesk.msi"
    $configPath = "C:\Windows\ServiceProfiles\LocalService\AppData\Roaming\coatsinkrustdesk\config\coatsinkrustdesk.toml"
    $msiUrl = "https://github.com/CoatsinkIT/RustDesk/raw/refs/heads/main/coatsinkrustdesk.msi"
    
    # Get computer name and clean it
    $hostname = $env:COMPUTERNAME
    $cleanId = $hostname -replace '-', ''
    $uniquePassword = "$cleanId`xC0ats1nk.c0m"
    
    Write-Log "Computer Name: $hostname"
    Write-Log "Clean ID: $cleanId"
    
    # Create temp directory
    if (-not (Test-Path $tempDir)) {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        Write-Log "Created temporary directory"
    }
    
    # Download MSI
    Write-Log "Downloading RustDesk MSI..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
    
    if (Test-Path $msiPath) {
        # Stop service if running
        Write-Log "Stopping CoatsinkRustDesk service..."
        Stop-Service -Name "CoatsinkRustDesk" -Force -ErrorAction SilentlyContinue
        
        # Install MSI
        Write-Log "Installing RustDesk..."
        $process = Start-Process "msiexec.exe" -ArgumentList "/i `"$msiPath`" /quiet /norestart" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Log "MSI Installation successful"
            Start-Sleep -Seconds 5
            
            # Wait for config file
            $retryCount = 0
            while (-not (Test-Path $configPath) -and $retryCount -lt 10) {
                Start-Sleep -Seconds 2
                $retryCount++
            }
            
            if (Test-Path $configPath) {
                # Stop service for configuration
                Write-Log "Stopping service for configuration..."
                Stop-Service -Name "CoatsinkRustDesk" -Force
                Start-Sleep -Seconds 2
                
                # Read and update config
                Write-Log "Updating configuration..."
                $tomlContent = Get-Content -Path $configPath -Raw
                
                $newTomlContent = @"
id = '$cleanId'
password = '$uniquePassword'
$(($tomlContent -split "`n" | Select-Object -Skip 2) -join "`n")
"@
                
                # Backup original config
                Copy-Item -Path $configPath -Destination "$configPath.bak" -Force
                
                # Save new config
                $newTomlContent | Set-Content -Path $configPath -Force
                Write-Log "Configuration updated"
                
                # Start service
                Write-Log "Starting service..."
                Start-Service -Name "CoatsinkRustDesk"
                Start-Sleep -Seconds 2
                
                # Verify service
                $serviceStatus = Get-Service -Name "CoatsinkRustDesk"
                if ($serviceStatus.Status -eq 'Running') {
                    Write-Log "Installation completed successfully!"
                    Write-Log "RustDesk ID: $cleanId"
                    Write-Log "RustDesk Password: $uniquePassword"
                }
                else {
                    throw "Service failed to start after configuration"
                }
            }
            else {
                throw "Configuration file not found after installation"
            }
        }
        else {
            throw "Installation failed with exit code: $($process.ExitCode)"
        }
    }
    else {
        throw "Failed to download MSI file"
    }
}
catch {
    Write-Log "Error during installation: $_" -IsError
    Write-Log "Installation failed!" -IsError
}
finally {
    # Cleanup
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned up temporary files"
    }
    
    Write-Log "Script completed"
    Write-Host "`nPress any key to exit..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}
