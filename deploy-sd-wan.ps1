#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads, installs, and configures ZeroTier One on Windows, with re-installation support.
.DESCRIPTION
    This script performs the following actions:
    1. Checks if ZeroTier One is already installed.
    2. If installed, prompts the user if they wish to re-install.
       - If yes, it attempts to uninstall the existing version silently.
    3. If not installed, or if re-installation is chosen, it downloads the latest ZeroTier One MSI.
    4. Installs ZeroTier One silently in headless mode using ADDLOCAL=ALL.
    5. Prompts the user to input a ZeroTier Network ID.
    6. Attempts to join the network using 'zerotier-cli.bat join <NetworkID>'.
    7. Waits for a '200 join OK' response. If not received or if the command hangs for over 30 seconds, it re-prompts.
    8. Checks for and creates/sets the 'IPEnableRouter' registry value to 1.
    9. Outputs the ZeroTier Node ID, joined Network ID, and IP forwarding status.
    10. Prompts the user if they want to reboot now or later.
#>

# --- Script Configuration ---
$zeroTierDownloadUrl = "https://download.zerotier.com/dist/ZeroTier%20One.msi"
$msiPath = Join-Path -Path $env:TEMP -ChildPath "ZeroTierOne.msi"
$installLogPath = Join-Path -Path $env:TEMP -ChildPath "ZeroTierInstall.log"
$uninstallLogPath = Join-Path -Path $env:TEMP -ChildPath "ZeroTierUninstall.log"

# Attempt to find Program Files (respects system architecture for 32-bit PS on 64-bit OS)
$programFilesPath = ${env:ProgramFiles}
if (-not $programFilesPath) {
    $programFilesPath = ${env:ProgramW6432} # For 32-bit process on 64-bit OS
}
if (-not $programFilesPath) {
    $programFilesPath = "C:\Program Files" # Fallback, less ideal
}
$zeroTierCliPathDefault = Join-Path -Path $programFilesPath -ChildPath 'ZeroTier\One\zerotier-cli.bat'
$zeroTierCliPathX86Default = Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'ZeroTier\One\zerotier-cli.bat'

$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
$registryName = "IPEnableRouter"

# --- Function to Get ZeroTier Installation Status and ProductCode ---
Function Get-ZeroTierInstallationInfo {
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $zeroTierProduct = $null
    foreach ($keyPath in $uninstallKeys) {
        if (Test-Path $keyPath) {
            $zeroTierProduct = Get-ChildItem -Path $keyPath |
                               Get-ItemProperty |
                               Where-Object { $_.DisplayName -like "ZeroTier One*" } |
                               Select-Object -First 1 -Property DisplayName, DisplayVersion, ProductCode, UninstallString
            if ($zeroTierProduct) { break }
        }
    }
    return $zeroTierProduct # Contains DisplayName, DisplayVersion, ProductCode, UninstallString
}

# --- Function to Uninstall ZeroTier Silently ---
Function Uninstall-ZeroTier {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ProductInfo
    )

    Write-Host "Attempting to uninstall existing ZeroTier One (Version: $($ProductInfo.DisplayVersion))..."
    $uninstallCommand = ""
    $arguments = ""

    if ($ProductInfo.ProductCode) {
        Write-Host "Using ProductCode: $($ProductInfo.ProductCode) for uninstallation."
        $uninstallCommand = "msiexec.exe"
        $arguments = "/x $($ProductInfo.ProductCode) /qn /norestart /L*V `"$uninstallLogPath`""
    } elseif ($ProductInfo.UninstallString) {
        Write-Host "Using UninstallString: $($ProductInfo.UninstallString)"
        # Attempt to make it silent if possible
        $uninstallCommand = ($ProductInfo.UninstallString -split ' ')[0]
        $baseArguments = $ProductInfo.UninstallString -replace "^`"$uninstallCommand`"\s*|^$uninstallCommand\s*", ""
        # Common silent flags for MSI, some InstallShield installers might use /s, /S, /silent, /q, /quiet
        if ($ProductInfo.UninstallString -match 'msiexec') {
            $arguments = "$baseArguments /qn /norestart /L*V `"$uninstallLogPath`""
        } else {
            # For non-msiexec uninstallers, silent flags vary. This is a best guess.
            $arguments = "$baseArguments /S /qn /norestart" # Common silent flags
            Write-Warning "Silent uninstallation for non-MSI UninstallString is a best guess. Monitor the process."
        }
    } else {
        Write-Error "Could not determine uninstallation method for ZeroTier One."
        return $false
    }

    try {
        Write-Host "Executing: $uninstallCommand $arguments"
        $process = Start-Process -FilePath $uninstallCommand -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) { # 3010 = success, reboot required
            Write-Host "ZeroTier One uninstalled successfully (Exit Code: $($process.ExitCode))."
            if ($process.ExitCode -eq 3010) {
                Write-Warning "A reboot is required to complete the uninstallation of the previous version."
            }
            # Wait a bit for system to settle after uninstall
            Start-Sleep -Seconds 10
            return $true
        } else {
            Write-Error "ZeroTier uninstallation failed. Exit code: $($process.ExitCode)."
            Write-Error "Please check the uninstallation log: $uninstallLogPath (if created by msiexec)"
            return $false
        }
    }
    catch {
        Write-Error "An error occurred during ZeroTier uninstallation: $($_.Exception.Message)"
        return $false
    }
}

# --- Function to Download ZeroTier ---
Function Download-ZeroTier {
    Write-Host "Downloading ZeroTier One MSI..."
    try {
        Invoke-WebRequest -Uri $zeroTierDownloadUrl -OutFile $msiPath -ErrorAction Stop
        Write-Host "Download complete: $msiPath"
    }
    catch {
        Write-Error "Failed to download ZeroTier MSI. Error: $($_.Exception.Message)"
        Write-Error "Please check the URL or your internet connection and try again."
        Write-Error "Script will exit."
        Read-Host "Press Enter to exit."
        Exit 1
    }
}

# --- Function to Install ZeroTier Silently ---
Function Install-ZeroTier {
    Write-Host "Installing ZeroTier One silently..."
    $msiArgs = @(
        "/i `"$msiPath`""
        "/qn"                          # Quiet, no UI
        "RUN_SERVICE=1"                # Ensure service is set to run
        "START_SERVICE_AFTER_INSTALL=1"# Start service after install
        "ADDLOCAL=ALL"                 # Install all core features (Fix for 1603 Error 2711)
        "/norestart"                   # Do not automatically restart
        "/L*V `"$installLogPath`""     # Verbose logging
    )
    try {
        $process = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) { # 3010 means success, reboot required
            Write-Host "ZeroTier One installed successfully (Exit Code: $($process.ExitCode))."
            if ($process.ExitCode -eq 3010) {
                Write-Warning "A reboot is recommended for the new ZeroTier installation to take full effect."
                $Global:RebootNeededFromInstall = $true
            }
            # Brief pause to ensure service is fully started
            Write-Host "Waiting for ZeroTier service to initialize..."
            Start-Sleep -Seconds 15

            # Determine actual CLI path after installation
            if (Test-Path $zeroTierCliPathDefault) {
                $Global:zeroTierCliPath = $zeroTierCliPathDefault
            } elseif (Test-Path $zeroTierCliPathX86Default) {
                $Global:zeroTierCliPath = $zeroTierCliPathX86Default
            } else {
                Write-Error "ZeroTier CLI not found at expected locations after installation: $zeroTierCliPathDefault or $zeroTierCliPathX86Default"
                Write-Error "Please check the installation log: $installLogPath"
                # Attempt to locate it more broadly in ProgramData (common for ZT service path)
                $ztServicePath = (Get-Service ZeroTierOneService -ErrorAction SilentlyContinue).PathName -replace '"',''
                if ($ztServicePath) {
                    $potentialCliPath = Join-Path (Split-Path $ztServicePath) "zerotier-cli.bat"
                    if (Test-Path $potentialCliPath) {
                        $Global:zeroTierCliPath = $potentialCliPath
                        Write-Warning "ZeroTier CLI found at non-standard path: $Global:zeroTierCliPath"
                    } else {
                        Read-Host "Press Enter to exit."
                        Exit 1
                    }
                } else {
                     Read-Host "Press Enter to exit."
                     Exit 1
                }
            }
            Write-Host "ZeroTier CLI will be used from: $Global:zeroTierCliPath"
            return $true
        } else {
            Write-Error "ZeroTier installation failed. Exit code: $($process.ExitCode)."
            Write-Error "Please check the installation log: $installLogPath"
            return $false
        }
    }
    catch {
        Write-Error "An error occurred during ZeroTier installation: $($_.Exception.Message)"
        Write-Error "Please check the installation log: $installLogPath"
        return $false
    }
}

# --- Function to Join ZeroTier Network ---
Function Join-ZeroTierNetwork {
    $networkJoined = $false
    $joinedNetworkId = $null
    while (-not $networkJoined) {
        $networkId = Read-Host -Prompt "Please enter the ZeroTier Network ID to join"
        if (-not ($networkId -match '^[0-9a-fA-F]{16}$')) {
            Write-Warning "Invalid Network ID format. It should be 16 hexadecimal characters."
            continue
        }

        Write-Host "Attempting to join network $networkId using $($Global:zeroTierCliPath)..."
        $job = Start-Job -ScriptBlock {
            param($cliP, $netId)
            & $cliP join $netId
        } -ArgumentList $Global:zeroTierCliPath

        if (Wait-Job $job -Timeout 30) {
            $output = Receive-Job $job
            Write-Host "Join command output: $output"
            if ($output -match "200 join OK") {
                Write-Host "Successfully joined network $networkId." -ForegroundColor Green
                $networkJoined = $true
                $joinedNetworkId = $networkId
            } else {
                Write-Warning "Failed to join network $networkId. Response: $output"
                Write-Warning "Please check the Network ID and ensure the ZeroTier service is running and you are authorized on the network."
                $ztService = Get-Service -Name "ZeroTierOneService" -ErrorAction SilentlyContinue
                if ($ztService -and $ztService.Status -ne "Running") {
                    Write-Warning "ZeroTier service is not running. Attempting to start it..."
                    Start-Service -Name "ZeroTierOneService" -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 5
                }
            }
        } else {
            Write-Warning "Joining network $networkId timed out after 30 seconds."
            Write-Warning "Please check your internet connection and the ZeroTier network authorization."
            Get-Job | Where-Object {$_.Id -eq $job.Id} | Stop-Job -Force
        }
        Get-Job | Where-Object {$_.Id -eq $job.Id} | Remove-Job -Force
    }
    return $joinedNetworkId
}

# --- Function to Configure IP Forwarding ---
Function Configure-IPForwarding {
    Write-Host "Checking and configuring IPEnableRouter (IP Forwarding)..."
    try {
        $currentValue = Get-ItemProperty -Path $registryPath -Name $registryName -ErrorAction SilentlyContinue
        if ($currentValue -and $currentValue.$registryName -eq 1) {
            Write-Host "IPEnableRouter is already set to 1."
            return "Enabled (Value: 1)"
        } else {
            Set-ItemProperty -Path $registryPath -Name $registryName -Value 1 -Type DWord -Force
            Write-Host "IPEnableRouter has been set to 1. A reboot is typically required for this to take full effect."
            $Global:RebootNeededFromConfig = $true
            return "Enabled (Value: 1) - Change applied"
        }
    }
    catch {
        Write-Error "Failed to configure IPEnableRouter. Error: $($_.Exception.Message)"
        return "Configuration failed"
    }
}

# --- Function to Get ZeroTier Node ID ---
Function Get-ZeroTierNodeId {
    Write-Host "Retrieving ZeroTier Node ID..."
    try {
        Start-Sleep -Seconds 5
        $nodeInfoOutput = & $Global:zeroTierCliPath info
        $match = $nodeInfoOutput | Select-String -Pattern '200 info ([0-9a-fA-F]{10})'
        if ($match) {
            $nodeId = $match.Matches[0].Groups[1].Value
            Write-Host "ZeroTier Node ID: $nodeId" -ForegroundColor Green
            return $nodeId
        } else {
            Write-Warning "Could not parse Node ID from 'zerotier-cli info' output: $nodeInfoOutput"
            $networksOutput = & $Global:zeroTierCliPath listnetworks -j
            $networksJson = $networksOutput | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($networksJson) {
                foreach ($net in $networksJson) {
                    if ($net.portDeviceName -and $net.address) {
                        $nodeIdFromNet = $net.portDeviceName
                         if ($nodeIdFromNet -match '^[0-9a-fA-F]{10}$') {
                            Write-Host "ZeroTier Node ID (derived from listnetworks): $nodeIdFromNet" -ForegroundColor Green
                            return $nodeIdFromNet
                         }
                    }
                }
            }
            Write-Warning "Could not determine Node ID via alternative methods."
            return "Unknown"
        }
    }
    catch {
        Write-Error "Failed to retrieve ZeroTier Node ID. Error: $($_.Exception.Message)"
        return "Error retrieving"
    }
}

# --- Main Script Execution ---
$Global:zeroTierCliPath = $null # Will be set after install or if already installed
$Global:RebootNeededFromInstall = $false
$Global:RebootNeededFromConfig = $false
$installationPerformed = $false

Write-Host "Starting ZeroTier Installation and Configuration Script..."

$existingInstall = Get-ZeroTierInstallationInfo

if ($existingInstall) {
    Write-Host "ZeroTier One (Version: $($existingInstall.DisplayVersion)) is already installed."
    # Set CLI path if found in existing install, for potential use even if not reinstalling
    if (Test-Path $zeroTierCliPathDefault) { $Global:zeroTierCliPath = $zeroTierCliPathDefault }
    elseif (Test-Path $zeroTierCliPathX86Default) { $Global:zeroTierCliPath = $zeroTierCliPathX86Default }
    else { # Try to find it via service path if standard paths fail
        $ztServicePath = (Get-Service ZeroTierOneService -ErrorAction SilentlyContinue).PathName -replace '"',''
        if ($ztServicePath) {
            $potentialCliPath = Join-Path (Split-Path $ztServicePath) "zerotier-cli.bat"
            if (Test-Path $potentialCliPath) { $Global:zeroTierCliPath = $potentialCliPath }
        }
    }
    if ($Global:zeroTierCliPath) {Write-Host "Existing ZeroTier CLI found at: $($Global:zeroTierCliPath)"}


    $choice = Read-Host "Would you like to remove the existing version and re-install the latest? (yes/no)"
    if ($choice -match '^y(es)?$') {
        if (Uninstall-ZeroTier -ProductInfo $existingInstall) {
            Download-ZeroTier
            if (Install-ZeroTier) {
                $installationPerformed = $true
            } else {
                Write-Error "Failed to install ZeroTier after uninstallation. Script will exit."
                Read-Host "Press Enter to exit."
                Exit 1
            }
        } else {
            Write-Error "Failed to uninstall existing ZeroTier. Script will exit."
            Read-Host "Press Enter to exit."
            Exit 1
        }
    } else {
        Write-Host "Re-installation skipped. Proceeding with configuration steps if possible..."
        # If not reinstalling, we assume the existing installation is what we want to configure
        if (-not $Global:zeroTierCliPath) {
            Write-Error "ZeroTier CLI path could not be determined for the existing installation. Cannot proceed with configuration."
            Read-Host "Press Enter to exit."
            Exit 1
        }
        $installationPerformed = $false # Indicate we are working with an existing, un-modified install
    }
} else {
    Write-Host "ZeroTier One is not currently installed."
    Download-ZeroTier
    if (Install-ZeroTier) {
        $installationPerformed = $true
    } else {
        Write-Error "Failed to install ZeroTier. Script will exit."
        Read-Host "Press Enter to exit."
        Exit 1
    }
}

# Proceed with configuration if ZeroTier is (now) installed and CLI path is known
if ($Global:zeroTierCliPath) {
    $joinedNetwork = Join-ZeroTierNetwork
    $nodeIdentity = Get-ZeroTierNodeId
    $ipForwardingStatus = Configure-IPForwarding

    Write-Host "`n--- Configuration Summary ---" -ForegroundColor Cyan
    Write-Host "ZeroTier Node ID     : $nodeIdentity"
    Write-Host "Joined Network ID    : $joinedNetwork"
    Write-Host "IPEnableRouter Status: $ipForwardingStatus"

    if ($Global:RebootNeededFromInstall -or $Global:RebootNeededFromConfig) {
        Write-Host "`nA reboot is recommended for all changes to take full effect."
        $rebootChoice = ''
        while ($rebootChoice -notmatch '^(y(es)?|n(o)?)$') {
            $rebootChoice = (Read-Host -Prompt "Do you want to reboot now? (yes/no)").ToLower()
        }

        if ($rebootChoice -match '^y(es)?$') {
            Write-Host "Rebooting now..."
            Restart-Computer -Force
        } else {
            Write-Host "Please reboot your computer later to apply all changes."
        }
    } else {
        Write-Host "`nConfiguration complete. No immediate reboot explicitly required by this script's actions, but consider one if issues arise."
    }
} elseif ($installationPerformed) {
    # This case means installation was attempted and reported success, but CLI path somehow wasn't set.
    Write-Error "Installation was performed but ZeroTier CLI path could not be determined. Configuration cannot proceed."
} else {
    # This case means ZeroTier was found, user chose not to reinstall, but CLI path still couldn't be found OR initial state was 'not installed' and install function failed to set CLI path.
    Write-Warning "ZeroTier installation was either skipped or failed in a way that the CLI path could not be set. Cannot proceed with network configuration."
}


Write-Host "Script finished."
Read-Host "Press Enter to exit."
