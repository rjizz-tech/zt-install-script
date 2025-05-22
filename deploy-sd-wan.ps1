#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Downloads, installs, and configures ZeroTier One on Windows, with re-installation support, logging, and status checks.
.DESCRIPTION
    This script performs the following actions:
    1.  Starts transcript logging to sd-wan-install.log in the %TEMP% directory.
    2.  Checks if ZeroTier One is already installed using a robust registry check.
    3.  If installed, prompts the user if they wish to re-install.
        - If yes, it attempts to uninstall the existing version silently.
    4.  If not installed, or if re-installation is chosen, it downloads the latest ZeroTier One MSI.
    5.  Installs ZeroTier One silently in headless mode using ZTHEADLESS=Yes.
    6.  Prompts the user to input a ZeroTier Network ID, trimming any whitespace.
    7.  Attempts to join the network using 'zerotier-cli.bat join <NetworkID>' with detailed job output.
    8.  Waits for a '200 join OK' response. If not received or if the command hangs for over 30 seconds, it re-prompts.
    9.  Checks for and creates/sets the 'IPEnableRouter' registry value to 1.
    10. Outputs the ZeroTier Node ID, joined Network ID, and IP forwarding status.
    11. After successful join, checks 'zerotier-cli listnetworks' for ACCESS_DENIED status for the joined network.
    12. Prompts the user if they want to reboot now or later, based on script actions.
    13. Stops transcript logging.
.NOTES
    Version: 1.5
    Requires: Administrator privileges to install software and modify the registry.
              Internet connectivity to download ZeroTier and join the network.
#>

# --- Script Configuration ---
$scriptLogPath = Join-Path -Path $env:TEMP -ChildPath "sd-wan-install.log"
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
    $programFilesPath = "C:\Program Files" # Fallback
}
$zeroTierCliPathDefault = Join-Path -Path $programFilesPath -ChildPath 'ZeroTier\One\zerotier-cli.bat'
$zeroTierCliPathX86Default = Join-Path -Path ${env:ProgramFiles(x86)} -ChildPath 'ZeroTier\One\zerotier-cli.bat'

$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
$registryName = "IPEnableRouter"

# --- Function to Get ZeroTier Installation Status and ProductCode ---
Function Get-ZeroTierInstallationInfo {
    Write-Verbose "(Get-ZeroTierInstallationInfo): Function started."
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $zeroTierProduct = $null
    try {
        foreach ($keyPath in $uninstallKeys) {
            Write-Verbose "(Get-ZeroTierInstallationInfo): Checking registry path: $keyPath"
            if (Test-Path $keyPath) {
                Write-Verbose "(Get-ZeroTierInstallationInfo): Path exists: $keyPath"
                
                $productRegistryKeys = Get-ChildItem -Path $keyPath -ErrorAction SilentlyContinue
                if ($null -eq $productRegistryKeys) {
                    Write-Verbose "(Get-ZeroTierInstallationInfo): No subkeys found under $keyPath or error listing them."
                    continue
                }

                foreach ($productRegKey in $productRegistryKeys) {
                    Write-Verbose "(Get-ZeroTierInstallationInfo): Processing subkey: $($productRegKey.PSChildName)"
                    try {
                        $properties = $productRegKey | Get-ItemProperty -ErrorAction Stop 
                        
                        if ($properties.PSObject.Properties['DisplayName'] -and $properties.DisplayName -like "ZeroTier One*") {
                            Write-Verbose "(Get-ZeroTierInstallationInfo): Found matching product: $($properties.DisplayName)"
                            $zeroTierProduct = $properties | Select-Object -Property DisplayName, DisplayVersion, ProductCode, UninstallString
                            break 
                        }
                    } catch {
                        Write-Warning "(Get-ZeroTierInstallationInfo): Error processing a specific registry entry '$($productRegKey.Name)'. Error: $($_.Exception.Message)."
                        Write-Verbose "(Get-ZeroTierInstallationInfo): Faulty Key Path: $($productRegKey.PSPath)"
                    }
                } 
                
                if ($zeroTierProduct) {
                    Write-Verbose "(Get-ZeroTierInstallationInfo): ZeroTier product found and processed. Breaking from outer loop."
                    break 
                }
            } else {
                Write-Verbose "(Get-ZeroTierInstallationInfo): Path does not exist: $keyPath"
            }
        } 
    } catch {
        Write-Error "(Get-ZeroTierInstallationInfo): An unexpected error occurred during overall processing: $($_.Exception.Message)"
    }
    Write-Verbose "(Get-ZeroTierInstallationInfo): Function finished. Product found: $($zeroTierProduct -ne $null)"
    return $zeroTierProduct
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
        $uninstallExecutable = ($ProductInfo.UninstallString -split ' ')[0].Replace('"','')
        $baseArguments = $ProductInfo.UninstallString -replace "^`"$uninstallExecutable`"\s*|^$uninstallExecutable\s*", ""
        
        $uninstallCommand = $uninstallExecutable
        if ($uninstallCommand -match 'msiexec') {
            $arguments = "$baseArguments /qn /norestart /L*V `"$uninstallLogPath`""
        } else {
            $arguments = "$baseArguments /S /qn /norestart" 
            Write-Warning "Silent uninstallation for non-MSI UninstallString is a best guess. Monitor the process."
        }
    } else {
        Write-Error "Could not determine uninstallation method for ZeroTier One."
        return $false
    }

    try {
        Write-Host "Executing: `"$uninstallCommand`" $arguments"
        $process = Start-Process -FilePath $uninstallCommand -ArgumentList $arguments -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) { 
            Write-Host "ZeroTier One uninstalled successfully (Exit Code: $($process.ExitCode))."
            if ($process.ExitCode -eq 3010) {
                Write-Warning "A reboot is required to complete the uninstallation of the previous version."
                $Global:RebootNeededFromInstall = $true
            }
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
        return $true
    }
    catch {
        Write-Error "Failed to download ZeroTier MSI. Error: $($_.Exception.Message)"
        Write-Error "Please check the URL or your internet connection and try again."
        return $false
    }
}

# --- Function to Install ZeroTier Silently ---
Function Install-ZeroTier {
    Write-Host "Installing ZeroTier One silently..."
    $msiArgs = @(
        "/i `"$msiPath`""
        "/qn"                          
        "RUN_SERVICE=1"                
        "START_SERVICE_AFTER_INSTALL=1"
        "ZTHEADLESS=Yes"               # Use ZTHEADLESS property for headless install
        "/norestart"                   
        "/L*V `"$installLogPath`""     
    )
    try {
        $process = Start-Process msiexec.exe -ArgumentList $msiArgs -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) { 
            Write-Host "ZeroTier One installed successfully (Exit Code: $($process.ExitCode))."
            if ($process.ExitCode -eq 3010) {
                Write-Warning "A reboot is recommended for the new ZeroTier installation to take full effect."
                $Global:RebootNeededFromInstall = $true
            }
            Write-Host "Waiting for ZeroTier service to initialize..."
            Start-Sleep -Seconds 15

            if (Test-Path $zeroTierCliPathDefault) {
                $Global:zeroTierCliPath = $zeroTierCliPathDefault
            } elseif (Test-Path $zeroTierCliPathX86Default) {
                $Global:zeroTierCliPath = $zeroTierCliPathX86Default
            } else {
                $ztServicePath = (Get-Service ZeroTierOneService -ErrorAction SilentlyContinue).PathName -replace '"',''
                if ($ztServicePath) {
                    $potentialCliPath = Join-Path (Split-Path $ztServicePath) "zerotier-cli.bat"
                    if (Test-Path $potentialCliPath) {
                        $Global:zeroTierCliPath = $potentialCliPath
                        Write-Warning "ZeroTier CLI found at non-standard path: $Global:zeroTierCliPath"
                    } else {
                        Write-Error "ZeroTier CLI not found at expected locations or derived service path after installation."
                        Write-Error "Please check the installation log: $installLogPath"
                        return $false 
                    }
                } else {
                     Write-Error "ZeroTier service not found after install, cannot derive CLI path."
                     return $false 
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
        $networkIdInput = Read-Host -Prompt "Please enter the ZeroTier Network ID to join"
        $networkId = $networkIdInput.Trim() 

        if (-not ($networkId -match '^[0-9a-fA-F]{16}$')) {
            Write-Warning "Invalid Network ID format. It should be 16 hexadecimal characters. You entered (raw): '$networkIdInput', after trim: '$networkId'"
            continue
        }

        Write-Host "Attempting to join network '$networkId' using CLI at '$($Global:zeroTierCliPath)'..."
        
        $scriptBlockContent = {
            param($cliP, $netIdParam)
            # These Write-Output lines are for job context debugging, they will appear in the transcript.
            Write-Output "Job received CLI Path: '$cliP'"
            Write-Output "Job received Network ID: '$netIdParam'"
            $output = & $cliP join $netIdParam 2>&1 
            Write-Output "Job command output: $output" 
            return $output 
        }
        
        $job = Start-Job -ScriptBlock $scriptBlockContent -ArgumentList $Global:zeroTierCliPath, $networkId

        if (Wait-Job $job -Timeout 30) {
            $rawOutputFromJob = Receive-Job $job
            $outputLines = @()
            if ($rawOutputFromJob -is [array]) {
                $outputLines = $rawOutputFromJob
            } else {
                $outputLines += $rawOutputFromJob
            }

            $fullOutputString = ($outputLines | ForEach-Object {$_.ToString()}) -join [Environment]::NewLine
            Write-Host "--- Job Output Start (captured by script) ---"
            Write-Host $fullOutputString
            Write-Host "--- Job Output End (captured by script) ---"

            if ($fullOutputString -match "200 join OK") {
                Write-Host "Successfully joined network $networkId." -ForegroundColor Green
                $networkJoined = $true
                $joinedNetworkId = $networkId
            } else {
                $cliErrorMessage = $fullOutputString # Default to full output
                if ($fullOutputString -match "zerotier-one_cli.exe: invalid network id: `"$($networkId)`"") { 
                    $cliErrorMessage = "ZeroTier CLI reported: invalid network id '$networkId'"
                } elseif ($fullOutputString -match "invalid network id") {
                     $cliErrorMessage = "ZeroTier CLI reported: invalid network id (check formatting or authorization)"
                }

                Write-Warning "Failed to join network $networkId. Response from CLI: $cliErrorMessage"
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
        $nodeInfoOutput = & $Global:zeroTierCliPath info 2>&1
        $match = $nodeInfoOutput | Select-String -Pattern '200 info ([0-9a-fA-F]{10})'
        if ($match) {
            $nodeId = $match.Matches[0].Groups[1].Value
            Write-Host "ZeroTier Node ID: $nodeId" -ForegroundColor Green
            return $nodeId
        } else {
            Write-Warning "Could not parse Node ID from 'zerotier-cli info' output: $nodeInfoOutput"
            $networksOutput = & $Global:zeroTierCliPath listnetworks -j 2>&1
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

# Remove explicit VerbosePreference setting for cleaner logs, relies on default or user set.
# $Global:VerbosePreference = "Continue" 

# Start Transcript Logging
try {
    Start-Transcript -Path $scriptLogPath -Append -Force -ErrorAction SilentlyContinue
} catch {
    Write-Warning "Failed to start transcript logging to '$scriptLogPath'. Error: $($_.Exception.Message)"
}

Write-Host "Script execution started. Initializing variables."

$Global:zeroTierCliPath = $null 
$Global:RebootNeededFromInstall = $false
$Global:RebootNeededFromConfig = $false
$Global:ScriptCrashed = $false
$performFullConfiguration = $false 
$finalJoinedNetworkId = $null # To store the successfully joined network ID for ACCESS_DENIED check

try {
    Write-Host "Starting ZeroTier Installation and Configuration Script..."

    $existingInstall = Get-ZeroTierInstallationInfo
    
    if ($existingInstall -ne $null) {
        Write-Host "ZeroTier One (Version: $($existingInstall.DisplayVersion)) is already installed."
        
        if (Test-Path $zeroTierCliPathDefault) { $Global:zeroTierCliPath = $zeroTierCliPathDefault }
        elseif (Test-Path $zeroTierCliPathX86Default) { $Global:zeroTierCliPath = $zeroTierCliPathX86Default }
        else { 
            $ztServicePath = (Get-Service ZeroTierOneService -ErrorAction SilentlyContinue).PathName -replace '"',''
            if ($ztServicePath) {
                $potentialCliPath = Join-Path (Split-Path $ztServicePath) "zerotier-cli.bat"
                if (Test-Path $potentialCliPath) { $Global:zeroTierCliPath = $potentialCliPath }
            }
        }
        if ($Global:zeroTierCliPath) {Write-Host "Existing ZeroTier CLI found at: $($Global:zeroTierCliPath)"}

        $choice = ''
        while ($choice -notmatch '^(y(es)?|n(o)?)$') {
            $choice = (Read-Host "Would you like to remove the existing version and re-install the latest? (yes/no)").ToLower()
        }

        if ($choice -match '^y(es)?$') {
            if (Uninstall-ZeroTier -ProductInfo $existingInstall) {
                if (Download-ZeroTier) {
                    if (Install-ZeroTier) {
                        $performFullConfiguration = $true
                    } else {
                        Write-Error "Failed to install ZeroTier after uninstallation."
                    }
                } else {
                     Write-Error "Failed to download ZeroTier after uninstallation."
                }
            } else {
                Write-Error "Failed to uninstall existing ZeroTier."
            }
        } else {
            Write-Host "Re-installation skipped. Proceeding with configuration steps for existing installation if CLI path is known."
            if (-not $Global:zeroTierCliPath) {
                Write-Error "ZeroTier CLI path could not be determined for the existing installation. Cannot proceed with configuration."
            } else {
                $performFullConfiguration = $true 
            }
        }
    } else {
        Write-Host "ZeroTier One is not currently installed. Proceeding with fresh installation."
        if (Download-ZeroTier) {
            if (Install-ZeroTier) {
                $performFullConfiguration = $true
            } else {
                Write-Error "Failed to install ZeroTier."
            }
        } else {
             Write-Error "Failed to download ZeroTier."
        }
    }

    if ($performFullConfiguration -and $Global:zeroTierCliPath) {
        $finalJoinedNetworkId = Join-ZeroTierNetwork # Store the joined network ID
        $nodeIdentity = Get-ZeroTierNodeId
        $ipForwardingStatus = Configure-IPForwarding

        Write-Host "`n--- Configuration Summary ---" -ForegroundColor Cyan
        Write-Host "ZeroTier Node ID     : $nodeIdentity"
        Write-Host "Joined Network ID    : $finalJoinedNetworkId"
        Write-Host "IPEnableRouter Status: $ipForwardingStatus"

        # Check for ACCESS_DENIED status
        if ($finalJoinedNetworkId) {
            Write-Host "Checking network status for $finalJoinedNetworkId..."
            try {
                $networksListJson = & $Global:zeroTierCliPath listnetworks -j 2>&1
                $networksList = $networksListJson | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($networksList) {
                    $joinedNetworkInfo = $networksList | Where-Object {$_.nwid -eq $finalJoinedNetworkId} | Select-Object -First 1
                    if ($joinedNetworkInfo) {
                        Write-Host "Status for network $finalJoinedNetworkId : $($joinedNetworkInfo.status)"
                        if ($joinedNetworkInfo.status -eq "ACCESS_DENIED") {
                            Write-Host "-------------------------------------------------------------------" -ForegroundColor Yellow
                            Write-Host "Network join is successful, ready for GBS IT to configure." -ForegroundColor Yellow
                            Write-Host "-------------------------------------------------------------------" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Warning "Could not find information for network $finalJoinedNetworkId in listnetworks output."
                    }
                } else {
                    Write-Warning "Failed to parse 'listnetworks -j' output or it was empty. JSON Output: $networksListJson"
                }
            } catch {
                Write-Warning "Could not check network status via 'listnetworks -j'. Error: $($_.Exception.Message)"
            }
        }

    } elseif ($performFullConfiguration -and -not $Global:zeroTierCliPath) {
        Write-Error "Installation was indicated as successful or skipped to use existing, but ZeroTier CLI path could not be determined. Configuration cannot proceed."
    } else {
        Write-Warning "ZeroTier installation/download failed or was skipped without a valid existing CLI. Network configuration did not proceed."
    }
    
    Write-Host "Main script logic within try block completed."

} catch {
    $Global:ScriptCrashed = $true 
    Write-Error "CRITICAL SCRIPT ERROR: An unhandled exception occurred in the main script block:"
    Write-Error "Error Type: $($_.Exception.GetType().FullName)"
    Write-Error "Error Message: $($_.Exception.Message)"
    Write-Error "Stack Trace: $($_.Exception.StackTrace | Out-String)" 
    Write-Error "Script Line Number: $($_.InvocationInfo.ScriptLineNumber)"
    Write-Error "Faulting Line Content: $($_.InvocationInfo.PositionMessage)"
} finally {
    Write-Host "Execution reached the 'finally' block."
    
    if ($Global:RebootNeededFromInstall -or $Global:RebootNeededFromConfig) {
        Write-Host "`nA reboot is recommended for all changes to take full effect."
        $rebootChoiceFinal = ''
        if (-not $Global:ScriptCrashed -or $performFullConfiguration) {
            while ($rebootChoiceFinal -notmatch '^(y(es)?|n(o)?)$') {
                $rebootChoiceFinal = (Read-Host -Prompt "Do you want to reboot now? (yes/no)").ToLower()
            }
            if ($rebootChoiceFinal -match '^y(es)?$') {
                Write-Host "Rebooting now..."
                Restart-Computer -Force
            } else {
                Write-Host "Please reboot your computer later to apply all changes."
            }
        } else {
             Write-Warning "Script encountered issues, automatic reboot prompt skipped. Please review messages and reboot manually if needed."
        }
    } elseif (-not $Global:ScriptCrashed -and $performFullConfiguration) {
        Write-Host "`nConfiguration attempt complete. No immediate reboot explicitly required by this script's actions, but consider one if issues arise."
    }

    if ($Global:ScriptCrashed) {
        Read-Host "CRITICAL ERROR OCCURRED. Review messages above. Press Enter to exit script."
    } else {
        Read-Host "Script execution finished or was exited. Press Enter to close window." 
    }

    # Stop Transcript Logging
    Stop-Transcript -ErrorAction SilentlyContinue
}
