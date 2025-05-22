# --- Function to Join ZeroTier Network ---
Function Join-ZeroTierNetwork {
    $networkJoined = $false
    $joinedNetworkId = $null
    while (-not $networkJoined) {
        $networkIdInput = Read-Host -Prompt "Please enter the ZeroTier Network ID to join"
        $networkId = $networkIdInput.Trim() # TRIM WHITESPACE

        if (-not ($networkId -match '^[0-9a-fA-F]{16}$')) {
            Write-Warning "Invalid Network ID format. It should be 16 hexadecimal characters. You entered (raw): '$networkIdInput', after trim: '$networkId'"
            continue
        }

        Write-Host "Attempting to join network '$networkId' using CLI at '$($Global:zeroTierCliPath)'..."
        
        # For debugging, let's see what the job receives
        $scriptBlockContent = {
            param($cliP, $netIdParam)
            Write-Output "Job received CLI Path: '$cliP'"
            Write-Output "Job received Network ID: '$netIdParam'"
            # Construct the command and arguments carefully
            $output = & $cliP join $netIdParam 2>&1 # Capture stderr as well
            Write-Output "Job command output: $output" # Output the full result from the job
            return $output # Return the output
        }
        
        $job = Start-Job -ScriptBlock $scriptBlockContent -ArgumentList $Global:zeroTierCliPath, $networkId

        if (Wait-Job $job -Timeout 30) {
            $rawOutputFromJob = Receive-Job $job
            
            # Process the array of output lines from the job
            $outputLines = @()
            if ($rawOutputFromJob -is [array]) {
                $outputLines = $rawOutputFromJob
            } else {
                $outputLines += $rawOutputFromJob
            }

            # Join lines and display for debugging
            $fullOutputString = ($outputLines | ForEach-Object {$_.ToString()}) -join [Environment]::NewLine
            Write-Host "--- Job Output Start ---"
            Write-Host $fullOutputString
            Write-Host "--- Job Output End ---"

            # Check for the success condition in the joined output
            if ($fullOutputString -match "200 join OK") {
                Write-Host "Successfully joined network $networkId." -ForegroundColor Green
                $networkJoined = $true
                $joinedNetworkId = $networkId
            } else {
                # Extract the relevant error from the CLI if possible, otherwise show full output
                $cliErrorMessage = $fullOutputString
                if ($fullOutputString -match "zerotier-one_cli.exe: invalid network id: `"$($networkId)`"") { # More specific error match
                    $cliErrorMessage = "ZeroTier CLI reported: invalid network id '$networkId'"
                } elseif ($fullOutputString -match "invalid network id") {
                     $cliErrorMessage = "ZeroTier CLI reported: invalid network id (check formatting or authorization)"
                }

                Write-Warning "Failed to join network $networkId. Response: $cliErrorMessage"
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
