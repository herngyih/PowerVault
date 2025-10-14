# Read VM names from file (one per line)
$VMs = Get-Content C:\scripts\computers.txt

# Optional: ensure no blank lines or whitespace
$VMs = $VMs | Where-Object { $_.Trim() -ne "" }

# Initialize result array
$results = @()

# Initialize counters
$total = $VMs.Count
$index = 0

foreach ($vm in $VMs) {
    $index++
    Write-Host "$index out of $total - Copying installer to $vm..." -NoNewline
    $session = $null
    try {
        # Establish remote session
        $session = New-PSSession -ComputerName $vm -ErrorAction Stop

        # Ensure target directory exists on remote machine
        Invoke-Command -Session $session -ScriptBlock {
            New-Item -ItemType Directory -Path "C:\NessusAgent" -Force | Out-Null
        }

        # Copy installer to remote machine
        Copy-Item -Path "E:\NessusAgent-10.9.0-x64.msi" -Destination "C:\NessusAgent\NessusAgent-10.9.0-x64.msi" -ToSession $session -Force
        Write-Host "Installing..." -NoNewline
        # Run silent Nessus Agent install over session
        Invoke-Command -Session $session -ScriptBlock {
            Start-Process msiexec.exe -ArgumentList '/i "C:\NessusAgent\NessusAgent-10.9.0-x64.msi" NESSUS_GROUPS="Windows-Servers" NESSUS_SERVER="10.10.2.185:8834" NESSUS_KEY=5c1abba71c32ee4c84cb46cc171eed3f8c195a7d7267f333ffe2364545fd4b5b /qn' -Wait -NoNewWindow
        }

        # Log success
        Write-Host "Done"
        $results += [PSCustomObject]@{
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            VM        = $vm
            Result    = "pass"
        }
    }
    catch {
        Write-Host "Fail"
        # Log failure with error message
        $results += [PSCustomObject]@{
            Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            VM        = $vm
            Result    = "fail: $($_.Exception.Message)"
        }
    }
    finally {
        # Always remove session to avoid leaks
        if ($session) {
            Remove-PSSession $session -ErrorAction SilentlyContinue
        }
    }
}

# Output result table
$results | Format-Table -AutoSize

# Optional: export to CSV
# $results | Export-Csv "C:\scripts\nessus-deploy-results.csv" -NoTypeInformation
