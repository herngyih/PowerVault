function Get-AdUserByEmployeeID {
    param (
        [Parameter(Mandatory)]
        [string]$EmployeeID
    )

    try {
        # Use single quotes around the value to ensure proper AD filter parsing
        $userResult = Get-ADUser -Filter "employeeID -eq '$EmployeeID'" -Properties employeeID, UserPrincipalName
        return $userResult
    }
    catch {
        Write-Warning "Error fetching user with employeeID: $EmployeeID — $($_.Exception.Message)"
        return $null
    }
}

$objReport = @()

$users = Get-Content .\users.txt | Where-Object { $_.Trim() -ne '' }

for ($i = 0; $i -lt $users.Count; $i++) {
    $currentID = $users[$i].Trim()
    Write-Host "Fetching user $($i + 1) of $($users.Count): $currentID"

    try {
        $fetchedResult = Get-AdUserByEmployeeID -EmployeeID $currentID

        if ($fetchedResult) {
            $objReport += [PSCustomObject]@{
                Name        = $fetchedResult.Name
                EmployeeID  = $fetchedResult.EmployeeID
                Email       = $fetchedResult.UserPrincipalName
                Status      = 'Found'
            }
        }
        else {
            $objReport += [PSCustomObject]@{
                Name        = 'NotFound'
                EmployeeID  = $currentID
                Email       = 'NotFound'
                Status      = 'No Match'
            }
        }
    }
    catch {
        $objReport += [PSCustomObject]@{
            Name        = 'BadInput'
            EmployeeID  = $currentID
            Email       = 'BadInput'
            Status      = 'Error'
        }
    }
}

# Display summary table
$objReport | Format-Table -AutoSize

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$objReport | Export-Csv .\UserLookupReport_$timestamp.csv -NoTypeInformation
