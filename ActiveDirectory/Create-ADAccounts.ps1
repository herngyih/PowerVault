function New-RandomPassword {
    param (
        [int]$Length = 12
    )

    $upper = [char[]]'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $lower = [char[]]'abcdefghijklmnopqrstuvwxyz'
    $numbers = [char[]]'0123456789'
    $special = [char[]]'!@#$%^&*()-_=+[]{}|;:,.<>?'

    # Ensure at least one from each category
    $mandatory = @(
        $upper | Get-Random
        $lower | Get-Random
        $numbers | Get-Random
        $special | Get-Random
    )

    # Fill the rest with a random mix of all allowed chars
    $allChars = $upper + $lower + $numbers + $special
    $remaining = $Length - $mandatory.Count
    $randomChars = 1..$remaining | ForEach-Object { $allChars | Get-Random }

    # Combine and shuffle
    $finalPassword = ($mandatory + $randomChars) | Get-Random -Count $Length
    return -join $finalPassword
}

function Create-User{
    param(
        [string]$engName, 
        [string] $givenName,
        [string] $surName,
        [string] $combineName,
        [string] $engJobTitle,
        [string] $engJobDept,
        [string] $subCompany,
        [string] $enAccount,
        [string] $mobilePhone,
        [securestring] $password
    )
    
   
    $UPN = $enAccount + "@hyva.com"
    $OU = "OU=Workday,OU=86-General,OU=Departments,OU=Regular Users,OU=Users,OU=Yangzhou,OU=China,OU=REGION 1,OU=GLOBAL,OU=HYVAGROUP,DC=hyvagroup,DC=com"
    
  New-ADUser `
    -Name $engName `
    -GivenName $givenName `
    -Surname $surName `
    -DisplayName $combineName `
    -SamAccountName $enAccount `
    -UserPrincipalName $UPN `
    -MobilePhone $mobilePhone `
    -Title $engJobTitle `
    -Department $engJobDept `
    -Company $subCompany `
    -City "Yangzhou" `
    -State "Jiangsu" `
    -Country "CN" `
    -Path $OU `
    -AccountPassword $password `
    -ChangePasswordAtLogon $true `
    -Enabled $true
}

$report = @()
$users = Import-Excel '.\O365 User List.xlsx' -Sheet 'duplicate'

foreach($user in $users){
    try{
    Write-Host "Creating user: $($user.'EN Name')" -ForegroundColor Cyan  
    $plainPassword = New-RandomPassword
    $securePassword = ConvertTo-SecureString $plainPassword -AsPlainText -Force   

         Create-User $user.'EN Name' `
                    $user.GivenName `
                    $user.Surname `
                    $user.'Combine Name' `
                    $user.'EN JobTitle' `
                    $user.'Dept.' `
                    $user.'Sub Company' `
                    $user.'EN Account' `
                    $user.'Mobile Correct' `
                    $securePassword
        
         $report += [PSCustomObject]@{
            Name = $user.'Combine Name'
            UserName =  $user.'EN Account'
            Password = $plainPassword
            Status = "Pass"
         }

    }
    catch{
         $report += [PSCustomObject]@{
            Name = $user.'Combine Name'
            UserName =  $user.'EN Account'
            Password = "NA"
            Status = "Failed: $($_.Exception.Message)"
         }
    }
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFile = ".\UserCreationReport_$timestamp.xlsx"

# Export the report to Excel
$report | Export-Excel -Path $reportFile -WorksheetName "Results" -AutoSize

Write-Host "Report exported to $reportFile" -ForegroundColor Green