# Path to your IIS SMTP log file
$logFile = "C:\Windows\System32\LogFiles\SMTPSVC1\ex250825.log"

# Path to export CSV for user adm_herngyih
$outputCsv = "C:\Users\adm_herngyih\Desktop\SMTP_ParsedLogs_Full.csv"

# Read the log file
$logEntries = Get-Content $logFile

# Parse and objectify relevant log lines
$parsedLogs = $logEntries | ForEach-Object {
    $line = $_
    $parts = $line -split "\s+"

    # Ensure the line has at least the fixed columns (date, time, sourceIP, eventType, server)
    if ($parts.Count -lt 5) { return }

    # Extract fixed fields
    $timestamp = $parts[0] + " " + $parts[1]
    $sourceIP = $parts[2]
    $eventType = $parts[3]
    $server = $parts[4]

    # Capture the rest of the line as Message
    $message = ($parts[5..($parts.Count - 1)] -join " ")

    # Check for authentication failure
    $authFailed = $false
    if ($message -match "535\+5\.7\.139") {
        $authFailed = $true
    }

    # Create custom object
    [PSCustomObject]@{
        Timestamp  = $timestamp
        SourceIP   = $sourceIP
        EventType  = $eventType
        Server     = $server
        Message    = $message
        AuthFailed = $authFailed
    }
}

# Display as table
$parsedLogs | Format-Table -AutoSize

# Export to CSV
$parsedLogs | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. CSV saved at $outputCsv"
