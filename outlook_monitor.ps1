param(
    [int]$CheckInterval = 5,
    [string]$QueueFile = "C:\bots\notification_queue.txt"
)

# Function to strip hyperlinks/URLs from email body
function Strip-Urls {
    param([string]$Body)

    if (!$Body) { return "" }

    # Remove common URL patterns
    $urlPatterns = @(
        'https?://[^\s]+',
        'www\.[^\s]+',
        '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b',
        '[A-Za-z]:\\[^\s]*',
        '\\\\[^\s]+'
    )

    $result = $Body

    foreach ($pattern in $urlPatterns) {
        $result = $result -replace $pattern, '[link removed]'
    }

    return $result
}

# Function to strip email signatures
function Strip-EmailSignature {
    param([string]$Body)

    if (!$Body) { return "" }

    $lines = $Body -split "`r?`n"

    if ($lines.Count -lt 3) {
        return $Body.Trim()
    }

    $checkLines = [Math]::Min($lines.Count, 10)
    $signatureStartIndex = -1

    for ($i = $lines.Count - $checkLines; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()

        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $isSignature = $false

        if ($line -match "^[-_]{2,}$") {
            $isSignature = $true
        }
        elseif ($line -match "(?i)\b(best regards|regards|sincerely|thank you|thanks|kind regards|warm regards|yours sincerely|yours truly|cheers|best wishes|with best wishes|cordially|looking forward|sent from|on behalf of)\b") {
            $isSignature = $true
        }
        elseif ($line -match "(?i)\b(phone|tel|mobile|fax|email|www\.|\.com|\.org|\.net)\b.*[:=]") {
            $isSignature = $true
        }
        elseif ($line -match "(?i)\b(director|manager|coordinator|specialist|analyst|administrator|officer|president|ceo|cto|cfo)\b") {
            if ($i -ge $lines.Count - 3) {
                $isSignature = $true
            }
        }

        if ($isSignature) {
            $signatureStartIndex = $i
            break
        }
    }

    if ($signatureStartIndex -ge 0) {
        $contentLines = $lines[0..($signatureStartIndex - 1)]
    }
    else {
        $contentLines = $lines
    }

    $result = $contentLines -join " "
    $result = $result -replace "\s+", " "
    $result = $result.Trim()

    return $result
}

# Function to add notification to queue
function Add-Notification {
    param([string]$Message, [string]$Source = "email")

    $timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss.fff"
    $notification = @{
        timestamp = $timestamp
        message = $Message
        source = $Source
        priority = "normal"
        spoken = $false
    }

    $json = $notification | ConvertTo-Json -Compress
    Add-Content -Path $QueueFile -Value $json -Encoding UTF8
    Write-Host "Added notification: $($Message.Substring(0, [Math]::Min(50, $Message.Length)))..."
}

# Function to monitor Outlook
function Start-OutlookMonitor {
    Write-Host "Starting Outlook email monitor..."
    Write-Host "Queue file: $QueueFile"
    Write-Host "Check interval: $CheckInterval seconds"
    Write-Host "Press Ctrl+C to stop"

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)

        Write-Host "Connected to Outlook successfully"
    }
    catch {
        Write-Error "Failed to connect to Outlook: $_"
        Write-Host "Make sure Outlook is installed and running"
        Write-Host "You may need to run PowerShell as Administrator"
        return
    }

    $seenEmails = @{}
    $seenFile = "C:\bots\outlook_seen.txt"
    
    if (Test-Path $seenFile) {
        try {
            $loaded = Get-Content $seenFile | ConvertFrom-Json
            if ($loaded) {
                foreach ($property in $loaded.PSObject.Properties) {
                    $seenEmails[$property.Name] = $property.Value
                }
            }
            Write-Host "Loaded $($seenEmails.Count) previously seen emails"
        }
        catch {
            $seenEmails = @{}
            Write-Host "Could not load seen emails file, starting fresh"
        }
    }

    $emailCount = 0

    while ($true) {
        try {
            $mostRecentUnread = $null
            $maxReceivedTime = [DateTime]::MinValue

            foreach ($item in $inbox.Items) {
                if ($item.Unread -and $item.ReceivedTime -gt $maxReceivedTime) {
                    $mostRecentUnread = $item
                    $maxReceivedTime = $item.ReceivedTime
                }
            }

            $newEmails = 0

            if ($mostRecentUnread) {
                $item = $mostRecentUnread

                try {
                    $entryId = $item.EntryID
                    
                    if ([string]::IsNullOrEmpty($entryId)) {
                        $sender = $item.SenderName
                        if (!$sender) { $sender = $item.SenderEmailAddress }
                        if (!$sender) { $sender = "Unknown" }
                        
                        $subject = $item.Subject
                        if (!$subject) { $subject = "No Subject" }
                        
                        $entryId = "$sender|$subject|$($item.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss'))"
                        Write-Warning "Using composite key for email without EntryID: $entryId"
                    }

                    if ($seenEmails.ContainsKey($entryId)) {
                        continue
                    }

                    $sender = $item.SenderName
                    if (!$sender) { $sender = $item.SenderEmailAddress }
                    if (!$sender) { $sender = "Unknown Sender" }

                    $subject = $item.Subject
                    if (!$subject) { $subject = "No Subject" }

                    $body = $item.Body
                    if ($body) {
                        $body = Strip-Urls -Body $body
                        $body = Strip-EmailSignature -Body $body
                    }

                    $message = "New email from $sender`: $subject"
                    if ($body -and $body.Trim()) {
                        $message += " - $($body.Trim())"
                    }

                    Add-Notification -Message $message -Source "email"

                    $seenEmails[$entryId] = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"

                    $newEmails++
                    $emailCount++

                    Write-Host "Processed most recent unread email from $sender"

                }
                catch {
                    Write-Warning "Error processing email: $_"
                }
            }

            if ($newEmails -gt 0) {
                Write-Host "Processed $newEmails new emails (total: $emailCount)"
            }

            if (($emailCount % 20 -eq 0 -and $emailCount -gt 0) -or ($seenEmails.Count -gt 0 -and $seenEmails.Count % 50 -eq 0)) {
                try {
                    $seenEmails | ConvertTo-Json -Compress | Set-Content $seenFile -Encoding UTF8 -Force
                    Write-Host "Saved $($seenEmails.Count) seen emails to disk"
                }
                catch {
                    Write-Warning "Could not save seen emails: $_"
                }
            }

        }
        catch {
            Write-Warning "Error checking emails: $_"
        }

        Start-Sleep -Seconds $CheckInterval
    }
}

# Run the monitor
Start-OutlookMonitor
