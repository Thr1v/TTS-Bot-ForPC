param(
    [int]$CheckInterval = 10,  # Check every 10 seconds instead of 30
    [string]$QueueFile = "C:\bots\notification_queue.txt"  # Absolute path
)

# Function to strip email signatures
function Strip-EmailSignature {
    param([string]$Body)

    if (!$Body) { return "" }

    # Split into lines
    $lines = $Body -split "`r?`n"

    # If email is short, return as-is
    if ($lines.Count -lt 3) {
        return $Body.Trim()
    }

    # Only check the last 10 lines for signatures (signatures are usually at the end)
    $checkLines = $lines.Count
    if ($checkLines -gt 10) {
        $checkLines = 10
    }

    $signatureStartIndex = -1

    # Check last few lines for signature patterns
    for ($i = $lines.Count - $checkLines; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()

        # Skip empty lines
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        # Check for signature indicators
        $isSignature = $false

        # Check for signature dashes (most reliable indicator)
        if ($line -match "^[-—_]{2,}$") {
            $isSignature = $true
        }

        # Check for signature keywords (case insensitive, whole words)
        elseif ($line -match "(?i)\b(best regards|regards|sincerely|thank you|thanks|kind regards|warm regards|yours sincerely|yours truly|cheers|best wishes|with best wishes|cordially|looking forward|sent from|on behalf of)\b") {
            $isSignature = $true
        }

        # Check for contact info patterns (more specific)
        elseif ($line -match "(?i)\b(phone|tel|mobile|fax|email|www\.|\.com|\.org|\.net)\b.*[:=]") {
            $isSignature = $true
        }

        # Check for job titles followed by company info
        elseif ($line -match "(?i)\b(director|manager|coordinator|specialist|analyst|administrator|officer|president|ceo|cto|cfo)\b") {
            # Only consider it a signature if it's in the last 3 lines
            if ($i -ge $lines.Count - 3) {
                $isSignature = $true
            }
        }

        if ($isSignature) {
            $signatureStartIndex = $i
            break
        }
    }

    # If we found a signature, return content before it
    if ($signatureStartIndex -ge 0) {
        $contentLines = $lines[0..($signatureStartIndex - 1)]
    }
    else {
        # No signature found, return all content
        $contentLines = $lines
    }

    # Join the lines back
    $result = $contentLines -join " "

    # Clean up extra whitespace
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

    # Try to connect to Outlook
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)  # 6 = Inbox

        Write-Host "Connected to Outlook successfully"
    }
    catch {
        Write-Error "Failed to connect to Outlook: $_"
        Write-Host "Make sure Outlook is installed and running"
        Write-Host "You may need to run PowerShell as Administrator"
        return
    }

    # Track seen emails
    $seenEmails = @{}

    # Load previously seen emails
    $seenFile = "C:\bots\outlook_seen.txt"
    if (Test-Path $seenFile) {
        try {
            $loaded = Get-Content $seenFile | ConvertFrom-Json
            # Convert PSCustomObject back to hashtable
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
            # Get unread emails
            $unreadItems = $inbox.Items | Where-Object { $_.Unread -eq $true }

            $newEmails = 0
            foreach ($item in $unreadItems) {
                try {
                    $entryId = $item.EntryID

                    # Skip if already seen
                    if ($seenEmails.ContainsKey($entryId)) {
                        continue
                    }

                    # Extract email details
                    $sender = $item.SenderName
                    if (!$sender) { $sender = $item.SenderEmailAddress }
                    if (!$sender) { $sender = "Unknown Sender" }

                    $subject = $item.Subject
                    if (!$subject) { $subject = "No Subject" }

                    # Get body (first 200 chars) and strip signature
                    $body = $item.Body
                    if ($body) {
                        $body = Strip-EmailSignature -Body $body
                        if ($body.Length -gt 200) {
                            $body = $body.Substring(0, 197) + "..."
                        }
                    }

                    # Create notification message
                    $message = "New email from $sender`: $subject"
                    if ($body -and $body.Trim()) {
                        $message += " - $($body.Trim())"
                    }

                    # Add to queue
                    Add-Notification -Message $message -Source "email"

                    # Mark as seen
                    $seenEmails[$entryId] = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"

                    $newEmails++
                    $emailCount++

                }
                catch {
                    Write-Warning "Error processing email: $_"
                }
            }

            if ($newEmails -gt 0) {
                Write-Host "Processed $newEmails new emails (total: $emailCount)"
            }

            # Save seen emails periodically
            if ($emailCount % 5 -eq 0 -and $emailCount -gt 0) {
                try {
                    $seenEmails | ConvertTo-Json | Set-Content $seenFile -Encoding UTF8 -Force
                }
                catch {
                    Write-Warning "Could not save seen emails: $_"
                }
            }

        }
        catch {
            Write-Warning "Error checking emails: $_"
        }

        # Wait before next check
        Start-Sleep -Seconds $CheckInterval
    }
}

# Run the monitor
Start-OutlookMonitor
