# TTS Bot with Auto-Reading Notifications

This enhanced TTS bot can automatically read notifications from log files or email, making it perfect for monitoring systems, alerts, or staying updated on important information.

## Features

- **Text-to-Speech**: Convert text to speech with multiple voice options
- **Auto-Reading**: Automatically read notifications from monitored sources
- **Log File Monitoring**: Watch log files for new entries
- **Email Monitoring**: Check for new emails via PowerShell Outlook integration
- **Audio Playback Controls**: Play, pause, stop, and rewind generated audio
- **Voice Input**: Use speech recognition to input text
- **Smart Signature Stripping**: Automatically removes email signatures and contact info

## Setup

### 1. Install Dependencies

```bash
pip install pyttsx3 pygame gtts speechrecognition
```

### 2. Configure Notification Monitoring

Edit `notification_config.ini` to set up monitoring:

```ini
[general]
notification_queue = notification_queue.txt
check_interval = 30

[log_monitoring]
enabled = true
log_files = system.log,application.log
position_files = system.pos,application.pos

[email_monitoring]
enabled = true
protocol = pop3
server = outlook.office365.com
port = 995
email_address = your.email@outlook.com
password = your_password
check_interval = 300
```

### 3. Email Monitoring Setup

The system uses PowerShell to directly monitor Outlook, bypassing network restrictions.

#### Requirements
- **Outlook must be installed** and running on your system
- **PowerShell execution policy** should allow script running

#### Quick Start
1. **Configure email settings** in `notification_config.ini` (see above)
2. **Run the Outlook monitor**:
   ```powershell
   .\outlook_monitor.ps1
   ```
3. **Start the TTS bot** with auto-reading enabled

#### Manual Email Addition
If automatic monitoring fails, you can manually add emails:

```bash
python add_email_gui.py
```

This opens a GUI where you can enter email details to be spoken by the TTS bot.

### 4. What Gets Read

The system extracts and speaks:
- **Sender**: Name of the email sender
- **Subject**: The email subject line
- **Body**: The complete email content, with signatures and contact information automatically removed

**Smart Signature Detection:**
- Removes content after signature dashes (--)
- Strips common closing phrases ("Best regards", "Sincerely", etc.)
- Removes contact information and job titles
- Preserves the full message content up to the signature

### Testing Email Monitoring

You can test the email monitoring by:

1. Configuring your email settings in `notification_config.ini`
2. Running `.\outlook_monitor.ps1` - it will start monitoring for new emails
3. Sending yourself a test email
4. The TTS bot will automatically read the email (with signature removed)

### Troubleshooting

- **Outlook COM error**: Make sure Outlook is installed and running
- **No emails detected**: Check that Outlook has access to your inbox
- **Signatures not removed**: The detection is heuristic - complex signatures may need manual adjustment
- **PowerShell execution**: Run PowerShell as Administrator if you get execution policy errors
- **Email connection fails**: Verify your email credentials in the config file

## Usage

### Running the TTS Bot

```bash
python tts-bot.py
```

### Running the Email Monitor

```powershell
.\outlook_monitor.ps1
```

### Auto-Reading Setup

1. Start the TTS bot: `python tts-bot.py`
2. Start the email monitor: `.\outlook_monitor.ps1`
3. In the TTS bot, check "Enable Auto-Reading"
4. Set the check interval (how often to look for new notifications)
5. The bot will automatically speak new emails as they arrive

### Manual Email Addition

If you receive emails that the automatic monitor misses, you can add them manually:

```bash
python add_email_gui.py
```

This opens a GUI where you can enter:
- Sender name
- Email subject
- Message content
- Priority level

The TTS bot will immediately speak the manually added email.

### Manual Notification Testing

You can manually add notifications to test the system:

```python
from notification_monitor import NotificationMonitor
monitor = NotificationMonitor()
monitor.add_notification("Test notification", source="manual")
```

## How It Works

1. **Outlook Monitor** (PowerShell) watches your Outlook inbox for unread emails
2. When new emails are found, signatures are automatically stripped
3. Email content is written to `notification_queue.txt` as JSON
4. **TTS Bot** periodically checks the queue file for new notifications
5. Unread notifications are automatically spoken using the selected voice
6. After speaking, notifications are marked as read

## File Structure

```
├── tts-bot.py                 # Main TTS application
├── outlook_monitor.ps1        # PowerShell Outlook monitoring service
├── add_email_gui.py           # Manual email addition GUI
├── notification_config.ini    # Configuration file
├── notification_queue.txt     # Notification queue (auto-created)
├── outlook_seen.txt          # Tracks processed emails (auto-created)
└── README.md                  # This file
```

## Log File Monitoring

To monitor a log file:

1. Add the log file path to `log_files` in the config
2. Add a corresponding position file to `position_files`
3. The monitor will remember where it left off and only read new content

Example:
```
log_files = C:\Logs\system.log,C:\Logs\app.log
position_files = system.pos,app.pos
```

## Troubleshooting

- **No voices available**: The bot falls back to online Google voices
- **Outlook COM error**: Make sure Outlook is installed and running
- **No emails detected**: Check that Outlook has access to your inbox
- **Signatures not removed**: The detection is heuristic - complex signatures may need manual adjustment
- **PowerShell execution**: Run PowerShell as Administrator if you get execution policy errors
- **Email connection fails**: Verify your email credentials in the config file
- **Full email not read**: The system now reads complete emails and only cuts at signatures

## Security Note

Email passwords are stored in plain text in the config file. For production use, consider using environment variables or a secure credential store.
