# TTS Bot with Auto-Reading Notifications

This enhanced TTS bot can automatically read notifications from log files or email, making it perfect for monitoring systems, alerts, or staying updated on important information.

## Features

- **Text-to-Speech**: Convert text to speech with multiple voice options
- **Auto-Reading**: Automatically read notifications from monitored sources
- **Log File Monitoring**: Watch log files for new entries
- **Email Monitoring**: Check for new emails and read them aloud
- **Audio Playback Controls**: Play, pause, stop, and rewind generated audio
- **Voice Input**: Use speech recognition to input text

## Setup

### 1. Install Dependencies

```bash
pip install pyttsx3 pygame gtts speechrecognition imapclient
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
enabled = false
imap_server = imap.gmail.com
email_address = your.email@gmail.com
password = your_app_password
check_interval = 300
last_check_file = email_last_check.txt
```

### 3. For Email Monitoring

1. Enable 2-factor authentication on your email account
2. Generate an app password (for Gmail: Settings > Security > App passwords)
3. Use the app password in the config (not your regular password)
4. Set `enabled = true` in the `[email_monitoring]` section

## Usage

### Running the TTS Bot

```bash
python tts-bot.py
```

### Running the Notification Monitor

```bash
python notification_monitor.py
```

### Auto-Reading Setup

1. Start both the TTS bot and notification monitor
2. In the TTS bot, check "Enable Auto-Reading"
3. Set the check interval (how often to look for new notifications)
4. The bot will automatically speak new notifications as they arrive

### Manual Notification Testing

You can manually add notifications to test the system:

```python
from notification_monitor import NotificationMonitor
monitor = NotificationMonitor()
monitor.add_notification("Test notification", source="manual")
```

## How It Works

1. **Notification Monitor** watches log files and emails for new content
2. When new content is found, it's written to `notification_queue.txt` as JSON
3. **TTS Bot** periodically checks the queue file for new notifications
4. Unread notifications are automatically spoken using the selected voice
5. After speaking, notifications are marked as read

## File Structure

```
├── tts-bot.py                 # Main TTS application
├── notification_monitor.py    # Notification monitoring service
├── notification_config.ini    # Configuration file
├── notification_queue.txt     # Notification queue (auto-created)
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
- **Email connection fails**: Check your app password and IMAP settings
- **Log files not monitored**: Ensure file paths are correct and accessible
- **Auto-reading not working**: Make sure both services are running

## Security Note

Email passwords are stored in plain text in the config file. For production use, consider using environment variables or a secure credential store.
