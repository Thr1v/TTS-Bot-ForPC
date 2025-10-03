#!/usr/bin/env python3
# tts_gui.py
# GUI Text-to-Speech tool using pyttsx3 (offline) + pygame for playback controls.

import os
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import time
import json
from pathlib import Path

import pyttsx3
import pygame

try:
    from gtts import gTTS
    GTTS_AVAILABLE = True
except ImportError:
    GTTS_AVAILABLE = False

try:
    import speech_recognition as sr
    SPEECH_REC_AVAILABLE = True
except ImportError:
    SPEECH_REC_AVAILABLE = False


class TTSApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Text-to-Speech (pyttsx3) — Voice & File Player")
        self.master.geometry("720x420")

        # State
        self.voices = []
        self.selected_voice_id = None
        self.selected_file_path = tk.StringVar(value="")
        self.rate_var = tk.IntVar(value=180)    # sensible default
        self.volume_var = tk.DoubleVar(value=0.9)
        self.current_audio_path = None
        self.is_generating = False
        self.is_speaking_live = False
        self.tts_engine = "pyttsx3"  # Default to offline
        self.conversation_history = []  # For conversational mode
        self.stop_speech_flag = False  # Flag to stop current speech

        # Auto-reading state
        self.auto_reading_enabled = tk.BooleanVar(value=False)
        self.auto_reading_interval = tk.IntVar(value=10)  # seconds
        self.notification_queue = "notification_queue.txt"
        self.last_notification_check = 0
        self.auto_reading_thread = None
        self.auto_reading_active = False

        # Init pygame mixer for playback
        pygame.mixer.init()

        # Build UI
        self._build_ui()

        # Load voices (in a thread to avoid UI stall if backend is slow)
        threading.Thread(target=self._load_voices, daemon=True).start()

    # ---------- UI ----------
    def _build_ui(self):
        # For now, just build the TTS interface directly
        self._build_tts_ui(self.master)

    def _build_tts_ui(self, parent):
        pad = {"padx": 10, "pady": 8}

        # Voice selection
        voice_frame = ttk.LabelFrame(parent, text="Voice")
        voice_frame.pack(fill="x", **pad)

        self.voice_combo = ttk.Combobox(voice_frame, state="readonly")
        self.voice_combo.pack(side="left", fill="x", expand=True, padx=10, pady=10)
        self.voice_combo.bind("<<ComboboxSelected>>", self._on_voice_selected)

        self.refresh_btn = ttk.Button(voice_frame, text="Refresh Voices", command=self._refresh_voices)
        self.refresh_btn.pack(side="left", padx=10, pady=10)

        # Rate & Volume
        rv_frame = ttk.LabelFrame(parent, text="Settings")
        rv_frame.pack(fill="x", **pad)

        ttk.Label(rv_frame, text="Rate").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        self.rate_scale = ttk.Scale(rv_frame, from_=100, to=250, variable=self.rate_var, orient="horizontal")
        self.rate_scale.grid(row=0, column=1, sticky="ew", padx=10, pady=6)

        ttk.Label(rv_frame, text="Volume").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        self.volume_scale = ttk.Scale(rv_frame, from_=0.0, to=1.0, variable=self.volume_var, orient="horizontal")
        self.volume_scale.grid(row=1, column=1, sticky="ew", padx=10, pady=6)

        rv_frame.columnconfigure(1, weight=1)

        # Auto-reading controls
        auto_frame = ttk.LabelFrame(parent, text="Auto-Reading")
        auto_frame.pack(fill="x", **pad)

        self.auto_read_check = ttk.Checkbutton(auto_frame, text="Enable Auto-Reading",
                                               variable=self.auto_reading_enabled,
                                               command=self._toggle_auto_reading)
        self.auto_read_check.pack(side="left", padx=10, pady=10)

        ttk.Label(auto_frame, text="Check every").pack(side="left", padx=(10,5), pady=10)
        self.interval_spin = ttk.Spinbox(auto_frame, from_=5, to=300, textvariable=self.auto_reading_interval,
                                        width=5, command=self._update_auto_reading_interval)
        self.interval_spin.pack(side="left", padx=(0,5), pady=10)
        ttk.Label(auto_frame, text="seconds").pack(side="left", padx=(0,10), pady=10)

        self.clear_notifications_btn = ttk.Button(auto_frame, text="Clear Notifications",
                                                 command=self._clear_notifications)
        self.clear_notifications_btn.pack(side="right", padx=10, pady=10)

        self.play_last_btn = ttk.Button(auto_frame, text="Play Last Notification",
                                       command=self._play_last_notification)
        self.play_last_btn.pack(side="right", padx=(0, 10), pady=10)

        # File selection
        file_frame = ttk.LabelFrame(parent, text="Text Input")
        file_frame.pack(fill="x", **pad)

        # Text input area
        self.text_area = tk.Text(file_frame, height=4, wrap=tk.WORD)
        self.text_area.pack(fill="both", expand=True, padx=10, pady=(10,5))
        
        # File selection buttons
        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill="x", padx=10, pady=(0,10))
        
        self.file_entry = ttk.Entry(file_btn_frame, textvariable=self.selected_file_path)
        self.file_entry.pack(side="left", fill="x", expand=True)

        self.browse_btn = ttk.Button(file_btn_frame, text="Load File", command=self._browse_file)
        self.browse_btn.pack(side="left", padx=(5,0))
        
        if SPEECH_REC_AVAILABLE:
            self.voice_input_btn = ttk.Button(file_btn_frame, text="🎤 Voice Input", command=self._voice_input)
            self.voice_input_btn.pack(side="left", padx=(5,0))

        # Actions: Live Speak & Generate
        action_frame = ttk.Frame(parent)
        action_frame.pack(fill="x", **pad)

        self.speak_btn = ttk.Button(action_frame, text="Speak Live (no file saved)", command=self._speak_live)
        self.speak_btn.pack(side="left", padx=10, pady=10)

        self.generate_btn = ttk.Button(action_frame, text="Generate Audio", command=self._generate_audio)
        self.generate_btn.pack(side="left", padx=10, pady=10)

        self.skip_btn = ttk.Button(action_frame, text="Skip/Stop Speech", command=self._skip_speech)
        self.skip_btn.pack(side="left", padx=10, pady=10)

        self.status_var = tk.StringVar(value="Ready.")
        self.status_lbl = ttk.Label(parent, textvariable=self.status_var, foreground="#555")
        self.status_lbl.pack(fill="x", padx=12)

        # Playback controls
        playback = ttk.LabelFrame(parent, text="Playback Controls")
        playback.pack(fill="x", **pad)

        self.play_btn = ttk.Button(playback, text="Play", command=self._play_audio, state="disabled")
        self.pause_btn = ttk.Button(playback, text="Pause", command=self._pause_audio, state="disabled")
        self.unpause_btn = ttk.Button(playback, text="Unpause", command=self._unpause_audio, state="disabled")
        self.stop_btn = ttk.Button(playback, text="Stop", command=self._stop_audio, state="disabled")
        self.rewind_btn = ttk.Button(playback, text="Rewind", command=self._rewind_audio, state="disabled")

        for w in (self.play_btn, self.pause_btn, self.unpause_btn, self.stop_btn, self.rewind_btn):
            w.pack(side="left", padx=8, pady=10)

        # Footer note
        note = ttk.Label(parent, text="Tip: Use \"Generate Audio\" to create a file you can replay with the controls.\n"
                                       "Speak Live uses the voice directly without saving a file.\n"
                                       "Use \"Skip/Stop Speech\" to interrupt current playback (including auto-reading).",
                         foreground="#666", justify="left")
        note.pack(fill="x", padx=12, pady=(0, 10))

    def _skip_speech(self):
        """Stop current speech playback"""
        self.stop_speech_flag = True
        self.is_speaking_live = False
        
        # Stop pygame mixer if playing
        try:
            if pygame.mixer.music.get_busy():
                pygame.mixer.music.stop()
        except:
            pass
            
        self._set_status("Speech stopped by user")

    def _voice_input(self):
        """Capture speech input and add to text area"""
        if not SPEECH_REC_AVAILABLE:
            messagebox.showerror("Feature Unavailable", "Speech recognition not installed.\nInstall with: pip install SpeechRecognition")
            return
        
        def worker():
            try:
                self._set_status("Listening... Speak now!")
                self.voice_input_btn.config(state="disabled")
                
                recognizer = sr.Recognizer()
                with sr.Microphone() as source:
                    # Adjust for ambient noise
                    recognizer.adjust_for_ambient_noise(source, duration=0.5)
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)
                
                # Recognize speech
                text = recognizer.recognize_google(audio)
                
                # Add to text area
                current_text = self.text_area.get("1.0", tk.END).strip()
                if current_text:
                    new_text = current_text + "\n\n" + text
                else:
                    new_text = text
                
                self.text_area.delete("1.0", tk.END)
                self.text_area.insert("1.0", new_text)
                
                self._set_status(f"Voice input: '{text[:50]}...'")
                
            except sr.WaitTimeoutError:
                self._set_status("Voice input timed out")
            except sr.UnknownValueError:
                self._set_status("Could not understand audio")
                messagebox.showwarning("Speech Recognition", "Could not understand the audio. Please try again.")
            except sr.RequestError as e:
                self._set_status(f"Speech recognition error: {e}")
                messagebox.showerror("Speech Recognition Error", f"Could not request results: {e}")
            except Exception as e:
                self._set_status(f"Voice input error: {e}")
                messagebox.showerror("Voice Input Error", str(e))
            finally:
                self.voice_input_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    def _build_chat_ui(self, parent):
        """Build the conversational AI chat interface"""
        # Chat interface not yet implemented
        pass

        # Footer note
        note = ttk.Label(self.master, text="Tip: Use “Generate Audio” to create a WAV you can replay with the controls.\n"
                                           "“Speak Live” uses the voice directly without saving a file.",
                         foreground="#666", justify="left")
        note.pack(fill="x", padx=12, pady=(0, 10))

        # Handle window close
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)

    def _load_voices(self, max_retries=3):
        """Load both online and offline voices"""
        self._set_status("Loading voices…")
        
        voices_loaded = False
        self.voices = []
        
        # Load online voices first (Google TTS)
        if GTTS_AVAILABLE:
            online_voices = [
                {"name": "Google English (US)", "id": "en", "type": "online"},
                {"name": "Google English (UK)", "id": "en-uk", "type": "online"},
                {"name": "Google English (AU)", "id": "en-au", "type": "online"},
                {"name": "Google Spanish", "id": "es", "type": "online"},
                {"name": "Google French", "id": "fr", "type": "online"},
                {"name": "Google German", "id": "de", "type": "online"},
                {"name": "Google Italian", "id": "it", "type": "online"},
                {"name": "Google Portuguese", "id": "pt", "type": "online"},
                {"name": "Google Japanese", "id": "ja", "type": "online"},
                {"name": "Google Korean", "id": "ko", "type": "online"},
                {"name": "Google Chinese", "id": "zh", "type": "online"}
            ]
            
            self.voices.extend(online_voices)
            voices_loaded = True
            self._set_status(f"Loaded {len(online_voices)} online voices.")
        
        # Always try to add system voices (pyttsx3) as well
        for attempt in range(max_retries):
            try:
                # Try pyttsx3 with careful error handling
                try:
                    import platform
                    if platform.system() == "Windows":
                        try:
                            engine = pyttsx3.init('sapi5')
                        except Exception:
                            engine = pyttsx3.init()  # Fall back to auto-detect
                    else:
                        engine = pyttsx3.init()
                    
                    voices = engine.getProperty("voices")
                    
                    if voices and len(voices) > 0:
                        offline_count = 0
                        for v in voices:
                            try:
                                # Handle problematic voice objects more carefully
                                if hasattr(v, 'name') and v.name:
                                    name = v.name
                                elif hasattr(v, 'id') and v.id:
                                    name = v.id
                                else:
                                    name = f"Voice {len(self.voices)}"
                                
                                # Safely get the ID
                                voice_id = getattr(v, 'id', name)
                                
                                self.voices.append({"name": name, "id": voice_id, "type": "offline"})
                                offline_count += 1
                            except Exception:
                                continue
                        
                        if offline_count > 0:
                            voices_loaded = True
                            engine.stop()
                            self._set_status(f"Loaded {len(self.voices)} total voices ({len(online_voices) if 'online_voices' in locals() else 0} online, {offline_count} system).")
                            break
                
                except Exception as e:
                    if attempt < max_retries - 1:
                        import time
                        time.sleep(1)
                        continue
                    else:
                        # If pyttsx3 fails but we have online voices, that's still good
                        if self.voices:
                            voices_loaded = True
                            self._set_status(f"Loaded {len(self.voices)} voices (online only).")
                            break
                        else:
                            # Last resort: create dummy voices
                            self.voices = [
                                {"name": "Default Voice", "id": "default", "type": "offline"}
                            ]
                            voices_loaded = True
                            self._set_status("Using fallback voice (TTS engines unavailable).")
                            break
            
            except Exception:
                if attempt == max_retries - 1:
                    # If we have online voices, don't override with dummy
                    if not self.voices:
                        self.voices = [
                            {"name": "Default Voice", "id": "default", "type": "offline"}
                        ]
                        voices_loaded = True
                        self._set_status("Using fallback voice (TTS engines unavailable).")
        
        if voices_loaded and self.voices:
            # Update UI on main thread
            names = [f"{i}: {v['name']}" for i, v in enumerate(self.voices)]
            self.master.after(0, lambda: self._populate_voice_combo(names))
            return
        else:
            error_msg = "Failed to load any voices"
            self._set_status(error_msg)
            self.master.after(0, lambda: messagebox.showerror("Voice Loading Error", 
                f"Could not load voices.\n\n{error_msg}\n\n"
                "Try:\n• Restarting the application\n• Checking system TTS settings\n"
                "• Installing additional TTS engines\n• Running: pip install pyttsx3"))

    def _populate_voice_combo(self, names):
        self.voice_combo["values"] = names
        if names:
            self.voice_combo.current(0)
            self.selected_voice_id = self.voices[0]["id"]

    def _on_voice_selected(self, _event=None):
        idx = self.voice_combo.current()
        if 0 <= idx < len(self.voices):
            voice = self.voices[idx]
            self.selected_voice_id = voice["id"]
            self.tts_engine = voice.get("type", "offline")

    def _refresh_voices(self):
        threading.Thread(target=self._load_voices, daemon=True).start()

    # ---------- File handling ----------
    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Choose a plain text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            self.selected_file_path.set(path)

    def _read_text(self):
        # First check if there's text in the text area
        text_content = self.text_area.get("1.0", tk.END).strip()
        if text_content:
            return text_content
        
        # Otherwise, read from file
        path = self.selected_file_path.get().strip()
        if not path:
            raise ValueError("Please enter text or choose a text file first.")
        if not os.path.exists(path):
            raise FileNotFoundError(f"File not found: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return f.read()

    # ---------- TTS core ----------
    def _configure_engine(self, engine: pyttsx3.Engine):
        """Configure pyttsx3 engine for offline voices"""
        # Voice
        if self.selected_voice_id and self.tts_engine == "offline":
            try:
                engine.setProperty("voice", self.selected_voice_id)
            except Exception:
                pass
        # Rate & Volume (only for offline)
        try:
            engine.setProperty("rate", int(self.rate_var.get()))
        except Exception:
            pass
        try:
            engine.setProperty("volume", float(self.volume_var.get()))
        except Exception:
            pass

    def _generate_online_audio(self, text: str, output_path: str):
        """Generate audio using Google TTS"""
        if not GTTS_AVAILABLE:
            raise Exception("Google TTS not available. Install with: pip install gtts")
        
        try:
            tts = gTTS(text=text, lang=self.selected_voice_id, slow=False)
            tts.save(output_path)
        except Exception as e:
            raise Exception(f"Google TTS failed: {e}")

    def _speak_live(self):
        if self.is_speaking_live or self.is_generating:
            return
        try:
            text = self._read_text()
        except Exception as e:
            messagebox.showerror("File error", str(e))
            return

        def worker():
            self.is_speaking_live = True
            self._set_status("Speaking live…")
            try:
                if self.tts_engine == "online":
                    # For online voices, generate temp audio and play it
                    import tempfile
                    with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
                        temp_path = tmp.name
                    
                    self._generate_online_audio(text, temp_path)
                    pygame.mixer.music.load(temp_path)
                    pygame.mixer.music.play()
                    
                    # Wait for playback to finish
                    while pygame.mixer.music.get_busy():
                        import time
                        time.sleep(0.1)
                    
                    # Cleanup
                    try:
                        os.unlink(temp_path)
                    except:
                        pass
                    
                    self._set_status("Done speaking.")
                else:
                    # Offline speaking
                    engine = pyttsx3.init()
                    self._configure_engine(engine)
                    engine.say(text)
                    engine.runAndWait()
                    engine.stop()
                    self._set_status("Done speaking.")
                    
            except Exception as e:
                self._set_status(f"Error speaking: {e}")
                messagebox.showerror("TTS Error", f"Could not speak.\n\n{e}")
            finally:
                self.is_speaking_live = False

        threading.Thread(target=worker, daemon=True).start()

    def _generate_audio(self):
        if self.is_generating or self.is_speaking_live:
            return
        try:
            text = self._read_text()
        except Exception as e:
            messagebox.showerror("File error", str(e))
            return

        def worker():
            self.is_generating = True
            self._set_status("Generating audio…")
            
            # Determine file extension based on TTS engine
            ext = "mp3" if self.tts_engine == "online" else "wav"
            tmpdir = tempfile.gettempdir()
            out_path = os.path.join(tmpdir, f"tts_output.{ext}")
            
            try:
                if self.tts_engine == "online":
                    self._generate_online_audio(text, out_path)
                else:
                    # Offline generation
                    engine = pyttsx3.init()
                    self._configure_engine(engine)
                    engine.save_to_file(text, out_path)
                    engine.runAndWait()
                    engine.stop()
                
                self.current_audio_path = out_path
                # Enable playback controls
                self._enable_playback_controls(True)
                self._set_status(f"Audio generated: {out_path}")
                
            except Exception as e:
                self._set_status(f"Error generating audio: {e}")
                messagebox.showerror("TTS Error", f"Could not generate audio.\n\n{e}")
            finally:
                self.is_generating = False

        threading.Thread(target=worker, daemon=True).start()

    # ---------- Playback via pygame ----------
    def _ensure_loaded(self):
        if not self.current_audio_path or not os.path.exists(self.current_audio_path):
            messagebox.showwarning("No audio", "Generate audio first to enable playback.")
            return False
        return True

    def _play_audio(self):
        if not self._ensure_loaded():
            return
        try:
            pygame.mixer.music.load(self.current_audio_path)
            pygame.mixer.music.play()
            self._set_status("Playing audio…")
        except Exception as e:
            self._set_status(f"Playback error: {e}")
            messagebox.showerror("Playback Error", str(e))

    def _pause_audio(self):
        try:
            pygame.mixer.music.pause()
            self._set_status("Paused.")
        except Exception as e:
            self._set_status(f"Pause error: {e}")

    def _unpause_audio(self):
        try:
            pygame.mixer.music.unpause()
            self._set_status("Playing…")
        except Exception as e:
            self._set_status(f"Unpause error: {e}")

    def _stop_audio(self):
        try:
            pygame.mixer.music.stop()
            self._set_status("Stopped.")
        except Exception as e:
            self._set_status(f"Stop error: {e}")

    def _rewind_audio(self):
        try:
            pygame.mixer.music.rewind()  # jumps to start
            self._set_status("Rewound to start.")
        except Exception as e:
            self._set_status(f"Rewind error: {e}")

    def _enable_playback_controls(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for btn in (self.play_btn, self.pause_btn, self.unpause_btn, self.stop_btn, self.rewind_btn):
            btn.config(state=state)

    # ---------- Auto-Reading Methods ----------
    def _toggle_auto_reading(self):
        """Enable or disable auto-reading of notifications"""
        if self.auto_reading_enabled.get():
            self._start_auto_reading()
        else:
            self._stop_auto_reading()

    def _update_auto_reading_interval(self):
        """Update the auto-reading check interval"""
        if self.auto_reading_active:
            self._stop_auto_reading()
            self._start_auto_reading()

    def _start_auto_reading(self):
        """Start the auto-reading thread"""
        if self.auto_reading_active:
            return

        self.auto_reading_active = True
        self.auto_reading_thread = threading.Thread(target=self._auto_reading_worker, daemon=True)
        self.auto_reading_thread.start()
        self._set_status(f"Auto-reading enabled (checking every {self.auto_reading_interval.get()}s)")

    def _stop_auto_reading(self):
        """Stop the auto-reading thread"""
        self.auto_reading_active = False
        if self.auto_reading_thread:
            self.auto_reading_thread.join(timeout=2)
        self._set_status("Auto-reading disabled")

    def _auto_reading_worker(self):
        """Background worker for auto-reading notifications"""
        while self.auto_reading_active:
            try:
                self._check_notifications()
            except Exception as e:
                print(f"Auto-reading error: {e}")
            time.sleep(self.auto_reading_interval.get())

    def _check_notifications(self):
        """Check for new notifications and speak them"""
        if not os.path.exists(self.notification_queue):
            return

        try:
            notifications = []
            # Try different encodings in case the file has encoding issues
            encodings_to_try = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'latin-1']
            
            content = None
            for encoding in encodings_to_try:
                try:
                    with open(self.notification_queue, 'r', encoding=encoding) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                print(f"Could not decode notification file with any known encoding")
                return
            
            for line in content.splitlines():
                line = line.strip()
                if line:
                    try:
                        notification = json.loads(line)
                        if not notification.get('spoken', False):
                            notifications.append(notification)
                    except json.JSONDecodeError:
                        continue

            # Speak new notifications
            for notification in notifications:
                if not self.auto_reading_active or self.stop_speech_flag:  # Check if we should stop
                    break

                message = notification['message']
                source = notification.get('source', 'unknown')
                priority = notification.get('priority', 'normal')

                # Add source prefix for context
                if source.startswith('log:'):
                    prefix = f"Log update: "
                elif source == 'email':
                    prefix = f"Email: "
                else:
                    prefix = f"Notification: "

                full_message = f"{prefix}{message}"

                # Speak the notification
                self._speak_text_live(full_message)

                # Check if speech was interrupted
                if self.stop_speech_flag:
                    break

                # Mark as spoken
                notification['spoken'] = True
                self._update_notification_status(notification)

                # Brief pause between notifications
                if self.auto_reading_active and not self.stop_speech_flag:
                    time.sleep(2)

        except Exception as e:
            print(f"Error checking notifications: {e}")

    def _update_notification_status(self, notification):
        """Update a notification's spoken status in the queue file"""
        try:
            # Try different encodings to read the file
            encodings_to_try = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'latin-1']
            
            content = None
            encoding_used = None
            for encoding in encodings_to_try:
                try:
                    with open(self.notification_queue, 'r', encoding=encoding) as f:
                        content = f.read()
                    encoding_used = encoding
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                print(f"Could not decode notification file with any known encoding")
                return
            
            # Read all notifications
            notifications = []
            for line in content.splitlines():
                line = line.strip()
                if line:
                    try:
                        notif = json.loads(line)
                        # Update the matching notification
                        if (notif.get('timestamp') == notification.get('timestamp') and
                            notif.get('message') == notification.get('message')):
                            notif['spoken'] = True
                        notifications.append(notif)
                    except json.JSONDecodeError:
                        continue

            # Write back all notifications with UTF-8 encoding
            with open(self.notification_queue, 'w', encoding='utf-8') as f:
                for notif in notifications:
                    f.write(json.dumps(notif) + '\n')

        except Exception as e:
            print(f"Error updating notification status: {e}")

    def _clear_notifications(self):
        """Clear all notifications from the queue"""
        try:
            if os.path.exists(self.notification_queue):
                # Create backup
                backup_name = f"{self.notification_queue}.backup"
                if os.path.exists(backup_name):
                    os.remove(backup_name)
                os.rename(self.notification_queue, backup_name)

            # Create empty queue
            Path(self.notification_queue).touch()
            self._set_status("Notifications cleared (backup created)")

        except Exception as e:
            self._set_status(f"Error clearing notifications: {e}")

    def _speak_text_live(self, text):
        """Speak text directly without saving to file"""
        if not text.strip():
            return

        try:
            self.is_speaking_live = True
            self.stop_speech_flag = False  # Reset stop flag
            self._set_status("Speaking...")

            # Find the selected voice
            selected_voice = None
            if self.selected_voice_id:
                for voice in self.voices:
                    if voice['id'] == self.selected_voice_id:
                        selected_voice = voice
                        break

            if selected_voice and selected_voice['type'] == 'online':
                # Use Google TTS
                tts = gTTS(text=text, lang=selected_voice['id'], slow=False)
                with tempfile.NamedTemporaryFile(delete=False, suffix='.mp3') as temp_file:
                    temp_file_path = temp_file.name
                    tts.save(temp_file_path)

                # Play the audio
                pygame.mixer.music.load(temp_file_path)
                pygame.mixer.music.play()

                # Wait for playback to finish
                while pygame.mixer.music.get_busy() and self.is_speaking_live and not self.stop_speech_flag:
                    pygame.time.wait(100)

                # Stop if interrupted
                if self.stop_speech_flag:
                    pygame.mixer.music.stop()

                # Clean up
                try:
                    os.unlink(temp_file_path)
                except:
                    pass

            else:
                # Use pyttsx3 (fallback) - check for stop flag
                if not self.stop_speech_flag:
                    engine = pyttsx3.init()
                    if selected_voice:
                        try:
                            engine.setProperty('voice', selected_voice['id'])
                        except:
                            pass

                    engine.setProperty('rate', self.rate_var.get())
                    engine.setProperty('volume', self.volume_var.get())
                    engine.say(text)
                    engine.runAndWait()

            if self.stop_speech_flag:
                self._set_status("Speech interrupted")
            else:
                self._set_status("Speech completed.")

        except Exception as e:
            self._set_status(f"Speech error: {e}")
        finally:
            self.is_speaking_live = False
            self.stop_speech_flag = False

    # ---------- Helpers ----------
    def _set_status(self, msg: str):
        self.status_var.set(msg)

    def _on_close(self):
        try:
            pygame.mixer.music.stop()
            pygame.mixer.quit()
        except Exception:
            pass
        self.master.destroy()

    def _get_last_notification(self):
        """Get the most recent notification from the queue"""
        if not os.path.exists(self.notification_queue):
            return None

        try:
            # Try different encodings to read the file
            encodings_to_try = ['utf-8', 'utf-16', 'utf-16-le', 'utf-16-be', 'latin-1']
            
            content = None
            for encoding in encodings_to_try:
                try:
                    with open(self.notification_queue, 'r', encoding=encoding) as f:
                        content = f.read()
                    break
                except UnicodeDecodeError:
                    continue
            
            if content is None:
                print(f"Could not decode notification file with any known encoding")
                return None
            
            notifications = []
            for line in content.splitlines():
                line = line.strip()
                if line:
                    try:
                        notification = json.loads(line)
                        notifications.append(notification)
                    except json.JSONDecodeError:
                        continue

            if notifications:
                # Return the most recent notification (last in file)
                return notifications[-1]
            return None

        except Exception as e:
            print(f"Error reading last notification: {e}")
            return None

    def _play_last_notification(self):
        """Play the most recent notification"""
        last_notif = self._get_last_notification()
        if not last_notif:
            self._set_status("No notifications found")
            messagebox.showinfo("No Notifications", "There are no notifications in the queue.")
            return

        message = last_notif['message']
        source = last_notif.get('source', 'unknown')

        # Add source prefix for context
        if source.startswith('log:'):
            prefix = "Log update: "
        elif source == 'email':
            prefix = "Email: "
        else:
            prefix = "Notification: "

        full_message = f"{prefix}{message}"

        # Speak the notification in a separate thread to avoid blocking UI
        self._set_status(f"Speaking last notification: {message[:50]}...")
        threading.Thread(target=self._speak_text_live, args=(full_message,), daemon=True).start()


def main():
    root = tk.Tk()
    # Use a nicer theme if available
    try:
        style = ttk.Style()
        if "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    app = TTSApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
