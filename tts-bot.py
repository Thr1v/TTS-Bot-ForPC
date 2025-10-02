#!/usr/bin/env python3
# tts_gui.py
# GUI Text-to-Speech tool using pyttsx3 (offline) + pygame for playback controls.

import os
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

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
                                       "Speak Live uses the voice directly without saving a file.",
                         foreground="#666", justify="left")
        note.pack(fill="x", padx=12, pady=(0, 10))

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
        """Load voices with online voices as primary option"""
        self._set_status("Loading voices…")
        
        voices_loaded = False
        
        # Always try to add online voices first (since they work reliably)
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
            
            self.voices = online_voices
            voices_loaded = True
            self._set_status(f"Loaded {len(online_voices)} online voices.")
        
        # Try pyttsx3 as secondary option (only if online voices failed)
        if not voices_loaded:
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
                            self.voices = []
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
                                except Exception:
                                    continue
                            
                            if self.voices:
                                voices_loaded = True
                                engine.stop()
                                self._set_status(f"Loaded {len(self.voices)} system voices.")
                                break
                    
                    except Exception as e:
                        if attempt < max_retries - 1:
                            import time
                            time.sleep(1)
                            continue
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
                        # Last resort: create dummy voices
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
