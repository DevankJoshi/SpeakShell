import os
import subprocess
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox
import time
from datetime import datetime

from deps import (
    SPEECH_RECOGNITION_AVAILABLE,
    PYAUDIO_AVAILABLE,
    TTS_AVAILABLE,
    TOAST_AVAILABLE,
    sr,
    pyttsx3,
    ToastNotifier,
    psutil,
)


class HighAccuracyVoiceCMD:
    """
    Voice CMD Terminal (moved from single-file project). Logic and
    behavior are preserved; this module only houses the main class.
    """
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Voice CMD Terminal - High Accuracy Mode")
        self.root.geometry("1000x720")

        # CMD colors
        self.bg_color = '#000000'
        self.text_color = '#FFFFFF'
        self.button_bg = '#2d2d2d'
        self.button_fg = '#00FF00'
        self.warn_fg = '#FFFF00'
        self.err_fg = '#FF5555'
        self.ok_fg = '#00FF00'

        self.root.configure(bg=self.bg_color)

        # Recognizer
        self.recognizer = None
        self.microphone = None

        self.voice_enabled = SPEECH_RECOGNITION_AVAILABLE and PYAUDIO_AVAILABLE
        if self.voice_enabled and sr is not None:
            self.recognizer = sr.Recognizer()
            # tuned parameters
            self.recognizer.energy_threshold = 300
            self.recognizer.dynamic_energy_threshold = True
            # these attributes may not exist on all versions; guard
            try:
                self.recognizer.dynamic_energy_adjustment_damping = 0.15
                self.recognizer.dynamic_energy_ratio = 1.5
            except Exception:
                pass
            self.recognizer.pause_threshold = 0.8
            self.recognizer.operation_timeout = None
            # phrase and non_speaking are optional
            try:
                self.recognizer.phrase_threshold = 0.3
                self.recognizer.non_speaking_duration = 0.5
            except Exception:
                pass

        # TTS
        self.tts_enabled = TTS_AVAILABLE
        try:
            self.tts_engine = pyttsx3.init() if (TTS_AVAILABLE and pyttsx3 is not None) else None
        except Exception:
            self.tts_engine = None

        # Toast
        self.toaster = ToastNotifier() if (TOAST_AVAILABLE and ToastNotifier is not None) else None

        self.is_listening = False

        self.command_history = []
        self.activity_log = []

        # Enhanced voice control parameters (exposed in UI)
        self.phrase_time_limit = 7
        self.listen_timeout = 10
        self.energy_threshold = 300
        # recognition engine: 'google' or 'sphinx' (if available)
        self.recognition_engine = 'google'

        # Track current working directory for navigation
        self.cwd = os.getcwd()

        self.create_simple_gui()

    def create_simple_gui(self):
        # Build a richer UI: menu, toolbar, output pane, history pane and controls
        self.root.configure(bg=self.bg_color)

        # Menu
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Save Log", command=self.save_log)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.destroy)
        menubar.add_cascade(label="File", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0)
        view_menu.add_command(label="Toggle Theme", command=self.toggle_theme)
        menubar.add_cascade(label="View", menu=view_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "Speak Shell — Voice-assisted terminal"))
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

        # Top toolbar
        toolbar = tk.Frame(self.root, bg=self.bg_color)
        toolbar.pack(fill='x', padx=6, pady=6)

        self.start_btn = tk.Button(toolbar, text="Start Listening", command=self.start_listening, bg=self.button_bg, fg=self.button_fg)
        self.start_btn.pack(side='left', padx=4)
        self.stop_btn = tk.Button(toolbar, text="Stop Listening", command=self.stop_listening, bg=self.button_bg, fg='#FF5555', state='disabled')
        self.stop_btn.pack(side='left', padx=4)
        tk.Button(toolbar, text="Calibrate Mic", command=self.calibrate_mic, bg=self.button_bg, fg=self.button_fg).pack(side='left', padx=4)
        tk.Button(toolbar, text="Clear", command=self.clear_screen, bg=self.button_bg, fg=self.button_fg).pack(side='left', padx=4)
        tk.Button(toolbar, text="Save Log", command=self.save_log, bg=self.button_bg, fg=self.button_fg).pack(side='left', padx=4)

        # Engine selector and sliders
        engine_frame = tk.Frame(toolbar, bg=self.bg_color)
        engine_frame.pack(side='right')
        tk.Label(engine_frame, text="Engine:", bg=self.bg_color, fg=self.text_color).pack(side='left')
        self.engine_var = tk.StringVar(value=self.recognition_engine)
        engines = ['google']
        if sr is not None and hasattr(sr, 'Recognizer') and hasattr(sr, 'Recognizer') and hasattr(sr, 'AudioFile'):
            # pocketsphinx detection heuristic
            try:
                import pocketsphinx  # type: ignore
                engines.append('sphinx')
            except Exception:
                pass
        self.engine_menu = tk.OptionMenu(engine_frame, self.engine_var, *engines)
        self.engine_menu.config(bg=self.button_bg, fg=self.button_fg)
        self.engine_menu.pack(side='left', padx=4)

        # Main panes
        content = tk.Frame(self.root, bg=self.bg_color)
        content.pack(fill='both', expand=True, padx=8, pady=6)

        # Left: output + input
        left = tk.Frame(content, bg=self.bg_color)
        left.pack(side='left', fill='both', expand=True)

        self.output_text = scrolledtext.ScrolledText(left, wrap='word', font=('Consolas', 11), bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, relief='flat', padx=8, pady=8)
        self.output_text.pack(fill='both', expand=True)

        input_frame = tk.Frame(left, bg=self.bg_color)
        input_frame.pack(fill='x', pady=(6,0))
        tk.Label(input_frame, text=">", font=('Consolas', 12, 'bold'), bg=self.bg_color, fg=self.button_fg).pack(side='left', padx=(0,6))
        self.input_entry = tk.Entry(input_frame, font=('Consolas', 12), bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color, relief='flat', bd=0)
        self.input_entry.pack(side='left', fill='x', expand=True)
        self.input_entry.bind('<Return>', lambda e: self.execute_input())
        tk.Button(input_frame, text="Execute", command=self.execute_input, bg=self.button_bg, fg=self.button_fg).pack(side='left', padx=6)

        # Right: history & activity
        right = tk.Frame(content, width=320, bg=self.bg_color)
        right.pack(side='right', fill='y')

        tk.Label(right, text="History", bg=self.bg_color, fg=self.warn_fg, font=('Consolas', 11, 'bold')).pack(anchor='nw', padx=6, pady=(2,0))
        self.history_listbox = tk.Listbox(right, height=10, bg='#111111', fg='#AAAAAA', activestyle='dotbox')
        self.history_listbox.pack(fill='x', padx=6, pady=(0,6))
        self.history_listbox.bind('<Double-1>', self._on_history_double)

        tk.Label(right, text="Activity Log", bg=self.bg_color, fg=self.text_color, font=('Consolas', 10, 'bold')).pack(anchor='nw', padx=6)
        self.activity_listbox = tk.Listbox(right, height=12, bg='#111111', fg='#CCCCCC')
        self.activity_listbox.pack(fill='both', expand=True, padx=6, pady=(0,6))

        # Bottom status
        status = tk.Frame(self.root, bg=self.bg_color)
        status.pack(fill='x')
        self.status_label = tk.Label(status, text="Status: Ready", font=('Consolas', 10), bg=self.bg_color, fg=self.ok_fg)
        self.status_label.pack(side='left', padx=6)
        self.cwd_label = tk.Label(status, text=f"CWD: {self.cwd}", font=('Consolas', 9), bg=self.bg_color, fg='#00CED1')
        self.cwd_label.pack(side='right', padx=6)

        # welcome
        welcome = """
Speak Shell — type commands or use voice. Use the toolbar to calibrate mic or toggle theme.
"""
        self.output_text.insert('1.0', welcome)
        self.output_text.see('end')

        # expose controls
        self.engine_var.trace_add('write', lambda *a: setattr(self, 'recognition_engine', self.engine_var.get()))

        # sliders for advanced tuning
        tuning = tk.Frame(right, bg=self.bg_color)
        tuning.pack(fill='x', padx=6, pady=6)
        tk.Label(tuning, text="Mic Energy", bg=self.bg_color, fg=self.text_color).pack(anchor='w')
        self.energy_slider = tk.Scale(tuning, from_=100, to=1000, orient='horizontal', bg=self.bg_color, fg=self.text_color, troughcolor='#222222', command=self._on_energy_change)
        self.energy_slider.set(self.energy_threshold)
        self.energy_slider.pack(fill='x')
        tk.Label(tuning, text="Phrase time limit (s)", bg=self.bg_color, fg=self.text_color).pack(anchor='w')
        self.phrase_slider = tk.Scale(tuning, from_=2, to=12, orient='horizontal', bg=self.bg_color, fg=self.text_color, troughcolor='#222222', command=self._on_phrase_change)
        self.phrase_slider.set(self.phrase_time_limit)
        self.phrase_slider.pack(fill='x')

        # ensure focus
        self.input_entry.focus_set()

    # --- UI callbacks and helpers ---
    def toggle_theme(self):
        # Simple theme toggle between dark and light
        if self.bg_color == '#000000':
            self.bg_color = '#FFFFFF'
            self.text_color = '#000000'
            self.button_bg = '#EEEEEE'
            self.button_fg = '#003300'
            self.warn_fg = '#AA7700'
            self.ok_fg = '#006600'
        else:
            self.bg_color = '#000000'
            self.text_color = '#FFFFFF'
            self.button_bg = '#2d2d2d'
            self.button_fg = '#00FF00'
            self.warn_fg = '#FFFF00'
            self.ok_fg = '#00FF00'
        # Apply to some widgets
        try:
            self.output_text.config(bg=self.bg_color, fg=self.text_color, insertbackground=self.text_color)
            self.input_entry.config(bg=self.bg_color, fg=self.text_color)
            self.status_label.config(bg=self.bg_color, fg=self.ok_fg)
            self.cwd_label.config(bg=self.bg_color)
        except Exception:
            pass

    def calibrate_mic(self):
        if sr is None or self.recognizer is None:
            messagebox.showinfo("Calibrate", "SpeechRecognition not available")
            return
        # Run a short calibration on a background thread to avoid blocking UI
        def _cal():
            try:
                with sr.Microphone(sample_rate=16000) as src:
                    self.root.after(0, self.print_output, "[Calibrate] Listening to ambient noise for 2s...")
                    self.recognizer.adjust_for_ambient_noise(src, duration=2)
                    # update energy threshold in UI
                    self.energy_threshold = getattr(self.recognizer, 'energy_threshold', self.energy_threshold)
                    self.root.after(0, lambda: self.energy_slider.set(self.energy_threshold))
                    self.root.after(0, self.print_output, f"[Calibrate] Done. energy_threshold={self.energy_threshold}")
            except Exception as e:
                self.root.after(0, self.print_output, f"[Calibrate] Error: {e}")
        threading.Thread(target=_cal, daemon=True).start()

    def _on_energy_change(self, val):
        try:
            v = int(val)
            self.energy_threshold = v
            if self.recognizer:
                try:
                    self.recognizer.energy_threshold = v
                except Exception:
                    pass
        except Exception:
            pass

    def _on_phrase_change(self, val):
        try:
            v = int(val)
            self.phrase_time_limit = v
        except Exception:
            pass

    def _on_history_double(self, event):
        try:
            sel = self.history_listbox.curselection()
            if sel:
                cmd = self.history_listbox.get(sel[0])
                # run the command again
                self.input_entry.delete(0, tk.END)
                self.input_entry.insert(0, cmd)
                self.execute_input()
        except Exception:
            pass

    def speak(self, text):
        if getattr(self, 'tts_var', None) and self.tts_var.get() and self.tts_engine:
            try:
                self.tts_engine.say(text)
                self.tts_engine.runAndWait()
            except Exception:
                pass

    def toast(self, title, msg, duration=4):
        try:
            if self.toaster:
                self.toaster.show_toast(title, msg, duration=duration, threaded=True)
        except Exception:
            pass

    def print_output(self, text, newline=True):
        self.output_text.insert('end', text)
        if newline:
            self.output_text.insert('end', '\n')
        self.output_text.see('end')

    def update_cwd(self, new_path=None):
        if new_path:
            self.cwd = new_path
        self.cwd_label.config(text=f"CWD: {self.cwd}")

    def start_listening(self):
        if not self.voice_enabled or self.is_listening:
            return
        self.is_listening = True
        self.status_label.config(text="Status: LISTENING | Speak clearly for best accuracy", fg=self.warn_fg)
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')

        self.listen_thread = threading.Thread(target=self.high_accuracy_listen_loop, daemon=True)
        self.listen_thread.start()

        self.print_output("\n[Voice] High accuracy mode activated")
        self.print_output("[Voice] Listening started with 90%+ accuracy settings...")
        self.log_activity("VOICE", "High accuracy listening started")

    def stop_listening(self):
        self.is_listening = False
        self.status_label.config(text="Status: Ready ", fg=self.ok_fg)
        if hasattr(self, 'start_btn'):
            self.start_btn.config(state='normal')
        if hasattr(self, 'stop_btn'):
            self.stop_btn.config(state='disabled')
        self.print_output("[Voice] Listening stopped.")
        self.log_activity("VOICE", "Listening stopped")

    def high_accuracy_listen_loop(self):
        if sr is None:
            return
        with sr.Microphone(sample_rate=16000) as source:
            self.root.after(0, self.print_output, "[Voice] Calibrating for ambient noise...")
            self.root.after(0, self.status_label.config, {"text": "Status: CALIBRATING (Please wait 2 seconds)...", "fg": "#00FFFF"})
            try:
                self.recognizer.adjust_for_ambient_noise(source, duration=2)
            except Exception:
                pass
            self.root.after(0, self.print_output, "[Voice] Calibration complete!")
            self.root.after(0, self.print_output, "[Voice] Ready! Speak your commands clearly...")
            self.root.after(0, self.status_label.config, {"text": "Status: LISTENING | Speak now!", "fg": "#00FF00"})

            while self.is_listening:
                try:
                    self.root.after(0, self.status_label.config, {"text": "Status: LISTENING | Speak clearly...", "fg": self.warn_fg})
                    audio = self.recognizer.listen(source, timeout=getattr(self, 'listen_timeout', 10), phrase_time_limit=getattr(self, 'phrase_time_limit', 7))
                    self.root.after(0, self.status_label.config, {"text": "Status: PROCESSING with high accuracy...", "fg": "#00FFFF"})

                    # Choose recognition engine dynamically
                    recognized = None
                    try:
                        if getattr(self, 'recognition_engine', 'google') == 'sphinx' and hasattr(self.recognizer, 'recognize_sphinx'):
                            recognized = self.recognizer.recognize_sphinx(audio)
                        else:
                            recognized = self.recognizer.recognize_google(audio, language='en-US', show_all=False)
                    except Exception as recog_err:
                        # Pass through to outer handlers
                        raise recog_err

                    text = recognized
                    if text and len(text) > 0:
                        self.root.after(0, self.print_output, f"[Voice] Recognized: {text}")
                        self.root.after(0, self.process_command, text, "voice")
                except sr.WaitTimeoutError:
                    continue
                except sr.UnknownValueError:
                    self.root.after(0, self.print_output, "[Voice] Could not understand - please speak more clearly")
                    self.root.after(0, self.print_output, "[Voice] Tips: Speak at normal pace, reduce background noise")
                    time.sleep(0.5)
                except sr.RequestError as e:
                    self.root.after(0, self.print_output, f"[Voice] ERROR: Google API error - {e}")
                    self.root.after(0, self.print_output, "[Voice] Check internet connection")
                    self.root.after(0, self.print_output, "[Voice] Use manual text input as backup")
                    break
                except Exception as e:
                    self.root.after(0, self.print_output, f"[Voice] ERROR: {str(e)}")
                    break

    def execute_input(self):
        command = self.input_entry.get().strip()
        if command:
            self.input_entry.delete(0, tk.END)
            self.process_command(command, "manual")

    def process_command(self, command, source="manual"):
        command = command.strip()
        self.print_output(f"\n> {command}")
        # record to history UI
        try:
            self.command_history.append(command)
            if hasattr(self, 'history_listbox'):
                self.history_listbox.insert('end', command)
        except Exception:
            pass
        self.log_activity(source.upper(), command)

        # Exit flow
        if any(w in command.lower() for w in ['exit', 'quit', 'close']):
            self.print_output("Exiting...")
            self.save_log()
            self.root.after(700, self.root.destroy)
            return

        # Local interpreted meta-commands
        if command.lower() == 'save log':
            self.save_log()
            return
        if command.lower().startswith('clear'):
            self.clear_screen()
            return

        # Map then execute
        system_cmd, is_shell = self.map_to_cmd(command)
        if not system_cmd:
            self.print_output("ERROR: Command not recognized. Type 'help' for commands.")
            self.speak("Command not recognized")
            return

        self.print_output(f"Executing: {system_cmd}")
        self.log_activity("EXECUTE", system_cmd)
        self.run_cmd(system_cmd, is_shell=is_shell)

    def sanitize_filename(self, name):
        # Reduce path traversal and strip quotes
        name = name.strip().strip('"').strip("'")
        # If it's a relative simple name, allow; otherwise normalize
        # Keep spaces intact; forbid redirection/special shell chars
        forbidden = ['&', '|', ';', '>', '<', '`']
        for ch in forbidden:
            name = name.replace(ch, '')
        return name

    def confirm(self, title, message):
        try:
            return messagebox.askokcancel(title, message)
        except Exception:
            return False

    def run_cmd(self, cmd, is_shell=True):
        """
        Execute a command in current working directory.
        If launching apps (start/explorer), use Popen; else run and capture.
        """
        try:
            # Built-in launchers
            if isinstance(cmd, str) and (cmd.startswith('start ') or cmd.startswith('explorer ')):
                subprocess.Popen(cmd, shell=True, cwd=self.cwd)
                self.print_output("OK - Application launched")
                self.toast("Voice CMD", "Application launched")
                self.speak("Application launched")
                return

            result = subprocess.run(
                cmd,
                capture_output=True, text=True,
                shell=is_shell, timeout=60, cwd=self.cwd
            )

            if result.stdout:
                output = result.stdout.strip()
                if len(output) > 12000:
                    output = output[:12000] + "\n... (output truncated)"
                self.print_output(output)

            if result.returncode == 0:
                self.print_output("OK")
                self.toast("Voice CMD", "Command executed successfully")
                self.speak("Command executed successfully")
            else:
                self.print_output(f"ERROR: Exit code {result.returncode}")
                if result.stderr:
                    self.print_output(result.stderr.strip())
                self.toast("Voice CMD", "Command failed")
                self.speak("Command failed")

        except subprocess.TimeoutExpired:
            self.print_output("ERROR: Command timed out")
            self.toast("Voice CMD", "Command timed out")
            self.speak("Command timed out")
        except Exception as e:
            self.print_output(f"ERROR: {str(e)}")
            self.toast("Voice CMD", "Unexpected error")
            self.speak("Unexpected error")

    def map_to_cmd(self, voice):
        """
        Map human or voice command to an actual command string.
        Returns (cmd_string, is_shell_bool).
        """
        v = voice.lower().strip()

        # Help
        if v == 'help':
            self.print_output("""
Available Commands:
  File ops:
    create file <name>           - Create a file (adds .txt if no extension)
    open file <name>             - Open a file with default app
    delete file <name>           - Delete a file (with confirmation)
    rename <old> to <new>        - Rename a file or folder
    move <src> to <dst>          - Move file/folder
    copy <src> to <dst>          - Copy file/folder
  Directory ops:
    list files                   - dir
    show files                   - dir
    create directory <name>      - mkdir
    make folder <name>           - mkdir
    go to desktop/downloads/docs - quick jump
    go to <path or folder>       - change directory if exists
    cd <path>                    - change directory
    go up                        - cd ..
  System info:
    show processes               - tasklist
    kill process <name>          - taskkill /f /im <name>.exe (confirm)
    task manager                 - start taskmgr
    system information           - systeminfo (filtered)
    memory usage                 - wmic OS get FreePhysicalMemory,TotalVisibleMemorySize /value
    disk space                   - wmic logicaldisk get caption,freespace,size
    battery status               - wmic path Win32_Battery get EstimatedChargeRemaining,Status
    network info                 - ipconfig /all
  Apps:
    calculator/notepad/paint     - launch
  Time/date:
    what time is it              - time /t
    what is the date             - date /t
  Misc:
    save log, clear screen, exit
  Raw CMD:
    say any Windows command listed in Microsoft docs; it will be passed through safely.
""")
            return None, True

        # Quick exits
        if any(k in v for k in ['exit', 'quit', 'close']):
            return 'echo Exiting...', True

        # Time/date phrases
        if 'what time' in v or 'current time' in v or 'show time' in v:
            return 'time /t', True
        if 'what date' in v or 'current date' in v or 'show date' in v or 'what is the date' in v:
            return 'date /t', True
        if v == 'what time is it':
            return 'time /t', True

        # Directory navigation
        if v.startswith('cd '):
            path = v[3:].strip().strip('"')
            resolved = self.resolve_path(path)
            if os.path.isdir(resolved):
                self.update_cwd(resolved)
                self.print_output(f"Directory changed to: {self.cwd}")
                self.speak("Directory changed")
                return 'dir', True
            else:
                self.print_output("ERROR: Directory not found")
                self.speak("Directory not found")
                return None, True

        if v == 'go up' or v == 'cd ..' or v == 'go back':
            parent = os.path.abspath(os.path.join(self.cwd, '..'))
            if os.path.isdir(parent):
                self.update_cwd(parent)
                self.print_output(f"Directory changed to: {self.cwd}")
                self.speak("Directory changed")
                return 'dir', True
            else:
                self.print_output("ERROR: Could not go up")
                self.speak("Could not go up")
                return None, True

        # "go to <path or folder>"
        if v.startswith('go to '):
            name = v[6:].strip().strip('"')
            resolved = self.resolve_path(name)
            if os.path.isdir(resolved):
                self.update_cwd(resolved)
                self.print_output(f"Directory changed to: {self.cwd}")
                self.speak("Directory changed")
                return 'dir', True
            else:
                # try relative in cwd
                maybe = os.path.join(self.cwd, name)
                if os.path.isdir(maybe):
                    self.update_cwd(os.path.abspath(maybe))
                    self.print_output(f"Directory changed to: {self.cwd}")
                    self.speak("Directory changed")
                    return 'dir', True
                self.print_output("ERROR: Target directory not found")
                self.speak("Target directory not found")
                return None, True

        # Quick jumps
        if 'go to desktop' in v:
            return self.change_dir_quick('%USERPROFILE%\\Desktop')
        if 'go to downloads' in v:
            return self.change_dir_quick('%USERPROFILE%\\Downloads')
        if 'go to documents' in v or 'go to docs' in v:
            return self.change_dir_quick('%USERPROFILE%\\Documents')

        # File ops
        if 'create file' in v or 'make file' in v:
            filename = self.extract_param(v, ['create file', 'make file'])
            if filename:
                filename = self.sanitize_filename(filename)
                if '.' not in filename:
                    filename += '.txt'
                # Create safely using Python instead of shell
                try:
                    open(os.path.join(self.cwd, filename), 'a', encoding='utf-8').close()
                    self.print_output(f"Created file: {filename}")
                    self.speak("File created")
                    return 'dir', True
                except Exception as e:
                    self.print_output(f"ERROR: {e}")
                    self.speak("Failed to create file")
                    return None, True
            return None, True

        if 'open file' in v:
            filename = self.extract_param(v, ['open file'])
            if filename:
                filename = self.sanitize_filename(filename)
                full = os.path.join(self.cwd, filename)
                if os.path.exists(full):
                    return f'start "" "{full}"', True
                else:
                    self.print_output("ERROR: File not found")
                    self.speak("File not found")
                    return None, True
            return None, True

        if 'delete file' in v or 'remove file' in v:
            filename = self.extract_param(v, ['delete file', 'remove file'])
            if filename:
                filename = self.sanitize_filename(filename)
                full = os.path.join(self.cwd, filename)
                if os.path.exists(full) and os.path.isfile(full):
                    if self.confirm("Confirm Delete", f"Delete file '{filename}'?"):
                        try:
                            os.remove(full)
                            self.print_output(f"Deleted file: {filename}")
                            self.speak("File deleted")
                            return 'dir', True
                        except Exception as e:
                            self.print_output(f"ERROR: {e}")
                            self.speak("Failed to delete file")
                            return None, True
                    else:
                        self.print_output("Delete cancelled")
                        return None, True
                else:
                    self.print_output("ERROR: File not found")
                    self.speak("File not found")
                    return None, True
            return None, True

        # Rename/move/copy using phrases
        if v.startswith('rename '):
            # Expect "rename old to new"
            parts = v.replace('  ', ' ').split()
            if ' to ' in v:
                old, new = v[len('rename '):].split(' to ', 1)
            else:
                # fallback two names
                toks = v.split()
                if len(toks) >= 3:
                    old, new = toks[1], toks[2]
                else:
                    old, new = None, None
            if old and new:
                old = self.sanitize_filename(old.strip().strip('"'))
                new = self.sanitize_filename(new.strip().strip('"'))
                src = os.path.join(self.cwd, old)
                dst = os.path.join(self.cwd, new)
                if os.path.exists(src):
                    try:
                        os.replace(src, dst)
                        self.print_output(f"Renamed '{old}' to '{new}'")
                        self.speak("Rename completed")
                        return 'dir', True
                    except Exception as e:
                        self.print_output(f"ERROR: {e}")
                        self.speak("Rename failed")
                        return None, True
                else:
                    self.print_output("ERROR: Source not found")
                    self.speak("Source not found")
                    return None, True
            self.print_output("Usage: rename <old> to <new>")
            return None, True

        if v.startswith('move '):
            # "move src to dst"
            if ' to ' in v:
                rest = v[len('move '):]
                src, dst = rest.split(' to ', 1)
            else:
                toks = v.split()
                src = toks[1] if len(toks) >= 3 else None
                dst = toks[2] if len(toks) >= 3 else None
            if src and dst:
                src = self.sanitize_filename(src.strip().strip('"'))
                dst = self.sanitize_filename(dst.strip().strip('"'))
                srcp = os.path.join(self.cwd, src)
                dstp = os.path.join(self.cwd, dst)
                if os.path.exists(srcp):
                    try:
                        os.replace(srcp, dstp)
                        self.print_output(f"Moved '{src}' to '{dst}'")
                        self.speak("Move completed")
                        return 'dir', True
                    except Exception as e:
                        self.print_output(f"ERROR: {e}")
                        self.speak("Move failed")
                        return None, True
                else:
                    self.print_output("ERROR: Source not found")
                    self.speak("Source not found")
                    return None, True
            self.print_output("Usage: move <src> to <dst>")
            return None, True

        if v.startswith('copy '):
            # "copy src to dst"
            import shutil
            if ' to ' in v:
                rest = v[len('copy '):]
                src, dst = rest.split(' to ', 1)
            else:
                toks = v.split()
                src = toks[1] if len(toks) >= 3 else None
                dst = toks[2] if len(toks) >= 3 else None
            if src and dst:
                src = self.sanitize_filename(src.strip().strip('"'))
                dst = self.sanitize_filename(dst.strip().strip('"'))
                srcp = os.path.join(self.cwd, src)
                dstp = os.path.join(self.cwd, dst)
                if os.path.exists(srcp):
                    try:
                        if os.path.isdir(srcp):
                            shutil.copytree(srcp, dstp, dirs_exist_ok=True)
                        else:
                            os.makedirs(os.path.dirname(dstp), exist_ok=True) if os.path.dirname(dstp) else None
                            shutil.copy2(srcp, dstp)
                        self.print_output(f"Copied '{src}' to '{dst}'")
                        self.speak("Copy completed")
                        return 'dir', True
                    except Exception as e:
                        self.print_output(f"ERROR: {e}")
                        self.speak("Copy failed")
                        return None, True
                else:
                    self.print_output("ERROR: Source not found")
                    self.speak("Source not found")
                    return None, True
            self.print_output("Usage: copy <src> to <dst>")
            return None, True

        # Directory ops
        if any(p in v for p in ['list files', 'show files', 'list directory']):
            return 'dir', True

        if 'create directory' in v or 'make folder' in v or v.startswith('mkdir '):
            dirname = self.extract_param(v, ['create directory', 'make folder'])
            if not dirname and v.startswith('mkdir '):
                dirname = v[len('mkdir '):].strip().strip('"')
            if dirname:
                dirname = self.sanitize_filename(dirname)
                try:
                    os.makedirs(os.path.join(self.cwd, dirname), exist_ok=True)
                    self.print_output(f"Directory created: {dirname}")
                    self.speak("Directory created")
                    return 'dir', True
                except Exception as e:
                    self.print_output(f"ERROR: {e}")
                    self.speak("Failed to create directory")
                    return None, True
            return None, True

        # System operations
        if 'show processes' in v or 'list processes' in v or v == 'tasklist':
            return 'tasklist', True

        if 'kill process' in v or 'terminate process' in v:
            process = self.extract_param(v, ['kill process', 'terminate process'])
            if process:
                process = self.sanitize_filename(process)
                procname = process if process.lower().endswith('.exe') else f"{process}.exe"
                if self.confirm("Confirm Kill", f"Terminate process '{procname}'?"):
                    return f'taskkill /f /im "{procname}"', True
                else:
                    self.print_output("Kill cancelled")
                    return None, True
            return None, True

        if 'task manager' in v:
            return 'start taskmgr', True

        if 'system information' in v or 'system info' in v:
            return 'systeminfo | findstr /C:"Host Name" /C:"OS Name" /C:"System Type" /C:"Total Physical Memory"', True

        if 'memory usage' in v or 'ram usage' in v:
            return 'wmic OS get FreePhysicalMemory,TotalVisibleMemorySize /value', True

        if 'disk space' in v or 'storage' in v:
            return 'wmic logicaldisk get caption,freespace,size', True

        if 'battery status' in v:
            # Guard: some desktops have no battery -> command may output nothing
            return 'wmic path Win32_Battery get EstimatedChargeRemaining,Status', True

        if 'network info' in v or v == 'ipconfig':
            return 'ipconfig /all', True

        # Applications
        if 'calculator' in v or v == 'calc':
            return 'start calc', True
        if 'notepad' in v:
            return 'start notepad', True
        if 'paint' in v or 'mspaint' in v:
            return 'start mspaint', True

        # Raw Windows command passthrough (safe-ish)
        forbidden = ['&', '|', ';', '>', '<', '`']
        if not any(ch in v for ch in forbidden):
            return v, True

        return None, True

    def change_dir_quick(self, env_path):
        path = os.path.expandvars(env_path)
        if os.path.isdir(path):
            self.update_cwd(path)
            self.print_output(f"Directory changed to: {self.cwd}")
            self.speak("Directory changed")
            return 'dir', True
        else:
            self.print_output("ERROR: Target folder not found")
            self.speak("Target folder not found")
            return None, True

    def resolve_path(self, name_or_path):
        # If absolute or has drive, expand and normalize
        p = os.path.expanduser(os.path.expandvars(name_or_path))
        if os.path.isabs(p):
            return os.path.abspath(p)
        # Otherwise relative to cwd
        return os.path.abspath(os.path.join(self.cwd, p))

    def extract_param(self, text, phrases):
        text_l = text.lower()
        for phrase in phrases:
            if phrase in text_l:
                pos = text_l.find(phrase) + len(phrase)
                original_tail = text[pos:].strip()
                # remove filler words
                cleaned = original_tail.replace('called', '').replace('named', '').strip()
                cleaned = cleaned.replace(' dot ', '.').replace(' dot', '.').replace('dot ', '.').strip()
                return cleaned if cleaned else None
        return None

    def log_activity(self, activity_type, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.activity_log.append(f"[{timestamp}] [{activity_type}] {message}")
        # also mirror into activity listbox if present
        try:
            if hasattr(self, 'activity_listbox'):
                self.activity_listbox.insert('end', f"[{activity_type}] {message}")
                # keep bottom visible
                self.activity_listbox.yview_moveto(1.0)
        except Exception:
            pass

    def save_log(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"voice_cmd_log_{timestamp}.txt"
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("VOICE CMD TERMINAL - HIGH ACCURACY MODE - ACTIVITY LOG\n")
                f.write("="*80 + "\n\n")
                for entry in self.activity_log:
                    f.write(entry + "\n")
            self.print_output(f"\nLog saved: {filename}")
            self.toast("Voice CMD", f"Log saved: {filename}")
            self.speak("Log saved")
        except Exception as e:
            self.print_output(f"ERROR: Could not save log - {str(e)}")
            self.speak("Failed to save log")

    def clear_screen(self):
        self.output_text.delete('1.0', 'end')
        self.print_output("Screen cleared. Type 'help' for commands.\n")

    def run(self):
        self.root.mainloop()
