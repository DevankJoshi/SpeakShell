# SpeakShell (Speak Shell)

SpeakShell is a lightweight voice-assisted command terminal for Windows with a simple GUI.
It uses SpeechRecognition (Google API) for high-accuracy voice recognition, optional TTS feedback
(pyttsx3), and safe command execution with confirmations. The original single-file project has been
split into modules for clarity and presentation — no logic changes were made.

## What I changed
- Split the original `final.py` into modules for better presentation and maintainability:
  - `final.py` — the entrypoint (runner) kept for backwards compatibility.
  - `voice_cmd.py` — contains the `HighAccuracyVoiceCMD` class (main app logic and GUI).
  - `deps.py` — centralizes optional imports and availability flags (SpeechRecognition, PyAudio, TTS, etc.).
  - `helpers.py` — small UI constants (colors, header) to make future presentation tweaks simple.
- Added this `README.md` with usage and file descriptions.

No runtime behavior or command mappings were changed — only reorganized for clarity.

## Requirements
- Python 3.8+
- Optional (for full functionality):
  - SpeechRecognition and PyAudio (voice input): `pip install SpeechRecognition pyaudio`
  - pyttsx3 (TTS): `pip install pyttsx3`
  - win10toast (Windows notifications): `pip install win10toast`
  - psutil (extra system info): `pip install psutil`

If optional dependencies are missing, the app falls back to text input and prints a notice in the UI.

## Run
Run the program using the new professional entrypoint `speak_shell.py` (a `final.py` wrapper is still available):

```bash
python speak_shell.py
# or for backward compatibility
python final.py
```

This will launch the GUI. Use the input field for typed commands, or click "Start Listening" if
voice dependencies are installed.

## Files
- `final.py` — program entrypoint. Prints startup info then instantiates the main app.
- `voice_cmd.py` — main application class and methods (GUI, command mapping, execution, voice loop).
- `deps.py` — detects availability of optional dependencies and exports flags/modules.
- `helpers.py` — UI constants and small utilities for future UI tweaks.

## Notes & Next steps
- I only reorganized the project and added the README. No logic or behavior was intentionally
  modified. If you'd like, I can:
  - Extract command handlers into separate modules for unit testing.
  - Add a small unit test suite for core text-command-to-action mapping.
  - Improve cross-platform compatibility (currently tailored to Windows `cmd` commands).

If you want any of these, tell me which and I'll implement them next.
# SpeakShell