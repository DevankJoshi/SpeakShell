"""
Dependency detection shim.

This module centralizes optional imports and exposes flags and modules
so the main application can remain unchanged in logic while moving
the import/availability checks to one place.
"""
SPEECH_RECOGNITION_AVAILABLE = False
PYAUDIO_AVAILABLE = False
TTS_AVAILABLE = False
TOAST_AVAILABLE = False

sr = None
pyaudio = None
pyttsx3 = None
ToastNotifier = None
psutil = None

try:
    import pyaudio as _pyaudio
    pyaudio = _pyaudio
    PYAUDIO_AVAILABLE = True
except Exception:
    pyaudio = None

try:
    import speech_recognition as _sr
    sr = _sr
    SPEECH_RECOGNITION_AVAILABLE = True
except Exception:
    sr = None

try:
    import pyttsx3 as _pyttsx3
    pyttsx3 = _pyttsx3
    TTS_AVAILABLE = True
except Exception:
    pyttsx3 = None

try:
    from win10toast import ToastNotifier as _ToastNotifier
    ToastNotifier = _ToastNotifier
    TOAST_AVAILABLE = True
except Exception:
    ToastNotifier = None

try:
    import psutil as _psutil
    psutil = _psutil
except Exception:
    psutil = None
