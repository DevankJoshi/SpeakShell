"""
Professional entrypoint for SpeakShell.

Run this file to launch the GUI application.
"""
from voice_cmd import HighAccuracyVoiceCMD


def main():
    print("=" * 60)
    print("Speak Shell")
    print("=" * 60)
    print("Optimizations:")
    print("  • Enhanced audio preprocessing")
    print("  • Dynamic energy threshold (auto noise adjustment)")
    print("  • Google Speech API with en-US language model")
    print("Enhancements:")
    print("  • TTS voice feedback, Windows notifications")
    print("  • Safer execution with confirmations")
    print("  • Extended navigation and system info")
    print("=" * 60)
    print()

    app = HighAccuracyVoiceCMD()
    app.run()


if __name__ == "__main__":
    main()
