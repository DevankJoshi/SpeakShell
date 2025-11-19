"""
Entry point (was: final.py) — now a thin runner that instantiates the
HighAccuracyVoiceCMD class (kept intact in `voice_cmd.py`).

This file preserves the original startup prints and acts as the main
script so existing workflows that run `final.py` continue to work.
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

    def run_cmd(self, cmd, is_shell=True):
        """
        Execute a command in current working directory.
        If launching apps (start/explorer), use Popen; else run and capture.
        """
        try:
            # Built-in launchers
            if cmd.startswith('start ') or cmd.startswith('explorer '):
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
        # Allow user to say a known Windows command directly (e.g., "dir", "chkdsk", "whoami", etc.)
        # For safety, block redirection and pipe characters.
        forbidden = ['&', '|', ';', '>', '<', '`']
        if not any(ch in v for ch in forbidden):
            # If it looks like a command (single word or typical cmd pattern), run it via shell
            # Examples: dir, whoami, ver, ipconfig, tree, where python, findstr ...
            # Keep current working directory context
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


if __name__ == "__main__":
    print("="*60)
    print("Speak Shell")
    print("="*60)
    print("Optimizations:")
    print("  • Enhanced audio preprocessing")
    print("  • Dynamic energy threshold (auto noise adjustment)")
    print("  • Google Speech API with en-US language model")
    print("Enhancements:")
    print("  • TTS voice feedback, Windows notifications")
    print("  • Safer execution with confirmations")
    print("  • Extended navigation and system info")
    print("="*60)
    print()

    app = HighAccuracyVoiceCMD()
    app.run()
