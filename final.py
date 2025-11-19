"""
Backwards-compatible thin wrapper so existing calls to `final.py` still work.
Prefer running `speak_shell.py` instead — it's the new professional entrypoint.
"""
from speak_shell import main


if __name__ == "__main__":
    main()

        # File ops
        if 'create file' in v or 'make file' in v:
            filename = self.extract_param(v, ['create file', 'make file'])
            if filename:
                filename = self.sanitize_filename(filename)
                if '.' not in filename:
                    filename += '.txt'
                """
                Backwards-compatible thin wrapper so existing calls to `final.py` still work.
                Prefer running `speak_shell.py` instead — it's the new professional entrypoint.
                """
                from speak_shell import main


                if __name__ == "__main__":
                    main()
                    return None, True
