"""
only.py — Launcher & CLI tool for InvenScan
Run:
    python only.py
"""

import sys
import app

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".bmp")):
        app.run_cli_ocr(sys.argv[1])
    else:
        print("\n" + "="*55)
        print(">> Starting InvenScan via only.py...")
        print("="*55)
        try:
            import webbrowser
            webbrowser.open("http://localhost:5000/form")
        except Exception:
            pass
        app.app.run(debug=False, host="0.0.0.0", port=5000)
