"""GUI-aware logger for text widget output."""
from __future__ import annotations

import time


class Logger:
    def __init__(self, tb):
        self.tb = tb
        try:
            self.tb.tag_configure("info", foreground="orange")
            self.tb.tag_configure("pack", foreground="blue")
            self.tb.tag_configure("error", foreground="red")
        except Exception:
            pass

    def log(self, msg: str) -> None:
        ts = time.strftime("%H:%M:%S")
        tag = None
        if "[INFO]" in msg:
            tag = "info"
        elif "[PACK]" in msg:
            tag = "pack"
        elif "ERROR" in msg:
            tag = "error"
        self.tb.configure(state="normal")
        try:
            self.tb.insert("end", f"[{ts}] {msg}\n", tag)
        except Exception:
            self.tb.insert("end", f"[{ts}] {msg}\n")
        self.tb.see("end")
        self.tb.configure(state="normal")
        self.tb.update_idletasks()

    def err(self, msg: str) -> None:
        self.log(f"ERROR: {msg}")
