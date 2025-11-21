"""GUI-aware logger for text widget output."""
from __future__ import annotations

import time


class Logger:
    COLORS = {
        "SYSTEM": "#0d47a1",   # blue
        "INFO": "#333333",     # black/gray
        "SUCCESS": "#2e7d32",  # green
        "WARNING": "#ef6c00",  # orange
        "ERROR": "#c62828",    # red
    }

    def __init__(self, tb):
        self.tb = tb
        try:
            for lvl, color in self.COLORS.items():
                self.tb.tag_configure(lvl.lower(), foreground=color)
        except Exception:
            pass

    @staticmethod
    def _format_exc(err: Exception) -> str:
        return f"{type(err).__name__}: {err}" if err else ""

    def _normalize(self, msg):
        if isinstance(msg, Exception):
            return self._format_exc(msg)
        return str(msg)

    def _detect_level(self, msg: str, level: str | None) -> str:
        if level:
            return level.upper()
        if "[ERROR]" in msg or msg.upper().startswith("ERROR"):
            return "ERROR"
        if "[SUCCESS]" in msg:
            return "SUCCESS"
        if "[WARN" in msg or "WARNING" in msg:
            return "WARNING"
        if "[PACK]" in msg or "[TMPBATCH]" in msg or "[PROGRESS]" in msg:
            return "SYSTEM"
        return "INFO"

    def _log(self, msg, level: str | None = None) -> None:
        txt = self._normalize(msg)
        lvl = self._detect_level(txt, level)
        ts = time.strftime("%H:%M:%S")
        tag = lvl.lower()
        self.tb.configure(state="normal")
        try:
            self.tb.insert("end", f"[{ts}] {txt}\n", tag)
        except Exception:
            self.tb.insert("end", f"[{ts}] {txt}\n")
        self.tb.see("end")
        self.tb.configure(state="normal")
        self.tb.update_idletasks()

    def log(self, msg: str) -> None:
        self._log(msg, "INFO")

    def log_system(self, msg) -> None:
        self._log(msg, "SYSTEM")

    def log_info(self, msg) -> None:
        self._log(msg, "INFO")

    def log_success(self, msg) -> None:
        self._log(msg, "SUCCESS")

    def log_warning(self, msg) -> None:
        self._log(msg, "WARNING")

    def log_error(self, msg) -> None:
        self._log(msg, "ERROR")

    def err(self, msg: str) -> None:
        self.log_error(msg)
