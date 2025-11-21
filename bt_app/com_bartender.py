"""BarTender COM automation wrapper."""
from __future__ import annotations

from typing import Any, Dict

try:
    from win32com.client import Dispatch
except Exception:  # pragma: no cover - pywin32 only on Windows
    Dispatch = None


class BT:
    def __init__(self, logger):
        self.logger = logger
        self.app = None

    def start(self):
        if not Dispatch:
            raise RuntimeError("pywin32 не установле")
        self.logger.log("Запуск BarTender COM...")
        self.app = Dispatch("BarTender.Application")
        self.app.Visible = False
        self.logger.log("BarTender COM запущен.")

    def stop(self):
        if self.app:
            try:
                self.logger.log("Завершение BarTender COM...")
                self.app.Quit(1)
            except Exception:
                pass
            self.app = None

    def open_format(self, path: str):
        self.logger.log(f"Открытие шаблона: {path}")
        fmt = self.app.Formats.Open(path, False, "")
        try:
            self.logger.log(f"NamedSubStrings: {[s.Name for s in fmt.NamedSubStrings]}")
        except Exception:
            pass
        return fmt

    def set_common_print_flags(self, fmt):
        for a, v in (("UseDatabase", False), ("SelectRecordsAtPrint", False), ("RecordRange", "1")):
            try:
                setattr(fmt, a, v)
            except Exception:
                pass

    def apply_fields(self, fmt, data: Dict[str, Any]) -> bool:
        names = set()
        try:
            names = {s.Name for s in fmt.NamedSubStrings}
        except Exception:
            pass
        payload = {k: v for k, v in data.items() if not k.startswith("_") and ((not names) or (k in names))}
        skipped = sorted(set(data.keys()) - set(payload.keys()) - {k for k in data if k.startswith("_")})
        if skipped:
            self.logger.log(f"Подстановка: пропущены поля (нет в шаблоне): {skipped}")
        cnt = 0
        for k, v in payload.items():
            try:
                fmt.SetNamedSubStringValue(k, str(v))
                cnt += 1
                continue
            except Exception:
                pass
            try:
                subs = getattr(fmt, "SubStrings", None)
                if subs:
                    subs(k).Value = str(v)
                    cnt += 1
            except Exception:
                pass
        self.logger.log(f"Подстановка полей: всего={len(payload)}, успешно={cnt}")
        return cnt > 0

    def export_preview(self, fmt, path: str) -> bool:
        try:
            fmt.ExportToFile(path, "PNG", 1, 300, 0)
            return True
        except Exception:
            return False
