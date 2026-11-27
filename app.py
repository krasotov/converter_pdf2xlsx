"""Окно: выбрать PDF и сохранить номенклатуру в Excel."""

from __future__ import annotations

import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from convert import convert_pdf, write_xlsx

BG = "#f4f6f8"
INK = "#1e293b"
MUTED = "#334155"
NAVY = "#1f4e79"


def _money(value) -> str:
    text = f"{value:,.2f}"
    return text.replace(",", " ").replace(".", ",")


def _clickable(parent: tk.Misc, text: str, command, bg: str, fg: str, font: tuple) -> tk.Label:
    label = tk.Label(parent, text=text, bg=bg, fg=fg, font=font, padx=16, pady=8, cursor="hand2")
    label.bind("<Button-1>", lambda _event: command())
    return label


class App(tk.Tk):
    def __init__(self, initial: str | None = None) -> None:
        super().__init__()
        self.title("PDF → Excel")
        self.geometry("580x390")
        self.minsize(540, 360)
        self.configure(bg=BG)
        self._busy = False
        self._last_file: Path | None = None

        footer = tk.Label(
            self,
            text="Разработано KT Digital",
            bg=BG,
            fg=INK,
            font=("Helvetica", 12),
        )
        footer.pack(side="bottom", pady=(0, 14))

        wrap = tk.Frame(self, bg=BG, padx=28, pady=22)
        wrap.pack(fill="both", expand=True)

        tk.Label(
            wrap,
            text="PDF → Excel",
            bg=BG,
            fg=NAVY,
            font=("Helvetica", 22, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            wrap,
            text="Выберите PDF. Программа перенесёт номенклатуру в Excel: код, название, количество, цену и сумму. Подходят товарные документы и УПД.",
            bg=BG,
            fg=INK,
            font=("Helvetica", 13),
            wraplength=510,
            justify="left",
            anchor="w",
        ).pack(fill="x", pady=(8, 18))

        self.pick = _clickable(
            wrap,
            "Выбрать PDF",
            self.pick_pdf,
            NAVY,
            "white",
            ("Helvetica", 14, "bold"),
        )
        self.pick.pack(anchor="w")

        self.status = tk.Label(
            wrap,
            text="Можно перетащить PDF на значок программы.",
            bg=BG,
            fg=MUTED,
            font=("Helvetica", 12),
            wraplength=510,
            justify="left",
            anchor="w",
        )
        self.status.pack(fill="x", pady=(18, 8))

        self.bar = tk.Canvas(wrap, height=8, bg="#d5dde6", highlightthickness=0)
        self.bar.pack(fill="x")
        self._bar_fill = self.bar.create_rectangle(0, 0, 0, 8, fill=NAVY, width=0)

        self.open_btn = _clickable(
            wrap,
            "Показать в папке",
            self.reveal,
            "#d9e2ec",
            INK,
            ("Helvetica", 12),
        )
        self.open_btn.pack(anchor="w", pady=(16, 0))
        self._set_open_enabled(False)

        try:
            self.createcommand("::tk::mac::OpenDocument", self._opened)
        except tk.TclError:
            pass
        if initial:
            self.after(200, lambda: self.convert(Path(initial)))

    def _set_open_enabled(self, enabled: bool) -> None:
        if enabled:
            self.open_btn.configure(fg=INK, cursor="hand2")
            self.open_btn.bind("<Button-1>", lambda _event: self.reveal())
        else:
            self.open_btn.configure(fg="#64748b", cursor="arrow")
            self.open_btn.unbind("<Button-1>")

    def _opened(self, *paths: str) -> None:
        for path in paths:
            if str(path).lower().endswith(".pdf"):
                self.convert(Path(path))
                return

    def set_status(self, text: str) -> None:
        self.status.configure(text=text, fg=INK)

    def set_progress(self, done: int, total: int) -> None:
        width = self.bar.winfo_width() or 500
        span = 0 if total <= 0 else int(width * min(done, total) / total)
        self.bar.coords(self._bar_fill, 0, 0, span, 8)

    def pick_pdf(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="Выберите PDF",
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")],
        )
        if path:
            self.convert(Path(path))

    def convert(self, pdf: Path) -> None:
        if self._busy:
            return
        if not pdf.exists():
            messagebox.showerror("Файл не найден", str(pdf))
            return
        default = pdf.with_name(f"Номенклатура_{pdf.stem}.xlsx")
        dest = filedialog.asksaveasfilename(
            title="Куда сохранить Excel",
            defaultextension=".xlsx",
            initialfile=default.name,
            initialdir=str(pdf.parent),
            filetypes=[("Excel", "*.xlsx")],
        )
        if not dest:
            return
        self._busy = True
        self.pick.configure(fg="#cbd5e1", cursor="arrow")
        self.pick.unbind("<Button-1>")
        self._set_open_enabled(False)
        self.set_status("Читаю документ… распознавание скана занимает около минуты.")
        self.set_progress(0, 1)

        def work() -> None:
            try:
                def progress(done: int, total: int, message: str) -> None:
                    self.after(0, lambda: self._on_progress(done, total, message))

                result = convert_pdf(pdf, progress=progress)
                write_xlsx(result, dest)
                self.after(0, lambda: self._on_done(Path(dest), result))
            except Exception as exc:
                self.after(0, lambda: self._on_fail(exc))

        threading.Thread(target=work, daemon=True).start()

    def _on_progress(self, done: int, total: int, message: str) -> None:
        self.set_status(message)
        self.set_progress(done, total)

    def _on_done(self, path: Path, result) -> None:
        self._busy = False
        self.pick.configure(fg="white", cursor="hand2")
        self.pick.bind("<Button-1>", lambda _event: self.pick_pdf())
        self._last_file = path
        self._set_open_enabled(True)
        self.set_progress(1, 1)
        line = f"Готово: {len(result.items)} поз., сумма {_money(result.total)} ₽"
        if result.invoice_no:
            line = f"Документ № {result.invoice_no}. " + line
        if result.warnings:
            line += "\n" + " ".join(result.warnings)
        self.set_status(line + f"\n{path}")

    def _on_fail(self, exc: Exception) -> None:
        self._busy = False
        self.pick.configure(fg="white", cursor="hand2")
        self.pick.bind("<Button-1>", lambda _event: self.pick_pdf())
        self.set_status("Не получилось прочитать этот PDF.")
        messagebox.showerror("Не удалось конвертировать", str(exc))

    def reveal(self) -> None:
        if not self._last_file or not self._last_file.exists():
            return
        path = str(self._last_file)
        if sys.platform == "win32":
            subprocess.run(["explorer", "/select,", path], check=False)
        else:
            subprocess.run(["open", "-R", path], check=False)


def main() -> None:
    initial = next((a for a in sys.argv[1:] if a.lower().endswith(".pdf")), None)
    App(initial).mainloop()


if __name__ == "__main__":
    main()
