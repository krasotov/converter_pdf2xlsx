"""Окно: выбрать PDF и сохранить номенклатуру в Excel."""

from __future__ import annotations

import subprocess
import sys
import threading
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox

from PIL import Image, ImageTk

from convert import convert_pdf, write_xlsx

BG = "#f4f6f8"
INK = "#1e293b"
MUTED = "#334155"
NAVY = "#1f4e79"


def _money(value) -> str:
    text = f"{value:,.2f}"
    return text.replace(",", " ").replace(".", ",")


def _asset(name: str) -> Path:
    meipass = getattr(sys, "_MEIPASS", None)
    root = Path(meipass) if meipass else Path(__file__).resolve().parent
    return root / name


def _clickable(parent: tk.Misc, text: str, command, bg: str, fg: str, font: tuple) -> tk.Label:
    label = tk.Label(parent, text=text, bg=bg, fg=fg, font=font, padx=16, pady=8, cursor="hand2")
    label.bind("<Button-1>", lambda _event: command())
    return label


def _output_path(folder: Path, stem: str, used: set[Path]) -> Path:
    path = folder / f"Номенклатура_{stem}.xlsx"
    number = 2
    while path in used:
        path = folder / f"Номенклатура_{stem}_{number}.xlsx"
        number += 1
    used.add(path)
    return path


class App(tk.Tk):
    def __init__(self, initial: list[str] | None = None) -> None:
        super().__init__()
        self.title("PDF → Excel")
        self.geometry("580x460")
        self.minsize(540, 420)
        self.configure(bg=BG)
        self._busy = False
        self._last_file: Path | None = None
        self._apply_icon()

        footer = tk.Label(
            self,
            text="Разработано KT Digital",
            bg=BG,
            fg=NAVY,
            font=("Helvetica", 12, "underline"),
            cursor="hand2",
        )
        footer.pack(side="bottom", pady=(0, 14))
        footer.bind("<Button-1>", lambda _event: webbrowser.open("https://ktdigital.ru"))

        wrap = tk.Frame(self, bg=BG, padx=28, pady=22)
        wrap.pack(fill="both", expand=True)

        head = tk.Frame(wrap, bg=BG)
        head.pack(fill="x")
        if self._logo_img is not None:
            tk.Label(head, image=self._logo_img, bg=BG).pack(side="left", padx=(0, 14))
        tk.Label(
            head,
            text="PDF → Excel",
            bg=BG,
            fg=NAVY,
            font=("Helvetica", 22, "bold"),
            anchor="w",
        ).pack(side="left")
        tk.Label(
            wrap,
            text="Выберите один или несколько PDF. Программа перенесёт номенклатуру в Excel: код, название, количество, цену и сумму. Подходят товарные документы и УПД.",
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
            text="Можно выбрать несколько файлов или перетащить их на значок программы.",
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
            files = [Path(path) for path in initial]
            self.after(200, lambda: self.convert(files))

    def _apply_icon(self) -> None:
        self._logo_img = None
        png = _asset("icon.png")
        if png.is_file():
            image = Image.open(png).convert("RGBA")
            self._icon_img = ImageTk.PhotoImage(image.resize((128, 128), Image.Resampling.LANCZOS))
            self._logo_img = ImageTk.PhotoImage(image.resize((64, 64), Image.Resampling.LANCZOS))
            self.iconphoto(True, self._icon_img)
        ico = _asset("icon.ico")
        if sys.platform == "win32" and ico.is_file():
            try:
                self.iconbitmap(str(ico))
            except tk.TclError:
                pass

    def _set_open_enabled(self, enabled: bool) -> None:
        if enabled:
            self.open_btn.configure(fg=INK, cursor="hand2")
            self.open_btn.bind("<Button-1>", lambda _event: self.reveal())
        else:
            self.open_btn.configure(fg="#64748b", cursor="arrow")
            self.open_btn.unbind("<Button-1>")

    def _opened(self, *paths: str) -> None:
        files = [Path(path) for path in paths if str(path).lower().endswith(".pdf")]
        if files:
            self.convert(files)

    def set_status(self, text: str) -> None:
        self.status.configure(text=text, fg=INK)

    def set_progress(self, done: int, total: int) -> None:
        width = self.bar.winfo_width() or 500
        span = 0 if total <= 0 else int(width * min(done, total) / total)
        self.bar.coords(self._bar_fill, 0, 0, span, 8)

    def pick_pdf(self) -> None:
        if self._busy:
            return
        paths = filedialog.askopenfilenames(
            title="Выберите PDF",
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")],
        )
        if paths:
            self.convert([Path(path) for path in paths])

    def convert(self, pdfs: list[Path]) -> None:
        if self._busy:
            return
        missing = [pdf for pdf in pdfs if not pdf.exists()]
        if missing:
            messagebox.showerror("Файл не найден", "\n".join(str(pdf) for pdf in missing))
            return
        if len(pdfs) == 1:
            pdf = pdfs[0]
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
            targets = [(pdf, Path(dest))]
        else:
            folder = filedialog.askdirectory(
                title="Куда сохранить Excel",
                initialdir=str(pdfs[0].parent),
            )
            if not folder:
                return
            used: set[Path] = set()
            targets = [(pdf, _output_path(Path(folder), pdf.stem, used)) for pdf in pdfs]
        self._busy = True
        self.pick.configure(fg="#cbd5e1", cursor="arrow")
        self.pick.unbind("<Button-1>")
        self._set_open_enabled(False)
        self.set_status("Читаю документы… распознавание скана занимает около минуты на файл.")
        self.set_progress(0, 1)

        def work() -> None:
            done = []
            errors = []
            total_files = len(targets)
            for index, (pdf, dest) in enumerate(targets):
                def progress(page_done: int, page_total: int, message: str, index=index, pdf=pdf) -> None:
                    prefix = f"Файл {index + 1} из {total_files}: {pdf.name}. "
                    frac = index / total_files
                    if page_total:
                        frac += min(page_done, page_total) / page_total / total_files
                    self.after(0, lambda f=frac, m=prefix + message: self._on_progress_frac(f, m))

                try:
                    result = convert_pdf(pdf, progress=progress)
                    write_xlsx(result, dest)
                    done.append((result, dest))
                except Exception as exc:
                    errors.append((pdf, exc))
            self.after(0, lambda: self._on_batch_done(done, errors))

        threading.Thread(target=work, daemon=True).start()

    def _on_progress(self, done: int, total: int, message: str) -> None:
        self.set_status(message)
        self.set_progress(done, total)

    def _on_progress_frac(self, frac: float, message: str) -> None:
        self.set_status(message)
        width = self.bar.winfo_width() or 500
        self.bar.coords(self._bar_fill, 0, 0, int(width * max(0.0, min(frac, 1.0))), 8)

    def _unlock(self) -> None:
        self._busy = False
        self.pick.configure(fg="white", cursor="hand2")
        self.pick.bind("<Button-1>", lambda _event: self.pick_pdf())

    def _summary(self, result, path: Path) -> str:
        line = f"Готово: {len(result.items)} поз., сумма {_money(result.total)} ₽"
        if result.invoice_no:
            line = f"Документ № {result.invoice_no}. " + line
        if result.warnings:
            line += "\n" + " ".join(result.warnings)
        return line + f"\n{path}"

    def _on_batch_done(self, done: list, errors: list) -> None:
        self._unlock()
        self.set_progress(1, 1)
        if done:
            self._last_file = done[0][1]
            self._set_open_enabled(True)
        if len(done) == 1 and not errors:
            self.set_status(self._summary(done[0][0], done[0][1]))
            return
        lines = [f"Готово файлов: {len(done)}."]
        for result, path in done:
            lines.append(f"{path.name}: {len(result.items)} поз., {_money(result.total)} ₽")
        if errors:
            lines.append("Не прочитаны: " + ", ".join(pdf.name for pdf, _exc in errors))
        self.set_status("\n".join(lines))
        if errors:
            messagebox.showerror(
                "Не все файлы прочитаны",
                "\n".join(f"{pdf.name}: {exc}" for pdf, exc in errors),
            )

    def reveal(self) -> None:
        if not self._last_file or not self._last_file.exists():
            return
        path = str(self._last_file)
        if sys.platform == "win32":
            subprocess.run(["explorer", "/select,", path], check=False)
        else:
            subprocess.run(["open", "-R", path], check=False)


def main() -> None:
    initial = [a for a in sys.argv[1:] if a.lower().endswith(".pdf")]
    App(initial).mainloop()


if __name__ == "__main__":
    main()
