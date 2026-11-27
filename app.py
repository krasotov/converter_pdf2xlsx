"""Окно: выбрать PDF и сохранить номенклатуру в Excel."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import urllib.request
import webbrowser
import zipfile
from pathlib import Path
from tkinter import filedialog, messagebox

from PIL import Image, ImageTk

from convert import convert_pdf, write_xlsx
from version import APP_VERSION

REPO = "krasotov/converter_pdf2xlsx"

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


def _version_key(value: str) -> tuple[int, ...]:
    numbers = []
    for piece in value.lstrip("vV").split("."):
        digits = ""
        for ch in piece:
            if ch.isdigit():
                digits += ch
            else:
                break
        numbers.append(int(digits or "0"))
    return tuple(numbers)


def _ps_quote(path: Path) -> str:
    return str(path).replace("'", "''")


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

        bottom = tk.Frame(self, bg=BG)
        bottom.pack(side="bottom", fill="x", padx=28, pady=(0, 14))
        self.update_btn = tk.Label(
            bottom,
            text="Обновить",
            bg=BG,
            fg=NAVY,
            font=("Helvetica", 12, "underline"),
            cursor="hand2",
        )
        self.update_btn.pack(side="left")
        self.update_btn.bind("<Button-1>", lambda _event: self.check_update())
        footer = tk.Label(
            bottom,
            text="Разработано KT Digital",
            bg=BG,
            fg=NAVY,
            font=("Helvetica", 12, "underline"),
            cursor="hand2",
        )
        footer.pack(side="right")
        footer.bind("<Button-1>", lambda _event: webbrowser.open("https://ktdigital.ru"))
        self.after(800, self._peek_update)

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

    def _peek_update(self) -> None:
        threading.Thread(target=lambda: self._fetch_release(install=False), daemon=True).start()

    def check_update(self) -> None:
        if self._busy:
            return
        self.set_status("Проверяю обновление…")
        threading.Thread(target=lambda: self._fetch_release(install=True), daemon=True).start()

    def _fetch_release(self, install: bool) -> None:
        try:
            request = urllib.request.Request(
                f"https://api.github.com/repos/{REPO}/releases/latest",
                headers={"User-Agent": "PDF2Excel", "Accept": "application/vnd.github+json"},
            )
            with urllib.request.urlopen(request, timeout=20) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except Exception:
            if install:
                self.after(0, lambda: self.set_status("Не удалось проверить обновление."))
            return
        remote = str(payload.get("tag_name", "")).lstrip("vV")
        if not remote or _version_key(remote) <= _version_key(APP_VERSION):
            if install:
                self.after(
                    0,
                    lambda: self.set_status(f"Установлена последняя версия {APP_VERSION}."),
                )
            return
        asset = next(
            (item for item in payload.get("assets", []) if item.get("name") == "PDF2Excel.zip"),
            None,
        )
        page = payload.get("html_url") or f"https://github.com/{REPO}/releases/latest"

        def offer() -> None:
            if not install:
                self.update_btn.configure(text=f"Обновить до {remote}")
                if not self._busy:
                    self.set_status(f"Доступна версия {remote}. Нажмите «Обновить до {remote}».")
                return
            if sys.platform == "win32" and asset:
                self._download_update(asset["browser_download_url"], remote)
            else:
                webbrowser.open(page)
                self.set_status(f"Доступна версия {remote}. Открыта страница загрузки.")

        self.after(0, offer)

    def _download_update(self, url: str, remote: str) -> None:
        self.set_status(f"Скачиваю версию {remote}…")

        def work() -> None:
            try:
                archive = Path(tempfile.gettempdir()) / "PDF2Excel-update.zip"
                urllib.request.urlretrieve(url, archive)
                folder = Path(tempfile.mkdtemp(prefix="pdf2excel-"))
                with zipfile.ZipFile(archive) as bundle:
                    bundle.extractall(folder)
                dest = Path(sys.executable).resolve().parent
                script = folder / "update.ps1"
                script.write_text(
                    "\n".join(
                        [
                            "$ErrorActionPreference = 'SilentlyContinue'",
                            f"$src = '{_ps_quote(folder)}'",
                            f"$dst = '{_ps_quote(dest)}'",
                            "$exe = Join-Path $dst 'PDF2Excel.exe'",
                            "for ($i = 0; $i -lt 40; $i++) {",
                            "    try {",
                            "        Copy-Item -LiteralPath (Join-Path $src 'PDF2Excel.exe') -Destination $exe -Force -ErrorAction Stop",
                            "        break",
                            "    } catch { Start-Sleep -Seconds 1 }",
                            "}",
                            "Copy-Item -LiteralPath (Join-Path $src '*') -Destination $dst -Recurse -Force",
                            "Start-Process -FilePath $exe",
                        ]
                    ),
                    encoding="utf-8",
                )
                subprocess.Popen(
                    [
                        "powershell",
                        "-NoProfile",
                        "-WindowStyle",
                        "Hidden",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        str(script),
                    ]
                )
                self.after(0, self.destroy)
            except Exception as exc:
                self.after(0, lambda: self.set_status(f"Не удалось скачать обновление. {exc}"))

        threading.Thread(target=work, daemon=True).start()

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
