"""PDF → таблица номенклатуры в Excel.

Берёт счета с колонками «Артикул», «Товары», «Кол-во», «Ед.», «Цена», «Сумма».
Обычный PDF читается своим текстом, скан — распознаванием macOS или Tesseract.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pymupdf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from PIL import Image

QTY_RE = re.compile(r"^(\d+(?:[.,]\d+)?)\s*([^\d\s]+)?$")
HEADER_SKIP = {
    "№",
    "nº",
    "no",
    "no.",
    "артикул",
    "товары",
    "товары (работы, услуги)",
    "наименование",
    "кол-во",
    "количество",
    "ед",
    "ед.",
    "цена",
    "сумма",
}


@dataclass
class Tok:
    text: str
    x: float
    y: float
    r: float
    b: float
    conf: float = 1.0

    @property
    def cx(self) -> float:
        return (self.x + self.r) / 2


@dataclass
class Item:
    n: int
    article: str
    name: str
    qty: Decimal | None
    unit: str
    price: Decimal | None
    total: Decimal | None


@dataclass
class Result:
    items: list[Item]
    supplier: str = ""
    buyer: str = ""
    invoice_no: str = ""
    invoice_date: str = ""
    stated_total: Decimal | None = None
    stated_vat: Decimal | None = None
    warnings: list[str] = field(default_factory=list)
    kind: str = "items"
    upd_rows: list = field(default_factory=list)
    stated_wo: Decimal | None = None

    @property
    def total(self) -> Decimal:
        return sum((it.total or Decimal(0) for it in self.items), Decimal("0"))


def parse_number(text: str) -> Decimal | None:
    raw = text.replace("\u00a0", " ").replace(" ", "")
    raw = re.sub(r"[^\d,.\-]", "", raw)
    if not raw or not re.search(r"\d", raw):
        return None
    if "," in raw and "." in raw:
        if raw.rfind(",") > raw.rfind("."):
            raw = raw.replace(".", "").replace(",", ".")
        else:
            raw = raw.replace(",", "")
    elif "," in raw:
        raw = raw.replace(",", ".")
    elif raw.count(".") > 1:
        raw = raw.replace(".", "")
    try:
        return Decimal(raw)
    except InvalidOperation:
        return None


def _frac_digits(text: str) -> int:
    m = re.search(r"[,.](\d+)", re.sub(r"[^\d,.]", "", text))
    return len(m.group(1)) if m else 0


def is_money(text: str) -> bool:
    return parse_number(text) is not None and bool(re.search(r"[,.]\d", text))


def norm_unit(unit: str) -> str:
    u = unit.lower().strip(" .")
    u = {"m": "м", "м2": "м2", "m2": "м2", "шт": "шт", "уп": "уп", "кг": "кг", "л": "л"}.get(u, u)
    return u


def parse_qty(text: str) -> tuple[Decimal | None, str]:
    raw = (
        text.replace("З", "3")
        .replace("•", "")
        .replace("М", "м")
        .replace("M", "м")
        .replace("Шт", "шт")
        .replace("м²", "м2")
        .strip()
    )
    m = QTY_RE.match(raw)
    if not m:
        return None, ""
    try:
        qty = Decimal(m.group(1).replace(",", "."))
    except InvalidOperation:
        return None, ""
    return qty, norm_unit(m.group(2) or "")


def _vision_ocr(image: Image.Image) -> list[Tok]:
    from ocrmac import ocrmac

    raw = ocrmac.OCR(
        image,
        language_preference=["ru-RU", "en-US"],
        recognition_level="accurate",
    ).recognize()
    out = []
    for text, conf, box in raw:
        x, y, w, h = box
        if not str(text).strip():
            continue
        out.append(
            Tok(str(text).strip(), float(x), float(1 - y - h), float(x + w), float(1 - y), float(conf))
        )
    return out


def _tesseract_exe() -> str | None:
    roots = [Path(sys.executable).resolve().parent]
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        roots.append(Path(meipass))
    candidates = [shutil.which("tesseract") or ""]
    for root in roots:
        candidates.append(str(root / "tesseract" / "tesseract.exe"))
        candidates.append(str(root / "Tesseract-OCR" / "tesseract.exe"))
    candidates.append(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return candidate
    return None


def _tesseract_ocr(image: Image.Image) -> list[Tok]:
    exe = _tesseract_exe()
    if not exe:
        raise RuntimeError("Не удалось прочитать скан: нет программы распознавания текста.")
    env = os.environ.copy()
    tessdata = Path(exe).resolve().parent / "tessdata"
    if tessdata.is_dir():
        env["TESSDATA_PREFIX"] = str(tessdata) + os.sep
    with tempfile.TemporaryDirectory() as tmp:
        img = Path(tmp) / "page.png"
        image.save(img)
        out = Path(tmp) / "out"
        subprocess.run(
            [exe, str(img), str(out), "-l", "rus+eng", "--psm", "6", "tsv"],
            check=True,
            capture_output=True,
            env=env,
        )
        toks = []
        w, h = image.size
        for line in out.with_suffix(".tsv").read_text(encoding="utf-8", errors="replace").splitlines()[1:]:
            parts = line.split("\t")
            if len(parts) < 12 or not parts[11].strip():
                continue
            try:
                conf = float(parts[10])
            except ValueError:
                conf = 0
            if conf < 20:
                continue
            left, top, ww, hh = (int(float(parts[i])) for i in (6, 7, 8, 9))
            toks.append(
                Tok(parts[11].strip(), left / w, top / h, (left + ww) / w, (top + hh) / h, conf / 100)
            )
        return toks


def ocr_image(image: Image.Image) -> list[Tok]:
    try:
        return _vision_ocr(image)
    except Exception:
        return _tesseract_ocr(image)


def text_layer_tokens(page: pymupdf.Page) -> list[Tok]:
    rect = page.rect
    toks = []
    for x0, y0, x1, y1, word, *_rest in page.get_text("words"):
        if str(word).strip():
            toks.append(Tok(str(word).strip(), x0 / rect.width, y0 / rect.height, x1 / rect.width, y1 / rect.height))
    return toks


def text_layer_is_useful(toks: list[Tok]) -> bool:
    labels = [t for t in toks if t.text.lower().rstrip(".") in {"сумма", "цена", "артикул"}]
    return any(t.x > 0.45 for t in labels)


def _is_footer_line(text: str) -> bool:
    return bool(re.match(r"Итого|Всего к оплате|Всего наименований|В том числе НДС", text))


def _continuation(text: str) -> bool:
    t = text.strip()
    if not t:
        return True
    if t[:1].islower() or t[:1] in {")", ",", "»", "-"}:
        return True
    if re.match(r"^(стекло|цвет |бухта |армир|никел|мм\b|шт[.)]|с уплот|с проклад|\(230)", t, re.I):
        return True
    return False


def _article_of(text: str) -> tuple[str, str]:
    text = text.strip(" |")
    m = re.match(r"^([A-Z0-9][A-Z0-9./\-]{2,})([А-Яа-яЁё].*)$", text)
    if m:
        return m.group(1), m.group(2).strip()
    parts = text.split(maxsplit=1)
    head = parts[0]
    if re.match(r"^[A-Z0-9][A-Z0-9./\-]{2,}$", head) and not re.search(r"[А-Яа-яЁё]", head):
        return head, parts[1].strip() if len(parts) > 1 else ""
    m = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{1,2})\.?\s*(.*)$", text)
    if m:
        return m.group(1), m.group(2).strip()
    return "", text


def _clean_name(name: str) -> str:
    name = re.sub(r"[|¦ỊịłŁšŚźîạồ]+", "", name)
    name = re.sub(r"\s+", " ", name).strip(" -–,;|")
    name = name.replace("армированние", "армирование")
    name = re.sub(r"\bпл\b", "пп", name)
    return name


def _money_rows(toks: list[Tok]) -> tuple[list[dict], Decimal | None, Decimal | None]:
    header_y = 0.0
    footer_y = 0.995
    stated_total = None
    stated_vat = None
    for t in toks:
        low = t.text.lower().rstrip(":")
        if low in {"сумма", "цена"} and t.x > 0.7:
            header_y = max(header_y, t.y)
        if _is_footer_line(t.text):
            footer_y = min(footer_y, t.y)
            if t.text.lower().startswith("итого"):
                # Сумма часто стоит отдельным фрагментом справа от слова «Итого».
                pass
    # Итог и НДС — числа правее подписи, на одной высоте.
    for t in toks:
        if not is_money(t.text) or t.cx < 0.85:
            continue
        near = [u for u in toks if abs(u.y - t.y) < 0.012]
        label = " ".join(u.text for u in near)
        if re.search(r"Итого", label) and stated_total is None:
            stated_total = parse_number(t.text)
        elif re.search(r"НДС", label) and stated_vat is None:
            stated_vat = parse_number(t.text)

    region = [t for t in toks if t.y > header_y + 0.006 and t.y < footer_y - 0.004]
    sums = [
        t
        for t in region
        if t.cx > 0.90 and is_money(t.text) and (parse_number(t.text) or 0) < Decimal("100000000")
    ]
    sums.sort(key=lambda t: t.y)
    merged: list[Tok] = []
    for t in sums:
        if merged and abs(t.y - merged[-1].y) < 0.006:
            if len(re.sub(r"\D", "", t.text)) > len(re.sub(r"\D", "", merged[-1].text)):
                merged[-1] = t
        else:
            merged.append(t)
    sums = merged
    if not sums:
        return [], stated_total, stated_vat

    buckets: list[dict] = [{"sum": s, "bits": []} for s in sums]

    def bucket_index(y: float) -> int:
        best = 0
        for i, s in enumerate(sums):
            if s.y <= y + 0.004:
                best = i
            else:
                break
        return best

    for t in sorted(region, key=lambda z: (z.y, z.x)):
        if t in sums or t.cx > 0.90:
            continue
        low = t.text.lower().strip(" :.")
        if low in HEADER_SKIP:
            continue
        i = bucket_index(t.y)
        if _continuation(t.text) and i > 0 and not buckets[i]["bits"]:
            buckets[i - 1]["bits"].append(t)
        else:
            buckets[i]["bits"].append(t)

    rows = []
    for b in buckets:
        arts: list[str] = []
        names: list[str] = []
        qty = price = None
        price_text = ""
        unit = ""
        for t in b["bits"]:
            if 0.78 < t.cx <= 0.905 and is_money(t.text):
                price = parse_number(t.text)
                price_text = t.text
                continue
            if 0.66 < t.cx < 0.78:
                q, u = parse_qty(t.text)
                if q is not None:
                    qty, unit = q, u
                    continue
            if t.cx < 0.16 and re.fullmatch(r"\d{1,3}", t.text):
                continue
            art, rest = _article_of(t.text)
            if art:
                arts.append(art.rstrip("."))
            if rest:
                names.append(rest)
        total = parse_number(b["sum"].text)
        sum_digits = _frac_digits(b["sum"].text)
        price_digits = _frac_digits(price_text)
        if qty is not None and price is not None and price_digits >= 2:
            calc = (qty * price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if total is None or sum_digits < 2 or abs(calc - total) <= Decimal("0.05"):
                total = calc
        if qty is not None and total is not None and qty != 0:
            implied = (total / qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if abs((implied * qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP) - total) <= Decimal("0.05"):
                price = implied
                total = (price * qty).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        rows.append(
            {
                "article": arts[0] if arts else "",
                "name": _clean_name(" ".join(names)),
                "qty": qty,
                "unit": unit,
                "price": price,
                "total": total,
            }
        )
    return rows, stated_total, stated_vat


def _meta(toks: list[Tok]) -> dict[str, str]:
    blob = " ".join(t.text for t in sorted(toks, key=lambda t: (round(t.y, 3), t.x)))
    meta = {"supplier": "", "buyer": "", "invoice_no": "", "invoice_date": ""}
    m = re.search(r"Сч[её]т\s+на\s+оплату\s*[№NºNo.#]*\s*(\d+)", blob, re.I)
    if m:
        meta["invoice_no"] = m.group(1)
    months = {
        "января": "01",
        "февраля": "02",
        "марта": "03",
        "апреля": "04",
        "мая": "05",
        "июня": "06",
        "июля": "07",
        "августа": "08",
        "сентября": "09",
        "октября": "10",
        "ноября": "11",
        "декабря": "12",
    }
    m = re.search(
        r"(\d{1,2})\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s+(\d{4})",
        blob,
        re.I,
    )
    if m:
        meta["invoice_date"] = f"{int(m.group(1)):02d}.{months[m.group(2).lower()]}.{m.group(3)}"
    m = re.search(r"Поставщик:\s*(.+?)\s+Покупатель:", blob)
    if m:
        meta["supplier"] = re.sub(r"\s+", " ", m.group(1)).strip(" ,")
    m = re.search(r"Покупатель:\s*(.+?)\s+Грузополучатель", blob)
    if m:
        meta["buyer"] = re.sub(r"\s+", " ", m.group(1)).strip(" ,")
    return meta


def convert_pdf(path: str | Path, progress=None) -> Result:
    path = Path(path)
    doc = pymupdf.open(path)
    sample = "\n".join(page.get_text("text") for page in doc)
    folded = sample.lower().replace("ё", "е")
    if "универсальный" in folded and "передаточн" in folded:
        from upd import parse_upd

        try:
            if progress:
                progress(0, 1, "Читаю УПД")
            return parse_upd(doc)
        finally:
            doc.close()
    if "счет на оплату" in folded:
        from bill import parse_bill

        try:
            if progress:
                progress(0, 1, "Читаю документ")
            return parse_bill(doc)
        finally:
            doc.close()
    rows: list[dict] = []
    meta: dict[str, str] = {}
    stated_total = None
    stated_vat = None
    warnings: list[str] = []
    n_pages = doc.page_count
    try:
        for i, page in enumerate(doc):
            if progress:
                progress(i, n_pages, f"Страница {i + 1} из {n_pages}")
            layer = text_layer_tokens(page)
            if text_layer_is_useful(layer):
                toks = layer
            else:
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                toks = ocr_image(image)
            if i == 0:
                meta = _meta(toks)
            page_rows, total_i, vat_i = _money_rows(toks)
            if not page_rows and i == 0:
                warnings.append("В файле не нашлась таблица с колонкой «Сумма».")
            if total_i is not None:
                stated_total = total_i
            if vat_i is not None:
                stated_vat = vat_i
            rows.extend(page_rows)
    finally:
        doc.close()
    if progress:
        progress(n_pages, n_pages, "Сборка таблицы")

    items = []
    for i, row in enumerate(rows, 1):
        if not row["name"] and row["total"] is None:
            continue
        items.append(
            Item(i, row["article"], row["name"], row["qty"], row["unit"], row["price"], row["total"])
        )
    for i, it in enumerate(items, 1):
        it.n = i
    result = Result(
        items,
        meta.get("supplier", ""),
        meta.get("buyer", ""),
        meta.get("invoice_no", ""),
        meta.get("invoice_date", ""),
        stated_total,
        stated_vat,
        warnings,
    )
    if not items:
        result.warnings.append("Позиции не распознаны.")
    elif stated_total is not None and abs(result.total - stated_total) > Decimal("1"):
        result.warnings.append(
            f"Сумма строк {result.total} не сходится с итогом документа {stated_total}. Проверьте файл."
        )
    return result


def write_xlsx(result: Result, dest: str | Path) -> Path:
    if result.kind == "upd":
        return _write_upd(result, dest)
    dest = Path(dest)
    wb = Workbook()
    ws = wb.active
    ws.title = "Номенклатура"
    navy = "1F4E79"
    title = Font(name="Calibri", bold=True, size=16, color=navy)
    sub = Font(name="Calibri", size=11, color="333333")
    head_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    cell = Font(name="Calibri", size=10)
    bold = Font(name="Calibri", bold=True, size=11)
    head_fill = PatternFill("solid", fgColor=navy)
    alt = PatternFill("solid", fgColor="F3F6F9")
    total_fill = PatternFill("solid", fgColor="E7EEF5")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    bits = ["Номенклатура"]
    if result.invoice_no:
        bits.append(f"№ {result.invoice_no}")
    if result.invoice_date:
        bits.append(f"от {result.invoice_date}")
    ws.merge_cells("A1:G1")
    ws["A1"] = " ".join(bits)
    ws["A1"].font = title
    ws.row_dimensions[1].height = 24

    info = []
    if result.supplier:
        info.append(f"Поставщик: {result.supplier}")
    if result.buyer:
        info.append(f"Покупатель: {result.buyer}")
    info.append(f"Позиций: {len(result.items)}")
    ws.merge_cells("A2:G2")
    ws["A2"] = "     ".join(info)
    ws["A2"].font = sub
    ws.row_dimensions[2].height = 18

    for col, h in enumerate(["№", "Артикул", "Наименование", "Кол-во", "Ед.", "Цена, руб.", "Сумма, руб."], 1):
        c = ws.cell(4, col, h)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = thin
    ws.row_dimensions[4].height = 22
    ws.freeze_panes = "A5"
    last = 4 + max(len(result.items), 1)
    ws.auto_filter.ref = f"A4:G{last}"

    for i, it in enumerate(result.items):
        r = 5 + i
        vals = [
            it.n,
            it.article,
            it.name,
            float(it.qty) if it.qty is not None else None,
            it.unit,
            float(it.price) if it.price is not None else None,
            float(it.total) if it.total is not None else None,
        ]
        for col, v in enumerate(vals, 1):
            c = ws.cell(r, col, v)
            c.font = cell
            c.border = thin
            c.alignment = Alignment(wrap_text=True, vertical="center") if col == 3 else center
            if i % 2:
                c.fill = alt
        ws.cell(r, 6).number_format = "#,##0.00"
        ws.cell(r, 7).number_format = "#,##0.00"
        ws.cell(r, 6).alignment = right
        ws.cell(r, 7).alignment = right
        ws.row_dimensions[r].height = 30

    tr = 5 + len(result.items)
    ws.merge_cells(start_row=tr, start_column=1, end_row=tr, end_column=5)
    for col in range(1, 8):
        ws.cell(tr, col).fill = total_fill
        ws.cell(tr, col).border = thin
        ws.cell(tr, col).font = bold
    ws.cell(tr, 1, "Итого").alignment = Alignment(horizontal="right", vertical="center")
    ws.cell(tr, 7, f"=SUM(G5:G{tr - 1})" if result.items else 0)
    ws.cell(tr, 7).number_format = "#,##0.00"
    if result.stated_vat is not None:
        ws.cell(tr + 1, 6, "В том числе НДС").font = sub
        ws.cell(tr + 1, 6).alignment = right
        ws.cell(tr + 1, 7, float(result.stated_vat)).number_format = "#,##0.00"
    ws.cell(tr + 2, 6, "Всего к оплате").font = bold
    ws.cell(tr + 2, 6).alignment = right
    ws.cell(tr + 2, 7, f"=G{tr}")
    ws.cell(tr + 2, 7).number_format = "#,##0.00"
    ws.cell(tr + 2, 7).font = bold

    for i, w in enumerate([6, 22, 78, 10, 8, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5, 0.2, 0.2)
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:4"
    wb.save(dest)
    return dest


def _write_upd(result: Result, dest: str | Path) -> Path:
    dest = Path(dest)
    wb = Workbook()
    ws = wb.active
    ws.title = "УПД"
    navy = "1F4E79"
    title = Font(name="Calibri", bold=True, size=16, color=navy)
    sub = Font(name="Calibri", size=11, color="333333")
    head_font = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
    cell = Font(name="Calibri", size=10)
    bold = Font(name="Calibri", bold=True, size=11)
    head_fill = PatternFill("solid", fgColor=navy)
    alt = PatternFill("solid", fgColor="F3F6F9")
    total_fill = PatternFill("solid", fgColor="E7EEF5")
    thin = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")
    bits = ["УПД"]
    if result.invoice_no:
        bits.append(f"№ {result.invoice_no}")
    if result.invoice_date:
        bits.append(f"от {result.invoice_date}")
    last_col = 15
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws["A1"] = " ".join(bits)
    ws["A1"].font = title
    ws.row_dimensions[1].height = 24
    info = []
    if result.supplier:
        info.append(f"Продавец: {result.supplier}")
    if result.buyer:
        info.append(f"Покупатель: {result.buyer}")
    info.append(f"Позиций: {len(result.upd_rows)}")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    ws["A2"] = "     ".join(info)
    ws["A2"].font = sub
    headers = [
        "№",
        "Код",
        "Наименование",
        "Ед. код",
        "Ед.",
        "Кол-во",
        "Цена",
        "Сумма без НДС",
        "Акциз",
        "Ставка",
        "НДС",
        "Сумма с НДС",
        "Страна код",
        "Страна",
        "ГТД",
    ]
    for col, header in enumerate(headers, 1):
        c = ws.cell(4, col, header)
        c.font = head_font
        c.fill = head_fill
        c.alignment = center
        c.border = thin
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:O{4 + max(len(result.upd_rows), 1)}"
    money_cols = {6, 7, 8, 11, 12}
    for i, row in enumerate(result.upd_rows):
        r = 5 + i
        vals = [
            row.n,
            row.code,
            row.name,
            row.unit_code,
            row.unit,
            float(row.qty) if row.qty is not None else None,
            float(row.price) if row.price is not None else None,
            float(row.sum_wo) if row.sum_wo is not None else None,
            row.excise,
            row.rate,
            float(row.vat) if row.vat is not None else None,
            float(row.sum_w) if row.sum_w is not None else None,
            row.country_code,
            row.country,
            row.gtd,
        ]
        for col, value in enumerate(vals, 1):
            c = ws.cell(r, col, value)
            c.font = cell
            c.border = thin
            c.alignment = Alignment(wrap_text=True, vertical="center") if col == 3 else center
            if col in money_cols:
                c.number_format = "#,##0.00"
                c.alignment = right
            if i % 2:
                c.fill = alt
        ws.row_dimensions[r].height = 32
    tr = 5 + len(result.upd_rows)
    for col in range(1, last_col + 1):
        ws.cell(tr, col).fill = total_fill
        ws.cell(tr, col).border = thin
        ws.cell(tr, col).font = bold
    ws.cell(tr, 1, "Всего к оплате")
    if result.stated_wo is not None:
        ws.cell(tr, 8, float(result.stated_wo)).number_format = "#,##0.00"
    if result.stated_vat is not None:
        ws.cell(tr, 11, float(result.stated_vat)).number_format = "#,##0.00"
    if result.stated_total is not None:
        ws.cell(tr, 12, float(result.stated_total)).number_format = "#,##0.00"
    for col, width in enumerate([6, 14, 46, 10, 8, 10, 12, 16, 14, 10, 12, 16, 12, 14, 18], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(0.4, 0.4, 0.5, 0.5, 0.2, 0.2)
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:4"
    wb.save(dest)
    return dest
