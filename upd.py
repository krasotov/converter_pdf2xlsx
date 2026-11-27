"""УПД с текстовым слоем → строки для Excel."""

from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal

import pymupdf

from convert import Item, Result, parse_number


@dataclass
class UpdRow:
    n: int
    code: str
    name: str
    unit_code: str
    unit: str
    qty: Decimal | None
    price: Decimal | None
    sum_wo: Decimal | None
    excise: str
    rate: str
    vat: Decimal | None
    sum_w: Decimal | None
    country_code: str
    country: str
    gtd: str


def is_upd(text: str) -> bool:
    folded = text.lower().replace("ё", "е")
    return "универсальный передаточный" in folded or "счет-фактура №" in folded


def _words(page: pymupdf.Page) -> list[dict]:
    width, height = page.rect.width, page.rect.height
    out = []
    for x0, y0, x1, y1, text, *_rest in page.get_text("words"):
        if not str(text).strip():
            continue
        out.append(
            {
                "t": str(text).strip(),
                "x": x0 / width,
                "y": y0 / height,
                "cx": ((x0 + x1) / 2) / width,
            }
        )
    return out


def _lines(words: list[dict]) -> list[list[dict]]:
    words = sorted(words, key=lambda w: (w["y"], w["x"]))
    lines: list[list[dict]] = []
    for word in words:
        if not lines or word["y"] - lines[-1][0]["y"] > 0.008:
            lines.append([word])
        else:
            lines[-1].append(word)
    return lines


def _cell(words: list[dict], x0: float, x1: float) -> str:
    picked = [w for w in words if x0 <= w["cx"] < x1]
    picked.sort(key=lambda w: w["x"])
    texts = [w["t"] for w in picked]
    if (
        len(texts) >= 2
        and re.fullmatch(r"\d{1,3}", texts[0])
        and re.fullmatch(r"\d{3}[.,]\d{2}", texts[1])
    ):
        return texts[0] + texts[1]
    return " ".join(texts).strip()


def _money(text: str) -> Decimal | None:
    if not text or text in {"-", "—", "Х", "X"}:
        return None
    return parse_number(text.replace(" ", ""))


def _meta(text: str) -> dict[str, str]:
    flat = re.sub(r"[ \t]+", " ", text)
    meta = {"no": "", "date": "", "seller": "", "buyer": "", "seller_inn": "", "buyer_inn": ""}
    m = re.search(r"Сч[её]т-фактура №\s*(\S+)\s+от\s+(\d{1,2}\s+\S+\s+\d{4})", flat, re.I)
    if m:
        meta["no"] = m.group(1)
        meta["date"] = re.sub(r"\s+", " ", m.group(2)).strip()
    m = re.search(r"Продавец\s+(.+?)\s+\(2\)", flat)
    if m:
        meta["seller"] = m.group(1).strip()
    m = re.search(r"Покупатель\s+(.+?)\s+\(6\)", flat)
    if m:
        meta["buyer"] = m.group(1).strip()
    m = re.search(r"ИНН/КПП продавца\s+(\S+)", flat)
    if m:
        meta["seller_inn"] = m.group(1)
    m = re.search(r"ИНН/КПП покупателя\s+(\S+)", flat)
    if m:
        meta["buyer_inn"] = m.group(1)
    return meta


def parse_upd(doc: pymupdf.Document) -> Result:
    rows: list[UpdRow] = []
    meta = {"no": "", "date": "", "seller": "", "buyer": "", "seller_inn": "", "buyer_inn": ""}
    stated_wo = stated_vat = stated_w = None
    for page in doc:
        text = page.get_text("text")
        page_meta = _meta(text)
        for key, value in page_meta.items():
            if value and not meta.get(key):
                meta[key] = value
        lines = _lines(_words(page))
        start = next((i for i, line in enumerate(lines) if any(w["t"] == "1а" for w in line)), None)
        end = next(
            (i for i, line in enumerate(lines) if any(w["t"] == "Всего" for w in line) and any(w["t"] == "оплате" for w in line)),
            None,
        )
        if start is None or end is None or end <= start:
            continue
        total_line = lines[end]
        stated_wo = _money(_cell(total_line, 0.50, 0.58)) or stated_wo
        stated_vat = _money(_cell(total_line, 0.64, 0.73)) or stated_vat
        stated_w = _money(_cell(total_line, 0.73, 0.82)) or stated_w

        blocks: list[list[list[dict]]] = []
        for line in lines[start + 1 : end]:
            if any(re.fullmatch(r"\d{5,}", w["t"]) and w["cx"] < 0.09 for w in line):
                blocks.append([line])
            elif blocks:
                blocks[-1].append(line)
        for block in blocks:
            first = block[0]
            name_bits = []
            for line in block:
                piece = _cell(line, 0.115, 0.30)
                if piece:
                    name_bits.append(piece)
            n_text = _cell(first, 0.09, 0.115)
            n = int(n_text) if n_text.isdigit() else len(rows) + 1
            rows.append(
                UpdRow(
                    n=n,
                    code=_cell(first, 0.0, 0.09),
                    name=re.sub(r"\s+", " ", " ".join(name_bits)).strip(),
                    unit_code=_cell(first, 0.31, 0.355),
                    unit=_cell(first, 0.355, 0.42),
                    qty=_money(_cell(first, 0.42, 0.46)),
                    price=_money(_cell(first, 0.46, 0.52)),
                    sum_wo=_money(_cell(first, 0.52, 0.568)),
                    excise=_cell(first, 0.568, 0.618) or "без акциза",
                    rate=_cell(first, 0.615, 0.66),
                    vat=_money(_cell(first, 0.66, 0.73)),
                    sum_w=_money(_cell(first, 0.73, 0.80)),
                    country_code=_cell(first, 0.80, 0.84),
                    country=_cell(first, 0.84, 0.90),
                    gtd=_cell(first, 0.90, 1.0),
                )
            )
    items = [
        Item(row.n, row.code, row.name, row.qty, row.unit, row.price, row.sum_w)
        for row in rows
    ]
    seller = meta["seller"]
    if meta["seller_inn"]:
        seller = f"{seller}, ИНН/КПП {meta['seller_inn']}".strip(", ")
    buyer = meta["buyer"]
    if meta["buyer_inn"]:
        buyer = f"{buyer}, ИНН {meta['buyer_inn']}".strip(", ")
    result = Result(
        items=items,
        supplier=seller,
        buyer=buyer,
        invoice_no=meta["no"],
        invoice_date=meta["date"],
        stated_total=stated_w,
        stated_vat=stated_vat,
    )
    result.kind = "upd"
    result.upd_rows = rows
    result.stated_wo = stated_wo
    if stated_w is not None and abs(result.total - stated_w) > Decimal("0.05"):
        result.warnings.append(
            f"Сумма строк {result.total} отличается от итога документа {stated_w}."
        )
    return result
