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


def _money(text: str) -> Decimal | None:
    if not text or text in {"-", "—", "Х", "X"}:
        return None
    return parse_number(text.replace(" ", ""))


_HEADERS = ["А", "1", "1а", "1б", "2", "2а", "3", "4", "5", "6", "7", "8", "9", "10", "10а", "11"]


def _meta(text: str) -> dict[str, str]:
    flat = re.sub(r"[ \t]+", " ", text)
    meta = {"no": "", "date": "", "seller": "", "buyer": "", "seller_inn": "", "buyer_inn": ""}
    m = re.search(r"Сч[её]т-фактура №:?\s*(\S+)\s+от\s+(\d{1,2}\s+\S+\s+\d{4})", flat, re.I)
    if m:
        meta["no"] = m.group(1)
        meta["date"] = re.sub(r"\s+", " ", m.group(2)).strip()
    m = re.search(r"Продавец:?\s+(.+?)\s+\(2\)", flat)
    if m:
        meta["seller"] = m.group(1).strip(" :")
    m = re.search(r"Покупатель:?\s+(.+?)\s+\(6\)", flat)
    if m:
        meta["buyer"] = m.group(1).strip(" :")
    m = re.search(r"ИНН/КПП продавца:?\s+(\S+)", flat)
    if m:
        meta["seller_inn"] = m.group(1)
    m = re.search(r"ИНН/КПП покупателя:?\s+(\S+)", flat)
    if m:
        meta["buyer_inn"] = m.group(1)
    return meta


def _is_code(text: str) -> bool:
    token = text.strip()
    if len(token) < 4 or not re.search(r"\d", token):
        return False
    return re.fullmatch(r"[\dA-Za-zА-Яа-яЁё][\w./-]{3,}", token) is not None


def _index_line(lines: list[list[dict]]) -> int | None:
    for i, line in enumerate(lines):
        tokens = {w["t"] for w in line}
        if {"1а", "2а", "11"} <= tokens:
            return i
    return None


def _header_pos(line: list[dict]) -> dict[str, float]:
    pos: dict[str, float] = {}
    for word in line:
        token = "А" if word["t"] == "A" else word["t"]
        if token in _HEADERS and token not in pos:
            pos[token] = word["cx"]
    return pos


def _bounds(pos: dict[str, float]) -> dict[str, tuple[float, float]]:
    keys = [key for key in _HEADERS if key in pos]
    bounds: dict[str, tuple[float, float]] = {}
    for i, key in enumerate(keys):
        left = 0.0 if i == 0 else (pos[keys[i - 1]] + pos[key]) / 2
        right = 1.0 if i + 1 == len(keys) else (pos[key] + pos[keys[i + 1]]) / 2
        bounds[key] = (left, right)
    return bounds


def _join_cell(words: list[dict]) -> str:
    ordered = sorted(words, key=lambda w: (w["y"], w["x"]))
    texts = [w["t"] for w in ordered]
    parts: list[str] = []
    i = 0
    while i < len(texts):
        if (
            i + 1 < len(texts)
            and re.fullmatch(r"\d{1,3}", texts[i])
            and re.fullmatch(r"\d{3}[.,]\d{2}", texts[i + 1])
        ):
            parts.append(texts[i] + texts[i + 1])
            i += 2
        else:
            parts.append(texts[i])
            i += 1
    text = " ".join(parts).strip()
    bits = text.split()
    if bits and all(re.fullmatch(r"[-—–]+", bit) for bit in bits):
        return "--"
    return text


def _column(words: list[dict], bounds: dict[str, tuple[float, float]], key: str) -> str:
    if key not in bounds:
        return ""
    left, right = bounds[key]
    return _join_cell([w for w in words if left <= w["cx"] < right])


def _row_from_block(block: list[list[dict]], layout: dict, rows: list[UpdRow]) -> UpdRow:
    bounds = layout["bounds"]
    pos = layout["pos"]
    words = [word for line in block for word in line]
    first = block[0]
    number = next(
        (
            int(w["t"])
            for w in first
            if w["t"].isdigit() and abs(w["cx"] - pos.get("1", 0.11)) < 0.03
        ),
        len(rows) + 1,
    )
    name = _join_cell(
        [
            w
            for w in words
            if layout["name_left"] <= w["cx"] < layout["name_right"]
            and not re.fullmatch(r"[-—–]+", w["t"])
        ]
    )
    return UpdRow(
        n=number,
        code=_column(first, bounds, "А"),
        name=re.sub(r"\s+", " ", re.sub(r"-\s+", "-", name)).strip(),
        unit_code=_column(words, bounds, "2"),
        unit=_column(words, bounds, "2а"),
        qty=_money(_column(words, bounds, "3")),
        price=_money(_column(words, bounds, "4")),
        sum_wo=_money(_column(words, bounds, "5")),
        excise=_column(words, bounds, "6") or "без акциза",
        rate=_column(words, bounds, "7"),
        vat=_money(_column(words, bounds, "8")),
        sum_w=_money(_column(words, bounds, "9")),
        country_code=_column(words, bounds, "10"),
        country=_column(words, bounds, "10а"),
        gtd=_column(words, bounds, "11"),
    )


def _is_page_chrome(line: list[dict]) -> bool:
    text = " ".join(w["t"] for w in line).lower().replace("ё", "е")
    if "счет-фактура" in text or "приложение" in text:
        return True
    if "универсальный" in text and "передаточ" in text:
        return True
    return text.startswith("продавец") or text.startswith("покупатель")


def _absorb_line(
    line: list[dict],
    layout: dict,
    pending: list[list[dict]] | None,
    rows: list[UpdRow],
) -> list[list[dict]] | None:
    code_right = layout["bounds"].get("А", (0.0, 0.1))[1]
    if any(_is_code(w["t"]) and w["cx"] < code_right for w in line):
        if pending:
            rows.append(_row_from_block(pending, layout, rows))
        return [line]
    if pending is not None and not _is_page_chrome(line):
        pending.append(line)
    return pending


def parse_upd(doc: pymupdf.Document) -> Result:
    rows: list[UpdRow] = []
    meta = {"no": "", "date": "", "seller": "", "buyer": "", "seller_inn": "", "buyer_inn": ""}
    stated_wo = stated_vat = stated_w = None
    layout: dict | None = None
    pending: list[list[dict]] | None = None
    for page in doc:
        text = page.get_text("text")
        page_meta = _meta(text)
        for key, value in page_meta.items():
            if value and not meta.get(key):
                meta[key] = value
        lines = _lines(_words(page))
        start = _index_line(lines)
        end = next(
            (
                i
                for i, line in enumerate(lines)
                if any(w["t"] == "Всего" for w in line) and any(w["t"] == "оплате" for w in line)
            ),
            None,
        )
        if start is not None:
            pos = _header_pos(lines[start])
            layout = {
                "pos": pos,
                "bounds": _bounds(pos),
                "name_right": (pos.get("1б", 0.30) + pos.get("2", 0.33)) / 2,
                "name_left": pos.get("1", 0.1) + 0.012,
            }
            if pending is not None:
                for line in lines[:start]:
                    pending = _absorb_line(line, layout, pending, rows)
            data_from = start + 1
        elif layout is None:
            continue
        else:
            data_from = 0
        data_to = end if end is not None else len(lines)
        for line in lines[data_from:data_to]:
            pending = _absorb_line(line, layout, pending, rows)
        if end is not None:
            if pending:
                rows.append(_row_from_block(pending, layout, rows))
                pending = None
            total_words = lines[end]
            bounds = layout["bounds"]
            stated_wo = _money(_column(total_words, bounds, "5")) or stated_wo
            stated_vat = _money(_column(total_words, bounds, "8")) or stated_vat
            stated_w = _money(_column(total_words, bounds, "9")) or stated_w
    if pending and layout is not None:
        rows.append(_row_from_block(pending, layout, rows))
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
    if not rows:
        result.warnings.append("Позиции не распознаны.")
    elif stated_w is not None and abs(result.total - stated_w) > Decimal("0.05"):
        result.warnings.append(
            f"Сумма строк {result.total} отличается от итога документа {stated_w}."
        )
    return result
