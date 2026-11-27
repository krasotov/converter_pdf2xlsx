"""Счёт на оплату с текстовым слоем → строки для Excel."""

from __future__ import annotations

import re
from decimal import Decimal

import pymupdf

from convert import Item, Result, parse_number

_LOOKALIKE = str.maketrans(
    {
        "A": "А",
        "B": "В",
        "C": "С",
        "E": "Е",
        "H": "Н",
        "K": "К",
        "M": "М",
        "O": "О",
        "P": "Р",
        "T": "Т",
        "X": "Х",
        "a": "а",
        "c": "с",
        "e": "е",
        "o": "о",
        "p": "р",
        "r": "г",
        "x": "х",
        "y": "у",
    }
)
_LOOK_LETTERS = set("ABCEHKMOPTXaceopxyr")


def is_bill(text: str) -> bool:
    folded = text.lower().replace("ё", "е")
    return "счет на оплату" in folded


def _words(page: pymupdf.Page) -> list[dict]:
    width, height = page.rect.width, page.rect.height
    out = []
    for x0, y0, x1, y1, text, *_rest in page.get_text("words"):
        token = str(text).strip()
        if not token:
            continue
        out.append(
            {
                "t": token,
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


def _join_money(words: list[dict]) -> str:
    texts = [w["t"] for w in sorted(words, key=lambda w: w["x"])]
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
    return " ".join(parts).strip()


def _header(lines: list[list[dict]]) -> dict[str, float] | None:
    for line in lines:
        tokens = {w["t"].lower().replace("ё", "е"): w["cx"] for w in line}
        qty = next((cx for name, cx in tokens.items() if name.startswith("кол")), None)
        price = tokens.get("цена")
        total = tokens.get("сумма")
        if qty is None or price is None or total is None:
            continue
        unit = next((cx for name, cx in tokens.items() if name.startswith("ед")), (qty + price) / 2)
        number = next((w["cx"] for w in line if w["t"] in {"№", "N", "No"}), 0.08)
        return {"n": number, "qty": qty, "unit": unit, "price": price, "sum": total}
    return None


def _in_band(words: list[dict], left: float, right: float) -> list[dict]:
    return [w for w in words if left <= w["cx"] < right]


def _meta(text: str) -> dict[str, str]:
    flat = re.sub(r"[ \t]+", " ", text)
    meta = {"no": "", "date": "", "seller": "", "buyer": ""}
    m = re.search(
        r"Сч[её]т на оплату №\s*(\S+)\s+от\s+(\d{1,2}\s+\S+\s+\d{4})",
        flat,
        re.I,
    )
    if m:
        meta["no"] = m.group(1)
        meta["date"] = re.sub(r"\s+", " ", m.group(2)).strip()
    m = re.search(
        r"Поставщик\s*(?:\(Исполнитель\):)?\s*(.+?)\s+Покупатель",
        flat,
        re.I | re.S,
    )
    if m:
        meta["seller"] = re.sub(r"\s+", " ", m.group(1)).strip(" :")
    m = re.search(
        r"Покупатель\s*(?:\(Заказчик\):)?\s*(.+?)\s+Основание",
        flat,
        re.I | re.S,
    )
    if m:
        meta["buyer"] = re.sub(r"\s+", " ", m.group(1)).strip(" :")
    return meta


def _qty_unit(text: str, unit_text: str) -> tuple[Decimal | None, str]:
    text = text.strip()
    match = re.match(r"([\d\s.,]+)\s+([^\d\s].*)$", text)
    if match:
        return parse_number(match.group(1)), unit_text or match.group(2).strip()
    return parse_number(text), unit_text


def _money_on_line(line: list[dict], left: float, right: float) -> Decimal | None:
    chunk = _join_money(_in_band(line, left, right))
    return parse_number(chunk) if chunk else None


def _repair_token(token: str) -> str:
    if any(ch.isdigit() for ch in token):
        return token
    letters = [ch for ch in token if ch.isalpha()]
    if not letters:
        return token
    has_cyrillic = any("а" <= ch.lower() <= "я" or ch.lower() == "ё" for ch in letters)
    if has_cyrillic or all(ch in _LOOK_LETTERS for ch in letters):
        return token.translate(_LOOKALIKE)
    return token


def _name_from(words: list[dict]) -> str:
    ordered = sorted(words, key=lambda w: (w["y"], w["x"]))
    return re.sub(r"\s+", " ", " ".join(_repair_token(w["t"]) for w in ordered)).strip()


def parse_bill(doc: pymupdf.Document) -> Result:
    rows: list[dict] = []
    meta = {"no": "", "date": "", "seller": "", "buyer": ""}
    stated_total = None
    stated_vat = None
    header: dict[str, float] | None = None
    for page in doc:
        text = page.get_text("text")
        page_meta = _meta(text)
        for key, value in page_meta.items():
            if value and not meta.get(key):
                meta[key] = value
        if stated_total is None:
            m = re.search(r"Всего к оплате:\s*([\d\s]+[.,]\d{2})", text)
            if m:
                stated_total = parse_number(m.group(1))
        if stated_vat is None:
            m = re.search(r"В том числе НДС[^:\n]*:\s*([\d\s]+[.,]\d{2})", text, re.I)
            if m:
                stated_vat = parse_number(m.group(1))
        words = _words(page)
        if not words:
            continue
        lines = _lines(words)
        found = _header(lines)
        if found is not None:
            header = found
        if header is None:
            continue
        number_right = header["n"] + 0.03
        name_left = header["n"] + 0.015
        name_right = header["qty"] - 0.02
        price_left = header["price"] - 0.02
        price_right = (header["price"] + header["sum"]) / 2
        page_rows: list[dict] = []
        current: dict | None = None
        for line in lines:
            if any(w["t"].rstrip(":").lower() == "итого" for w in line):
                if stated_total is None:
                    stated_total = _money_on_line(line, price_right, 1.0) or _money_on_line(line, price_left, 1.0)
                break
            number = next(
                (
                    int(w["t"])
                    for w in line
                    if w["t"].isdigit() and w["cx"] < number_right and int(w["t"]) < 1000
                ),
                None,
            )
            if number is not None:
                qty, unit = _qty_unit(_join_money(_in_band(line, name_right, price_left)), "")
                current = {
                    "n": number,
                    "name_words": _in_band(line, name_left, name_right),
                    "qty": qty,
                    "unit": unit,
                    "price": parse_number(_join_money(_in_band(line, price_left, price_right))),
                    "total": parse_number(_join_money(_in_band(line, price_right, 1.0))),
                }
                page_rows.append(current)
            elif current is not None:
                current["name_words"].extend(_in_band(line, name_left, name_right))
        for row in page_rows:
            row["name"] = _name_from(row["name_words"])
        rows.extend(page_rows)
    items = [
        Item(row["n"], "", row["name"], row["qty"], row["unit"], row["price"], row["total"])
        for row in rows
        if row["name"] or row["total"] is not None
    ]
    for index, item in enumerate(items, 1):
        item.n = index
    result = Result(
        items=items,
        supplier=meta["seller"],
        buyer=meta["buyer"],
        invoice_no=meta["no"],
        invoice_date=meta["date"],
        stated_total=stated_total,
        stated_vat=stated_vat,
    )
    if not items:
        result.warnings.append("Позиции не распознаны.")
    elif stated_total is not None and abs(result.total - stated_total) > Decimal("0.05"):
        result.warnings.append(
            f"Сумма строк {result.total} отличается от итога документа {stated_total}."
        )
    return result
