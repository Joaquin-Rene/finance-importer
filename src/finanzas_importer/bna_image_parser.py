from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from hashlib import sha1
from io import BytesIO
from pathlib import Path
import re
import shutil

import pandas as pd
from PIL import Image, ImageFilter, ImageOps

from .mp_parser import normalize_description
from .utils import parse_number_ar

try:
    import pytesseract
except Exception:  # pragma: no cover - runtime optional dependency
    pytesseract = None


@dataclass
class BnaParseResult:
    parsed_df: pd.DataFrame
    raw_text: str
    extracted_rows: int
    warnings: list[str]


_DATE_AMOUNT_PATTERN = re.compile(
    r"(?P<prefix>.*?)(?P<date>\d{1,2}/\d{1,2})(?P<middle>.*?)(?P<amount>[+-]?\s*\d[\d\.\s]*,\d{2})\s*$"
)


def _preprocess_image(source_bytes: bytes) -> Image.Image:
    image = Image.open(BytesIO(source_bytes)).convert("L")
    # Upscale and denoise to improve OCR on small mobile screenshots.
    image = image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)
    image = image.filter(ImageFilter.MedianFilter(size=3))
    image = ImageOps.autocontrast(image)
    image = image.point(lambda p: 255 if p > 150 else 0)
    return image


def _extract_text_with_ocr(source_bytes: bytes) -> str:
    if pytesseract is None:
        raise RuntimeError("OCR no disponible: falta paquete Python 'pytesseract'.")
    if not getattr(pytesseract.pytesseract, "tesseract_cmd", ""):
        candidates = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Users\Pilar\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"),
        ]
        for candidate in candidates:
            if candidate.exists():
                pytesseract.pytesseract.tesseract_cmd = str(candidate)
                break
    cmd_value = str(getattr(pytesseract.pytesseract, "tesseract_cmd", "")).strip()
    cmd_exists = bool(cmd_value and (Path(cmd_value).exists() or shutil.which(cmd_value)))
    if not cmd_exists:
        raise RuntimeError("OCR no disponible: no se encontro el ejecutable Tesseract en rutas conocidas.")
    image = _preprocess_image(source_bytes)
    try:
        available_langs: set[str] = set()
        try:
            available_langs = set(pytesseract.get_languages(config=""))
        except Exception:
            available_langs = set()
        # Prefer Spanish-aware OCR when installed, fallback to English-only.
        ocr_lang = "spa+eng" if {"spa", "eng"}.issubset(available_langs) else "eng"
        return pytesseract.image_to_string(image, lang=ocr_lang)
    except Exception as exc:
        raise RuntimeError(f"Fallo OCR de Tesseract: {exc}") from exc


def _parse_date_ddmm_current_year(date_ddmm: str) -> pd.Timestamp:
    day, month = date_ddmm.split("/")
    return pd.Timestamp(year=datetime.now().year, month=int(month), day=int(day))


def _normalize_lines(raw_text: str) -> list[str]:
    lines = [line.strip() for line in raw_text.splitlines()]
    return [line for line in lines if line]


def _sanitize_ocr_line(line: str) -> str:
    """Normalizes common OCR glitches affecting dd/mm tokens."""
    normalized = line.replace("—", "-").replace("–", "-")

    char_map = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "!": "1"})

    def _repl_double_day(match: re.Match[str]) -> str:
        d1 = match.group(1).translate(char_map)
        d2 = match.group(2).translate(char_map)
        m1 = match.group(3).translate(char_map)
        m2 = match.group(4).translate(char_map)
        return f"{d1}{d2}/{m1}{m2}"

    # Handles forms like "1 1/02", "I1/02", "O1/12".
    normalized = re.sub(
        r"(?<!\d)([0-3OoIl|!])\s*([0-9OoIl|!])\s*/\s*([0-1OoIl|!])\s*([0-9OoIl|!])(?!\d)",
        _repl_double_day,
        normalized,
    )

    def _repl_single_day(match: re.Match[str]) -> str:
        d1 = match.group(1).translate(char_map)
        m1 = match.group(2).translate(char_map)
        m2 = match.group(3).translate(char_map)
        return f"{d1}/{m1}{m2}"

    # Handles forms like "1 /02" and "I/02".
    normalized = re.sub(
        r"(?<!\d)([0-9OoIl|!])\s*/\s*([0-1OoIl|!])\s*([0-9OoIl|!])(?!\d)",
        _repl_single_day,
        normalized,
    )
    return normalized


def _maybe_fix_ocr_day_with_context(parsed_rows: list[dict[str, object]], warnings: list[str]) -> None:
    """Fixes likely OCR confusion in day prefix (0x vs 1x) using neighbor dates."""
    for idx, row in enumerate(parsed_rows):
        date_token = str(row.get("_date_token", ""))
        if not re.match(r"^0\d/\d{1,2}$", date_token):
            continue

        day = int(row["_day"])
        month = int(row["_month"])
        day_candidate = day + 10
        if day_candidate > 31:
            continue

        neighbor_days: list[int] = []
        for neighbor_idx in (idx - 1, idx + 1):
            if neighbor_idx < 0 or neighbor_idx >= len(parsed_rows):
                continue
            neighbor = parsed_rows[neighbor_idx]
            if int(neighbor["_month"]) == month:
                neighbor_days.append(int(neighbor["_day"]))

        if not neighbor_days:
            continue

        current_score = sum(abs(day - n_day) for n_day in neighbor_days)
        candidate_score = sum(abs(day_candidate - n_day) for n_day in neighbor_days)

        # Keep it conservative: only correct when candidate is clearly closer.
        if candidate_score + 2 <= current_score:
            row["_day"] = day_candidate
            row["date"] = pd.Timestamp(year=datetime.now().year, month=month, day=day_candidate)
            warnings.append(
                f"Correccion OCR aplicada en fecha: '{date_token}' -> '{day_candidate:02d}/{month:02d}' ({row.get('descripcion', 'Movimiento bancario')})."
            )


def parse_bna_image(source_bytes: bytes, source_name: str = "") -> BnaParseResult:
    raw_text = _extract_text_with_ocr(source_bytes)
    lines = _normalize_lines(raw_text)
    warnings: list[str] = []
    rows: list[dict[str, object]] = []
    pending_desc = ""

    for line in lines:
        clean_line = _sanitize_ocr_line(line)
        match = _DATE_AMOUNT_PATTERN.search(clean_line)
        if not match:
            if re.search(r"[A-Z]{2,}", clean_line.upper()) and not re.search(r"\d{1,2}/\d{1,2}", clean_line):
                pending_desc = clean_line.strip()
            continue

        desc_inline = (match.group("prefix") or "").strip()
        desc = desc_inline if desc_inline else pending_desc
        pending_desc = ""
        if not desc:
            desc = "Movimiento bancario"

        date_token = match.group("date")
        amount_token = match.group("amount")
        signed_amount = parse_number_ar(amount_token.replace(" ", ""))
        tipo = "Gasto" if str(amount_token).strip().startswith("-") else "Ingreso"
        if "CR INTERB" in desc.upper():
            tipo = "Ingreso"

        amount_abs = abs(float(signed_amount))
        day_value, month_value = (int(part) for part in date_token.split("/"))
        date_value = _parse_date_ddmm_current_year(date_token)
        desc_norm = normalize_description(desc)
        ref_seed = f"bank_capture|ocr|bna|{date_value.strftime('%Y-%m-%d')}|{tipo}|{amount_abs:.2f}|{desc_norm}"
        ref_hash = sha1(ref_seed.encode("utf-8")).hexdigest()[:12]
        mp_ref = f"cap_{ref_hash}"

        note = f"origen=captura_bancaria; banco=bna; img={source_name or 'captura'}; mp_ref={mp_ref}"
        rows.append(
            {
                "date": date_value,
                "tipo": tipo,
                "categoria": "Otros",
                "subcategoria": "",
                "regla_categoria": "CAPTURA_BANCARIA_OCR_V1",
                "descripcion": desc.strip(),
                "descripcion_norm": desc_norm,
                "monto": amount_abs,
                "mp_ref": mp_ref,
                "source_account": "Banco",
                "source_channel": "Captura",
                "compartido": "No",
                "source_note": note,
                "_date_token": date_token,
                "_day": day_value,
                "_month": month_value,
            }
        )

    if not rows:
        warnings.append(
            "No se detectaron movimientos con patron fecha+monto (dd/mm y monto con coma). Revisa calidad de imagen."
        )
    else:
        _maybe_fix_ocr_day_with_context(rows, warnings)

    parsed_df = pd.DataFrame(rows)
    if not parsed_df.empty:
        parsed_df = parsed_df.drop(columns=["_date_token", "_day", "_month"], errors="ignore")
        parsed_df = parsed_df.sort_values("date").reset_index(drop=True)

    return BnaParseResult(
        parsed_df=parsed_df,
        raw_text=raw_text,
        extracted_rows=int(parsed_df.shape[0]),
        warnings=warnings,
    )
