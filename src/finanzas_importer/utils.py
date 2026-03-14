from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re


def parse_number_ar(value: object) -> float:
    """Parsea numeros con formato AR (puntos de miles, coma decimal)."""
    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    text = text.replace("$", "").replace(" ", "")

    # Mantiene solo caracteres numericos relevantes.
    text = re.sub(r"[^0-9,.-]", "", text)
    if not text or text in {"-", ".", ",", "-.", "-,", ",.", ".,"}:
        return 0.0

    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    if not text or text in {"-", ".", "-."}:
        return 0.0

    return float(text)


def create_backup_path(finanzas_path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return finanzas_path.with_name(f"{finanzas_path.stem}.{ts}.bak{finanzas_path.suffix}")
