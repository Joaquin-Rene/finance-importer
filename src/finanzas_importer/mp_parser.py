from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import BinaryIO
import re
import unicodedata

import pandas as pd

from .utils import parse_number_ar

REQUIRED_COLUMNS = ["TRANSACTION_TYPE", "REFERENCE_ID", "TRANSACTION_NET_AMOUNT"]
DATE_CANDIDATES = [
    "RELEASE_DATE",
    "CREATION_DATE",
    "DATE",
    "TRANSACTION_DATE",
    "OPERATION_DATE",
]
FILTER_REASON_RENDIMIENTOS = "Rendimientos"
FILTER_REASON_SELF_TRANSFER = "Transferencia interna (self)"
_DATE_DOMINANT_DAY_THRESHOLD = 0.70
SELF_TRANSFER_TOKENS = ("joaquin", "rene", "matias")


@dataclass
class ParseResult:
    parsed_df: pd.DataFrame
    filtered_df: pd.DataFrame
    filtered_reasons: dict[str, int]
    self_transfer_examples: pd.DataFrame
    nadia_shared_count: int
    date_source_column: str
    suspicious_dates: bool
    suspicious_day_ratio: float


def _find_header_row(path_or_buffer: str | Path | BinaryIO) -> int:
    preview = pd.read_excel(path_or_buffer, header=None, dtype=str)

    for idx, row in preview.iterrows():
        values = [str(v).strip().upper() for v in row.tolist() if pd.notna(v)]
        if "RELEASE_DATE" in values:
            return int(idx)

    raise ValueError("No se encontro la fila header con RELEASE_DATE en el Excel de Mercado Pago.")


def _strip_accents(text: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch))


def normalize_description(text: object) -> str:
    value = _strip_accents(str(text).strip().lower())
    value = re.sub(r"^(pago|transferencia enviada|transferencia recibida)\s+", "", value)
    value = re.sub(r"[^a-z0-9\s]", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _clean_description(transaction_type: str) -> str:
    raw = str(transaction_type).strip()
    lowered = _strip_accents(raw.lower())
    if "payuaruber" in lowered:
        return "Uber"
    if "pago shell" in lowered:
        return "nafta moto"
    if "pago supermercados dia" in lowered or "pago dia tienda" in lowered:
        return "Supermercado dia"
    if "almacenes nicoleta" in lowered or "almacenesnico" in lowered:
        return "Almacen nicoleta"
    return raw


def _is_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    return "transferencia" in lowered or "tranferencia" in lowered


def _is_sent_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    return "transferencia enviada" in lowered or "tranferencia enviada" in lowered


def _is_self_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    return _is_transfer(lowered) and all(token in lowered for token in SELF_TRANSFER_TOKENS)


def _is_nadia_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    return _is_transfer(lowered) and "nadia" in lowered


def _is_nadia_sent_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    return _is_nadia_transfer(lowered) and _is_sent_transfer(lowered)


def _is_nadia_caceres_transfer(text: object) -> bool:
    lowered = _strip_accents(str(text).lower())
    lowered = re.sub(r"\s+", " ", lowered).strip()
    return _is_transfer(lowered) and (
        "nadia marisol caceres" in lowered
        or ("nadia" in lowered and "marisol" in lowered and "caceres" in lowered)
    )


def _resolve_tipo(transaction_type: object, amount: float) -> str:
    lowered = _strip_accents(str(transaction_type).lower())
    if (
        "transferencia enviada nadia" in lowered
        or "tranferencia enviada nadia" in lowered
        or _is_nadia_caceres_transfer(transaction_type)
    ):
        return "Gasto"
    return "Gasto" if amount < 0 else "Ingreso"


def infer_category_from_description(description: object, default_category: str = "Otros") -> tuple[str, str, str, str]:
    desc = str(description).strip()
    normalized = _strip_accents(desc.lower())

    if "ingreso principal" in normalized:
        return desc, "Ingresos", "Base", "MP_INCOME_BASE"
    if "proyecto independiente" in normalized:
        return desc, "Ingresos", "Proyecto", "MP_INCOME_PROJECT"
    if "gastos compartidos" in normalized:
        return desc, "Ingresos", "Compartido", "MP_SHARED_IN"
    if "ajuste demo" in normalized or "ajuste viaje demo" in normalized or "reintegro" in normalized:
        return desc, "Ingresos", "Ajuste", "MP_INCOME_ADJUSTMENT"
    if "vivienda mensual" in normalized:
        return desc, "Vivienda", "Fijo", "MP_HOUSING"
    if "fondo viaje" in normalized:
        return desc, "Ahorro", "Fondo viaje", "MP_SAVINGS"
    if "mercado semanal" in normalized or "compra de cercania" in normalized:
        return desc, "Comida", "Mercado", "MP_FOOD"
    if "salida casual" in normalized:
        return desc, "Comida", "Salida", "MP_DINING"
    if "salud y cuidado" in normalized:
        return desc, "Salud", "Cuidado", "MP_HEALTH"
    if "combustible" in normalized or "movilidad urbana" in normalized:
        return desc, "Transporte", "Movilidad", "MP_TRANSPORT"
    if "suscripcion media" in normalized or "suscripcion demo" in normalized:
        return desc, "Ocio", "Digital", "MP_DIGITAL"
    if "actividad fisica" in normalized:
        return desc, "Bienestar", "Habito", "MP_WELLNESS"
    if "regalo planificado" in normalized:
        return desc, "Regalos", "Planificado", "MP_GIFTS"
    if normalized.startswith("transferencia recibida"):
        return desc, "Ingresos", "Transferencia", "MP_TRANSFER_IN"
    if normalized.startswith("transferencia enviada"):
        return desc, "Transferencias", "Salida", "MP_TRANSFER_OUT"

    if normalized in {"cr interb", "cap inter"} or "interb" in normalized:
        return desc, "Ingresos", "Transferencia", "BANK_TRANSFER_IN"
    if normalized == "db transf":
        return desc, "Transferencias", "Salida", "BANK_TRANSFER_OUT"
    if normalized in {"cpra.sup", "cpra.merp", "cpra.dia"}:
        return desc, "Comida", "Consumo diario", "BANK_FOOD"
    if normalized.startswith("cpra.la"):
        return desc, "Compras", "Consumo general", "BANK_PURCHASE"
    if normalized.startswith("cpra."):
        return desc, "Compras", "Consumo general", "BANK_PURCHASE"

    if re.search(r"\bdia\b", normalized) or "dia tienda" in normalized or "supermercados dia" in normalized:
        return "Supermercado dia", "Comida", "", "DIA"
    if "shell" in normalized:
        return "nafta moto", "Transporte", "", "SHELL"
    if "payuaruber" in normalized or "uber" in normalized:
        return "uber", "Transporte", "", "UBER"
    if "almacenes nicoleta" in normalized or "almacenesnico" in normalized:
        return "Almacen nicoleta", "Comida", "", "NICOLETA"

    return desc, default_category, "", "DEFAULT"


def _parse_single_date(value: object) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return pd.to_datetime(float(value), unit="D", origin="1899-12-30", errors="coerce")
    text = str(value).strip()
    if text == "":
        return pd.NaT
    if re.fullmatch(r"\d+(\.0+)?", text):
        return pd.to_datetime(float(text), unit="D", origin="1899-12-30", errors="coerce")
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def _parse_date_series(df: pd.DataFrame, col: str) -> pd.Series:
    return df[col].map(_parse_single_date)


def _best_date_column(df: pd.DataFrame) -> tuple[str, pd.Series]:
    candidates = [c for c in DATE_CANDIDATES if c in df.columns]
    if not candidates:
        candidates = [c for c in df.columns if "DATE" in c.upper()]
    if not candidates:
        raise ValueError("No se encontro ninguna columna de fecha (RELEASE_DATE/CREATION_DATE/DATE).")

    best_col: str | None = None
    best_series: pd.Series | None = None
    best_score: tuple[int, float] | None = None
    for col in candidates:
        parsed = _parse_date_series(df, col)
        valid = parsed.dropna()
        valid_count = int(valid.shape[0])
        if valid_count == 0:
            score = (0, -1.0)
        else:
            dominant_share = float(valid.dt.day.value_counts(normalize=True).iloc[0])
            score = (valid_count, -dominant_share)
        if best_score is None or score > best_score:
            best_col = col
            best_series = parsed
            best_score = score

    if best_col is None or best_series is None:
        raise ValueError("No se pudieron parsear fechas validas desde el export de Mercado Pago.")
    return best_col, best_series


def parse_mercado_pago_excel(path_or_buffer: str | Path | BinaryIO) -> ParseResult:
    header_row = _find_header_row(path_or_buffer)

    df = pd.read_excel(path_or_buffer, header=header_row)
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Faltan columnas esperadas: {missing}")

    mp = df.copy()
    mp = mp.dropna(subset=REQUIRED_COLUMNS, how="all")
    date_source_column, parsed_dates = _best_date_column(mp)
    mp["date"] = parsed_dates
    if mp["date"].isna().all():
        raise ValueError("No se pudieron parsear fechas validas desde las columnas de fecha.")

    mp["net_amount"] = mp["TRANSACTION_NET_AMOUNT"].map(parse_number_ar)
    mp["tipo"] = mp.apply(lambda row: _resolve_tipo(row["TRANSACTION_TYPE"], row["net_amount"]), axis=1)
    mp["descripcion"] = mp["TRANSACTION_TYPE"].map(_clean_description)
    mp["categoria"] = ""
    mp["subcategoria"] = ""
    mp["regla_categoria"] = ""
    mp["compartido"] = "No"
    mp["monto"] = mp["net_amount"].abs().astype(float)
    mp["mp_ref"] = mp["REFERENCE_ID"].astype(str).str.strip()
    mp["descripcion_norm"] = mp["descripcion"].map(normalize_description)
    mp["source_account"] = "Mercado Pago"
    mp["source_channel"] = "Excel"

    mask_rendimientos = mp["descripcion"].astype(str).str.contains("rendimientos", case=False, na=False)
    mask_self_transfer = mp["TRANSACTION_TYPE"].map(_is_self_transfer)
    filtered_df = mp.loc[mask_rendimientos | mask_self_transfer, ["date", "tipo", "descripcion", "monto", "mp_ref"]].copy()
    filtered_df.loc[mask_self_transfer[filtered_df.index], "reason"] = FILTER_REASON_SELF_TRANSFER
    filtered_df.loc[mask_rendimientos[filtered_df.index], "reason"] = FILTER_REASON_RENDIMIENTOS
    self_transfer_examples = mp.loc[mask_self_transfer, ["date", "tipo", "descripcion", "monto", "mp_ref"]].copy()
    self_transfer_examples = self_transfer_examples.head(10).reset_index(drop=True)
    filtered_reasons = {
        FILTER_REASON_RENDIMIENTOS: int(mask_rendimientos.sum()),
        FILTER_REASON_SELF_TRANSFER: int(mask_self_transfer.sum()),
    }
    mp = mp.loc[~mask_rendimientos & ~mask_self_transfer].copy()

    categories = mp["descripcion"].map(lambda d: infer_category_from_description(d, default_category="Otros"))
    mp["descripcion"], mp["categoria"], mp["subcategoria"], mp["regla_categoria"] = zip(*categories)
    nadia_mask = mp["TRANSACTION_TYPE"].map(_is_nadia_transfer) | mp["TRANSACTION_TYPE"].map(_is_nadia_caceres_transfer)
    nadia_sent_mask = mp["TRANSACTION_TYPE"].map(_is_nadia_sent_transfer)
    nadia_shared_mask = nadia_mask & ~nadia_sent_mask
    mp.loc[nadia_shared_mask, "compartido"] = "S?"
    mp.loc[nadia_sent_mask, "compartido"] = "No"
    mp.loc[mp["TRANSACTION_TYPE"].map(_is_nadia_caceres_transfer), "tipo"] = "Gasto"
    nadia_shared_count = int(nadia_shared_mask.sum())

    std = mp[
        [
            "date",
            "tipo",
            "categoria",
            "subcategoria",
            "regla_categoria",
            "descripcion",
            "descripcion_norm",
            "monto",
            "mp_ref",
            "source_account",
            "source_channel",
            "compartido",
        ]
    ].copy()
    std = std.dropna(subset=["date"])
    std = std[std["mp_ref"] != ""]
    std = std.sort_values("date").reset_index(drop=True)

    day_distribution = std["date"].dt.day.value_counts(normalize=True)
    suspicious_ratio = float(day_distribution.iloc[0]) if not day_distribution.empty else 0.0
    suspicious_dates = suspicious_ratio > _DATE_DOMINANT_DAY_THRESHOLD

    return ParseResult(
        parsed_df=std,
        filtered_df=filtered_df.reset_index(drop=True),
        filtered_reasons=filtered_reasons,
        self_transfer_examples=self_transfer_examples,
        nadia_shared_count=nadia_shared_count,
        date_source_column=date_source_column,
        suspicious_dates=suspicious_dates,
        suspicious_day_ratio=suspicious_ratio,
    )
