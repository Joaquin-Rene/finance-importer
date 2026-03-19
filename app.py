from __future__ import annotations

from io import BytesIO
from pathlib import Path
from hashlib import sha1
from datetime import datetime
import re

import pandas as pd
import streamlit as st
from PIL import Image

from src.finanzas_importer.bna_image_parser import BnaParseResult, parse_bna_image
from src.finanzas_importer.analytics import (
    build_executive_summary,
    build_alerts,
    compute_month_projection,
    compute_monthly_kpis,
    load_finanzas_history,
    merge_history_with_pending,
)
from src.finanzas_importer.mp_parser import (
    ParseResult,
    infer_category_from_description,
    normalize_description,
    parse_mercado_pago_excel,
)
from src.finanzas_importer.ui_components import (
    format_preview_df,
    init_upload_state,
    inject_styles,
    is_path_writable,
    render_footer,
    render_hero,
    render_import_step,
    render_insights_step,
    render_mode_panel,
    render_review_step,
    render_preview_table,
    render_step_title,
    render_upload_dropzone_hint,
    render_upload_meta,
    setup_sidebar,
)
from src.finanzas_importer.workbook_writer import (
    DATE_FILTER_MODE_AFTER_MAX,
    DATE_FILTER_MODE_SKIP_EXISTING,
    ImportPlan,
    build_import_plan,
    import_into_finanzas_workbook,
)
from src.finanzas_importer.utils import parse_number_ar


def _build_bna_private_preview(image_bytes: bytes, crop_top_ratio: float = 0.34) -> Image.Image:
    image = Image.open(BytesIO(image_bytes))
    width, height = image.size
    top = int(height * crop_top_ratio)
    if top >= height:
        return image
    return image.crop((0, top, width, height))


def _serialize_bna_seed(seed_df: pd.DataFrame) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for _, row in seed_df.fillna("").iterrows():
        monto_raw = row.get("monto", 0.0)
        try:
            monto_value = abs(float(parse_number_ar(monto_raw)))
        except Exception:
            monto_value = 0.0
        is_income = str(row.get("tipo", "")).strip().lower().startswith("i")
        monto_signed = monto_value if is_income else -monto_value
        rows.append(
            {
                "fecha": str(row.get("fecha", "")).strip(),
                "descripcion": str(row.get("descripcion", "")).strip(),
                "monto": f"{monto_signed:.2f}".replace(".", ",") if monto_value else "",
            }
        )
    return rows or [{"fecha": "", "descripcion": "", "monto": ""}]


def _render_bna_manual_editor(seed_df: pd.DataFrame) -> pd.DataFrame:
    seed_signature = sha1(seed_df.to_csv(index=False).encode("utf-8")).hexdigest()
    if st.session_state.get("bna_manual_seed") != seed_signature:
        st.session_state.bna_manual_seed = seed_signature
        st.session_state.bna_manual_rows = _serialize_bna_seed(seed_df)

    rows: list[dict[str, object]] = list(st.session_state.get("bna_manual_rows", _serialize_bna_seed(seed_df)))

    updated_rows: list[dict[str, object]] = []
    remove_idx: int | None = None
    with st.container(border=True):
        st.markdown("<div class='manual-editor-scope'></div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='manual-editor-kicker'>Editable table</div>"
            "<div class='manual-editor-intro'>Review or complete each transaction before importing it.</div>",
            unsafe_allow_html=True,
        )
        header_cols = st.columns([1.1, 2.9, 1.2, 1.0, 0.9])
        header_cols[0].markdown("<div class='manual-editor-head'>Date</div>", unsafe_allow_html=True)
        header_cols[1].markdown("<div class='manual-editor-head'>Description</div>", unsafe_allow_html=True)
        header_cols[2].markdown("<div class='manual-editor-head'>Amount</div>", unsafe_allow_html=True)
        header_cols[3].markdown("<div class='manual-editor-head'>Type</div>", unsafe_allow_html=True)
        header_cols[4].markdown("<div class='manual-editor-head'>Action</div>", unsafe_allow_html=True)

        for idx, row in enumerate(rows):
            st.markdown("<div class='manual-editor-row'></div>", unsafe_allow_html=True)
            cols = st.columns([1.1, 2.9, 1.2, 1.0, 0.9])
            fecha = cols[0].text_input(
                f"Date {idx + 1}",
                value=str(row.get("fecha", "")),
                key=f"bna_fecha_{idx}",
                label_visibility="collapsed",
                placeholder="dd/mm",
            )
            descripcion = cols[1].text_input(
                f"Description {idx + 1}",
                value=str(row.get("descripcion", "")),
                key=f"bna_desc_{idx}",
                label_visibility="collapsed",
                placeholder="Description",
            )
            monto = cols[2].text_input(
                f"Amount {idx + 1}",
                value=str(row.get("monto", "")),
                key=f"bna_monto_{idx}",
                label_visibility="collapsed",
                placeholder="-0,00 / +0,00",
            )
            try:
                monto_num = float(parse_number_ar(monto))
            except Exception:
                monto_num = 0.0
            tipo = "Ingreso" if monto_num > 0 else "Gasto"
            tipo_class = "manual-type-ingreso" if tipo == "Ingreso" else "manual-type-gasto"
            cols[3].markdown(
                f"<div class='manual-type-chip {tipo_class}'><span class='tipo-dot'>&#9679;</span>{tipo}</div>",
                unsafe_allow_html=True,
            )
            if cols[4].button("Remove", key=f"bna_remove_{idx}", use_container_width=True):
                remove_idx = idx
            updated_rows.append({"fecha": fecha, "descripcion": descripcion, "monto": monto})

        action_cols = st.columns([1.1, 1.1, 3.8])
        if action_cols[0].button("Add row", key="bna_add_row", use_container_width=True):
            updated_rows.append({"fecha": "", "descripcion": "", "monto": ""})
            st.session_state.bna_manual_rows = updated_rows
            st.rerun()
        if action_cols[1].button("Restore OCR", key="bna_restore_seed", use_container_width=True):
            st.session_state.bna_manual_rows = _serialize_bna_seed(seed_df)
            st.rerun()

    if remove_idx is not None:
        updated_rows.pop(remove_idx)
        st.session_state.bna_manual_rows = updated_rows or [{"fecha": "", "descripcion": "", "monto": ""}]
        st.rerun()

    st.session_state.bna_manual_rows = updated_rows or [{"fecha": "", "descripcion": "", "monto": ""}]
    return pd.DataFrame(st.session_state.bna_manual_rows)


st.set_page_config(page_title="Financial Operations Importer", layout="wide")


inject_styles()
finanzas_path, date_filter_mode, default_category, show_technical, ahorro_meta_pct = setup_sidebar()
finanzas_ok, finanzas_msg = is_path_writable(finanzas_path)
render_hero(finanzas_ok, finanzas_msg)

if "import_mode" not in st.session_state:
    st.session_state.import_mode = "Excel file"

mode_cols = st.columns(2)
with mode_cols[0]:
    if st.button("Excel file", type="primary" if st.session_state.import_mode == "Excel file" else "secondary", width="stretch"):
        st.session_state.import_mode = "Excel file"
with mode_cols[1]:
    if st.button("Bank capture", type="primary" if st.session_state.import_mode == "Bank capture" else "secondary", width="stretch"):
        st.session_state.import_mode = "Bank capture"

mode = st.session_state.import_mode
render_mode_panel(mode)

init_upload_state()

if mode == "Excel file":
    render_step_title(1, "Load transaction file")
    st.markdown(
        "<div class='subtle'>Upload a .xlsx or .xls file to review it before importing. The current parser is optimized for Mercado Pago exports.</div>",
        unsafe_allow_html=True,
    )
    render_upload_dropzone_hint()

    source_bytes: bytes | None = None
    uploaded_file = st.file_uploader(
        "Drag and drop the transaction file here (.xlsx or .xls), or use Browse files",
        type=["xlsx", "xls"],
        key=st.session_state.uploader_key,
        help="If drag and drop does not work, drop the file inside the dotted uploader area.",
    )

    if uploaded_file is not None and st.session_state.upload_at is None:
        st.session_state.upload_at = pd.Timestamp.now().to_pydatetime()
    if uploaded_file is not None:
        source_bytes = uploaded_file.getvalue()

    if uploaded_file is not None:
        render_upload_meta(uploaded_file)
    else:
        st.markdown("**Alternative:** load from a local path if browser drag and drop fails.")
        local_path = st.text_input(
            "Transaction file path",
            value=str((Path.home() / "Downloads").resolve()),
            help="Paste the full path to the .xlsx or .xls file.",
        )
        if st.button("Load file from path"):
            local_file = Path(local_path.strip().strip('"'))
            if not local_file.exists() or not local_file.is_file():
                st.error("The path does not exist or is not a file.")
            elif local_file.suffix.lower() not in {".xlsx", ".xls"}:
                st.error("The file must be .xlsx or .xls.")
            else:
                source_bytes = local_file.read_bytes()
                st.session_state.upload_at = pd.Timestamp.now().to_pydatetime()
                st.success(f"File loaded from path: {local_file.name}")

    parse_result: ParseResult | None = None
    parsed_df: pd.DataFrame | None = None
    plan: ImportPlan | None = None
    plan_error: str | None = None
    analytics_df: pd.DataFrame | None = None

    if source_bytes is not None:
        try:
            parse_result = parse_mercado_pago_excel(BytesIO(source_bytes))
            parsed_df = parse_result.parsed_df

            if finanzas_path.strip() and Path(finanzas_path).exists():
                try:
                    plan = build_import_plan(parsed_df, Path(finanzas_path), date_filter_mode=date_filter_mode)
                except Exception as exc:
                    plan_error = str(exc)
            elif finanzas_path.strip():
                st.warning(f"Destination workbook not found at: {finanzas_path}")

            to_import_final = render_review_step(
                parsed_df=parsed_df,
                parse_result=parse_result,
                plan=plan,
                date_filter_mode=date_filter_mode,
                show_technical=show_technical,
                plan_error=plan_error,
            )
            if finanzas_path.strip() and Path(finanzas_path).exists():
                history_df = load_finanzas_history(Path(finanzas_path))
                pending_df = plan.to_import_df if plan is not None else parsed_df
                analytics_df = merge_history_with_pending(history_df=history_df, pending_df=pending_df)
                ref_date = analytics_df["date"].max() if not analytics_df.empty else None
                kpis = compute_monthly_kpis(analytics_df, ref_date=ref_date)
                projection = compute_month_projection(analytics_df, ref_date=ref_date)
                alerts = build_alerts(analytics_df, projection, ref_date=ref_date, ahorro_meta_pct=ahorro_meta_pct)
                executive_summary = build_executive_summary(alerts, kpis, projection)
                render_insights_step(
                    analytics_df=analytics_df,
                    kpis=kpis,
                    projection=projection,
                    alerts=alerts,
                    ref_date=ref_date,
                    executive_summary=executive_summary,
                    ahorro_meta_pct=ahorro_meta_pct,
                )

            render_import_step(
                parsed_df=parsed_df,
                parse_result=parse_result,
                to_import_final=to_import_final,
                finanzas_path=finanzas_path,
                date_filter_mode=date_filter_mode,
                default_category=default_category,
                importer_fn=import_into_finanzas_workbook,
            )

        except Exception as exc:
            st.error(f"No pudimos leer el archivo de movimientos: {exc}")
else:
    render_step_title(1, "Load bank capture")
    current_year = datetime.now().year
    st.markdown(
        f"<div class='subtle'>Upload up to 3 bank captures. Manual entry is assisted, with optional OCR preload when available. Dates in dd/mm use year {current_year}.</div>",
        unsafe_allow_html=True,
    )
    bna_files = st.file_uploader(
        "Drag bank captures here (.jpg/.jpeg/.png) - up to 3 files",
        type=["jpg", "jpeg", "png"],
        key="bna_capture_uploader",
        accept_multiple_files=True,
    )

    bna_result: BnaParseResult | None = None
    bna_plan: ImportPlan | None = None
    if "bna_date_mode" not in st.session_state:
        st.session_state.bna_date_mode = "Exclude existing dates"

    st.caption("Date filter for bank captures")
    bna_mode_cols = st.columns(2)
    with bna_mode_cols[0]:
        if st.button(
            "Exclude existing dates",
            type="primary" if st.session_state.bna_date_mode == "Exclude existing dates" else "secondary",
            width="stretch",
            key="bna_mode_existing",
        ):
            st.session_state.bna_date_mode = "Exclude existing dates"
    with bna_mode_cols[1]:
        if st.button(
            "Only after latest date",
            type="primary" if st.session_state.bna_date_mode == "Only after latest existing date" else "secondary",
            width="stretch",
            key="bna_mode_after",
        ):
            st.session_state.bna_date_mode = "Only after latest existing date"

    bna_date_mode_label = st.session_state.bna_date_mode
    bna_date_filter_mode = (
        DATE_FILTER_MODE_SKIP_EXISTING
        if bna_date_mode_label.startswith("Exclude")
        else DATE_FILTER_MODE_AFTER_MAX
    )

    parsed_df: pd.DataFrame | None = None
    if bna_files:
        selected_files = bna_files[:3]
        if len(bna_files) > 3:
            st.warning("Only the first 3 captures will be processed.")

        preview_cols = st.columns(len(selected_files))
        for idx, capture_file in enumerate(selected_files):
            with preview_cols[idx]:
                preview_image = _build_bna_private_preview(capture_file.getvalue())
                st.image(preview_image, caption=f"Private preview: {capture_file.name}", width=240)

        seed_rows: list[dict[str, object]] = []
        ocr_warnings: list[str] = []
        for capture_file in selected_files:
            image_bytes = capture_file.getvalue()
            try:
                bna_result = parse_bna_image(image_bytes, source_name=capture_file.name)
                ocr_warnings.extend(bna_result.warnings)
                if not bna_result.parsed_df.empty:
                    for _, row in bna_result.parsed_df.iterrows():
                        seed_rows.append(
                            {
                                "fecha": row["date"].strftime("%d/%m"),
                                "tipo": row["tipo"],
                                "descripcion": row["descripcion"],
                                "monto": row["monto"],
                            }
                        )
            except Exception as exc:
                ocr_warnings.append(f"[{capture_file.name}] OCR preload unavailable: {exc}")

        for warning in ocr_warnings:
            st.warning(warning)

        if seed_rows:
            seed_df = pd.DataFrame(seed_rows)
        else:
            seed_df = pd.DataFrame(
                [
                    {
                        "fecha": "",
                        "tipo": "Gasto",
                        "descripcion": "",
                        "monto": 0.0,
                    }
                ]
            )
        if not seed_df.empty:
            sort_dates = pd.to_datetime(
                seed_df["fecha"].astype(str).str.strip() + f"/{current_year}",
                dayfirst=True,
                errors="coerce",
            )
            seed_df = (
                seed_df.assign(_sort_date=sort_dates)
                .sort_values("_sort_date", na_position="last")
                .drop(columns=["_sort_date"])
                .reset_index(drop=True)
            )

        render_step_title(2, "Complete transactions manually")
        st.caption(f"Use dd/mm dates (year {current_year} is added automatically). Use `+` for income and `-` for expense. `CR INTERB` is treated as income.")
        edited_df = _render_bna_manual_editor(seed_df)

        rows: list[dict[str, object]] = []
        parse_errors: list[str] = []
        seen_bna_refs: set[str] = set()
        duplicate_rows_in_batch = 0
        for idx, row in edited_df.iterrows():
            fecha_raw = str(row.get("fecha", "")).strip()
            desc = str(row.get("descripcion", "")).strip()
            monto_raw = row.get("monto", 0.0)

            if not fecha_raw and not desc:
                continue
            if re.match(r"^\d{1,2}/\d{1,2}$", fecha_raw):
                fecha_raw = f"{fecha_raw}/{current_year}"

            fecha = pd.to_datetime(fecha_raw, dayfirst=True, errors="coerce")
            if pd.isna(fecha):
                parse_errors.append(f"Row {idx + 1}: invalid date '{row.get('fecha', '')}'.")
                continue

            try:
                monto_signed = float(parse_number_ar(monto_raw))
            except Exception:
                parse_errors.append(f"Row {idx + 1}: invalid amount '{monto_raw}'.")
                continue
            monto_val = abs(monto_signed)
            if monto_val <= 0:
                continue

            tipo = "Ingreso" if monto_signed > 0 else "Gasto"
            if "CR INTERB" in desc.upper():
                tipo = "Ingreso"

            desc_safe = desc if desc else "Bank transaction"
            desc_safe, categoria, subcategoria, regla = infer_category_from_description(desc_safe, default_category="Otros")
            desc_norm = normalize_description(desc_safe)
            ref_seed = f"bank_capture|manual|{fecha.strftime('%Y-%m-%d')}|{tipo}|{monto_val:.2f}|{desc_norm}"
            mp_ref = f"cap_{sha1(ref_seed.encode('utf-8')).hexdigest()[:12]}"
            if mp_ref in seen_bna_refs:
                duplicate_rows_in_batch += 1
                continue

            seen_bna_refs.add(mp_ref)
            note = f"source=bank_capture; bank=bna; mp_ref={mp_ref}"
            rows.append(
                {
                    "date": fecha,
                    "tipo": tipo,
                    "categoria": categoria,
                    "subcategoria": subcategoria,
                    "regla_categoria": f"CAPTURA_BANCARIA_MANUAL_{regla}",
                    "descripcion": desc_safe,
                    "descripcion_norm": desc_norm,
                    "monto": monto_val,
                    "mp_ref": mp_ref,
                    "source_account": "Banco",
                    "source_channel": "Captura",
                    "compartido": "No",
                    "source_note": note,
                }
            )

        for err in parse_errors:
            st.warning(err)

        if rows:
            parsed_df = pd.DataFrame(rows).sort_values("date").reset_index(drop=True)
            render_step_title(3, "Bank capture preview")
            render_preview_table(
                format_preview_df(parsed_df).head(100),
                ["date", "tipo_visual", "categoria", "descripcion", "monto", "compartido"],
                {
                    "date": "Date",
                    "tipo_visual": "Type",
                    "categoria": "Category",
                    "descripcion": "Description",
                    "monto": "Amount",
                    "compartido": "Shared",
                },
            )
            if duplicate_rows_in_batch > 0:
                st.info(f"{duplicate_rows_in_batch} duplicates were detected and skipped across captures.")
        else:
            st.info("Complete at least one valid row to preview and import.")

    if parsed_df is not None and not parsed_df.empty:
        if finanzas_path.strip() and Path(finanzas_path).exists():
            try:
                bna_plan = build_import_plan(parsed_df, Path(finanzas_path), date_filter_mode=bna_date_filter_mode)
            except Exception as exc:
                st.error(f"Could not build the bank capture import plan: {exc}")
        else:
            st.warning(f"Destination workbook not found at: {finanzas_path}")

        to_import = int(bna_plan.to_import_df.shape[0]) if bna_plan is not None else int(parsed_df.shape[0])
        dupes_ref = int(bna_plan.duplicate_mp_ref_df.shape[0]) if bna_plan is not None else 0
        dupes_key = int(bna_plan.duplicate_compound_df.shape[0]) if bna_plan is not None else 0
        st.info(
            f"Loaded: {int(parsed_df.shape[0])} | To import: {to_import} | Duplicate refs: {dupes_ref} | Duplicate keys: {dupes_key}"
        )

        confirm_bna = st.checkbox("I confirm the bank capture import", value=False)
        if st.button("Import bank capture", type="primary", disabled=not confirm_bna):
            import_df = bna_plan.to_import_df if bna_plan is not None else parsed_df
            try:
                result = import_into_finanzas_workbook(
                    import_df,
                    Path(finanzas_path),
                    date_filter_mode=bna_date_filter_mode,
                    default_category=default_category,
                )
                if result.note == "No rows to import":
                    st.info("There are no new rows to import from the bank capture.")
                else:
                    st.success(
                        "Bank capture import completed: "
                        f"{result.added_rows} added, "
                        f"{result.skipped_rows_by_date} skipped by date, "
                        f"{result.duplicate_rows_mp_ref} duplicate references, "
                        f"{result.duplicate_rows_compound_key} duplicate compound keys."
                    )
                    if result.backup_path is not None:
                        st.info("Backup created:")
                        st.code(str(result.backup_path))
            except Exception as exc:
                st.error(f"An error occurred while importing the bank capture into the destination workbook: {exc}")

render_footer()
