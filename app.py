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
            "<div class='manual-editor-kicker'>Tabla editable</div>"
            "<div class='manual-editor-intro'>Corrige o completa cada movimiento antes de importarlo.</div>",
            unsafe_allow_html=True,
        )
        header_cols = st.columns([1.1, 2.9, 1.2, 1.0, 0.9])
        header_cols[0].markdown("<div class='manual-editor-head'>Fecha</div>", unsafe_allow_html=True)
        header_cols[1].markdown("<div class='manual-editor-head'>Descripcion</div>", unsafe_allow_html=True)
        header_cols[2].markdown("<div class='manual-editor-head'>Monto</div>", unsafe_allow_html=True)
        header_cols[3].markdown("<div class='manual-editor-head'>Tipo</div>", unsafe_allow_html=True)
        header_cols[4].markdown("<div class='manual-editor-head'>Accion</div>", unsafe_allow_html=True)

        for idx, row in enumerate(rows):
            st.markdown("<div class='manual-editor-row'></div>", unsafe_allow_html=True)
            cols = st.columns([1.1, 2.9, 1.2, 1.0, 0.9])
            fecha = cols[0].text_input(
                f"Fecha {idx + 1}",
                value=str(row.get("fecha", "")),
                key=f"bna_fecha_{idx}",
                label_visibility="collapsed",
                placeholder="dd/mm",
            )
            descripcion = cols[1].text_input(
                f"Descripcion {idx + 1}",
                value=str(row.get("descripcion", "")),
                key=f"bna_desc_{idx}",
                label_visibility="collapsed",
                placeholder="Descripcion",
            )
            monto = cols[2].text_input(
                f"Monto {idx + 1}",
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
            if cols[4].button("Quitar", key=f"bna_remove_{idx}", use_container_width=True):
                remove_idx = idx
            updated_rows.append({"fecha": fecha, "descripcion": descripcion, "monto": monto})

        action_cols = st.columns([1.1, 1.1, 3.8])
        if action_cols[0].button("Agregar fila", key="bna_add_row", use_container_width=True):
            updated_rows.append({"fecha": "", "descripcion": "", "monto": ""})
            st.session_state.bna_manual_rows = updated_rows
            st.rerun()
        if action_cols[1].button("Restaurar OCR", key="bna_restore_seed", use_container_width=True):
            st.session_state.bna_manual_rows = _serialize_bna_seed(seed_df)
            st.rerun()

    if remove_idx is not None:
        updated_rows.pop(remove_idx)
        st.session_state.bna_manual_rows = updated_rows or [{"fecha": "", "descripcion": "", "monto": ""}]
        st.rerun()

    st.session_state.bna_manual_rows = updated_rows or [{"fecha": "", "descripcion": "", "monto": ""}]
    return pd.DataFrame(st.session_state.bna_manual_rows)


st.set_page_config(page_title="Importador de gastos e ingresos -> FINANZAS", layout="wide")


inject_styles()
finanzas_path, date_filter_mode, default_category, show_technical, ahorro_meta_pct = setup_sidebar()
finanzas_ok, finanzas_msg = is_path_writable(finanzas_path)
render_hero(finanzas_ok, finanzas_msg)

if "import_mode" not in st.session_state:
    st.session_state.import_mode = "Archivo Excel"

mode_cols = st.columns(2)
with mode_cols[0]:
    if st.button("Archivo Excel", type="primary" if st.session_state.import_mode == "Archivo Excel" else "secondary", width="stretch"):
        st.session_state.import_mode = "Archivo Excel"
with mode_cols[1]:
    if st.button("Captura bancaria", type="primary" if st.session_state.import_mode == "Captura bancaria" else "secondary", width="stretch"):
        st.session_state.import_mode = "Captura bancaria"

mode = st.session_state.import_mode
render_mode_panel(mode)

init_upload_state()

if mode == "Archivo Excel":
    render_step_title(1, "Cargar archivo de gastos e ingresos")
    st.markdown(
        "<div class='subtle'>Subi un archivo .xlsx o .xls para revisarlo antes de importar. Hoy el parseo esta preparado para exportes de Mercado Pago.</div>",
        unsafe_allow_html=True,
    )
    render_upload_dropzone_hint()

    source_bytes: bytes | None = None
    uploaded_file = st.file_uploader(
        "Arrastra y suelta el archivo de movimientos aqui (.xlsx o .xls), o usa Browse files",
        type=["xlsx", "xls"],
        key=st.session_state.uploader_key,
        help="Si no toma el arrastre, soltalo dentro del recuadro punteado del uploader.",
    )

    if uploaded_file is not None and st.session_state.upload_at is None:
        st.session_state.upload_at = pd.Timestamp.now().to_pydatetime()
    if uploaded_file is not None:
        source_bytes = uploaded_file.getvalue()

    if uploaded_file is not None:
        render_upload_meta(uploaded_file)
    else:
        st.markdown("**Alternativa:** cargar desde ruta local (si el drag del navegador falla).")
        local_path = st.text_input(
            "Ruta del archivo de movimientos",
            value=str((Path.home() / "Downloads").resolve()),
            help="Pega la ruta completa del archivo .xlsx/.xls.",
        )
        if st.button("Cargar archivo desde ruta"):
            local_file = Path(local_path.strip().strip('"'))
            if not local_file.exists() or not local_file.is_file():
                st.error("La ruta no existe o no es un archivo.")
            elif local_file.suffix.lower() not in {".xlsx", ".xls"}:
                st.error("El archivo debe ser .xlsx o .xls.")
            else:
                source_bytes = local_file.read_bytes()
                st.session_state.upload_at = pd.Timestamp.now().to_pydatetime()
                st.success(f"Archivo cargado desde ruta: {local_file.name}")

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
                st.warning(f"No encontramos FINANZAS.xlsx en: {finanzas_path}")

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
    render_step_title(1, "Cargar captura bancaria")
    current_year = datetime.now().year
    st.markdown(
        f"<div class='subtle'>Subi hasta 3 capturas bancarias. Carga manual asistida, con precarga OCR opcional cuando este disponible. Fecha dd/mm usa ano {current_year}.</div>",
        unsafe_allow_html=True,
    )
    bna_files = st.file_uploader(
        "Arrastra capturas bancarias (.jpg/.jpeg/.png) - maximo 3",
        type=["jpg", "jpeg", "png"],
        key="bna_capture_uploader",
        accept_multiple_files=True,
    )

    bna_result: BnaParseResult | None = None
    bna_plan: ImportPlan | None = None
    if "bna_date_mode" not in st.session_state:
        st.session_state.bna_date_mode = "No importar fechas ya existentes"

    st.caption("Filtro de fechas para capturas")
    bna_mode_cols = st.columns(2)
    with bna_mode_cols[0]:
        if st.button(
            "Excluir fechas existentes",
            type="primary" if st.session_state.bna_date_mode == "No importar fechas ya existentes" else "secondary",
            width="stretch",
            key="bna_mode_existing",
        ):
            st.session_state.bna_date_mode = "No importar fechas ya existentes"
    with bna_mode_cols[1]:
        if st.button(
            "Solo posteriores",
            type="primary" if st.session_state.bna_date_mode == "Solo importar movimientos posteriores a la ultima fecha" else "secondary",
            width="stretch",
            key="bna_mode_after",
        ):
            st.session_state.bna_date_mode = "Solo importar movimientos posteriores a la ultima fecha"

    bna_date_mode_label = st.session_state.bna_date_mode
    bna_date_filter_mode = (
        DATE_FILTER_MODE_SKIP_EXISTING
        if bna_date_mode_label.startswith("No importar")
        else DATE_FILTER_MODE_AFTER_MAX
    )

    parsed_df: pd.DataFrame | None = None
    if bna_files:
        selected_files = bna_files[:3]
        if len(bna_files) > 3:
            st.warning("Se tomaran solo las primeras 3 capturas.")

        preview_cols = st.columns(len(selected_files))
        for idx, capture_file in enumerate(selected_files):
            with preview_cols[idx]:
                preview_image = _build_bna_private_preview(capture_file.getvalue())
                st.image(preview_image, caption=f"Vista privada: {capture_file.name}", width=240)

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
                ocr_warnings.append(f"[{capture_file.name}] Precarga OCR no disponible: {exc}")

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

        st.markdown("#### Paso 2 - Completar movimientos (manual)")
        st.caption(f"Fecha en formato dd/mm (se completa ano {current_year}). Usa `+` para ingresos y `-` para gastos. `CR INTERB` se toma como Ingreso.")
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
                parse_errors.append(f"Fila {idx + 1}: fecha invalida '{row.get('fecha', '')}'.")
                continue

            try:
                monto_signed = float(parse_number_ar(monto_raw))
            except Exception:
                parse_errors.append(f"Fila {idx + 1}: monto invalido '{monto_raw}'.")
                continue
            monto_val = abs(monto_signed)
            if monto_val <= 0:
                continue

            tipo = "Ingreso" if monto_signed > 0 else "Gasto"
            if "CR INTERB" in desc.upper():
                tipo = "Ingreso"

            desc_safe = desc if desc else "Movimiento bancario"
            desc_norm = normalize_description(desc_safe)
            ref_seed = f"bank_capture|manual|{fecha.strftime('%Y-%m-%d')}|{tipo}|{monto_val:.2f}|{desc_norm}"
            mp_ref = f"cap_{sha1(ref_seed.encode('utf-8')).hexdigest()[:12]}"
            if mp_ref in seen_bna_refs:
                duplicate_rows_in_batch += 1
                continue

            seen_bna_refs.add(mp_ref)
            note = f"origen=captura_bancaria; banco=bna; mp_ref={mp_ref}"
            rows.append(
                {
                    "date": fecha,
                    "tipo": tipo,
                    "categoria": "Otros",
                    "subcategoria": "",
                    "regla_categoria": "CAPTURA_BANCARIA_MANUAL_V1",
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
            st.markdown("#### Paso 3 - Preview captura bancaria")
            render_preview_table(
                format_preview_df(parsed_df).head(100),
                ["date", "tipo_visual", "categoria", "descripcion", "monto", "compartido"],
                {
                    "date": "Fecha",
                    "tipo_visual": "Tipo",
                    "categoria": "Categoria",
                    "descripcion": "Descripcion",
                    "monto": "Monto",
                    "compartido": "Compartido",
                },
            )
            if duplicate_rows_in_batch > 0:
                st.info(f"Se detectaron y omitieron {duplicate_rows_in_batch} duplicados entre capturas.")
        else:
            st.info("Completa al menos una fila valida para previsualizar e importar.")

    if parsed_df is not None and not parsed_df.empty:
        if finanzas_path.strip() and Path(finanzas_path).exists():
            try:
                bna_plan = build_import_plan(parsed_df, Path(finanzas_path), date_filter_mode=bna_date_filter_mode)
            except Exception as exc:
                st.error(f"No se pudo calcular el plan de importacion de captura bancaria: {exc}")
        else:
            st.warning(f"No encontramos FINANZAS.xlsx en: {finanzas_path}")

        to_import = int(bna_plan.to_import_df.shape[0]) if bna_plan is not None else int(parsed_df.shape[0])
        dupes_ref = int(bna_plan.duplicate_mp_ref_df.shape[0]) if bna_plan is not None else 0
        dupes_key = int(bna_plan.duplicate_compound_df.shape[0]) if bna_plan is not None else 0
        st.info(
            f"Cargados: {int(parsed_df.shape[0])} | A importar: {to_import} | Duplicados ref: {dupes_ref} | Duplicados clave: {dupes_key}"
        )

        confirm_bna = st.checkbox("Confirmo importacion de movimientos de captura bancaria", value=False)
        if st.button("Importar captura bancaria", type="primary", disabled=not confirm_bna):
            import_df = bna_plan.to_import_df if bna_plan is not None else parsed_df
            try:
                result = import_into_finanzas_workbook(
                    import_df,
                    Path(finanzas_path),
                    date_filter_mode=bna_date_filter_mode,
                    default_category=default_category,
                )
                if result.note == "No rows to import":
                    st.info("No hay filas nuevas para importar desde la captura bancaria.")
                else:
                    st.success(
                        "Importacion de captura bancaria completada: "
                        f"{result.added_rows} agregadas, "
                        f"{result.skipped_rows_by_date} saltadas por fecha, "
                        f"{result.duplicate_rows_mp_ref} duplicadas por referencia, "
                        f"{result.duplicate_rows_compound_key} duplicadas por clave."
                    )
                    if result.backup_path is not None:
                        st.info("Backup generado:")
                        st.code(str(result.backup_path))
            except Exception as exc:
                st.error(f"Ocurrio un error importando la captura bancaria a FINANZAS.xlsx: {exc}")

render_footer()
