import streamlit as st
import anthropic
import base64
import json
import re
import io
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# ── Configuración de página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Facturas del Exterior",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS personalizado ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 1.5rem 2rem;
        border-radius: 10px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.8rem; }
    .main-header p { color: #b8d4f0; margin: 0.3rem 0 0; font-size: 0.95rem; }
    .step-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #2d6a9f;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        margin-bottom: 1rem;
    }
    .step-card h3 { color: #1e3a5f; margin: 0 0 0.5rem; font-size: 1rem; }
    .success-box {
        background: #f0fdf4;
        border: 1px solid #86efac;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #166534;
    }
    .warning-box {
        background: #fffbeb;
        border: 1px solid #fcd34d;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #92400e;
    }
    .info-box {
        background: #eff6ff;
        border: 1px solid #93c5fd;
        border-radius: 8px;
        padding: 0.8rem 1rem;
        color: #1e40af;
    }
    .metric-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
    }
    .metric-card .value { font-size: 1.6rem; font-weight: 700; color: #1e3a5f; }
    .metric-card .label { font-size: 0.8rem; color: #64748b; margin-top: 0.2rem; }
    .empresa-tag {
        display: inline-block;
        background: #dbeafe;
        color: #1e40af;
        border-radius: 12px;
        padding: 2px 10px;
        font-size: 0.8rem;
        font-weight: 600;
    }
    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# ── API Key ───────────────────────────────────────────────────────────────────
API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
if not API_KEY:
    st.error("⚠️ No se encontró la API key de Anthropic. Configúrala en Streamlit Secrets como `ANTHROPIC_API_KEY`.")
    st.stop()

client = anthropic.Anthropic(api_key=API_KEY)

# ── Session state ─────────────────────────────────────────────────────────────
if "tipo_cambio" not in st.session_state:
    st.session_state.tipo_cambio = {}
if "filas" not in st.session_state:
    st.session_state.filas = []
if "empresa_activa" not in st.session_state:
    st.session_state.empresa_activa = "EMPRESA1"
if "mes_activo" not in st.session_state:
    st.session_state.mes_activo = ""
if "nps_cruzados" not in st.session_state:
    st.session_state.nps_cruzados = {}

# ── Helpers ───────────────────────────────────────────────────────────────────
def pdf_to_base64(file_bytes: bytes) -> str:
    return base64.standard_b64encode(file_bytes).decode("utf-8")

def get_tc(fecha_str: str, tc_dict: dict) -> float:
    """Devuelve el tipo de cambio venta para una fecha dada."""
    try:
        d = datetime.strptime(fecha_str, "%Y-%m-%d")
        day = d.day
        return tc_dict.get(day, 0.0)
    except:
        return 0.0

def calcular_fila(usd: float, tc: float) -> dict:
    sin_igv = round(usd * tc, 2)
    igv_soles = round(sin_igv * 0.18, 2)
    igv_usd = round(usd * 0.18, 2)
    total_soles = round(sin_igv + igv_soles, 2)
    renta = round(sin_igv * 0.30, 2)
    renta_usd = round(usd * 0.30, 2)
    return {
        "sin_igv": sin_igv,
        "igv_soles": igv_soles,
        "igv_usd": igv_usd,
        "total_soles": total_soles,
        "renta": renta,
        "renta_usd": renta_usd,
    }

# ── Extracción con IA ─────────────────────────────────────────────────────────
def extraer_tipo_cambio(pdf_bytes: bytes) -> dict:
    """Lee el PDF de tipo de cambio SUNAT y devuelve {dia: tc_venta}."""
    b64 = pdf_to_base64(pdf_bytes)
    prompt = """Extrae el tipo de cambio VENTA de cada día del mes de este PDF de SUNAT.
Responde SOLO con JSON válido sin markdown, formato exacto:
{"1": 3.495, "2": 3.467, ...}
Solo incluye los días que tienen valor. Usa punto decimal."""
    
    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    text = resp.content[0].text.strip()
    text = re.sub(r"```json|```", "", text).strip()
    raw = json.loads(text)
    return {int(k): float(v) for k, v in raw.items()}

def extraer_invoice(pdf_bytes: bytes, tc_dict: dict, empresa: str) -> dict | None:
    """Extrae datos de un PDF de invoice usando Claude."""
    b64 = pdf_to_base64(pdf_bytes)
    prompt = f"""Extrae los datos de esta factura del exterior. Responde SOLO con JSON válido sin markdown:
{{
  "empresa": "{empresa}",
  "fecha": "YYYY-MM-DD",
  "proveedor": "nombre completo del emisor",
  "pais": "país del emisor en español",
  "domicilio": "dirección completa del emisor",
  "vat": "número VAT/RUC/Tax ID del emisor",
  "invoice": "número de factura exacto",
  "usd": número_en_dólares,
  "moneda_original": "USD o PEN u otra"
}}
IMPORTANTE: 
- Si el monto está en PEN (soles), conviértelo a USD dividiendo entre el tipo de cambio venta del día de la factura.
- El tipo de cambio disponible para este mes es por día: {json.dumps(tc_dict)}
- Si la moneda es PEN, busca el TC del día de la fecha y divide el monto PEN entre ese TC.
- El campo usd siempre debe ser el monto en dólares americanos."""

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    text = resp.content[0].text.strip()
    text = re.sub(r"```json|```", "", text).strip()
    try:
        data = json.loads(text)
        usd = float(data.get("usd", 0))
        fecha = data.get("fecha", "")
        tc = get_tc(fecha, tc_dict)
        calcs = calcular_fila(usd, tc)
        return {
            "empresa": empresa,
            "fecha": fecha,
            "proveedor": data.get("proveedor", ""),
            "pais": data.get("pais", ""),
            "domicilio": data.get("domicilio", ""),
            "vat": data.get("vat", ""),
            "invoice": data.get("invoice", ""),
            "nac": "16630-",
            "usd": usd,
            "tc": tc,
            **calcs,
            "estado": "OK",
        }
    except Exception as e:
        st.warning(f"No se pudo procesar un invoice: {e}")
        return None

def extraer_nps(pdf_bytes: bytes) -> dict | None:
    """Extrae el monto IGV y número NPS de un PDF de constancia NPS."""
    b64 = pdf_to_base64(pdf_bytes)
    prompt = """De esta constancia NPS de SUNAT extrae:
- El número NPS (Número de Pago SUNAT)
- El monto en soles (S/) del tributo IGV (código 1041)
- El período (YYYYMM)
Responde SOLO con JSON sin markdown:
{"nps": "0004803607473", "monto_igv": 7, "periodo": "202604", "tributo": "1041"}
Si el tributo es Renta (3062) igualmente inclúyelo con "tributo": "3062"."""

    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        messages=[{
            "role": "user",
            "content": [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": b64}},
                {"type": "text", "text": prompt}
            ]
        }]
    )
    text = resp.content[0].text.strip()
    text = re.sub(r"```json|```", "", text).strip()
    try:
        return json.loads(text)
    except:
        return None

def cruzar_nps_con_invoices(nps_list: list, filas: list) -> dict:
    """
    Cruza los NPS de IGV con los invoices por monto de IGV redondeado.
    Retorna {indice_fila: numero_orden}
    """
    resultado = {}
    filas_pendientes = list(range(len(filas)))
    
    for nps in nps_list:
        if nps.get("tributo") != "1041":
            continue
        monto_nps = int(round(nps["monto_igv"]))
        
        for idx in filas_pendientes:
            fila = filas[idx]
            igv_fila = int(round(fila["igv_soles"]))
            if igv_fila == monto_nps:
                resultado[idx] = nps["nps"]
                filas_pendientes.remove(idx)
                break
    
    return resultado

# ── Generación de Excel ───────────────────────────────────────────────────────
def generar_excel(filas: list, mes: str) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    empresas = list(dict.fromkeys(f["empresa"] for f in filas))

    # Estilos
    header_fill = PatternFill("solid", fgColor="1e3a5f")
    header_font = Font(bold=True, color="FFFFFF", size=9)
    total_fill = PatternFill("solid", fgColor="dbeafe")
    total_font = Font(bold=True, size=9)
    data_font = Font(size=9)
    thin = Side(style="thin", color="CBD5E1")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center")

    headers = [
        "Fecha Invoice", "Proveedor", "PAIS DE RESIDENCIA", "DOMICILIO",
        "RUC/VAT", "Invoice", "Nacionalización", "Monto USD",
        "Tipo de cambio\nSUNAT (venta)", "Monto en S/.\n(sin IGV)",
        "IGV (18%)\nen S/.", "IGV USD", "Total con\nIGV (S/.)",
        "Renta", "Renta USD"
    ]
    col_widths = [12, 28, 16, 35, 14, 22, 18, 10, 11, 13, 11, 10, 13, 11, 10]

    for empresa in empresas:
        ws = wb.create_sheet(title=f"{empresa}_{mes}")
        rows_emp = [f for f in filas if f["empresa"] == empresa]

        # Cabecera
        for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
            cell = ws.cell(row=1, column=col_idx, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
            ws.column_dimensions[get_column_letter(col_idx)].width = w
        ws.row_dimensions[1].height = 30

        # Datos
        for row_idx, fila in enumerate(rows_emp, 2):
            fecha_val = fila["fecha"]
            try:
                fecha_val = datetime.strptime(fila["fecha"], "%Y-%m-%d")
            except:
                pass

            valores = [
                fecha_val, fila["proveedor"], fila["pais"], fila["domicilio"],
                fila["vat"], fila["invoice"], fila.get("nac", "16630-"),
                fila["usd"], fila["tc"], fila["sin_igv"],
                fila["igv_soles"], fila["igv_usd"], fila["total_soles"],
                fila["renta"], fila["renta_usd"]
            ]

            for col_idx, val in enumerate(valores, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = data_font
                cell.border = border
                if col_idx == 1 and isinstance(val, datetime):
                    cell.number_format = "DD/MM/YYYY"
                    cell.alignment = center
                elif col_idx in [8, 9, 10, 11, 12, 13, 14, 15]:
                    cell.number_format = "#,##0.00"
                    cell.alignment = right
                else:
                    cell.alignment = Alignment(vertical="center", wrap_text=True)

        # Fila de totales
        tot_row = len(rows_emp) + 2
        ws.cell(row=tot_row, column=1, value="Totales").font = total_font
        ws.cell(row=tot_row, column=2, value="—").font = total_font
        for col_idx in range(1, 16):
            ws.cell(row=tot_row, column=col_idx).fill = total_fill
            ws.cell(row=tot_row, column=col_idx).border = border

        sum_cols = {8: "usd", 10: "sin_igv", 11: "igv_soles", 13: "total_soles", 14: "renta"}
        for col_idx, field in sum_cols.items():
            val = sum(f[field] for f in rows_emp)
            cell = ws.cell(row=tot_row, column=col_idx, value=round(val, 2))
            cell.font = total_font
            cell.number_format = "#,##0.00"
            cell.alignment = right

        ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ── UI Principal ──────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>🧾 Facturas del Exterior</h1>
    <p>Procesamiento automático de invoices · IGV no domiciliado · Retención de renta</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    
    empresa = st.text_input(
        "Nombre de empresa",
        value=st.session_state.empresa_activa,
        placeholder="Ej: ESSENTTA, SERPRESS...",
        help="Puedes cambiar esto para procesar otra empresa."
    )
    st.session_state.empresa_activa = empresa.upper().strip() if empresa else "EMPRESA"

    mes = st.text_input(
        "Mes / período",
        value=st.session_state.mes_activo,
        placeholder="Ej: ABRIL_2026, MAYO_2026...",
        help="Se usará como nombre de pestaña en el Excel."
    )
    st.session_state.mes_activo = mes.upper().strip() if mes else "MES_2026"

    st.divider()

    if st.session_state.filas:
        empresas_cargadas = list(dict.fromkeys(f["empresa"] for f in st.session_state.filas))
        st.markdown("**Datos en memoria:**")
        for emp in empresas_cargadas:
            n = sum(1 for f in st.session_state.filas if f["empresa"] == emp)
            st.markdown(f'<span class="empresa-tag">{emp}</span> · {n} invoices', unsafe_allow_html=True)
        
        st.divider()
        if st.button("🗑️ Limpiar todo", use_container_width=True):
            st.session_state.filas = []
            st.session_state.tipo_cambio = {}
            st.session_state.nps_cruzados = {}
            st.rerun()
    
    st.divider()
    st.markdown("**Cómo usar:**")
    st.markdown("""
1. Configura empresa y mes
2. Sube el PDF de tipo de cambio SUNAT
3. Sube los PDFs de invoices
4. Revisa y edita la tabla
5. Sube los NPS para llenar N° orden
6. Descarga el Excel
    """)

# ── Tabs principales ──────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📄 1. Tipo de cambio",
    "📥 2. Invoices",
    "✏️ 3. Revisar tabla",
    "🔢 4. N° Órdenes NPS",
    "📊 5. Exportar Excel"
])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Tipo de cambio
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="step-card"><h3>📄 Paso 1 — Tipo de cambio SUNAT</h3>Sube el PDF mensual de tipo de cambio que descargas de SUNAT. La IA extrae todos los valores automáticamente.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 1])
    with col1:
        pdf_tc = st.file_uploader("PDF Tipo de Cambio SUNAT", type=["pdf"], key="tc_uploader")
        
        if pdf_tc:
            with st.spinner("Extrayendo tipo de cambio..."):
                try:
                    tc_dict = extraer_tipo_cambio(pdf_tc.read())
                    st.session_state.tipo_cambio = tc_dict
                    st.markdown('<div class="success-box">✅ Tipo de cambio cargado correctamente.</div>', unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"Error al procesar: {e}")

    with col2:
        if st.session_state.tipo_cambio:
            st.markdown("**Tipo de cambio venta cargado:**")
            tc_df = pd.DataFrame([
                {"Día": k, "TC Venta": v}
                for k, v in sorted(st.session_state.tipo_cambio.items())
            ])
            st.dataframe(tc_df, use_container_width=True, hide_index=True, height=300)
        else:
            st.markdown('<div class="info-box">ℹ️ Aún no hay tipo de cambio cargado. Sube el PDF de SUNAT.</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Invoices
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown('<div class="step-card"><h3>📥 Paso 2 — Subir invoices</h3>Sube todos los PDFs de facturas del mes. La IA extrae automáticamente los datos de cada uno.</div>', unsafe_allow_html=True)

    if not st.session_state.tipo_cambio:
        st.markdown('<div class="warning-box">⚠️ Primero carga el tipo de cambio SUNAT en el Paso 1.</div>', unsafe_allow_html=True)
    else:
        pdfs_invoices = st.file_uploader(
            "PDFs de invoices (puedes subir varios a la vez)",
            type=["pdf"],
            accept_multiple_files=True,
            key="invoices_uploader"
        )

        if pdfs_invoices:
            empresa_actual = st.session_state.empresa_activa
            st.info(f"Se procesarán {len(pdfs_invoices)} invoice(s) para la empresa **{empresa_actual}**")
            
            if st.button("🤖 Procesar invoices con IA", type="primary", use_container_width=True):
                progress = st.progress(0)
                status = st.empty()
                nuevos = 0
                
                for i, pdf_file in enumerate(pdfs_invoices):
                    status.markdown(f"Procesando **{pdf_file.name}**...")
                    result = extraer_invoice(
                        pdf_file.read(),
                        st.session_state.tipo_cambio,
                        empresa_actual
                    )
                    if result:
                        st.session_state.filas.append(result)
                        nuevos += 1
                    progress.progress((i + 1) / len(pdfs_invoices))
                
                status.empty()
                progress.empty()
                st.markdown(f'<div class="success-box">✅ Se procesaron <strong>{nuevos}</strong> invoice(s) correctamente. Ve al Paso 3 para revisar.</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Revisar tabla
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown('<div class="step-card"><h3>✏️ Paso 3 — Revisar y editar</h3>Verifica los datos extraídos. Puedes editar cualquier celda directamente.</div>', unsafe_allow_html=True)

    if not st.session_state.filas:
        st.markdown('<div class="info-box">ℹ️ Aún no hay datos. Procesa los invoices en el Paso 2.</div>', unsafe_allow_html=True)
    else:
        filas = st.session_state.filas
        
        # Métricas resumen
        empresas_unicas = list(dict.fromkeys(f["empresa"] for f in filas))
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total invoices", len(filas))
        with col2:
            st.metric("Total USD", f"${sum(f['usd'] for f in filas):,.2f}")
        with col3:
            st.metric("Total S/. sin IGV", f"S/{sum(f['sin_igv'] for f in filas):,.2f}")
        with col4:
            st.metric("IGV total S/.", f"S/{sum(f['igv_soles'] for f in filas):,.2f}")

        st.divider()

        # Filtro por empresa
        if len(empresas_unicas) > 1:
            empresa_filtro = st.selectbox("Filtrar por empresa:", ["Todas"] + empresas_unicas)
        else:
            empresa_filtro = "Todas"

        filas_mostrar = filas if empresa_filtro == "Todas" else [f for f in filas if f["empresa"] == empresa_filtro]

        # Tabla editable
        df = pd.DataFrame(filas_mostrar)
        display_cols = ["empresa", "fecha", "proveedor", "pais", "vat", "invoice", "nac", "usd", "tc", "sin_igv", "igv_soles", "igv_usd", "total_soles", "renta", "renta_usd"]
        col_labels = {
            "empresa": "Empresa", "fecha": "Fecha", "proveedor": "Proveedor",
            "pais": "País", "vat": "RUC/VAT", "invoice": "N° Invoice",
            "nac": "Nacionalización", "usd": "USD", "tc": "T.C. Venta",
            "sin_igv": "S/. sin IGV", "igv_soles": "IGV S/.", "igv_usd": "IGV USD",
            "total_soles": "Total S/.", "renta": "Renta S/.", "renta_usd": "Renta USD"
        }
        df_display = df[display_cols].rename(columns=col_labels)
        
        edited_df = st.data_editor(
            df_display,
            use_container_width=True,
            num_rows="dynamic",
            hide_index=True,
            column_config={
                "USD": st.column_config.NumberColumn(format="%.2f"),
                "T.C. Venta": st.column_config.NumberColumn(format="%.3f"),
                "S/. sin IGV": st.column_config.NumberColumn(format="%.2f"),
                "IGV S/.": st.column_config.NumberColumn(format="%.2f"),
                "IGV USD": st.column_config.NumberColumn(format="%.2f"),
                "Total S/.": st.column_config.NumberColumn(format="%.2f"),
                "Renta S/.": st.column_config.NumberColumn(format="%.2f"),
                "Renta USD": st.column_config.NumberColumn(format="%.2f"),
            },
            key="tabla_editor"
        )

        if st.button("💾 Guardar cambios", type="primary"):
            inv_labels = {v: k for k, v in col_labels.items()}
            edited_back = edited_df.rename(columns=inv_labels)
            
            if empresa_filtro == "Todas":
                nuevas_filas = edited_back.to_dict("records")
            else:
                otras = [f for f in filas if f["empresa"] != empresa_filtro]
                nuevas_filas = otras + edited_back.to_dict("records")
            
            # Recalcular campos derivados si cambiaron USD o TC
            for fila in nuevas_filas:
                try:
                    usd = float(fila.get("usd", 0))
                    tc = float(fila.get("tc", 0))
                    calcs = calcular_fila(usd, tc)
                    fila.update(calcs)
                except:
                    pass
            
            st.session_state.filas = nuevas_filas
            st.success("✅ Cambios guardados.")
            st.rerun()

        # Agregar fila manual
        with st.expander("➕ Agregar invoice manualmente"):
            c1, c2, c3 = st.columns(3)
            with c1:
                m_empresa = st.text_input("Empresa", value=st.session_state.empresa_activa, key="m_emp")
                m_fecha = st.text_input("Fecha (YYYY-MM-DD)", key="m_fecha")
                m_proveedor = st.text_input("Proveedor", key="m_prov")
                m_pais = st.text_input("País", key="m_pais")
                m_domicilio = st.text_input("Domicilio", key="m_dom")
            with c2:
                m_vat = st.text_input("RUC/VAT", key="m_vat")
                m_invoice = st.text_input("N° Invoice", key="m_inv")
                m_usd = st.number_input("Monto USD", min_value=0.0, step=0.01, key="m_usd")
                m_tc = st.number_input("Tipo de cambio", min_value=0.0, step=0.001, format="%.3f", key="m_tc")
            with c3:
                if m_usd and m_tc:
                    calcs = calcular_fila(m_usd, m_tc)
                    st.metric("S/. sin IGV", f"{calcs['sin_igv']:,.2f}")
                    st.metric("IGV S/.", f"{calcs['igv_soles']:,.2f}")
                    st.metric("Total S/.", f"{calcs['total_soles']:,.2f}")
                    st.metric("Renta S/.", f"{calcs['renta']:,.2f}")
            
            if st.button("Agregar fila", key="btn_agregar"):
                if m_fecha and m_proveedor and m_usd:
                    calcs = calcular_fila(m_usd, m_tc)
                    st.session_state.filas.append({
                        "empresa": m_empresa.upper(),
                        "fecha": m_fecha,
                        "proveedor": m_proveedor,
                        "pais": m_pais,
                        "domicilio": m_domicilio,
                        "vat": m_vat,
                        "invoice": m_invoice,
                        "nac": "16630-",
                        "usd": m_usd,
                        "tc": m_tc,
                        **calcs,
                        "estado": "Manual"
                    })
                    st.success("✅ Fila agregada.")
                    st.rerun()

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — NPS
# ═══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown('<div class="step-card"><h3>🔢 Paso 4 — Cruzar NPS con invoices</h3>Sube las constancias NPS de IGV. El sistema identifica automáticamente qué número de orden corresponde a cada invoice por el monto de IGV.</div>', unsafe_allow_html=True)

    if not st.session_state.filas:
        st.markdown('<div class="info-box">ℹ️ Primero carga y revisa los invoices en los pasos anteriores.</div>', unsafe_allow_html=True)
    else:
        pdfs_nps = st.file_uploader(
            "PDFs de constancias NPS (solo IGV - tributo 1041)",
            type=["pdf"],
            accept_multiple_files=True,
            key="nps_uploader"
        )

        if pdfs_nps:
            st.info(f"Se procesarán {len(pdfs_nps)} constancia(s) NPS")
            
            if st.button("🔄 Cruzar NPS con invoices", type="primary", use_container_width=True):
                progress = st.progress(0)
                status = st.empty()
                nps_extraidos = []

                for i, pdf_file in enumerate(pdfs_nps):
                    status.markdown(f"Leyendo NPS **{pdf_file.name}**...")
                    resultado = extraer_nps(pdf_file.read())
                    if resultado:
                        nps_extraidos.append(resultado)
                    progress.progress((i + 1) / len(pdfs_nps))

                status.empty()
                progress.empty()

                # Solo IGV para cruzar
                nps_igv = [n for n in nps_extraidos if n.get("tributo") == "1041"]
                
                cruce = cruzar_nps_con_invoices(nps_igv, st.session_state.filas)
                
                # Aplicar N° órdenes
                for idx, nps_num in cruce.items():
                    st.session_state.filas[idx]["nac"] = f"16630-{nps_num}"
                
                st.session_state.nps_cruzados = cruce

                # Mostrar resultado
                st.markdown(f'<div class="success-box">✅ Se cruzaron <strong>{len(cruce)}</strong> de {len(st.session_state.filas)} invoices con número de orden.</div>', unsafe_allow_html=True)
                
                # Tabla resumen del cruce
                if cruce:
                    st.markdown("**Resultado del cruce:**")
                    resumen = []
                    for idx, nps_num in cruce.items():
                        fila = st.session_state.filas[idx]
                        resumen.append({
                            "Proveedor": fila["proveedor"],
                            "Invoice": fila["invoice"],
                            "IGV S/.": fila["igv_soles"],
                            "N° Orden asignado": f"16630-{nps_num}"
                        })
                    st.dataframe(pd.DataFrame(resumen), use_container_width=True, hide_index=True)

                sin_cruzar = [i for i in range(len(st.session_state.filas)) if i not in cruce]
                if sin_cruzar:
                    st.markdown("**⚠️ Invoices sin N° de orden (no se encontró NPS coincidente):**")
                    for idx in sin_cruzar:
                        f = st.session_state.filas[idx]
                        st.markdown(f"- {f['proveedor']} · {f['invoice']} · IGV S/. {f['igv_soles']:.2f}")

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 — Exportar
# ═══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.markdown('<div class="step-card"><h3>📊 Paso 5 — Exportar Excel</h3>Genera el Excel final con una pestaña por empresa, en el formato exacto de tu plantilla.</div>', unsafe_allow_html=True)

    if not st.session_state.filas:
        st.markdown('<div class="info-box">ℹ️ No hay datos para exportar. Completa los pasos anteriores.</div>', unsafe_allow_html=True)
    else:
        filas = st.session_state.filas
        empresas = list(dict.fromkeys(f["empresa"] for f in filas))
        mes = st.session_state.mes_activo or "MES_2026"

        # Resumen por empresa
        st.markdown("**Resumen por empresa:**")
        for emp in empresas:
            rows_emp = [f for f in filas if f["empresa"] == emp]
            tot_usd = sum(f["usd"] for f in rows_emp)
            tot_soles = sum(f["sin_igv"] for f in rows_emp)
            tot_igv = sum(f["igv_soles"] for f in rows_emp)
            tot_renta = sum(f["renta"] for f in rows_emp)
            sin_orden = sum(1 for f in rows_emp if f.get("nac", "16630-") == "16630-")
            
            with st.expander(f"🏢 {emp} · {len(rows_emp)} invoices", expanded=True):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total USD", f"${tot_usd:,.2f}")
                c2.metric("Total S/. sin IGV", f"S/{tot_soles:,.2f}")
                c3.metric("IGV Total S/.", f"S/{tot_igv:,.2f}")
                c4.metric("Retención Renta S/.", f"S/{tot_renta:,.2f}")
                if sin_orden > 0:
                    st.markdown(f'<div class="warning-box">⚠️ {sin_orden} invoice(s) aún sin número de orden (Nacionalización = "16630-"). Puedes descargar igual y completar después.</div>', unsafe_allow_html=True)

        st.divider()

        nombre_archivo = f"{'_'.join(empresas)}_{mes}.xlsx"
        
        if st.button("📥 Generar y descargar Excel", type="primary", use_container_width=True):
            with st.spinner("Generando Excel..."):
                excel_bytes = generar_excel(filas, mes)
            
            st.download_button(
                label="⬇️ Descargar Excel ahora",
                data=excel_bytes,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.markdown(f'<div class="success-box">✅ Excel generado: <strong>{nombre_archivo}</strong><br>Pestañas incluidas: {", ".join(f"{e}_{mes}" for e in empresas)}</div>', unsafe_allow_html=True)
