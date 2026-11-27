"""ERP académico para Pizza Pronto Italiana S.R.L.

Ejecutar en VS Code:
    pip install -r requirements_erp.txt
    streamlit run erp_pizza_pronto.py

Coloca el Excel del proyecto junto a este archivo, o indica su ruta en la
barra lateral. Las opciones de registro escriben nuevas filas en el Excel.
"""

from pathlib import Path
from datetime import date
import io
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ERP | Pizza Pronto", page_icon="🍕", layout="wide")

DEFAULT_XLSX = "proyecto_gestion_avance_t2 (1) (2).xlsx"
STATUS_OPTIONS = ["Pendiente", "Descartado", "Contactado", "Cotizado", "Venta"]
MONTHS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def clean(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def read_sheet(source, sheet, header=0):
    try:
        df = pd.read_excel(source, sheet_name=sheet, header=header, engine="openpyxl")
        # Treat empty strings and whitespace as missing, then remove entirely
        # blank rows/columns so the interface does not show Excel's empty areas.
        df = df.replace(r"^\s*$", pd.NA, regex=True)
        df = df.dropna(how="all").dropna(axis=1, how="all")
        df.columns = [clean(c) or f"Columna {i+1}" for i, c in enumerate(df.columns)]
        return df
    except Exception:
        return pd.DataFrame()


def value_after_label(df, pattern, default="No registrado"):
    for row in df.astype(str).itertuples(index=False, name=None):
        for value in row:
            if re.search(pattern, value, re.I):
                tail = re.split(r":", value, maxsplit=1)
                if len(tail) > 1 and tail[1].strip():
                    return tail[1].strip()
    return default


def save_new_row(path, sheet, row, startrow=None):
    """Append one row while retaining the existing workbook and its sheets."""
    from openpyxl import load_workbook
    wb = load_workbook(path)
    ws = wb[sheet]
    if startrow is None:
        ws.append(row)
    else:
        ws.append(row)
    wb.save(path)


def table(df, height=420):
    if df.empty:
        st.info("No hay registros para mostrar en esta hoja.")
    else:
        visible = df.replace(r"^\s*$", pd.NA, regex=True).dropna(how="all").dropna(axis=1, how="all").fillna("")
        st.dataframe(visible, use_container_width=True, hide_index=True, height=height)


def section_title(title, caption=None):
    st.title(title)
    if caption:
        st.caption(caption)


st.sidebar.title("🍕 Pizza Pronto")
st.sidebar.caption("Sistema ERP académico")
local_copy = Path(__file__).resolve().parent / DEFAULT_XLSX
downloads_copy = Path.home() / "Downloads" / DEFAULT_XLSX
default_path = local_copy if local_copy.exists() else downloads_copy
excel_path_text = st.sidebar.text_input("Ruta del archivo Excel", str(default_path))
excel_path = Path(excel_path_text).expanduser()
uploaded = st.sidebar.file_uploader("O selecciona un Excel", type=["xlsx"])
if uploaded is not None:
    workbook_source = io.BytesIO(uploaded.getvalue())
    writable = False
    st.sidebar.caption("Archivo cargado para consulta. Para registrar cambios, usa la ruta local del Excel.")
elif excel_path.exists():
    workbook_source = str(excel_path)
    writable = True
else:
    workbook_source = None
    writable = False
    st.sidebar.error("No se encuentra el Excel. Colócalo junto a este programa o cambia la ruta.")

if not writable:
    st.sidebar.info("La edición requiere abrir el archivo desde una ruta local.")

menu = st.sidebar.radio(
    "Módulos",
    ["Inicio", "Información", "CRM", "Producción", "Inventarios", "Compras y Proveedores", "Dashboard"],
)

if workbook_source is None:
    st.title("ERP | Pizza Pronto")
    st.warning("Selecciona el libro Excel del proyecto para consultar sus datos.")
    st.stop()

empresa = read_sheet(workbook_source, "PRESENTACION", header=None)
clientes = read_sheet(workbook_source, "Clientes")
proveedores = read_sheet(workbook_source, "PROVEEDORES", header=3)
ordenes = read_sheet(workbook_source, "ORDEN DE COMPRA Y SERVICIO", header=4)
maestro = read_sheet(workbook_source, "MAESTRO", header=1)
pmp = read_sheet(workbook_source, "PMP", header=0)
pap = read_sheet(workbook_source, "PAP", header=0)
pronosticos = read_sheet(workbook_source, "PRONOSTICOS", header=8)

company_name = value_after_label(empresa, r"Nombre de la empresa")
location = value_after_label(empresa, r"Ubicaci[oó]n")
sector = value_after_label(empresa, r"Rubro|Sector")

if menu == "Inicio":
    st.title("ERP | Pizza Pronto")
    st.subheader(company_name)
    st.caption(f"{location} · Sector: {sector}")
    st.markdown("Consulta integrada de ventas, producción, inventarios, compras y proveedores.")
    a, b, c, d = st.columns(4)
    total_sales = len(clientes) if not clientes.empty else 0
    unit_sales = pd.to_numeric(clientes.get("Cantidad", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    a.metric("Registros de ventas", f"{total_sales:,}")
    b.metric("Unidades vendidas", f"{unit_sales:,.0f}")
    c.metric("Materiales", f"{len(maestro):,}")
    d.metric("Órdenes de compra", f"{len(ordenes):,}")
    st.info("Elige un módulo en el menú lateral. Los datos se consultan desde el Excel del proyecto.")

elif menu == "Información":
    section_title("Información de la empresa", "Datos tomados de la hoja PRESENTACION del Excel.")
    c1, c2 = st.columns(2)
    c1.write(f"**Nombre comercial:** {company_name}")
    c1.write(f"**Ubicación:** {location}")
    c2.write(f"**Sector:** {sector}")
    c2.write("**Descripción:** Producción y venta de pizzas. La hoja de presentación no incluye una descripción desarrollada.")
    st.subheader("Productos principales")
    p1, p2 = st.columns(2)
    p1.markdown("### Pizza Hawaiana\nProducto registrado en la hoja PRODUCTOS.")
    p1.caption("La hoja PRODUCTOS indica que falta detallar actividades de preparación de masa, salsa, queso y otros componentes.")
    p2.markdown("### Pizza Americana\nProducto registrado en la hoja PRODUCTOS.")
    p2.caption("Los detalles del proceso deben completarse en el Excel para mostrar una ruta de fabricación sustentada.")
    st.subheader("Datos de empresa disponibles")
    st.write("**Responsable registrado:**", value_after_label(empresa, r"Nombre"))

elif menu == "CRM":
    section_title("CRM · Clientes y leads", "Los clientes y estados se cargan de la hoja Clientes.")
    if not clientes.empty:
        search = st.text_input("Buscar cliente", placeholder="Nombre, producto o estado")
        view = clientes.copy()
        if search:
            view = view[view.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)]
        table(view)
        counts = clientes.get("Estado", pd.Series(dtype=str)).fillna("Sin estado").value_counts()
        if not counts.empty:
            st.subheader("Clientes por estado")
            st.bar_chart(counts)
    st.subheader("Registrar lead o cliente")
    if not writable:
        st.info("Abre el Excel desde una ruta local para habilitar registros.")
    with st.form("lead_form", clear_on_submit=True):
        x1, x2, x3 = st.columns(3)
        lead_date = x1.date_input("Fecha", value=date.today())
        lead_name = x2.text_input("Nombre del cliente")
        lead_phone = x3.text_input("Teléfono")
        x4, x5, x6 = st.columns(3)
        lead_state = x4.selectbox("Estado", STATUS_OPTIONS)
        lead_product = x5.selectbox("Producto", ["Pizza Hawaiana", "Pizza Americana"])
        lead_qty = x6.number_input("Cantidad", min_value=0, step=1, value=0)
        channel = st.text_input("Canal de venta / asesor", placeholder="Ej.: WhatsApp / Asesor 1")
        submitted = st.form_submit_button("Guardar en Excel", disabled=not writable)
        if submitted:
            if not lead_name.strip():
                st.error("Escribe el nombre del cliente.")
            else:
                save_new_row(excel_path, "Clientes", [lead_date, lead_name.strip(), lead_phone.strip(), lead_state, lead_product, int(lead_qty), channel.strip()])
                st.success("Registro agregado a la hoja Clientes. Recarga la página para verlo.")
elif menu == "Producción":
    section_title("Producción", "Consultas resumidas de pronósticos, PAP, PMP, BOM, MRP y DRP.")
    tabs = st.tabs(["Pronósticos", "PAP", "PMP", "BOM / Materiales", "MRP", "DRP"])
    with tabs[0]:
        st.subheader("Pronósticos de ventas")
        table(pronosticos)
    with tabs[1]:
        st.subheader("Planes agregados de producción")
        st.caption("Vista de los cálculos anuales registrados en la hoja PAP.")
        table(pap, height=500)
    with tabs[2]:
        st.subheader("Plan maestro de producción")
        table(pmp)
        st.caption("El PMP del libro incluye requerimientos mensuales y bloques de producción semanal.")
    with tabs[3]:
        st.subheader("Maestro de materiales")
        table(maestro)
        st.markdown("**BOM Hawaina** y **BOM Americana** contienen listas de materiales y componentes en el Excel.")
        bom_product = st.selectbox("Lista de materiales", ["BOM HAWAINA", "BOM AMERICANA "])
        bom = read_sheet(workbook_source, bom_product, header=None)
        table(bom, height=300)
    with tabs[4]:
        st.subheader("MRP · Materiales y capacidades")
        sheet = st.selectbox("Hoja MRP", ["MRP PIZZA", "MRP 2 Pizza"])
        mrp = read_sheet(workbook_source, sheet, header=None)
        table(mrp, height=500)
    with tabs[5]:
        st.subheader("DRP · Distribución")
        sheet = st.selectbox("Producto", ["DRP PIZZA HAWAINA", "DRP PIZZA AMERICANA"])
        drp = read_sheet(workbook_source, sheet, header=None)
        table(drp, height=500)

elif menu == "Inventarios":
    section_title("Inventarios", "Maestro de materiales del Excel, con consulta de stock y registro de materiales.")
    if not maestro.empty:
        col_code = next((c for c in maestro.columns if "código" in c.lower() or "codigo" in c.lower()), maestro.columns[0])
        col_desc = next((c for c in maestro.columns if "descrip" in c.lower()), maestro.columns[min(1, len(maestro.columns)-1)])
        col_stock = next((c for c in maestro.columns if "inventario" in c.lower() or "stock" in c.lower()), None)
        query = st.text_input("Buscar material", placeholder="Código o descripción")
        view = maestro.copy()
        if query:
            view = view[view.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False)).any(axis=1)]
        if col_stock:
            st.metric("Unidades / materiales listados", f"{len(maestro):,}")
        table(view)
    st.subheader("Registrar nuevo material")
    with st.form("material_form", clear_on_submit=True):
        m1, m2, m3, m4 = st.columns(4)
        code = m1.text_input("Código")
        desc = m2.text_input("Descripción")
        unit = m3.text_input("Unidad", value="unidad")
        stock = m4.number_input("Inventario inicial", min_value=0.0, step=1.0)
        m5, m6, m7, m8 = st.columns(4)
        safety = m5.number_input("Stock de seguridad", min_value=0.0, step=1.0)
        lead = m6.number_input("Lead time (semanas)", min_value=0, step=1)
        lot = m7.text_input("Tamaño de lote", value="LXL")
        supplier = m8.text_input("Proveedor")
        save_material = st.form_submit_button("Guardar material", disabled=not writable)
        if save_material:
            if not code.strip() or not desc.strip():
                st.error("Completa al menos el código y la descripción.")
            else:
                save_new_row(excel_path, "MAESTRO", [code.strip(), desc.strip(), unit.strip(), stock, safety, lead, lot.strip(), supplier.strip()])
                st.success("Material agregado a MAESTRO. Recarga la página para verlo.")

elif menu == "Compras y Proveedores":
    section_title("Compras y proveedores", "Datos de las hojas ORDEN DE COMPRA Y SERVICIO y PROVEEDORES.")
    t1, t2 = st.tabs(["Órdenes", "Proveedores"])
    with t1:
        if not ordenes.empty:
            estado_col = next((c for c in ordenes.columns if "estado" in c.lower()), None)
            if estado_col:
                states = ["Todos"] + sorted(ordenes[estado_col].dropna().astype(str).unique().tolist())
                selected = st.selectbox("Filtrar por estado", states)
                view = ordenes if selected == "Todos" else ordenes[ordenes[estado_col].astype(str) == selected]
            else:
                view = ordenes
            table(view)
            amount_col = next((c for c in ordenes.columns if "monto" in c.lower()), None)
            if amount_col:
                amount = pd.to_numeric(ordenes[amount_col], errors="coerce").fillna(0).sum()
                st.metric("Monto total de órdenes", f"S/ {amount:,.2f}")
        else:
            st.info("La hoja de órdenes está vacía o no se pudo leer.")
    with t2:
        table(proveedores)

elif menu == "Dashboard":
    section_title("Dashboard", "Indicadores calculados con los registros del Excel. Las unidades se muestran en cada indicador.")
    units = pd.to_numeric(clientes.get("Cantidad", pd.Series(dtype=float)), errors="coerce").fillna(0)
    states = clientes.get("Estado", pd.Series(dtype=str)).fillna("").astype(str) if not clientes.empty else pd.Series(dtype=str)
    sales_count = int(states.str.casefold().eq("venta").sum())
    conversion = sales_count / len(clientes) * 100 if len(clientes) else 0
    month_col = pd.to_datetime(clientes.get("Fecha", pd.Series(dtype=object)), errors="coerce") if not clientes.empty else pd.Series(dtype="datetime64[ns]")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Conversión CRM", f"{conversion:.1f}%", help="Ventas / registros de clientes")
    c2.metric("Ventas registradas", f"{sales_count:,} pedidos")
    c3.metric("Unidades vendidas", f"{units.sum():,.0f} pizzas")
    c4.metric("Órdenes de compra", f"{len(ordenes):,} órdenes")
    st.subheader("Ventas por mes y producto")
    if not clientes.empty and "Producto" in clientes.columns and month_col.notna().any():
        chart_data = clientes.assign(_mes=month_col.dt.strftime("%b"), _cantidad=units)
        pivot = chart_data.pivot_table(index="_mes", columns="Producto", values="_cantidad", aggfunc="sum", fill_value=0)
        pivot = pivot.reindex([m for m in MONTHS if m in pivot.index])
        st.bar_chart(pivot)
    st.subheader("Producción planificada por mes (pizzas)")
    if not pmp.empty and len(pmp.columns) >= 13:
        prod = pmp[pmp.iloc[:, 0].astype(str).str.contains("Producción", case=False, na=False)]
        if not prod.empty:
            row = prod.iloc[0]
            chart = pd.Series([pd.to_numeric(row.iloc[i], errors="coerce") for i in range(1, 13)], index=MONTHS)
            st.bar_chart(chart.dropna())
        else:
            st.info("No se encontró una fila de producción mensual en PMP.")
    else:
        st.info("No se pudo leer el PMP para el gráfico mensual.")
    stock_col = next((c for c in maestro.columns if "inventario" in c.lower() or "stock" in c.lower()), None)
    if stock_col:
        stock_values = pd.to_numeric(maestro[stock_col], errors="coerce").fillna(0)
        st.metric("Stock de materiales", f"{stock_values.sum():,.0f} unidades registradas")
        st.caption("El Excel no define el mismo tipo de unidad para todos los materiales; el total es solo una referencia agregada.")
st.sidebar.divider()
st.sidebar.caption("Fuente: libro Excel del proyecto. Los módulos muestran datos disponibles en sus hojas.")
