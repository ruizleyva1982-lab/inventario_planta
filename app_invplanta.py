import streamlit as st
import pandas as pd
import json
import os
import io
from datetime import date, datetime

# ──────────────────────────────────────────────
# CONFIGURACIÓN GENERAL Y CONSTANTES
# ──────────────────────────────────────────────
INVENTARIO_PATH = "inventario.xlsx"
REGISTROS_PATH = "registros_conteo.json"
EXCEL_REGISTROS = "registros_conteo.xlsx"
CONTEOS = [1, 2, 3, 4, 5]  # Números de conteo configurados

# Unidades de medida frecuentes para los formularios
UM_SUGERIDAS = [
    "UNIDAD (BIENES)", "KILOGRAMO", "UND", "PQT x 100UND", "CAJA", 
    "BOLSA x 1KGS", "BOTELLA", "ROLLO", "LITRO", "PQT x 1000UND", 
    "GALON", "CAJA x 10KGS", "PQT x 50UND", "SACO x 25KGS", "PAR"
]

st.set_page_config(page_title="Sistema de Dosimetría e Inventario", page_icon="🧪", layout="wide")

# ──────────────────────────────────────────────
# FUNCIONES DE CARGA Y PERSISTENCIA
# ──────────────────────────────────────────────
@st.cache_data(ttl=60)
def cargar_inventario() -> pd.DataFrame:
    """
    Carga el inventario maestro desde inventario.xlsx, limpia columnas/valores 
    y crea una columna 'DISPLAY' para las búsquedas.
    """
    try:
        df = pd.read_excel(INVENTARIO_PATH)
        df.columns = [c.strip().upper() for c in df.columns]

        # Limpieza y conversión a texto
        df["CÓDIGO"] = df["CÓDIGO"].astype(str).str.strip()
        df["INSUMO"] = df["INSUMO"].astype(str).str.strip()
        df["UM"] = df["UM"].fillna("UNIDAD (BIENES)").astype(str).str.strip().str.upper()

        # Clave única para buscadores en Streamlit
        df["DISPLAY"] = df["CÓDIGO"] + " — " + df["INSUMO"]
        return df
    except Exception as e:
        st.error(f"⚠️ No se pudo cargar el maestro de inventario: {e}")
        return pd.DataFrame(columns=["CÓDIGO", "INSUMO", "UM", "DISPLAY"])


def guardar_inventario(df: pd.DataFrame) -> bool:
    """
    Guarda las columnas base del maestro en inventario.xlsx y limpia la caché.
    """
    try:
        cols_guardar = ["CÓDIGO", "INSUMO", "UM"]
        df_export = df[cols_guardar].copy()
        df_export.to_excel(INVENTARIO_PATH, index=False)
        cargar_inventario.clear()
        return True
    except PermissionError:
        st.error("⚠️ No se pudo guardar porque **inventario.xlsx** está abierto en Excel. Ciérralo e intentalo de nuevo.")
        return False
    except Exception as e:
        st.error(f"⚠️ Error al guardar el maestro de inventario: {e}")
        return False


def cargar_registros() -> dict:
    if os.path.exists(REGISTROS_PATH):
        try:
            with open(REGISTROS_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            st.error(f"Error al leer el archivo JSON de registros: {e}")
            return {}
    return {}


def guardar_registros(data: dict):
    try:
        with open(REGISTROS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        guardar_excel_registros(data)
    except Exception as e:
        st.error(f"Error al guardar los registros en JSON: {e}")


def guardar_excel_registros(data: dict):
    if not data:
        return
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        wb.remove(wb.active)
        fechas = sorted(set(v["fecha"] for v in data.values()), reverse=True)
        color_h = "1A3A5C"
        color_s = "2C6FB5"
        thin = Side(style="thin", color="CCCCCC")
        borde = Border(left=thin, right=thin, top=thin, bottom=thin)

        for fecha in fechas:
            regs = sorted([v for v in data.values() if v["fecha"] == fecha], key=lambda x: x["insumo"])
            ws = wb.create_sheet(title=fecha.replace("-", ""))
            
            # Encabezado principal
            ws.merge_cells("A1:I1")
            ws["A1"] = f"REGISTRO DE CONTEO — {fecha}"
            ws["A1"].font = Font(name="Arial", bold=True, size=13, color="FFFFFF")
            ws["A1"].fill = PatternFill("solid", fgColor=color_h)
            ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 28
            
            # Subencabezados
            encabezados = ["CÓDIGO", "INSUMO", "UM", "CONTEO 1", "CONTEO 2", "CONTEO 3", "CONTEO 4", "CONTEO 5", "TOTAL"]
            for col, h in enumerate(encabezados, 1):
                c = ws.cell(row=2, column=col, value=h)
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=color_s)
                c.alignment = Alignment(horizontal="center")
                c.border = borde
            ws.row_dimensions[2].height = 20

            # Filas de datos
            for ri, reg in enumerate(regs, 3):
                vals = [
                    reg.get("codigo", ""),
                    reg.get("insumo", ""),
                    reg.get("um", ""),
                    reg.get("mesas", {}).get("1", 0),
                    reg.get("mesas", {}).get("2", 0),
                    reg.get("mesas", {}).get("3", 0),
                    reg.get("mesas", {}).get("4", 0),
                    reg.get("mesas", {}).get("5", 0),
                    reg.get("total", 0)
                ]
                for ci, val in enumerate(vals, 1):
                    c = ws.cell(row=ri, column=ci, value=val)
                    c.font = Font(name="Arial", size=10, bold=(ci == 9), color=(color_h if ci == 9 else "000000"))
                    c.border = borde
                    c.alignment = Alignment(horizontal="left" if ci == 2 else "center")
                if ri % 2 == 0:
                    for ci in range(1, 10):
                        ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor="EEF2F7")

            # Fila de Totales
            uf = 2 + len(regs) + 1
            ws.merge_cells(f"A{uf}:C{uf}")
            ws[f"A{uf}"] = "TOTAL GENERAL"
            ws[f"A{uf}"].font = Font(name="Arial", bold=True, color="FFFFFF")
            ws[f"A{uf}"].fill = PatternFill("solid", fgColor=color_h)
            ws[f"A{uf}"].alignment = Alignment(horizontal="center")
            
            for ci in range(4, 10):
                l = get_column_letter(ci)
                c = ws.cell(row=uf, column=ci, value=f"=SUM({l}3:{l}{uf-1})")
                c.font = Font(name="Arial", bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor=color_h)
                c.alignment = Alignment(horizontal="center")
                c.border = borde

            for ci, w in enumerate([14, 40, 15, 12, 12, 12, 12, 12, 12], 1):
                ws.column_dimensions[get_column_letter(ci)].width = w

        wb.save(EXCEL_REGISTROS)
    except PermissionError:
        st.error("⚠️ No se pudo actualizar `registros_conteo.xlsx` porque está abierto por otro programa.")
    except Exception as e:
        st.error(f"Error inesperado al generar Excel de registros: {e}")


def excel_bytes(df: pd.DataFrame) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    
    wb = Workbook()
    ws = wb.active
    color_s = "2C6FB5"
    thin = Side(style="thin", color="CCCCCC")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    cols = df.columns.tolist()

    for ci, col in enumerate(cols, 1):
        c = ws.cell(row=1, column=ci, value=col)
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=color_s)
        c.alignment = Alignment(horizontal="center")
        c.border = borde

    for ri, row in enumerate(df.itertuples(index=False), 2):
        for ci, val in enumerate(row, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = Font(name="Arial", size=10)
            c.border = borde
            c.alignment = Alignment(horizontal="left" if ci == 2 else "center")
        if ri % 2 == 0:
            for ci in range(1, len(cols) + 1):
                ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor="EEF2F7")

    for ci in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 20

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ──────────────────────────────────────────────
# ESTILOS VISUALES
# ──────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #f5f7fa; }
    .block-container { padding-top: 1.5rem; }
    h1 { color: #1a3a5c; }
    h2, h3 { color: #2c5282; }
    .metric-box {
        background: white; border-radius: 10px; padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08); text-align: center;
    }
    .total-box {
        background: linear-gradient(135deg,#1a3a5c,#2c6fb5); color: white;
        border-radius: 12px; padding: 20px; text-align: center;
        font-size: 2rem; font-weight: 700;
        box-shadow: 0 4px 12px rgba(44,111,181,0.3);
    }
    div[data-testid="stTabs"] button { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# CABECERA Y ESTRUCTURA PRINCIPAL
# ──────────────────────────────────────────────
st.title("Sistema de Inventario en Planta")
st.markdown("---")

tab1, tab2, tab3 = st.tabs([
    "📋 Registro de Conteo",
    "🔍 Consulta por Fecha",
    "⚙️ Gestión de Inventario"
])

# ══════════════════════════════════════════════
# TAB 1 — REGISTRO DE CONTEO
# ══════════════════════════════════════════════
with tab1:
    df_inv = cargar_inventario()
    registros = cargar_registros()
    st.subheader("📋 Registro de Conteo General")

    col_fecha, col_buscar = st.columns([2, 4])
    with col_fecha:
        fecha_sel = st.date_input("📅 Fecha de conteo", value=date.today(), key="fecha_reg")
        fecha_str = fecha_sel.strftime("%Y-%m-%d")
    
    with col_buscar:
        opcion_sel = st.selectbox(
            "🔍 Buscar por Código o Insumo",
            ["-- Seleccione un insumo --"] + df_inv["DISPLAY"].tolist(),
            key="insumo_sel",
            help="Escribe el código o el nombre del insumo para buscar"
        )

    if opcion_sel != "-- Seleccione un insumo --":
        fila = df_inv[df_inv["DISPLAY"] == opcion_sel].iloc[0]
        codigo = fila["CÓDIGO"]
        insumo_nombre = fila["INSUMO"]
        um_inventario = fila["UM"]

        clave = f"{fecha_str}__{codigo}"
        existente = registros.get(clave, {})
        mesas_previas = existente.get("mesas", {str(c_num): 0 for c_num in CONTEOS})

        um_guardada = existente.get("um", um_inventario)
        
        opciones_um_dinamicas = list(dict.fromkeys([
            um_guardada, um_inventario, "UNIDAD (BIENES)", "KILOGRAMO", "UND", 
            "ROLLO", "CAJA", "BOLSA", "LITRO", "MOLDES", "PLANCHA"
        ]))
        idx_um = opciones_um_dinamicas.index(um_guardada) if um_guardada in opciones_um_dinamicas else 0

        c1, c2, c3 = st.columns(3)
        c1.markdown(f"<div class='metric-box'>🔑 <b>Código</b><br>{codigo}</div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='metric-box'>📦 <b>Insumo</b><br>{insumo_nombre}</div>", unsafe_allow_html=True)
        with c3:
            um = st.selectbox("⚖️ Unidad de Medida (UM)", opciones_um_dinamicas, index=idx_um, key=f"um_sel_{clave}")

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 🔢 Desglose de Conteos")
        
        cols_conteos = st.columns(5)
        valores_conteo = {}
        for i, c_num in enumerate(CONTEOS):
            with cols_conteos[i]:
                v = st.number_input(
                    f"Conteo {c_num}",
                    min_value=0.0,
                    value=float(mesas_previas.get(str(c_num), 0)),
                    step=0.5,
                    key=f"conteo_{c_num}_{clave}"
                )
                valores_conteo[str(c_num)] = v

        total = sum(valores_conteo.values())
        st.markdown(f"<div class='total-box'>TOTAL DE CONTEO: {total:,.2f} {um}</div>", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        col_guard, col_eliminar = st.columns([3, 1])
        with col_guard:
            if st.button("💾 Guardar Registro", use_container_width=True, type="primary"):
                registros[clave] = {
                    "fecha": fecha_str,
                    "codigo": codigo,
                    "insumo": insumo_nombre,
                    "um": um,
                    "mesas": valores_conteo,
                    "total": total,
                    "updated": datetime.now().isoformat()
                }
                guardar_registros(registros)
                st.success(f"✅ Registro guardado para **[{codigo}] {insumo_nombre}** el **{fecha_str}**")
                st.info("📊 Archivo de registros actualizado correctamente.")
                st.balloons()

        with col_eliminar:
            if clave in registros:
                if st.button("🗑️ Eliminar Registro", use_container_width=True, type="secondary"):
                    del registros[clave]
                    guardar_registros(registros)
                    st.warning(f"⚠️ Registro eliminado para **[{codigo}] {insumo_nombre}**")
                    st.rerun()

    st.markdown("---")
    registros_dia = [v for v in registros.values() if v.get("fecha") == fecha_str]
    
    if registros_dia:
        st.markdown(f"### 📑 Registros guardados del día: {fecha_str} ({len(registros_dia)} insumos)")
        df_dia = pd.DataFrame(registros_dia)[["codigo", "insumo", "mesas", "total", "um"]]
        
        for c_num in CONTEOS:
            df_dia[f"Conteo {c_num}"] = df_dia["mesas"].apply(lambda x: x.get(str(c_num), 0))
            
        df_dia = df_dia.drop(columns=["mesas"]).rename(columns={
            "codigo": "Código",
            "insumo": "Insumo",
            "total": "Total",
            "um": "UM"
        })
        
        columnas_ordenadas = ["Código", "Insumo", "UM"] + [f"Conteo {c_num}" for c_num in CONTEOS] + ["Total"]
        st.dataframe(df_dia[columnas_ordenadas].sort_values("Insumo"), use_container_width=True, hide_index=True)
    else:
        st.info(f"📭 No hay registros para el {fecha_str}")

# ══════════════════════════════════════════════
# TAB 2 — CONSULTA POR FECHA
# ══════════════════════════════════════════════
with tab2:
    st.subheader("🔍 Consulta de Inventario por Fecha")
    registros = cargar_registros()

    if not registros:
        st.info("📭 Aún no hay registros guardados en la base de datos.")
    else:
        fecha_consulta = st.date_input("📅 Seleccione la fecha", value=date.today(), key="fecha_consulta")
        fecha_consulta_str = fecha_consulta.strftime("%Y-%m-%d")
        resultados = [v for v in registros.values() if v.get("fecha") == fecha_consulta_str]

        if resultados:
            st.success(f"✅ **{len(resultados)}** registros encontrados para la fecha **{fecha_consulta_str}**")
            df_res = pd.DataFrame(resultados)
            
            for c_num in CONTEOS:
                df_res[f"Conteo {c_num}"] = df_res["mesas"].apply(lambda x: x.get(str(c_num), 0))
                
            df_res = df_res.drop(columns=["mesas", "updated"], errors="ignore").rename(columns={
                "fecha": "Fecha",
                "codigo": "Código",
                "insumo": "Insumo",
                "total": "Total",
                "um": "UM"
            })
            
            cols_res = ["Código", "Insumo", "UM"] + [f"Conteo {c_num}" for c_num in CONTEOS] + ["Total"]
            df_res = df_res[cols_res].sort_values("Insumo")
            
            st.dataframe(df_res, use_container_width=True, hide_index=True)
            
            c1, c2 = st.columns(2)
            c1.metric("📦 Insumos contados", len(resultados))
            c2.metric("⚖️ Total acumulado", f"{df_res['Total'].sum():,.2f}")
            
            st.download_button(
                "📥 Descargar Excel del Día",
                data=excel_bytes(df_res),
                file_name=f"conteo_{fecha_consulta_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning(f"📭 No hay registros para el **{fecha_consulta_str}**")

# ══════════════════════════════════════════════
# TAB 3 — GESTIÓN DE INVENTARIO
# ══════════════════════════════════════════════
with tab3:
    st.subheader("⚙️ Mantenimiento de Catálogo de Insumos")
    df_inv = cargar_inventario()

    subtab1, subtab2, subtab3 = st.tabs([
        "➕ Agregar Insumo",
        "✏️ Editar Insumo",
        "🗑️ Eliminar Insumo"
    ])

    # ------------------------------------------
    # SUBTAB 1: AGREGAR INSUMO
    # ------------------------------------------
    with subtab1:
        st.markdown("#### ➕ Registrar nuevo producto en `inventario.xlsx`")
        
        ultimo_num = len(df_inv) + 1
        codigo_sugerido = f"E100{ultimo_num:04d}"

        with st.form("form_nuevo_insumo", clear_on_submit=True):
            col_a1, col_a2 = st.columns([1, 2])
            with col_a1:
                nuevo_codigo = st.text_input("🔑 Código", value=codigo_sugerido, help="Debe ser un código único")
            with col_a2:
                nuevo_insumo = st.text_input("📦 Nombre del Insumo", help="Ejemplo: HARINA DE TRIGO 25KG")

            col_a3, col_a4 = st.columns(2)
            with col_a3:
                nueva_um_select = st.selectbox("⚖️ Unidad de Medida (UM)", UM_SUGERIDAS)
            with col_a4:
                nueva_um_custom = st.text_input("✍️ O escribe una UM personalizada (Opcional)")

            btn_crear = st.form_submit_button("➕ Guardar Nuevo Insumo", type="primary", use_container_width=True)

        if btn_crear:
            nuevo_codigo = nuevo_codigo.strip().upper()
            nuevo_insumo = nuevo_insumo.strip().upper()
            um_final = nueva_um_custom.strip().upper() if nueva_um_custom.strip() else nueva_um_select.upper()

            if not nuevo_codigo or not nuevo_insumo:
                st.error("⚠️ El código y el nombre del insumo no pueden estar vacíos.")
            elif nuevo_codigo in df_inv["CÓDIGO"].values:
                st.error(f"⚠️ El código **{nuevo_codigo}** ya existe en el inventario.")
            else:
                nueva_fila = pd.DataFrame([{"CÓDIGO": nuevo_codigo, "INSUMO": nuevo_insumo, "UM": um_final}])
                df_inv_actualizado = pd.concat([df_inv, nueva_fila], ignore_index=True)
                
                if guardar_inventario(df_inv_actualizado):
                    st.success(f"✅ Insumo **[{nuevo_codigo}] {nuevo_insumo}** creado exitosamente.")
                    st.rerun()

    # ------------------------------------------
    # SUBTAB 2: EDITAR INSUMO
    # ------------------------------------------
    with subtab2:
        st.markdown("#### ✏️ Editar producto existente en `inventario.xlsx`")
        
        insumo_editar_sel = st.selectbox(
            "🔍 Seleccionar insumo a editar",
            ["-- Seleccione un insumo --"] + df_inv["DISPLAY"].tolist(),
            key="edit_insumo_sel"
        )

        if insumo_editar_sel != "-- Seleccione un insumo --":
            fila_edit = df_inv[df_inv["DISPLAY"] == insumo_editar_sel].iloc[0]
            codigo_edit = fila_edit["CÓDIGO"]
            insumo_edit = fila_edit["INSUMO"]
            um_edit = fila_edit["UM"]

            with st.form("form_editar_insumo"):
                st.info(f"Modificando insumo con código: **{codigo_edit}**")
                edit_insumo_nombre = st.text_input("📦 Nombre del Insumo", value=insumo_edit)
                
                idx_um_edit = UM_SUGERIDAS.index(um_edit) if um_edit in UM_SUGERIDAS else 0
                edit_um_select = st.selectbox("⚖️ Unidad de Medida (UM)", UM_SUGERIDAS, index=idx_um_edit)
                edit_um_custom = st.text_input("✍️ O escribir UM personalizada", value="" if um_edit in UM_SUGERIDAS else um_edit)

                btn_guardar_edit = st.form_submit_button("💾 Guardar Cambios", type="primary", use_container_width=True)

            if btn_guardar_edit:
                edit_insumo_nombre = edit_insumo_nombre.strip().upper()
                um_edit_final = edit_um_custom.strip().upper() if edit_um_custom.strip() else edit_um_select.upper()

                if not edit_insumo_nombre:
                    st.error("⚠️ El nombre del insumo no puede estar vacío.")
                else:
                    idx_m = df_inv[df_inv["CÓDIGO"] == codigo_edit].index[0]
                    df_inv.loc[idx_m, "INSUMO"] = edit_insumo_nombre
                    df_inv.loc[idx_m, "UM"] = um_edit_final

                    if guardar_inventario(df_inv):
                        st.success(f"✅ Insumo **[{codigo_edit}]** actualizado correctamente.")
                        st.rerun()

    # ------------------------------------------
    # SUBTAB 3: ELIMINAR INSUMO
    # ------------------------------------------
    with subtab3:
        st.markdown("#### 🗑️ Eliminar insumo de `inventario.xlsx`")
        
        insumo_del_sel = st.selectbox(
            "🔍 Seleccionar insumo a eliminar",
            ["-- Seleccione un insumo --"] + df_inv["DISPLAY"].tolist(),
            key="del_insumo_sel"
        )

        if insumo_del_sel != "-- Seleccione un insumo --":
            fila_del = df_inv[df_inv["DISPLAY"] == insumo_del_sel].iloc[0]
            codigo_del = fila_del["CÓDIGO"]
            insumo_del = fila_del["INSUMO"]
            um_del = fila_del["UM"]

            st.warning(f"⚠️ ¿Estás seguro de que deseas eliminar **[{codigo_del}] {insumo_del}** ({um_del})?")
            st.error("Esta acción eliminará el producto del catálogo maestro `inventario.xlsx`.")

            if st.button("🗑️ Confirmar y Eliminar Insumo", type="primary", use_container_width=True):
                df_inv_filtrado = df_inv[df_inv["CÓDIGO"] != codigo_del]
                
                if guardar_inventario(df_inv_filtrado):
                    st.success(f"✅ Insumo **[{codigo_del}] {insumo_del}** eliminado correctamente.")
                    st.rerun()
