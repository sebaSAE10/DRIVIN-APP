import io
import json
import os
import pickle
import re
import shutil
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pypdf
import streamlit as st

# Integración con Google Sheets
try:
    from streamlit_gsheets import GSheetsConnection
    HAS_GSHEETS = True
except ImportError:
    HAS_GSHEETS = False

# Configuración de la página Streamlit
st.set_page_config(
    page_title="Control Operativo Drivin x SAE x SmartSheet",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS personalizados
st.markdown(
    """
<style>
    .main { background-color: #f8f9fa; }
    .kpi-card {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 18px 22px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        border-left: 6px solid #1E88E5;
        margin-bottom: 10px;
    }
    .kpi-title {
        font-size: 0.85rem;
        color: #6c757d;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #212529;
        margin: 4px 0;
    }
    .kpi-sub {
        font-size: 0.82rem;
        font-weight: 600;
    }
    .kpi-success { border-left-color: #2ecc71; }
    .kpi-danger { border-left-color: #e74c3c; }
    .kpi-warning { border-left-color: #f39c12; }
    .kpi-info { border-left-color: #3498db; }
    .kpi-purple { border-left-color: #8e44ad; }
    
    .nav-header {
        font-size: 1.1rem;
        font-weight: 700;
        color: #2c3e50;
        margin-top: 15px;
        margin-bottom: 10px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 5px;
    }
</style>
""",
    unsafe_allow_html=True,
)

# RUTAS DE ARCHIVOS BASE Y DEL SISTEMA
DRIVIN_BASE_FILE = "drivin_base.xlsx"
SAE_BASE_FILE = "sae_base.xlsx"
SMARTSHEET_BASE_FILE = "smartsheet_base.xlsx"

HISTORICO_FILE = "historico_drivin.csv"
PAPELERA_HISTORICO_FILE = "papelera_historico.csv"
REPORTES_FILE = "reportes_fo_problemas.csv"
PROBLEMAS_CONFIG_FILE = "opciones_problemas.json"
COORDINADORES_CONFIG_FILE = "opciones_coordinadores.json"
CAPTURAS_DIR = "capturas_reportes"
RESPALDOS_MES_DIR = "respaldos_mensuales"
HISTORICO_CIERRES_FILE = "historico_cierres_mensuales.csv"

CRUCE_CACHE_FILE = "ultimo_cruce.parquet"
TOTALES_CACHE_FILE = "totales_cache.pkl"
RANKING_CACHE_FILE = "ranking_cache.parquet"
INFO_ACTUALIZACION_FILE = "ultima_actualizacion.txt"

# Crear carpetas del sistema si no existen
for directorio in [CAPTURAS_DIR, RESPALDOS_MES_DIR]:
    if not os.path.exists(directorio):
        os.makedirs(directorio)

# INICIALIZACIÓN DE NAVEGACIÓN EN SESSION STATE
if "modulo_activo" not in st.session_state:
    st.session_state.modulo_activo = "INICIO"

CLAVE_AUTORIZADA_CIERRE = "autorizado2026"


# --- FUNCIONES DE SANITIZACIÓN Y PARSEO ---
def limpiar_fo_str(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "nan", "None"]:
        return ""
    txt = str(val).strip().replace(" ", "")
    if txt.endswith(".0"):
        txt = txt[:-2]
    try:
        return str(int(float(txt)))
    except (ValueError, TypeError):
        return txt


def formatear_fecha_str(val):
    if pd.isna(val) or str(val).strip() in ["", "-", "nan", "None"]:
        return "-"
    txt = str(val).strip()
    try:
        dt = pd.to_datetime(txt, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        pass
    return txt


def buscar_valor_columna_flexible(row, palabras_clave):
    for col in row.index:
        col_clean = str(col).strip().lower()
        for kw in palabras_clave:
            if (
                kw in col_clean
                and pd.notna(row[col])
                and str(row[col]).strip() not in ["", "-", "nan", "None"]
            ):
                return row[col]
    return None


def parsear_datetime_robusto(fecha_str):
    if pd.isna(fecha_str) or str(fecha_str).strip() in ["-", "", "nan", "None"]:
        return None
    txt = str(fecha_str).strip()
    for fmt in [
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M",
    ]:
        try:
            return datetime.strptime(txt, fmt)
        except ValueError:
            pass
    dt_pd = pd.to_datetime(txt, dayfirst=True, errors="coerce")
    if pd.notna(dt_pd):
        return dt_pd.to_pydatetime()
    return None


def calcular_fin_real_automatico(fin_plan_str, desviacion_str):
    dt_plan = parsear_datetime_robusto(fin_plan_str)
    if (
        not dt_plan
        or pd.isna(desviacion_str)
        or str(desviacion_str).strip() in ["-", "", "nan", "None"]
    ):
        return "-"

    try:
        texto = str(desviacion_str).replace("+", "").strip().lower()
        horas = 0
        minutos = 0

        m_h = re.search(r"(\d+)\s*h", texto)
        if m_h:
            horas = int(m_h.group(1))

        m_m = re.search(r"(\d+)\s*m", texto)
        if m_m:
            minutos = int(m_m.group(1))

        if not m_h and not m_m:
            m_num = re.search(r"(\d+)", texto)
            if m_num:
                horas = int(m_num.group(1))

        dt_real = dt_plan + timedelta(hours=horas, minutes=minutos)
        return dt_real.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return "-"


def calcular_desviacion_automatica(fin_plan_str, fin_real_str):
    dt_plan = parsear_datetime_robusto(fin_plan_str)
    dt_real = parsear_datetime_robusto(fin_real_str)

    if not dt_plan or not dt_real:
        return "-"

    diferencia = dt_real - dt_plan
    total_minutos = int(diferencia.total_seconds() // 60)

    if total_minutos <= 0:
        return "-"

    hrs = total_minutos // 60
    mins = total_minutos % 60
    return f"+{hrs}h {mins:02d}m"


# --- OPCIONES DINÁMICAS Y CONFIGURACIONES ---
OPCIONES_PROBLEMAS_DEFECTO = [
    "FO informada tarde",
    "FO con fecha errónea",
    "FO enviada tarde",
    "FO sin plan",
    "Otros",
]

def cargar_opciones_problemas():
    if os.path.exists(PROBLEMAS_CONFIG_FILE):
        try:
            with open(PROBLEMAS_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    return list(OPCIONES_PROBLEMAS_DEFECTO)


def guardar_opciones_problemas(lista_opciones):
    try:
        with open(PROBLEMAS_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(lista_opciones, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Error al guardar opciones de problemas: {e}")


OPCIONES_COORDINADORES_DEFECTO = [
    "Alison Paulina Palma Acuña",
    "Andrea Jacqueline Astorga Gil",
    "Antonio Enrique Rojas Colina",
    "Braulio Oscar Cabezas Araya",
    "Camila Stephanie Arancibia Gamboa",
    "Carolina Andrea Leiva Ramirez",
    "Danna Corey-All Hun Tello",
    "Elizabeth Del Carmen Ortiz Roa",
    "Franco Aharon Taucano Aballay",
    "Jennifer Magdalena Díaz Olmedo",
    "Karina Alejandra Valencia Muñoz",
    "Nicol Andrea Martinez Marambio",
    "Patricia Andrea Cuevas Arenas",
    "Yheri Eber Araya González",
]

def cargar_opciones_coordinadores():
    if os.path.exists(COORDINADORES_CONFIG_FILE):
        try:
            with open(COORDINADORES_CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list) and len(data) > 0:
                    return data
        except Exception:
            pass
    return list(OPCIONES_COORDINADORES_DEFECTO)


def guardar_opciones_coordinadores(lista_opciones):
    try:
        with open(COORDINADORES_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(lista_opciones, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Error al guardar opciones de coordinadores: {e}")


# --- GESTIÓN DE REPORTES Y PERSISTENCIA (GOOGLE SHEETS / CSV LOCAL) ---
def cargar_reportes_fo():
    if HAS_GSHEETS and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        try:
            conn = st.connection("gsheets", type=GSheetsConnection)
            df = conn.read(ttl=0)
            if df is not None and not df.empty:
                df = df.astype(str)
                df["FO"] = df["FO"].apply(limpiar_fo_str)
                if "Coordinador_Komatsu" not in df.columns:
                    df["Coordinador_Komatsu"] = "-"
                return df
        except Exception as e:
            st.warning(f"⚠️ Leyendo respaldos locales por problemas de conexión con Google Sheets: {e}")

    if os.path.exists(REPORTES_FILE):
        df = pd.read_csv(REPORTES_FILE, dtype=str)
        df["FO"] = df["FO"].apply(limpiar_fo_str)
        if "Coordinador_Komatsu" not in df.columns:
            df["Coordinador_Komatsu"] = "-"
        return df

    return pd.DataFrame(
        columns=[
            "Fecha_Reporte",
            "FO",
            "Sucursal",
            "Tipo_Problema",
            "Coordinador_Komatsu",
            "Responsable",
            "Detalle_Observacion",
            "Tipo_Respaldo",
            "Archivo_Pantallazo",
        ]
    )


def guardar_dataframe_reportes(df_reportes_completo):
    df_reportes_completo = df_reportes_completo.astype(str)
    df_reportes_completo.to_csv(REPORTES_FILE, index=False)

    if HAS_GSHEETS and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        try:
            conn = st.connection("gsheets", type=GSheetsConnection)
            conn.update(data=df_reportes_completo)
            return True
        except Exception as e:
            st.error(f"Error al actualizar datos en Google Sheets: {e}")
            return False
    return True


def guardar_reporte_fo(fo, sucursal, problema, coord_komatsu, responsable, observacion, file_obj):
    df_rep = cargar_reportes_fo()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    fo_limpia = limpiar_fo_str(fo)

    nombre_archivo = "Sin respaldo"
    tipo_res = "Sin archivo"

    if file_obj is not None:
        ext = os.path.splitext(file_obj.name)[1].lower()
        nombre_archivo = f"FO_{fo_limpia}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
        path_file = os.path.join(CAPTURAS_DIR, nombre_archivo)
        with open(path_file, "wb") as f:
            f.write(file_obj.getbuffer())

        tipo_res = "Imagen / Foto" if ext in [".png", ".jpg", ".jpeg"] else ("Documento PDF" if ext == ".pdf" else "Correo / Mensaje")

    nuevo_rep = {
        "Fecha_Reporte": fecha_actual,
        "FO": fo_limpia,
        "Sucursal": sucursal,
        "Tipo_Problema": problema,
        "Coordinador_Komatsu": coord_komatsu if coord_komatsu else "-",
        "Responsable": responsable,
        "Detalle_Observacion": observacion,
        "Tipo_Respaldo": tipo_res,
        "Archivo_Pantallazo": nombre_archivo,
    }

    df_rep = pd.concat([df_rep, pd.DataFrame([nuevo_rep])], ignore_index=True)
    guardar_dataframe_reportes(df_rep)
    reaplicar_cruce_con_reportes()
    return True


def editar_reporte_fo(fo_target, sucursal, problema, coord_komatsu, responsable, observacion, file_obj=None):
    df_rep = cargar_reportes_fo()
    fo_target_limpia = limpiar_fo_str(fo_target)
    idx_matches = df_rep[df_rep["FO"] == fo_target_limpia].index

    if len(idx_matches) > 0:
        idx = idx_matches[-1]
        df_rep.at[idx, "Sucursal"] = sucursal
        df_rep.at[idx, "Tipo_Problema"] = problema
        df_rep.at[idx, "Coordinador_Komatsu"] = coord_komatsu if coord_komatsu else "-"
        df_rep.at[idx, "Responsable"] = responsable
        df_rep.at[idx, "Detalle_Observacion"] = observacion

        if file_obj is not None:
            ext = os.path.splitext(file_obj.name)[1].lower()
            nombre_archivo = f"FO_{fo_target_limpia}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
            path_file = os.path.join(CAPTURAS_DIR, nombre_archivo)
            with open(path_file, "wb") as f:
                f.write(file_obj.getbuffer())

            df_rep.at[idx, "Archivo_Pantallazo"] = nombre_archivo
            df_rep.at[idx, "Tipo_Respaldo"] = (
                "Imagen / Foto" if ext in [".png", ".jpg", ".jpeg"] else ("Documento PDF" if ext == ".pdf" else "Correo / Mensaje")
            )

        guardar_dataframe_reportes(df_rep)
        reaplicar_cruce_con_reportes()
        return True
    return False


def eliminar_reporte_fo(fo_target):
    df_rep = cargar_reportes_fo()
    fo_target_limpia = limpiar_fo_str(fo_target)
    if fo_target_limpia in df_rep["FO"].values:
        df_rep = df_rep[df_rep["FO"] != fo_target_limpia]
        guardar_dataframe_reportes(df_rep)
        reaplicar_cruce_con_reportes()
        return True
    return False


# --- CACHÉ Y MANEJO DE HISTÓRICO ---
def guardar_cache_global(totales, df_cruce, df_ranking):
    try:
        if not df_cruce.empty:
            df_cruce.to_parquet(CRUCE_CACHE_FILE, index=False)
        if df_ranking is not None and not df_ranking.empty:
            df_ranking.to_parquet(RANKING_CACHE_FILE, index=False)
        with open(TOTALES_CACHE_FILE, "wb") as f:
            pickle.dump(totales, f)

        fecha_hora_str = datetime.now().strftime("%d/%m/%Y a las %H:%M hrs")
        with open(INFO_ACTUALIZACION_FILE, "w", encoding="utf-8") as f:
            f.write(fecha_hora_str)
    except Exception as e:
        st.error(f"Error guardando archivos locales de caché: {e}")


def cargar_cache_global():
    totales = None
    df_cruce = pd.DataFrame()
    df_ranking = pd.DataFrame()
    fecha_act = None

    if os.path.exists(TOTALES_CACHE_FILE):
        try:
            with open(TOTALES_CACHE_FILE, "rb") as f:
                totales = pickle.load(f)
        except Exception:
            totales = None

    if os.path.exists(CRUCE_CACHE_FILE):
        try:
            df_cruce = pd.read_parquet(CRUCE_CACHE_FILE)
        except Exception:
            df_cruce = pd.DataFrame()

    if os.path.exists(RANKING_CACHE_FILE):
        try:
            df_ranking = pd.read_parquet(RANKING_CACHE_FILE)
        except Exception:
            df_ranking = pd.DataFrame()

    if os.path.exists(INFO_ACTUALIZACION_FILE):
        try:
            with open(INFO_ACTUALIZACION_FILE, "r", encoding="utf-8") as f:
                fecha_act = f.read().strip()
        except Exception:
            fecha_act = None

    return totales, df_cruce, df_ranking, fecha_act


def cargar_historico():
    if os.path.exists(HISTORICO_FILE):
        df = pd.read_csv(HISTORICO_FILE)
        df["Fecha_Registro"] = pd.to_datetime(df["Fecha_Registro"]).dt.date
        for col in ["FO_Reportadas", "FO_Con_Evento", "Retrasos_Reportados"]:
            if col not in df.columns:
                df[col] = 0
        return df.sort_values(by="Fecha_Registro")
    return pd.DataFrame(
        columns=[
            "Fecha_Registro", "FO_Totales", "Integradas", "A_Tiempo", "Con_Retraso",
            "Vigentes", "Vencidas", "Rechazadas", "FO_Reportadas", "FO_Con_Evento",
            "Retrasos_Reportados", "Pct_A_Tiempo", "Pct_Vencidas"
        ]
    )


def cargar_papelera():
    if os.path.exists(PAPELERA_HISTORICO_FILE):
        return pd.read_csv(PAPELERA_HISTORICO_FILE)
    return pd.DataFrame()


def ejecutar_cierre_de_mes(nombre_mes, usuario_resp):
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    carpeta_destino = os.path.join(RESPALDOS_MES_DIR, f"Cierre_{nombre_mes}_{stamp}")
    os.makedirs(carpeta_destino, exist_ok=True)

    df_cruce_actual = st.session_state.get("df_cruce", pd.DataFrame())
    df_reps_actual = cargar_reportes_fo()
    df_hist_actual = cargar_historico()

    if not df_cruce_actual.empty:
        df_cruce_actual.to_excel(os.path.join(carpeta_destino, "Cruce_Vencidas_Vigentes.xlsx"), index=False)
    if not df_reps_actual.empty:
        df_reps_actual.to_excel(os.path.join(carpeta_destino, "Reportes_FO_Problemas.xlsx"), index=False)
    if not df_hist_actual.empty:
        df_hist_actual.to_excel(os.path.join(carpeta_destino, "Historico_Diario.xlsx"), index=False)

    if os.path.exists(CAPTURAS_DIR) and len(os.listdir(CAPTURAS_DIR)) > 0:
        dest_capturas = os.path.join(carpeta_destino, "capturas_reportes")
        shutil.copytree(CAPTURAS_DIR, dest_capturas, dirs_exist_ok=True)

    reg_cierre = {
        "Fecha_Cierre": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Nombre_Periodo": nombre_mes,
        "Usuario_Responsable": usuario_resp,
        "Ruta_Carpeta_Respaldo": carpeta_destino,
        "Total_FO_Procesadas": len(df_cruce_actual),
        "Total_Reportes": len(df_reps_actual),
    }

    df_cierres = pd.DataFrame()
    if os.path.exists(HISTORICO_CIERRES_FILE):
        df_cierres = pd.read_csv(HISTORICO_CIERRES_FILE)
    df_cierres = pd.concat([df_cierres, pd.DataFrame([reg_cierre])], ignore_index=True)
    df_cierres.to_csv(HISTORICO_CIERRES_FILE, index=False)

    # RESET COMPLETO DE MEMORIA Y HOJA
    st.session_state.df_cruce = pd.DataFrame()
    st.session_state.df_ranking = pd.DataFrame()
    st.session_state.totales = {
        "total_fo": 0, "integradas": 0, "a_tiempo": 0, "con_retraso": 0,
        "vigentes": 0, "vencidas": 0, "rechazadas": 0
    }

    df_vacio_reps = pd.DataFrame(
        columns=[
            "Fecha_Reporte", "FO", "Sucursal", "Tipo_Problema", "Coordinador_Komatsu",
            "Responsable", "Detalle_Observacion", "Tipo_Respaldo", "Archivo_Pantallazo"
        ]
    )
    guardar_dataframe_reportes(df_vacio_reps)

    for arch_cache in [CRUCE_CACHE_FILE, TOTALES_CACHE_FILE, RANKING_CACHE_FILE]:
        if os.path.exists(arch_cache):
            try:
                os.remove(arch_cache)
            except Exception:
                pass

    st.session_state.ultima_fecha_act = f"Mes {nombre_mes} cerrado el {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    return True


# --- REAPLICAR REPORTES A TABLAS EN MEMORIA ---
def reaplicar_cruce_con_reportes():
    if "df_cruce" in st.session_state and not st.session_state.df_cruce.empty:
        df_cruce_p = st.session_state.df_cruce.copy()
        
        for col_del in ["Estado_Reporte", "Detalle_Reporte", "FO_KEY"]:
            if col_del in df_cruce_p.columns:
                df_cruce_p = df_cruce_p.drop(columns=[col_del])

        df_cruce_p["FO_KEY"] = df_cruce_p["FO"].astype(str).apply(limpiar_fo_str)

        df_reps_p = cargar_reportes_fo()
        if not df_reps_p.empty:
            df_reps_p["FO_KEY"] = df_reps_p["FO"].astype(str).apply(limpiar_fo_str)

            df_reps_grouped = (
                df_reps_p.groupby("FO_KEY")
                .agg({
                    "Tipo_Problema": "last",
                    "Coordinador_Komatsu": "last",
                    "Responsable": "last",
                    "Detalle_Observacion": "last",
                })
                .reset_index()
            )
            df_reps_grouped["Estado_Reporte"] = "🚨 REPORTADA"
            df_reps_grouped["Detalle_Reporte"] = (
                "Motivo: "
                + df_reps_grouped["Tipo_Problema"].astype(str)
                + " | Coord: "
                + df_reps_grouped["Coordinador_Komatsu"].astype(str)
                + " | Resp: "
                + df_reps_grouped["Responsable"].astype(str)
                + " | Obs: "
                + df_reps_grouped["Detalle_Observacion"].astype(str)
            )

            df_cruce_p = pd.merge(
                df_cruce_p,
                df_reps_grouped[["FO_KEY", "Estado_Reporte", "Detalle_Reporte"]],
                on="FO_KEY",
                how="left",
            )
            df_cruce_p["Estado_Reporte"] = df_cruce_p["Estado_Reporte"].fillna("Sin Reporte")
            df_cruce_p["Detalle_Reporte"] = df_cruce_p["Detalle_Reporte"].fillna("-")
        else:
            df_cruce_p["Estado_Reporte"] = "Sin Reporte"
            df_cruce_p["Detalle_Reporte"] = "-"

        if "FO_KEY" in df_cruce_p.columns:
            df_cruce_p = df_cruce_p.drop(columns=["FO_KEY"])

        cols_order = [
            "FO", "Estado_Reporte", "Fin plan", "Fin real", "Desviación",
            "Conductor", "Detalle_Reporte", "SUCURSAL", "RESPONSABLE",
            "Inicio Planificado", "Fin Planificado", "Estado de la ruta",
            "Evento", "ORIGEN", "DESTINO", "FOLIO"
        ]
        cols_exist = [c for c in cols_order if c in df_cruce_p.columns]
        df_cruce_p = df_cruce_p[cols_exist]

        st.session_state.df_cruce = df_cruce_p
        guardar_cache_global(st.session_state.totales, df_cruce_p, st.session_state.df_ranking)


# --- EXPORTACIÓN A UN SOLO EXCEL CON RESULTADOS ---
def generar_excel_unificado_completo(df_cruce):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Pestaña 1: Vencidas y Vigentes (Base Unificada Completa)
        df_cruce.to_excel(writer, sheet_name="Vencidas_y_Vigentes", index=False)
        
        # Pestaña 2: RESULTADOS (Tabla cruzada por sucursal)
        if not df_cruce.empty and "SUCURSAL" in df_cruce.columns and "Estado de la ruta" in df_cruce.columns:
            df_cruce_clean = df_cruce.copy()
            df_cruce_clean["SUCURSAL_CLEAN"] = df_cruce_clean["SUCURSAL"].astype(str).str.strip().str.title()
            tabla_res = pd.crosstab(
                df_cruce_clean["SUCURSAL_CLEAN"],
                df_cruce_clean["Estado de la ruta"],
                margins=True,
                margins_name="Total general"
            ).reset_index()
            tabla_res.to_excel(writer, sheet_name="RESULTADOS", index=False)

        # Pestaña 3: Finalizadas con Retraso (Filtrado)
        df_retrasos = df_cruce[df_cruce["Estado de la ruta"].str.contains("Retraso", na=False)]
        if not df_retrasos.empty:
            df_retrasos.to_excel(writer, sheet_name="Finalizadas_Con_Retraso", index=False)
        else:
            pd.DataFrame(columns=df_cruce.columns).to_excel(writer, sheet_name="Finalizadas_Con_Retraso", index=False)

    return output.getvalue()


def generar_excel_solo_tabla(df_tabla, nombre_hoja="Datos"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
    return output.getvalue()


def color_filas_estado(val):
    val_str = str(val)
    if "Vencida" in val_str and "REPORTADA" not in val_str:
        return "background-color: #fbe9e7; color: #c62828; font-weight: bold;"
    elif val_str == "Vigente":
        return "background-color: #ffffff; color: #212529;"
    elif "Retraso" in val_str and "REPORTADA" not in val_str:
        return "background-color: #fff3e0; color: #ef6c00; font-weight: bold;"
    elif val_str == "Rechazada":
        return "background-color: #e1bee7; color: #4a148c; font-weight: bold;"
    elif "REPORTADA" in val_str:
        return "background-color: #ffe0b2; color: #e65100; font-weight: bold;"
    return ""


# --- PROCESAMIENTO DE ARCHIVOS SAE / SMARTSHEET / DRIVIN ---
def procesar_excel_sae(file_excel):
    try:
        df = pd.read_excel(file_excel, header=2)
        cols_indices = [0, 2, 3, 11, 29]
        df_clean = df.iloc[:, cols_indices].copy()
        df_clean.columns = ["SUCURSAL", "ORIGEN", "DESTINO", "FOLIO", "FO"]
        df_clean["FO"] = df_clean["FO"].apply(limpiar_fo_str)
        df_clean["FOLIO"] = df_clean["FOLIO"].apply(limpiar_fo_str)
        return df_clean.dropna(subset=["FO"]).drop_duplicates(subset=["FO"])
    except Exception:
        return pd.DataFrame()


def procesar_smartsheet(file_sm):
    try:
        df = pd.read_excel(file_sm)
        df_clean = df.iloc[:, [0, 10, 11]].copy()
        df_clean.columns = ["FO", "ORIGEN_SM", "DESTINO_SM"]

        def clean_desc(val):
            if pd.isna(val):
                return "PENDIENTE EN GO TRUCK"
            text = str(val).strip()
            text_limpio = re.sub(r"^(descripción:|descripcion:)\s*", "", text, flags=re.IGNORECASE).strip()
            return text_limpio if text_limpio else "PENDIENTE EN GO TRUCK"

        df_clean["FO"] = df_clean["FO"].apply(limpiar_fo_str)
        df_clean["ORIGEN_SM"] = df_clean["ORIGEN_SM"].apply(clean_desc)
        df_clean["DESTINO_SM"] = df_clean["DESTINO_SM"].apply(clean_desc)
        return df_clean.dropna(subset=["FO"]).drop_duplicates(subset=["FO"])
    except Exception:
        return pd.DataFrame()


def procesar_drivin_excel(file_excel_drivin):
    totales = {
        "total_fo": 0, "integradas": 0, "a_tiempo": 0, "con_retraso": 0,
        "vigentes": 0, "vencidas": 0, "rechazadas": 0
    }
    lista_fos = []
    df_ranking = pd.DataFrame()

    xls = pd.ExcelFile(file_excel_drivin)
    sheet_names_map = {str(s).strip().lower(): s for s in xls.sheet_names}

    sheet_vv = None
    for k in sheet_names_map:
        if ("vencid" in k and "vigend" in k) or "vencidas" in k:
            sheet_vv = sheet_names_map[k]
            break

    if sheet_vv:
        df_vv = pd.read_excel(xls, sheet_name=sheet_vv)
        df_vv = df_vv.dropna(subset=["FO"])
        for _, row in df_vv.iterrows():
            inicio_p = formatear_fecha_str(buscar_valor_columna_flexible(row, ["inicio plan", "inicio planificado", "inicio"]))
            fin_p = formatear_fecha_str(buscar_valor_columna_flexible(row, ["fin plan", "fin planificado", "fin estimad"]))
            val_fin_real = buscar_valor_columna_flexible(row, ["fin real", "fin de entrega real", "llegada real", "cierre real", "fecha real"])
            fin_r = formatear_fecha_str(val_fin_real)

            lista_fos.append({
                "FO": limpiar_fo_str(row["FO"]),
                "RESPONSABLE": str(row.get("Responsable KCC", row.get("RESPONSABLE", "-"))),
                "Fin plan": fin_p,
                "Fin real": fin_r,
                "Inicio Planificado": inicio_p,
                "Fin Planificado": fin_p,
                "Conductor": str(row.get("Conductor", "-")),
                "Estado de la ruta": str(row.get("Estado de la ruta", row.get("Estado", "Vigente"))).strip(),
                "Evento": "-" if pd.isna(row.get("Evento")) else str(row.get("Evento")),
                "Desviación": str(row.get("Desviación", "-")),
                "SUCURSAL": str(row.get("SUCURSAL", "PENDIENTE EN GO TRUCK")).strip().title(),
                "ORIGEN": str(row.get("ORIGEN", "PENDIENTE EN GO TRUCK")),
                "DESTINO": str(row.get("DESTINO", "PENDIENTE EN GO TRUCK")),
                "FOLIO": limpiar_fo_str(row.get("FOLIO", "")),
            })

    sheet_ret = None
    for k in sheet_names_map:
        if "retraso" in k or "finalizad" in k:
            sheet_ret = sheet_names_map[k]
            break

    if sheet_ret and not sheet_vv:
        df_ret = pd.read_excel(xls, sheet_name=sheet_ret)
        df_ret = df_ret.dropna(subset=["FO"])

        for _, row in df_ret.iterrows():
            inicio_p = formatear_fecha_str(buscar_valor_columna_flexible(row, ["inicio plan", "inicio planificado", "inicio"]))
            fin_p = formatear_fecha_str(buscar_valor_columna_flexible(row, ["fin plan", "fin planificado", "fin estimad"]))
            val_fin_real = buscar_valor_columna_flexible(row, ["fin real", "fin de entrega real", "llegada real", "cierre real", "fecha real"])
            fin_r = formatear_fecha_str(val_fin_real)

            desviacion_val = "-"
            val_desv = buscar_valor_columna_flexible(row, ["desviacion", "desfase", "delay"])
            if val_desv:
                desviacion_val = str(val_desv).strip()

            if fin_r == "-" and fin_p != "-" and desviacion_val != "-":
                fin_r = calcular_fin_real_automatico(fin_p, desviacion_val)

            if desviacion_val == "-" and fin_p != "-" and fin_r != "-":
                desviacion_val = calcular_desviacion_automatica(fin_p, fin_r)

            lista_fos.append({
                "FO": limpiar_fo_str(row["FO"]),
                "RESPONSABLE": str(row.get("Responsable KCC", row.get("RESPONSABLE", "-"))),
                "Fin plan": fin_p,
                "Fin real": fin_r,
                "Inicio Planificado": inicio_p,
                "Fin Planificado": fin_p,
                "Conductor": str(row.get("Conductor", "-")),
                "Estado de la ruta": str(row.get("Estado de la ruta", "Finalizada con Retraso")).strip(),
                "Evento": "-" if pd.isna(row.get("Evento")) else str(row.get("Evento")),
                "Desviación": desviacion_val,
                "SUCURSAL": str(row.get("SUCURSAL", "PENDIENTE EN GO TRUCK")).strip().title(),
                "ORIGEN": str(row.get("ORIGEN", "PENDIENTE EN GO TRUCK")),
                "DESTINO": str(row.get("DESTINO", "PENDIENTE EN GO TRUCK")),
                "FOLIO": limpiar_fo_str(row.get("FOLIO", "")),
            })

    sheet_rank = sheet_names_map.get("ranking")
    if sheet_rank:
        df_ranking = pd.read_excel(xls, sheet_name=sheet_rank)

    sheet_tot = sheet_names_map.get("hoja4") or sheet_names_map.get("hoja 4") or sheet_names_map.get("totales")
    if sheet_tot:
        df_tot = pd.read_excel(xls, sheet_name=sheet_tot, header=None)
        dict_t = {}
        for _, row in df_tot.iterrows():
            if pd.notna(row[0]) and pd.notna(row[1]):
                dict_t[str(row[0]).strip().upper()] = row[1]

        totales["total_fo"] = int(dict_t.get("FO TOTAL", 0))
        totales["a_tiempo"] = int(dict_t.get("A TIEMPO", 0))
        totales["con_retraso"] = int(dict_t.get("CON RETRASO", 0))
        totales["vigentes"] = int(dict_t.get("VIGENTES", 0))
        totales["vencidas"] = int(dict_t.get("VENCIDAS", 0))
        totales["integradas"] = totales["total_fo"]

    df_fos = pd.DataFrame(lista_fos)
    if not df_fos.empty:
        df_fos["Prioridad"] = df_fos["Estado de la ruta"].map({"Vencida": 1, "Rechazada": 2, "Finalizada con Retraso": 3, "Vigente": 4}).fillna(5)
        df_fos = df_fos.sort_values(by="Prioridad").drop_duplicates(subset=["FO"], keep="first").drop(columns=["Prioridad"])

    if totales["total_fo"] == 0 and not df_fos.empty:
        totales["integradas"] = len(df_fos)
        totales["vencidas"] = len(df_fos[df_fos["Estado de la ruta"].str.contains("Vencida", na=False)])
        totales["vigentes"] = len(df_fos[df_fos["Estado de la ruta"] == "Vigente"])
        totales["con_retraso"] = len(df_fos[df_fos["Estado de la ruta"].str.contains("Retraso", na=False)])
        totales["total_fo"] = totales["integradas"]

    return totales, df_fos, df_ranking


def procesar_drivin_pdf(file_pdf_drivin):
    totales = {
        "total_fo": 0, "integradas": 0, "a_tiempo": 0, "con_retraso": 0,
        "vigentes": 0, "vencidas": 0, "rechazadas": 0
    }
    lista_fos = []

    reader = pypdf.PdfReader(file_pdf_drivin)
    texto_completo = ""
    for page in reader.pages:
        texto_completo += (page.extract_text() or "") + "\n"

    lineas = texto_completo.split("\n")
    seccion_actual = "Vigente"

    for i, line in enumerate(lineas):
        line_up = line.upper()
        if "RUTAS QUE REQUIEREN" in line_up or "RUTAS VENCIDAS" in line_up or "FO VENCIDAS" in line_up or ("VENCIDAS" in line_up and "A TIEMPO" not in line_up):
            seccion_actual = "Vencida"
        elif "FINALIZADAS CON RETRASO" in line_up:
            seccion_actual = "Finalizada con Retraso"
        elif "RECHAZADA" in line_up or "RECHAZADAS" in line_up:
            seccion_actual = "Rechazada"
        elif "RUTAS VIGENTES" in line_up or "FO VIGENTES" in line_up:
            seccion_actual = "Vigente"

        m_fo = re.search(r"\b(6100\d{6})\b", line)
        if m_fo:
            fo_code = limpiar_fo_str(m_fo.group(1))
            bloque_texto = " ".join(lineas[max(0, i - 2) : min(len(lineas), i + 3)])
            estado_linea = seccion_actual
            if "VENCIDA" in bloque_texto.upper():
                estado_linea = "Vencida"
            elif "RECHAZADA" in bloque_texto.upper():
                estado_linea = "Rechazada"

            fechas = re.findall(r"(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})", bloque_texto)
            inicio_plan = fechas[0] if len(fechas) > 0 else "-"
            fin_plan = fechas[1] if len(fechas) > 1 else "-"
            fin_real = fechas[2] if len(fechas) > 2 and "Retraso" in estado_linea else "-"

            desfase_txt = "-"
            m_desf = re.search(r"\b(\d+\s*d[ií]as?\s*\d+\s*hrs?|\d+\s*hrs?|\+\d+h\s*\d+m)\b", bloque_texto, re.IGNORECASE)
            if m_desf:
                desfase_txt = m_desf.group(1)

            if fin_real == "-" and fin_plan != "-" and desfase_txt != "-":
                fin_real = calcular_fin_real_automatico(fin_plan, desfase_txt)

            if desfase_txt == "-" and fin_plan != "-" and fin_real != "-":
                desfase_txt = calcular_desviacion_automatica(fin_plan, fin_real)

            resp_kcc = "-"
            m_kcc = re.search(r"\b(Danna\s+[A-Za-z\-]+|Nicol\s+[A-Za-z\s]+|Yheri\s+[A-Za-z\s]+|Carolina\s+[A-Za-z\s]+|Karina\s+[A-Za-z\s]+|Patricia\s+[A-Za-z\s]+|Elizabeth\s+[A-Za-z\s]+|Diego\s+[A-Za-z\s]+|Camila\s+[A-Za-z\s]+|Franco\s+[A-Za-z\s]+)\b", bloque_texto)
            if m_kcc:
                resp_kcc = m_kcc.group(1)

            lista_fos.append({
                "FO": fo_code,
                "RESPONSABLE": resp_kcc,
                "Fin plan": fin_plan,
                "Fin real": fin_real,
                "Inicio Planificado": inicio_plan,
                "Fin Planificado": fin_plan,
                "Conductor": "-",
                "Estado de la ruta": estado_linea,
                "Evento": "-",
                "Desviación": desfase_txt,
                "SUCURSAL": "PENDIENTE EN GO TRUCK",
                "ORIGEN": "PENDIENTE EN GO TRUCK",
                "DESTINO": "PENDIENTE EN GO TRUCK",
                "FOLIO": "",
            })

    df_fos = pd.DataFrame(lista_fos)
    if not df_fos.empty:
        df_fos["Prioridad"] = df_fos["Estado de la ruta"].map({"Vencida": 1, "Rechazada": 2, "Finalizada con Retraso": 3, "Vigente": 4}).fillna(5)
        df_fos = df_fos.sort_values(by="Prioridad").drop_duplicates(subset=["FO"], keep="first").drop(columns=["Prioridad"])

    totales["integradas"] = len(df_fos)
    totales["vencidas"] = len(df_fos[df_fos["Estado de la ruta"].str.contains("Vencida", na=False)]) if not df_fos.empty else 0
    totales["vigentes"] = len(df_fos[df_fos["Estado de la ruta"] == "Vigente"]) if not df_fos.empty else 0
    totales["con_retraso"] = len(df_fos[df_fos["Estado de la ruta"].str.contains("Retraso", na=False)]) if not df_fos.empty else 0
    totales["rechazadas"] = len(df_fos[df_fos["Estado de la ruta"] == "Rechazada"]) if not df_fos.empty else 0
    totales["total_fo"] = totales["integradas"]

    return totales, df_fos, pd.DataFrame()


def ejecutar_cruce_completo(file_drivin, file_sae=None, file_smartsheet=None):
    if hasattr(file_drivin, "name") and file_drivin.name.endswith(".pdf"):
        totales_p, df_fos_p, df_rank_p = procesar_drivin_pdf(file_drivin)
    else:
        totales_p, df_fos_p, df_rank_p = procesar_drivin_excel(file_drivin)

    if totales_p["total_fo"] == 0 or totales_p["total_fo"] < len(df_fos_p):
        totales_p["total_fo"] = len(df_fos_p) if not df_fos_p.empty else 650
        totales_p["integradas"] = totales_p["total_fo"]

    df_sae_p = procesar_excel_sae(file_sae) if file_sae is not None else pd.DataFrame()
    df_sm_p = procesar_smartsheet(file_smartsheet) if file_smartsheet is not None else pd.DataFrame()

    df_cruce_p = pd.DataFrame()
    if not df_fos_p.empty:
        df_cruce_p = df_fos_p.copy()

        if not df_sae_p.empty:
            df_cruce_p = pd.merge(df_cruce_p, df_sae_p, on="FO", how="left", suffixes=('', '_sae'))
        
        if not df_sm_p.empty:
            df_cruce_p = pd.merge(df_cruce_p, df_sm_p, on="FO", how="left")
            if "ORIGEN_SM" in df_cruce_p.columns:
                df_cruce_p["ORIGEN"] = df_cruce_p["ORIGEN"].fillna(df_cruce_p["ORIGEN_SM"])
                df_cruce_p["DESTINO"] = df_cruce_p["DESTINO"].fillna(df_cruce_p["DESTINO_SM"])
                df_cruce_p = df_cruce_p.drop(columns=["ORIGEN_SM", "DESTINO_SM"])

        df_cruce_p["SUCURSAL"] = df_cruce_p["SUCURSAL"].fillna("PENDIENTE EN GO TRUCK").astype(str).str.strip().str.title()
        df_cruce_p["ORIGEN"] = df_cruce_p["ORIGEN"].fillna("PENDIENTE EN GO TRUCK")
        df_cruce_p["DESTINO"] = df_cruce_p["DESTINO"].fillna("PENDIENTE EN GO TRUCK")
        df_cruce_p["FOLIO"] = df_cruce_p["FOLIO"].apply(limpiar_fo_str)

        # Integración de reportes desde Google Sheets / CSV
        df_reps_p = cargar_reportes_fo()
        if not df_reps_p.empty:
            df_cruce_p["FO_KEY"] = df_cruce_p["FO"].astype(str).apply(limpiar_fo_str)
            df_reps_p["FO_KEY"] = df_reps_p["FO"].astype(str).apply(limpiar_fo_str)

            df_reps_grouped = (
                df_reps_p.groupby("FO_KEY")
                .agg({
                    "Tipo_Problema": "last",
                    "Coordinador_Komatsu": "last",
                    "Responsable": "last",
                    "Detalle_Observacion": "last",
                })
                .reset_index()
            )
            df_reps_grouped["Estado_Reporte"] = "🚨 REPORTADA"
            df_reps_grouped["Detalle_Reporte"] = (
                "Motivo: "
                + df_reps_grouped["Tipo_Problema"].astype(str)
                + " | Coord. Komatsu: "
                + df_reps_grouped["Coordinador_Komatsu"].astype(str)
                + " | Resp: "
                + df_reps_grouped["Responsable"].astype(str)
                + " | Obs: "
                + df_reps_grouped["Detalle_Observacion"].astype(str)
            )

            df_cruce_p = pd.merge(
                df_cruce_p,
                df_reps_grouped[["FO_KEY", "Estado_Reporte", "Detalle_Reporte"]],
                on="FO_KEY",
                how="left",
            )
            df_cruce_p["Estado_Reporte"] = df_cruce_p["Estado_Reporte"].fillna("Sin Reporte")
            df_cruce_p["Detalle_Reporte"] = df_cruce_p["Detalle_Reporte"].fillna("-")
            
            if "FO_KEY" in df_cruce_p.columns:
                df_cruce_p = df_cruce_p.drop(columns=["FO_KEY"])
        else:
            df_cruce_p["Estado_Reporte"] = "Sin Reporte"
            df_cruce_p["Detalle_Reporte"] = "-"

        cols_order = [
            "FO", "Estado_Reporte", "Fin plan", "Fin real", "Desviación",
            "Conductor", "Detalle_Reporte", "SUCURSAL", "RESPONSABLE",
            "Inicio Planificado", "Fin Planificado", "Estado de la ruta",
            "Evento", "ORIGEN", "DESTINO", "FOLIO"
        ]
        cols_exist = [c for c in cols_order if c in df_cruce_p.columns]
        df_cruce_p = df_cruce_p[cols_exist]

    return totales_p, df_cruce_p, df_rank_p


# --- CARGA Y SINCRONIZACIÓN INICIAL DE MEMORIA ---
cache_totales, cache_df_cruce, cache_df_ranking, ultima_fecha_act = cargar_cache_global()

if (cache_df_cruce is None or cache_df_cruce.empty) and os.path.exists(DRIVIN_BASE_FILE):
    f_sae_b = SAE_BASE_FILE if os.path.exists(SAE_BASE_FILE) else None
    f_sm_b = SMARTSHEET_BASE_FILE if os.path.exists(SMARTSHEET_BASE_FILE) else None
    try:
        cache_totales, cache_df_cruce, cache_df_ranking = ejecutar_cruce_completo(
            DRIVIN_BASE_FILE, f_sae_b, f_sm_b
        )
        guardar_cache_global(cache_totales, cache_df_cruce, cache_df_ranking)
        ultima_fecha_act = datetime.now().strftime("%d/%m/%Y a las %H:%M hrs")
    except Exception as e:
        st.warning(f"No se pudieron cargar los archivos base predeterminados: {e}")

if "df_cruce" not in st.session_state or st.session_state.df_cruce.empty:
    st.session_state.totales = cache_totales
    st.session_state.df_cruce = cache_df_cruce
    st.session_state.df_ranking = cache_df_ranking
    st.session_state.ultima_fecha_act = ultima_fecha_act


# --- BARRA LATERAL (PANEL ADMINISTRATIVO) ---
with st.sidebar:
    st.title(" Panel Drivin")
    st.caption("Operación Komatsu Cummins - SAE")

    if st.session_state.modulo_activo != "INICIO":
        if st.button("🏠 Menú Principal / Inicio", use_container_width=True, type="primary"):
            st.session_state.modulo_activo = "INICIO"
            st.rerun()
        st.divider()

    # Indicador de estado de la conexión a la base de datos de Google Sheets
    if HAS_GSHEETS and "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        st.success("🟢 **Google Sheets Conectado:** Las FO se guardan permanentemente en la nube.")
    else:
        st.warning("🔴 **Modo Offline/Local:** Para persistencia en la nube, configura Secrets en Streamlit Cloud.")

    fecha_act_display = st.session_state.get("ultima_fecha_act", ultima_fecha_act)
    if fecha_act_display:
        st.info(f"🕒 **Última actualización local:**\n{fecha_act_display}")

    if st.button("🔄 Sincronizar con Google Sheets", use_container_width=True):
        reaplicar_cruce_con_reportes()
        st.success("✅ Sistema sincronizado correctamente.")
        st.rerun()

    st.divider()

    # --- BOTÓN DE DESCARGA UNIFICADA COMPLETA EN PANEL ADMINISTRADOR ---
    df_cruce_admin = st.session_state.get("df_cruce", pd.DataFrame())
    if not df_cruce_admin.empty:
        st.subheader("📥 Exportación de Datos Operativos")
        bytes_excel_unificado = generar_excel_unificado_completo(df_cruce_admin)
        st.download_button(
            label="📊 Descargar Base Completa (V&V + Retrasos en Excel)",
            data=bytes_excel_unificado,
            file_name=f"Base_Operativa_Unificada_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            help="Descarga un solo archivo Excel con 3 pestañas: Vencidas/Vigentes, RESULTADOS y Finalizadas con Retraso."
        )
        st.divider()

    # --- MÓDULO DE CIERRE DE MES ---
    st.subheader("🔒 Cierre de Mes (Cero / Reset)")
    with st.expander("🚨 Ejecutar Cierre Mensual"):
        st.warning("⚠️ Esta opción guardará un snapshot del periodo y limpiará los indicadores para el nuevo mes.")
        nombre_mes_in = st.text_input("Mes/Año a cerrar (Ej: Enero_2026):", key="txt_mes_cierre")
        user_cierre_in = st.text_input("Usuario Responsable:", key="txt_user_cierre")
        pass_cierre_in = st.text_input("Clave de Seguridad Autorizada:", type="password", key="pass_cierre")

        if st.button("🔥 EJECUTAR CIERRE DE MES", type="primary", use_container_width=True):
            if pass_cierre_in != CLAVE_AUTORIZADA_CIERRE:
                st.error("🔑 Clave de seguridad incorrecta. Acción no autorizada.")
            elif not nombre_mes_in or not user_cierre_in:
                st.warning("⚠️ Debes indicar el nombre del periodo y el usuario responsable.")
            else:
                if ejecutar_cierre_de_mes(nombre_mes_in.strip(), user_cierre_in.strip()):
                    st.success(f"🎉 Periodo '{nombre_mes_in}' cerrado exitosamente. El sistema ha iniciado de 0.")
                    st.rerun()

    st.divider()

    # --- RESTAURACIÓN MANUAL DE COPIAS DE SEGURIDAD ---
    st.subheader("📥 Cargar Respaldo Manual (CSV/Excel)")
    with st.expander("📂 Importar Copia Externa"):
        uploaded_recovery = st.file_uploader("Subir respaldo previo:", type=["csv", "xlsx"])
        if st.button("🚀 Forzar Carga a Google Sheets", use_container_width=True):
            if uploaded_recovery is not None:
                try:
                    df_rec = pd.read_csv(uploaded_recovery, dtype=str) if uploaded_recovery.name.endswith(".csv") else pd.read_excel(uploaded_recovery, sheet_name=0, dtype=str)
                    df_rec["FO"] = df_rec["FO"].apply(limpiar_fo_str)
                    guardar_dataframe_reportes(df_rec)
                    reaplicar_cruce_con_reportes()
                    st.success("✅ Respaldo procesado e importado exitosamente.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al importar el archivo: {e}")

    st.divider()

    st.subheader("⚙️ Parámetros de Operación")
    servicios_mensuales_admin = st.number_input(
        "📦 Total Servicios Mensuales Operativos:",
        min_value=1,
        value=650,
        step=10,
        help="Indica el total estimado de servicios del mes para calculos proporcionales.",
    )

    # Configuración dinámica de opciones de problemas
    st.markdown("##### ⚙️ Lista de Problemas FO")
    opciones_actuales_prob = cargar_opciones_problemas()

    nueva_opcion_prob = st.text_input("Agregar tipo de problema:", key="txt_nuevo_prob")
    if st.button("➕ Añadir Problema", use_container_width=True):
        txt_clean = nueva_opcion_prob.strip()
        if txt_clean and txt_clean not in opciones_actuales_prob:
            opciones_actuales_prob.append(txt_clean)
            guardar_opciones_problemas(opciones_actuales_prob)
            st.success(f"✅ Añadido: '{txt_clean}'")
            st.rerun()

    opcion_eliminar_prob = st.selectbox(
        "Eliminar problema de lista:",
        options=["-- Seleccionar --"] + opciones_actuales_prob,
        key="sb_elim_prob"
    )
    if st.button("🗑️ Eliminar Problema", use_container_width=True):
        if opcion_eliminar_prob != "-- Seleccionar --":
            opciones_actuales_prob.remove(opcion_eliminar_prob)
            guardar_opciones_problemas(opciones_actuales_prob)
            st.success(f"🗑️ Eliminado: '{opcion_eliminar_prob}'")
            st.rerun()

    # Configuración dinámica de coordinadores Komatsu
    st.divider()
    st.markdown("##### ⚙️ Lista de Coordinadores Komatsu / KCC")
    opciones_actuales_coord = cargar_opciones_coordinadores()

    nuevo_coord = st.text_input("Agregar coordinador:", key="txt_nuevo_coord")
    if st.button("➕ Añadir Coordinador", use_container_width=True):
        txt_coord_clean = nuevo_coord.strip()
        if txt_coord_clean and txt_coord_clean not in opciones_actuales_coord:
            opciones_actuales_coord.append(txt_coord_clean)
            guardar_opciones_coordinadores(opciones_actuales_coord)
            st.success(f"✅ Añadido: '{txt_coord_clean}'")
            st.rerun()

    coord_eliminar = st.selectbox(
        "Eliminar coordinador de lista:",
        options=["-- Seleccionar --"] + opciones_actuales_coord,
        key="sb_elim_coord"
    )
    if st.button("🗑️ Eliminar Coordinador", use_container_width=True):
        if coord_eliminar != "-- Seleccionar --":
            opciones_actuales_coord.remove(coord_eliminar)
            guardar_opciones_coordinadores(opciones_actuales_coord)
            st.success(f"🗑️ Eliminado: '{coord_eliminar}'")
            st.rerun()

    st.divider()

    st.subheader("📂 Cargar Planillas Operativas")
    formato_drivin = st.radio(
        "Tipo de Archivo Drivin:",
        ["📊 Excel Drivin (.xlsx)", "📄 PDF Drivin (.pdf)"],
        index=0,
    )

    uploaded_drivin_excel = None
    uploaded_drivin_pdf = None

    if "Excel" in formato_drivin:
        uploaded_drivin_excel = st.file_uploader("1. Drivin Excel (.xlsx)", type=["xlsx", "xls"])
    else:
        uploaded_drivin_pdf = st.file_uploader("1. Drivin PDF (.pdf)", type=["pdf"])

    uploaded_excel_sae = st.file_uploader("2. Planilla SAE EDP (.xlsx)", type=["xlsx", "xls"])
    uploaded_smartsheet = st.file_uploader("3. Planilla SmartSheet (.xlsx)", type=["xlsx", "xls"])

    btn_procesar = st.button("🚀 Procesar Carga de Información", type="primary", use_container_width=True)

    if btn_procesar:
        f_drivin_sel = uploaded_drivin_excel if "Excel" in formato_drivin else uploaded_drivin_pdf
        if f_drivin_sel is None and not os.path.exists(DRIVIN_BASE_FILE):
            st.error("⚠️ Debes adjuntar el archivo base de Drivin.")
        else:
            try:
                f_d = f_drivin_sel if f_drivin_sel is not None else DRIVIN_BASE_FILE
                f_sae = uploaded_excel_sae if uploaded_excel_sae is not None else (SAE_BASE_FILE if os.path.exists(SAE_BASE_FILE) else None)
                f_sm = uploaded_smartsheet if uploaded_smartsheet is not None else (SMARTSHEET_BASE_FILE if os.path.exists(SMARTSHEET_BASE_FILE) else None)

                totales_p, df_cruce_p, df_rank_p = ejecutar_cruce_completo(f_d, f_sae, f_sm)

                if totales_p["total_fo"] == 0 or totales_p["total_fo"] < len(df_cruce_p):
                    totales_p["total_fo"] = len(df_cruce_p) if not df_cruce_p.empty else servicios_mensuales_admin
                    totales_p["integradas"] = totales_p["total_fo"]

                guardar_cache_global(totales_p, df_cruce_p, df_rank_p)

                st.session_state.totales = totales_p
                st.session_state.df_cruce = df_cruce_p
                st.session_state.df_ranking = df_rank_p
                st.session_state.ultima_fecha_act = datetime.now().strftime("%d/%m/%Y a las %H:%M hrs")
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar la información: {e}")


# --- ENCABEZADO Y KPIS PRINCIPALES DE LA PÁGINA ---
st.title("🚚 Dashboard Operativo Drivin")

fecha_encabezado = st.session_state.get("ultima_fecha_act", ultima_fecha_act)
if fecha_encabezado:
    st.caption(f"📅 **Información actualizada al:** {fecha_encabezado}")

df_cruce = st.session_state.get("df_cruce", pd.DataFrame())
df_ranking = st.session_state.get("df_ranking", pd.DataFrame())
df_hist = cargar_historico()

col_m1, col_m2 = st.columns([1, 2])

with col_m1:
    tipo_vista = st.radio(
        "Modo de Visualización:",
        ["Día Específico", "Mes Completo Acumulado"],
        horizontal=True,
    )

with col_m2:
    if tipo_vista == "Día Específico":
        if not df_hist.empty:
            fechas_disponibles = [str(f) for f in df_hist["Fecha_Registro"].tolist()]
            fechas_disponibles.insert(0, "📅 Carga Actual Local")
            fecha_consultada = st.selectbox("🔍 Seleccionar Día:", options=fechas_disponibles)
        else:
            fecha_consultada = "📅 Carga Actual Local"
    else:
        fecha_consultada = "Mes Completo Acumulado"
        st.write("📊 Consolidando todos los registros guardados del mes.")

# CALCULOS DE KPIS
if tipo_vista == "Mes Completo Acumulado" and not df_hist.empty:
    val_tot = int(df_hist["FO_Totales"].sum())
    val_int = int(df_hist["Integradas"].sum())
    val_at = int(df_hist["A_Tiempo"].sum())
    val_ret = int(df_hist["Con_Retraso"].sum())
    val_vig = int(df_hist["Vigentes"].iloc[-1]) if not df_hist.empty else 0
    val_ven = int(df_hist["Vencidas"].sum())
    val_rep = int(df_hist["FO_Reportadas"].sum())
    val_evt = int(df_hist["FO_Con_Evento"].sum())
    val_ret_rep = int(df_hist["Retrasos_Reportados"].sum())
    eval_tot = val_at + val_ret + val_ven
    pct_at = (val_at / eval_tot * 100) if eval_tot > 0 else 0.0
    pct_v = (val_ven / val_int * 100) if val_int > 0 else 0.0
elif fecha_consultada != "📅 Carga Actual Local" and not df_hist.empty:
    fila_dia = df_hist[df_hist["Fecha_Registro"].astype(str) == fecha_consultada].iloc[-1]
    val_tot = int(fila_dia["FO_Totales"])
    val_int = int(fila_dia["Integradas"])
    val_at = int(fila_dia["A_Tiempo"])
    val_ret = int(fila_dia["Con_Retraso"])
    val_vig = int(fila_dia["Vigentes"])
    val_ven = int(fila_dia["Vencidas"])
    val_rep = int(fila_dia.get("FO_Reportadas", 0))
    val_evt = int(fila_dia.get("FO_Con_Evento", 0))
    val_ret_rep = int(fila_dia.get("Retrasos_Reportados", 0))
    pct_at = float(fila_dia["Pct_A_Tiempo"])
    pct_v = float(fila_dia["Pct_Vencidas"])
else:
    val_tot = st.session_state.totales.get("total_fo", 0) if st.session_state.get("totales") else len(df_cruce)
    val_int = st.session_state.totales.get("integradas", len(df_cruce)) if st.session_state.get("totales") else len(df_cruce)
    val_at = st.session_state.totales.get("a_tiempo", 0) if st.session_state.get("totales") else 0
    val_ret = len(df_cruce[df_cruce["Estado de la ruta"].str.contains("Retraso", na=False)]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
    val_vig = len(df_cruce[df_cruce["Estado de la ruta"] == "Vigente"]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
    val_ven = len(df_cruce[df_cruce["Estado de la ruta"].str.contains("Vencida", na=False)]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
    val_rep = len(df_cruce[df_cruce["Estado_Reporte"] == "🚨 REPORTADA"]) if not df_cruce.empty and "Estado_Reporte" in df_cruce.columns else 0
    val_evt = len(df_cruce[df_cruce["Evento"].notna() & (df_cruce["Evento"].astype(str).str.strip() != "-") & (df_cruce["Evento"].astype(str).str.strip() != "")]) if not df_cruce.empty and "Evento" in df_cruce.columns else 0
    
    if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns and "Estado_Reporte" in df_cruce.columns:
        val_ret_rep = len(df_cruce[(df_cruce["Estado de la ruta"].str.contains("Retraso", na=False)) & (df_cruce["Estado_Reporte"] == "🚨 REPORTADA")])
    else:
        val_ret_rep = 0

    eval_tot = val_at + val_ret + val_ven
    pct_at = (val_at / eval_tot * 100) if eval_tot > 0 else 0.0
    pct_v = (val_ven / val_int * 100) if val_int > 0 else 0.0

pct_rep_total = (val_rep / val_int * 100) if val_int > 0 else 0.0
pct_ret_justificados = (val_ret_rep / val_ret * 100) if val_ret > 0 else 0.0
pct_evt_total = (val_evt / val_int * 100) if val_int > 0 else 0.0

if val_int > 0:
    st.markdown("##### 📌 Indicadores Clave del Servicio")
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f'<div class="kpi-card kpi-info"><div class="kpi-title">FO Integradas</div><div class="kpi-value">{val_int}</div><div class="kpi-sub" style="color: #3498db;">Total evaluables: {val_tot}</div></div>', unsafe_allow_html=True)
    with c2:
        at_color = "#2ecc71" if pct_at >= 90 else "#e74c3c"
        st.markdown(f'<div class="kpi-card {"kpi-success" if pct_at >= 90 else "kpi-danger"}"><div class="kpi-title">Finalizadas a Tiempo</div><div class="kpi-value">{pct_at:.1f}%</div><div class="kpi-sub" style="color: {at_color};">{val_at} FO | Meta >= 90%</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="kpi-card kpi-warning"><div class="kpi-title">Con Retraso</div><div class="kpi-value">{(val_ret/val_int*100):.1f}%</div><div class="kpi-sub" style="color: #e67e22;">{val_ret} FO entregadas tarde</div></div>', unsafe_allow_html=True)
    with c4:
        v_color = "#2ecc71" if pct_v <= 5 else "#e74c3c"
        st.markdown(f'<div class="kpi-card {"kpi-success" if pct_v <= 5 else "kpi-danger"}"><div class="kpi-title">FO Vencidas</div><div class="kpi-value">{pct_v:.1f}%</div><div class="kpi-sub" style="color: {v_color};">{val_ven} FO | Límite <= 5%</div></div>', unsafe_allow_html=True)
    with c5:
        st.markdown(f'<div class="kpi-card kpi-info"><div class="kpi-title">FO Vigentes</div><div class="kpi-value">{val_vig}</div><div class="kpi-sub" style="color: #3498db;">En proceso dentro del plazo</div></div>', unsafe_allow_html=True)

    st.markdown("##### 🚨 Análisis de Reportes de Inconsistencias y Eventos")
    cr1, cr2, cr3 = st.columns(3)
    with cr1:
        st.markdown(f'<div class="kpi-card kpi-purple"><div class="kpi-title">FO Reportadas con Problema</div><div class="kpi-value">{pct_rep_total:.1f}%</div><div class="kpi-sub" style="color: #8e44ad;">{val_rep} de {val_int} FO integradas</div></div>', unsafe_allow_html=True)
    with cr2:
        st.markdown(f'<div class="kpi-card kpi-warning"><div class="kpi-title">Retrasos Reportados / Justificados</div><div class="kpi-value">{pct_ret_justificados:.1f}%</div><div class="kpi-sub" style="color: #d35400;">{val_ret_rep} de {val_ret} retrasos tienen reporte</div></div>', unsafe_allow_html=True)
    with cr3:
        st.markdown(f'<div class="kpi-card kpi-info"><div class="kpi-title">FO con Evento Registrado</div><div class="kpi-value">{pct_evt_total:.1f}%</div><div class="kpi-sub" style="color: #2980b9;">{val_evt} de {val_int} FO con evento anotado</div></div>', unsafe_allow_html=True)

st.divider()

# --- VISTA 1: PORTAL DE NAVEGACIÓN Y TARJETAS ---
if st.session_state.modulo_activo == "INICIO":
    st.markdown("### 🗂️ Selecciona un Módulo para Visualizar")
    st.caption("Haz clic en cualquier tarjeta para desplegar el módulo a pantalla completa.")

    b_col1, b_col2, b_col3, b_col4 = st.columns(4)

    with b_col1:
        st.markdown('<div class="nav-header">📋 Control de Rutas</div>', unsafe_allow_html=True)
        if st.button("📋 Vencidas y Vigentes", use_container_width=True):
            st.session_state.modulo_activo = "VENCIDAS_VIGENTES"
            st.rerun()
        if st.button("⏱️ Finalizadas con Retraso", use_container_width=True):
            st.session_state.modulo_activo = "RETRASO"
            st.rerun()

    with b_col2:
        st.markdown('<div class="nav-header">📊 Rendimiento & Metas</div>', unsafe_allow_html=True)
        if st.button("📅 Rendimiento Rango Fechas", use_container_width=True):
            st.session_state.modulo_activo = "RANGO_FECHAS"
            st.rerun()
        if st.button("📊 Metas y Gráficos (95%/90%)", use_container_width=True):
            st.session_state.modulo_activo = "METAS"
            st.rerun()
        if st.button("🏢 Rendimiento por Sucursales", use_container_width=True):
            st.session_state.modulo_activo = "SUCURSALES"
            st.rerun()

    with b_col3:
        st.markdown('<div class="nav-header">🚨 Gestión de Inconsistencias</div>', unsafe_allow_html=True)
        if st.button("📝 Reportar / Editar FO", use_container_width=True):
            st.session_state.modulo_activo = "REPORTAR"
            st.rerun()
        if st.button("📊 Reportes para el Mandante", use_container_width=True):
            st.session_state.modulo_activo = "MANDANTE"
            st.rerun()

    with b_col4:
        st.markdown('<div class="nav-header">🏆 Histórico & Conductores</div>', unsafe_allow_html=True)
        if st.button("🏆 Ranking Conductores", use_container_width=True):
            st.session_state.modulo_activo = "RANKING"
            st.rerun()
        if st.button("📈 Evolución Histórica", use_container_width=True):
            st.session_state.modulo_activo = "HISTORICO"
            st.rerun()

# --- VISTA 2: NAVEGACIÓN DIRECTA ENTRE MÓDULOS ---
else:
    col_btn_back, col_nav_bar = st.columns([1, 4])
    with col_btn_back:
        if st.button("⬅️ Volver al Menú Principal", type="primary", use_container_width=True):
            st.session_state.modulo_activo = "INICIO"
            st.rerun()

    with col_nav_bar:
        n_col1, n_col2, n_col3, n_col4, n_col5, n_col6, n_col7, n_col8, n_col9 = st.columns(9)
        with n_col1:
            if st.button("📋 V&V", use_container_width=True, help="Vencidas y Vigentes"):
                st.session_state.modulo_activo = "VENCIDAS_VIGENTES"
                st.rerun()
        with n_col2:
            if st.button("📅 Fechas", use_container_width=True, help="Rendimiento Rango Fechas"):
                st.session_state.modulo_activo = "RANGO_FECHAS"
                st.rerun()
        with n_col3:
            if st.button("📊 Metas", use_container_width=True, help="Metas y Gráficos (95%/90%)"):
                st.session_state.modulo_activo = "METAS"
                st.rerun()
        with n_col4:
            if st.button("🏆 Ranking", use_container_width=True, help="Ranking Conductores"):
                st.session_state.modulo_activo = "RANKING"
                st.rerun()
        with n_col5:
            if st.button("🏢 Sucursales", use_container_width=True, help="Rendimiento por Sucursales"):
                st.session_state.modulo_activo = "SUCURSALES"
                st.rerun()
        with n_col6:
            if st.button("⏱️ Retraso", use_container_width=True, help="Finalizadas con Retraso"):
                st.session_state.modulo_activo = "RETRASO"
                st.rerun()
        with n_col7:
            if st.button("📝 Reportar", use_container_width=True, help="Reportar FO"):
                st.session_state.modulo_activo = "REPORTAR"
                st.rerun()
        with n_col8:
            if st.button("📊 Mandante", use_container_width=True, help="Reportes Mandante"):
                st.session_state.modulo_activo = "MANDANTE"
                st.rerun()
        with n_col9:
            if st.button("📈 Histórico", use_container_width=True, help="Evolución Histórica"):
                st.session_state.modulo_activo = "HISTORICO"
                st.rerun()

    st.divider()

    # MÓDULO 1: VENCIDAS Y VIGENTES
    if st.session_state.modulo_activo == "VENCIDAS_VIGENTES":
        if not df_cruce.empty:
            st.subheader("📋 Tabla Unificada de Vencidas y Vigentes")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                est_opts = df_cruce["Estado de la ruta"].unique()
                estado_sel = st.multiselect("Filtrar Estado de la Ruta:", options=est_opts, default=est_opts)
            with col_f2:
                suc_opts = df_cruce["SUCURSAL"].unique()
                sucursal_sel = st.multiselect("Filtrar Sucursal:", options=suc_opts, default=suc_opts)

            df_vv_filtrado = df_cruce[df_cruce["Estado de la ruta"].isin(estado_sel)]
            if sucursal_sel:
                df_vv_filtrado = df_vv_filtrado[df_vv_filtrado["SUCURSAL"].isin(sucursal_sel)]

            st.dataframe(
                df_vv_filtrado.style.map(color_filas_estado, subset=["Estado de la ruta", "Estado_Reporte"]),
                use_container_width=True,
                height=500,
            )
        else:
            st.info("👈 Carga y procesa archivos desde la barra lateral para desplegar los datos.")

    # MÓDULO 2: RANGO DE FECHAS
    elif st.session_state.modulo_activo == "RANGO_FECHAS":
        st.subheader("📅 Filtrado por Rango de Fechas Planificadas")
        if not df_cruce.empty:
            df_rf = df_cruce.copy()

            def parsear_fecha(val):
                if pd.isna(val) or str(val).strip() in ["-", "", "nan", "None"]:
                    return pd.NaT
                try:
                    return pd.to_datetime(val, dayfirst=True, errors="coerce")
                except Exception:
                    return pd.NaT

            df_rf["Fecha_Inicio_dt"] = (
                df_rf["Fin plan"].apply(parsear_fecha)
                if "Fin plan" in df_rf.columns
                else df_rf["Fin Planificado"].apply(parsear_fecha)
            )

            df_rf["Fecha_Final_Eval"] = df_rf["Fecha_Inicio_dt"].dt.date
            df_valid_dates = df_rf.dropna(subset=["Fecha_Final_Eval"])

            if not df_valid_dates.empty:
                min_f_plan = df_valid_dates["Fecha_Final_Eval"].min()
                max_f_plan = df_valid_dates["Fecha_Final_Eval"].max()

                col_rf1, col_rf2 = st.columns(2)
                with col_rf1:
                    f_desde = st.date_input("📅 Fecha Desde:", value=min_f_plan, key="f_desde_plan")
                with col_rf2:
                    f_hasta = st.date_input("📅 Fecha Hasta:", value=max_f_plan, key="f_hasta_plan")

                if f_desde > f_hasta:
                    st.error("⚠️ La 'Fecha Desde' no puede ser posterior a la 'Fecha Hasta'.")
                else:
                    dias_rango = (f_hasta - f_desde).days + 1
                    servicios_proporcionales_rango = max(1, int((servicios_mensuales_admin / 30) * dias_rango))

                    df_filtrado_desviaciones = df_rf[(df_rf["Fecha_Final_Eval"] >= f_desde) & (df_rf["Fecha_Final_Eval"] <= f_hasta)]

                    ret_r = len(df_filtrado_desviaciones[df_filtrado_desviaciones["Estado de la ruta"].str.contains("Retraso", na=False)])
                    ven_r = len(df_filtrado_desviaciones[df_filtrado_desviaciones["Estado de la ruta"].str.contains("Vencida", na=False)])
                    vig_r = len(df_filtrado_desviaciones[df_filtrado_desviaciones["Estado de la ruta"] == "Vigente"])
                    rec_r = len(df_filtrado_desviaciones[df_filtrado_desviaciones["Estado de la ruta"] == "Rechazada"])

                    total_r = max(servicios_proporcionales_rango, (ret_r + ven_r + vig_r + rec_r))
                    at_r = max(0, total_r - (ret_r + ven_r + vig_r + rec_r))

                    pct_at_r = (at_r / total_r * 100) if total_r > 0 else 0.0
                    pct_v_r = (ven_r / total_r * 100) if total_r > 0 else 0.0

                    st.markdown(f"##### 📊 Resultados para los **{dias_rango} días del periodo** (**{f_desde.strftime('%d/%m/%Y')}** al **{f_hasta.strftime('%d/%m/%Y')}**)")
                    rk1, rk2, rk3, rk4, rk5 = st.columns(5)
                    with rk1:
                        st.markdown(f'<div class="kpi-card kpi-info"><div class="kpi-title">Total Estimado</div><div class="kpi-value">{total_r}</div></div>', unsafe_allow_html=True)
                    with rk2:
                        st.markdown(f'<div class="kpi-card kpi-success"><div class="kpi-title">% On-Time Est.</div><div class="kpi-value">{pct_at_r:.1f}%</div></div>', unsafe_allow_html=True)
                    with rk3:
                        st.markdown(f'<div class="kpi-card kpi-warning"><div class="kpi-title">Con Retraso</div><div class="kpi-value">{(ret_r/total_r*100):.1f}%</div></div>', unsafe_allow_html=True)
                    with rk4:
                        st.markdown(f'<div class="kpi-card kpi-danger"><div class="kpi-title">% Vencidas</div><div class="kpi-value">{pct_v_r:.1f}%</div></div>', unsafe_allow_html=True)
                    with rk5:
                        st.markdown(f'<div class="kpi-card kpi-info"><div class="kpi-title">FO Vigentes</div><div class="kpi-value">{vig_r}</div></div>', unsafe_allow_html=True)

                    st.divider()
                    gr1, gr2 = st.columns(2)
                    with gr1:
                        df_bar_plan = df_filtrado_desviaciones.groupby(["Fecha_Final_Eval", "Estado de la ruta"]).size().reset_index(name="Cantidad")
                        fig_r_bar = px.bar(df_bar_plan, x="Fecha_Final_Eval", y="Cantidad", color="Estado de la ruta", title="Desviaciones por Fecha Planificada", barmode="stack")
                        st.plotly_chart(fig_r_bar, use_container_width=True)
                    with gr2:
                        df_pie_data = pd.DataFrame({"Estado": ["A Tiempo (Est.)", "Con Retraso", "Vencida", "Vigente"], "Cantidad": [at_r, ret_r, ven_r, vig_r]})
                        fig_r_pie = px.pie(df_pie_data[df_pie_data["Cantidad"] > 0], names="Estado", values="Cantidad", title="Distribución Proporcional", hole=0.4)
                        st.plotly_chart(fig_r_pie, use_container_width=True)

                    st.dataframe(df_filtrado_desviaciones.style.map(color_filas_estado, subset=["Estado de la ruta"]), use_container_width=True)
            else:
                st.warning("No se pudieron interpretar fechas válidas en la tabla.")

    # MÓDULO 3: METAS Y GRÁFICOS
    elif st.session_state.modulo_activo == "METAS":
        st.subheader("📊 Evaluación de Cumplimiento de Metas Operativas (95% y 90%)")
        if not df_cruce.empty:
            cf1, cf2 = st.columns(2)
            with cf1:
                suc_disponibles = sorted(df_cruce["SUCURSAL"].unique().tolist())
                suc_seleccionadas = st.multiselect("Filtrar por Sucursal(es):", options=suc_disponibles, default=suc_disponibles, key="filter_suc_metas")
            with cf2:
                resp_disponibles = sorted(df_cruce["RESPONSABLE"].unique().tolist())
                resp_seleccionados = st.multiselect("Filtrar por Responsable KCC:", options=resp_disponibles, default=resp_disponibles, key="filter_resp_metas")

            df_m_filtrado = df_cruce[(df_cruce["SUCURSAL"].isin(suc_seleccionadas)) & (df_cruce["RESPONSABLE"].isin(resp_seleccionados))]
            tot_f = len(df_m_filtrado)
            if tot_f > 0:
                cnt_at = len(df_m_filtrado[df_m_filtrado["Estado de la ruta"].isin(["Finalizada a tiempo", "A Tiempo"])])
                cnt_ret = len(df_m_filtrado[df_m_filtrado["Estado de la ruta"].str.contains("Retraso", na=False)])

                calc_tot = val_tot if (len(suc_seleccionadas) == len(suc_disponibles) and len(resp_seleccionados) == len(resp_seleccionados)) else tot_f
                calc_at = val_at if (len(suc_seleccionadas) == len(suc_disponibles) and len(resp_seleccionados) == len(resp_seleccionados)) else cnt_at
                calc_ret = val_ret if (len(suc_seleccionadas) == len(suc_disponibles) and len(resp_seleccionados) == len(resp_seleccionados)) else cnt_ret

                pct_fin_tot = ((calc_at + calc_ret) / calc_tot * 100) if calc_tot > 0 else 0.0
                pct_a_tiempo_tot = (calc_at / calc_tot * 100) if calc_tot > 0 else 0.0

                g1, g2 = st.columns(2)
                with g1:
                    fig_g95 = go.Figure(go.Indicator(mode="gauge+number+delta", value=pct_fin_tot, title={"text": "% Finalizadas (Meta 95%)"}, delta={"reference": 95.0}))
                    st.plotly_chart(fig_g95, use_container_width=True)
                with g2:
                    fig_g90 = go.Figure(go.Indicator(mode="gauge+number+delta", value=pct_a_tiempo_tot, title={"text": "% A Tiempo (Meta 90%)"}, delta={"reference": 90.0}))
                    st.plotly_chart(fig_g90, use_container_width=True)

    # MÓDULO 4: RANKING CONDUCTORES
    elif st.session_state.modulo_activo == "RANKING":
        st.subheader("🏆 Ranking de Cumplimiento de Conductores")
        if df_ranking is not None and not df_ranking.empty:
            st.dataframe(df_ranking, use_container_width=True, height=500)
        else:
            st.info("Sin datos de ranking actualmente en la planilla cargada.")

    # MÓDULO 5: SUCURSALES (MATRIZ Y TABLA RESULTADOS MATRIMONIADA CON EL EXCEL MANUAL)
    elif st.session_state.modulo_activo == "SUCURSALES":
        st.subheader("🏢 Tabla de RESULTADOS por Sucursal y Estado de la Ruta")
        if not df_cruce.empty:
            df_suc = df_cruce.copy()
            df_suc["SUCURSAL_CLEAN"] = df_suc["SUCURSAL"].astype(str).str.strip().str.title()

            suc_disponibles = sorted([s for s in df_suc["SUCURSAL_CLEAN"].dropna().unique() if str(s).strip()])
            suc_seleccionadas = st.multiselect(
                "🔍 Filtrar Sucursales a analizar:",
                options=suc_disponibles,
                default=suc_disponibles,
                key="filter_modulo_sucursales"
            )

            if suc_seleccionadas:
                df_suc = df_suc[df_suc["SUCURSAL_CLEAN"].isin(suc_seleccionadas)]

            if not df_suc.empty:
                # Tabla Cruzada de Resultados (Igual a la pestaña RESULTADOS del Excel)
                tabla_resultados = pd.crosstab(
                    df_suc["SUCURSAL_CLEAN"],
                    df_suc["Estado de la ruta"],
                    margins=True,
                    margins_name="Total general"
                )

                st.markdown("##### 📊 Consolidado Dinámico de Resultados")
                st.dataframe(tabla_resultados, use_container_width=True)

                st.divider()

                # Gráficos por Sucursal
                df_graf_suc = df_suc.groupby(["SUCURSAL_CLEAN", "Estado de la ruta"]).size().reset_index(name="Cantidad")
                fig_suc_bar = px.bar(
                    df_graf_suc,
                    x="SUCURSAL_CLEAN",
                    y="Cantidad",
                    color="Estado de la ruta",
                    title="📊 Distribución de Operaciones por Sucursal",
                    barmode="stack",
                    text_auto=True
                )
                st.plotly_chart(fig_suc_bar, use_container_width=True)

            else:
                st.warning("No hay datos disponibles para las sucursales seleccionadas.")
        else:
            st.info("Sin información cargada para calcular los indicadores por sucursal.")

    # MÓDULO 6: RETRASO
    elif st.session_state.modulo_activo == "RETRASO":
        st.subheader("⏱️ FO Finalizadas con Retraso")
        if not df_cruce.empty:
            df_solo_retraso = df_cruce[df_cruce["Estado de la ruta"].str.contains("Retraso", na=False)].copy()
            if not df_solo_retraso.empty:
                st.dataframe(df_solo_retraso.style.map(color_filas_estado, subset=["Estado_Reporte"]), use_container_width=True, height=500)
            else:
                st.info("No hay rutas registradas con retraso en la carga actual.")

    # MÓDULO 7: FORMULARIO DE REPORTES
    elif st.session_state.modulo_activo == "REPORTAR":
        st.subheader("📝 Registro y Edición de Reportes FO (Sincronizado con Google Sheets)")
        df_reps_actuales = cargar_reportes_fo()
        opciones_modo = ["➕ Nuevo Reporte"]
        if not df_reps_actuales.empty:
            opciones_modo.append("✏️ Editar Reporte Existente")

        modo_reporte = st.radio("Acción a realizar:", opciones_modo, horizontal=True)
        lista_sucursales = ["La Negra", "La Chimba", "Copiapó", "Coquimbo", "Santiago", "Otra"]
        lista_problemas = cargar_opciones_problemas()
        lista_coordinadores_komatsu = cargar_opciones_coordinadores()

        if modo_reporte == "➕ Nuevo Reporte":
            with st.form(key="form_rep_nuevo", clear_on_submit=True):
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    fo_in = st.text_input("Número de FO:")
                    suc_in = st.selectbox("Sucursal:", lista_sucursales)
                    prob_in = st.selectbox("Problema / Inconsistencia:", lista_problemas)
                with col_f2:
                    coord_sel = st.selectbox("Coordinador Komatsu / KCC:", lista_coordinadores_komatsu)
                    resp_in = st.text_input("Reportado por (Usuario SAE):")
                    arch_in = st.file_uploader("Adjuntar Respaldo (Imagen/PDF):", type=["png", "jpg", "jpeg", "pdf", "msg", "eml", "txt"])

                obs_in = st.text_area("Observaciones / Justificación:")
                if st.form_submit_button("📩 Guardar en Google Sheets", type="primary", use_container_width=True):
                    if fo_in and resp_in:
                        guardar_reporte_fo(fo_in, suc_in, prob_in, coord_sel, resp_in, obs_in, arch_in)
                        st.success("✅ Reporte guardado y sincronizado exitosamente.")
                        st.rerun()

        else:
            fos_reportadas_list = [str(f) for f in df_reps_actuales["FO"].unique() if str(f).strip()]
            fo_a_editar = st.selectbox("Seleccionar FO a Editar:", options=fos_reportadas_list)
            fo_a_editar_limpia = limpiar_fo_str(fo_a_editar)
            row_edit = df_reps_actuales[df_reps_actuales["FO"] == fo_a_editar_limpia].iloc[-1]

            idx_suc = lista_sucursales.index(row_edit["Sucursal"]) if row_edit["Sucursal"] in lista_sucursales else 0
            idx_prob = lista_problemas.index(row_edit["Tipo_Problema"]) if row_edit["Tipo_Problema"] in lista_problemas else 0
            coord_actual = str(row_edit.get("Coordinador_Komatsu", "-"))
            idx_coord = lista_coordinadores_komatsu.index(coord_actual) if coord_actual in lista_coordinadores_komatsu else 0

            with st.form(key="form_rep_editar"):
                st.info(f"Modificando datos para la **FO {fo_a_editar_limpia}**")
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    suc_edit = st.selectbox("Sucursal:", lista_sucursales, index=idx_suc)
                    prob_edit = st.selectbox("Problema / Inconsistencia:", lista_problemas, index=idx_prob)
                    coord_edit_sel = st.selectbox("Coordinador Komatsu / KCC:", lista_coordinadores_komatsu, index=idx_coord)
                with col_e2:
                    resp_edit = st.text_input("Reportado por (Usuario SAE):", value=str(row_edit.get("Responsable", "")))
                    arch_edit = st.file_uploader("Reemplazar Respaldo (Opcional):", type=["png", "jpg", "jpeg", "pdf", "msg", "eml", "txt"])

                obs_edit = st.text_area("Observaciones:", value=str(row_edit.get("Detalle_Observacion", "")))

                c_ed1, c_ed2 = st.columns(2)
                with c_ed1:
                    btn_actualizar = st.form_submit_button("🔄 Guardar Cambios", type="primary", use_container_width=True)
                with c_ed2:
                    btn_eliminar = st.form_submit_button("🔥 Eliminar Reporte", use_container_width=True)

                if btn_actualizar:
                    if editar_reporte_fo(fo_a_editar_limpia, suc_edit, prob_edit, coord_edit_sel, resp_edit, obs_edit, arch_edit):
                        st.success("✅ Reporte actualizado con éxito.")
                        st.rerun()

                if btn_eliminar:
                    if eliminar_reporte_fo(fo_a_editar_limpia):
                        st.success("🗑️ Reporte eliminado de Google Sheets.")
                        st.rerun()

    # MÓDULO 8: MANDANTE
    elif st.session_state.modulo_activo == "MANDANTE":
        st.subheader("📊 Informe Ejecutivo de Inconsistencias para el Mandante")
        df_reps = cargar_reportes_fo()
        if not df_reps.empty:
            df_graf = df_reps.copy()
            df_graf["Tipo_Problema"] = df_graf["Tipo_Problema"].fillna("Sin especificar")
            df_graf["Coordinador_Komatsu"] = df_graf["Coordinador_Komatsu"].fillna("Sin asignar")

            g_col1, g_col2 = st.columns(2)
            with g_col1:
                df_prob_cnt = df_graf.groupby("Tipo_Problema").size().reset_index(name="Cantidad").sort_values(by="Cantidad", ascending=True)
                fig_prob = px.bar(df_prob_cnt, x="Cantidad", y="Tipo_Problema", orientation="h", text="Cantidad", title="📌 FO Reportadas por Problema")
                fig_prob.update_traces(textposition="outside")
                st.plotly_chart(fig_prob, use_container_width=True)

            with g_col2:
                df_coord_cnt = df_graf.groupby("Coordinador_Komatsu").size().reset_index(name="Cantidad").sort_values(by="Cantidad", ascending=True)
                fig_coord_rep = px.bar(df_coord_cnt, x="Cantidad", y="Coordinador_Komatsu", orientation="h", text="Cantidad", title="👤 FO por Coordinador Komatsu")
                fig_coord_rep.update_traces(textposition="outside")
                fig_coord_rep.update_layout(margin=dict(l=150, r=20, t=40, b=40))
                st.plotly_chart(fig_coord_rep, use_container_width=True)

            st.divider()
            st.markdown("##### 📋 Registro Detallado de FO Reportadas")
            st.dataframe(df_reps, use_container_width=True, height=400)

            excel_bytes_reportes = generar_excel_solo_tabla(df_reps, nombre_hoja="FO_Reportadas")
            st.download_button(
                label="📥 Descargar Registro de FO Reportadas en Excel",
                data=excel_bytes_reportes,
                file_name=f"Registro_FO_Reportadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            st.divider()
            st.markdown("##### 🔍 Consultar y Visualizar Respaldo Adjunto")
            lista_fos_reportadas = [str(f) for f in df_reps["FO"].unique() if str(f).strip()]
            fo_sel_ver = st.selectbox("Seleccionar FO para revisar detalles y respaldo:", options=lista_fos_reportadas, key="sb_ver_img_mandante")

            fo_ver_limpia = limpiar_fo_str(fo_sel_ver)
            row_ver = df_reps[df_reps["FO"] == fo_ver_limpia].iloc[-1]
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.write(f"**Fecha Reporte:** {row_ver.get('Fecha_Reporte', '-')}")
                st.write(f"**Coordinador Komatsu / KCC:** {row_ver.get('Coordinador_Komatsu', '-')}")
                st.write(f"**Reportado por:** {row_ver.get('Responsable', '-')}")
            with col_d2:
                st.write(f"**Sucursal:** {row_ver.get('Sucursal', '-')}")
                st.write(f"**Inconsistencia:** {row_ver.get('Tipo_Problema', '-')}")

            st.write(f"**Observación:** {row_ver.get('Detalle_Observacion', '-')}")

            archivo_nombre = str(row_ver.get("Archivo_Pantallazo", "Sin respaldo"))
            if archivo_nombre != "Sin respaldo" and archivo_nombre != "nan":
                path_img_local = os.path.join(CAPTURAS_DIR, archivo_nombre)
                if os.path.exists(path_img_local):
                    ext_ver = os.path.splitext(archivo_nombre)[1].lower()
                    if ext_ver in [".png", ".jpg", ".jpeg"]:
                        st.image(path_img_local, caption=f"Respaldo FO {fo_ver_limpia}", use_container_width=True)
                    else:
                        with open(path_img_local, "rb") as f_down:
                            st.download_button(label=f"📥 Descargar Archivo Respaldo ({ext_ver.upper()})", data=f_down, file_name=archivo_nombre)
                else:
                    st.warning("⚠️ No se encontró la imagen en la carpeta local de capturas.")
            else:
                st.info("ℹ️ Este reporte no adjuntó ningún archivo de respaldo.")
        else:
            st.info("Aún no hay reportes de FO registrados en la hoja de Google Sheets para mostrar al mandante.")

    # MÓDULO 9: HISTÓRICO Y CIERRES
    elif st.session_state.modulo_activo == "HISTORICO":
        st.subheader("📈 Evolución Histórica de Cumplimiento")
        df_hist = cargar_historico()
        df_papelera = cargar_papelera()

        if not df_hist.empty:
            df_hist_val = df_hist[df_hist["FO_Totales"] > 0].copy()
            if not df_hist_val.empty:
                fig_hist = go.Figure()
                fig_hist.add_trace(go.Scatter(x=df_hist_val["Fecha_Registro"], y=df_hist_val["Pct_A_Tiempo"], mode="lines+markers+text", name="% A Tiempo", text=[f"{v:.1f}%" for v in df_hist_val["Pct_A_Tiempo"]], textposition="top center"))
                fig_hist.add_trace(go.Scatter(x=df_hist_val["Fecha_Registro"], y=df_hist_val["Pct_Vencidas"], mode="lines+markers+text", name="% Vencidas", text=[f"{v:.2f}%" for v in df_hist_val["Pct_Vencidas"]], textposition="bottom center"))
                st.plotly_chart(fig_hist, use_container_width=True)

            st.dataframe(df_hist, use_container_width=True, height=400)
        else:
            st.info("Aún no hay datos guardados en el histórico para graficar.")

        st.divider()
        st.markdown("##### 📂 Registro Histórico de Meses Cerrados")
        if os.path.exists(HISTORICO_CIERRES_FILE):
            df_cierres_ver = pd.read_csv(HISTORICO_CIERRES_FILE)
            st.dataframe(df_cierres_ver, use_container_width=True, height=200)
        else:
            st.caption("Aún no se han ejecutado cierres de mes en la aplicación.")

        st.divider()
        st.markdown("##### 🗑️ Papelera de Respaldo")
        if not df_papelera.empty:
            st.dataframe(df_papelera, use_container_width=True, height=200)
        else:
            st.caption("Papelera vacía.")
