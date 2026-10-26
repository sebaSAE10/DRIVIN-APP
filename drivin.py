import io
import json
import os
import pickle
import re
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pypdf
import streamlit as st

# Configuración de la página Streamlit
st.set_page_config(
    page_title="Control Operativo Drivin x SAE x SmartSheet",
    page_icon="🚚",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilos CSS
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
</style>
""",
    unsafe_allow_html=True,
)

# RUTAS DE ARCHIVOS BASE EN GITHUB Y ARCHIVOS DEL SISTEMA
DRIVIN_BASE_FILE = "drivin_base.xlsx"
SAE_BASE_FILE = "sae_base.xlsx"
SMARTSHEET_BASE_FILE = "smartsheet_base.xlsx"

HISTORICO_FILE = "historico_drivin.csv"
PAPELERA_HISTORICO_FILE = "papelera_historico.csv"
REPORTES_FILE = "reportes_fo_problemas.csv"
PROBLEMAS_CONFIG_FILE = "opciones_problemas.json"
COORDINADORES_CONFIG_FILE = "opciones_coordinadores.json"
CAPTURAS_DIR = "capturas_reportes"

CRUCE_CACHE_FILE = "ultimo_cruce.parquet"
TOTALES_CACHE_FILE = "totales_cache.pkl"
RANKING_CACHE_FILE = "ranking_cache.parquet"
INFO_ACTUALIZACION_FILE = "ultima_actualizacion.txt"

if not os.path.exists(CAPTURAS_DIR):
  os.makedirs(CAPTURAS_DIR)


# --- MANEJO DE OPCIONES DINÁMICAS DE PROBLEMAS ---
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


# --- MANEJO DE OPCIONES DINÁMICAS DE COORDINADORES KOMATSU / KCC ---
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


# --- FUNCIONES LOCALES DE LIMPIEZA Y CÁLCULO ---
def limpiar_fo_str(val):
  if pd.isna(val) or str(val).strip() in ["", "-", "nan", "None"]:
    return ""
  txt = str(val).strip()
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


# --- MANEJO DE CACHÉ Y ARCHIVOS LOCALES ---
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
    st.error(f"Error guardando archivos locales: {e}")


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
          "Fecha_Registro",
          "FO_Totales",
          "Integradas",
          "A_Tiempo",
          "Con_Retraso",
          "Vigentes",
          "Vencidas",
          "Rechazadas",
          "FO_Reportadas",
          "FO_Con_Evento",
          "Retrasos_Reportados",
          "Pct_A_Tiempo",
          "Pct_Vencidas",
      ]
  )


def cargar_papelera():
  if os.path.exists(PAPELERA_HISTORICO_FILE):
    return pd.read_csv(PAPELERA_HISTORICO_FILE)
  return pd.DataFrame()


def guardar_nuevo_registro(registro):
  df_hist = cargar_historico()
  fecha_str = str(registro["Fecha_Registro"])
  if fecha_str in df_hist["Fecha_Registro"].astype(str).values:
    st.toast(f"⚠️ El día {fecha_str} ya existe en el histórico.", icon="⚠️")
    return False
  else:
    df_hist = pd.concat([df_hist, pd.DataFrame([registro])], ignore_index=True)
    df_hist["Fecha_Registro"] = pd.to_datetime(df_hist["Fecha_Registro"]).dt.date
    df_hist = df_hist.sort_values(by="Fecha_Registro")
    df_hist.to_csv(HISTORICO_FILE, index=False)
    st.toast(f"✅ Se guardó el día {fecha_str}.", icon="✅")
    return True


def actualizar_registro_existente(registro):
  df_hist = cargar_historico()
  fecha_str = str(registro["Fecha_Registro"])
  if fecha_str in df_hist["Fecha_Registro"].astype(str).values:
    idx = df_hist[df_hist["Fecha_Registro"].astype(str) == fecha_str].index[0]
    for key, value in registro.items():
      df_hist.at[idx, key] = value
    df_hist["Fecha_Registro"] = pd.to_datetime(df_hist["Fecha_Registro"]).dt.date
    df_hist = df_hist.sort_values(by="Fecha_Registro")
    df_hist.to_csv(HISTORICO_FILE, index=False)
    st.toast(f"🔄 Se actualizó el día {fecha_str}.", icon="🔄")
    return True
  else:
    st.toast(f"❌ No existe un registro previo para {fecha_str}.", icon="❌")
    return False


def eliminar_registro_historico(fecha_eliminar_str, clave_ingresada):
  if clave_ingresada != "123":
    st.toast("🔑 Clave incorrecta. No se eliminó el registro.", icon="🔑")
    return False

  df_hist = cargar_historico()
  fecha_eliminar_str = str(fecha_eliminar_str)

  if fecha_eliminar_str not in df_hist["Fecha_Registro"].astype(str).values:
    st.toast(f"❌ No se encontró el día {fecha_eliminar_str} en el histórico.", icon="❌")
    return False

  fila_elim = df_hist[
      df_hist["Fecha_Registro"].astype(str) == fecha_eliminar_str
  ].copy()
  fila_elim["Fecha_Eliminacion"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

  df_papelera = cargar_papelera()
  df_papelera = pd.concat([df_papelera, fila_elim], ignore_index=True)
  df_papelera.to_csv(PAPELERA_HISTORICO_FILE, index=False)

  df_hist = df_hist[df_hist["Fecha_Registro"].astype(str) != fecha_eliminar_str]
  df_hist.to_csv(HISTORICO_FILE, index=False)

  st.toast(f"🗑️ Se eliminó el día {fecha_eliminar_str}.", icon="🗑️")
  return True


def cargar_reportes_fo():
  if os.path.exists(REPORTES_FILE):
    df = pd.read_csv(REPORTES_FILE, dtype={"FO": str})
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


def guardar_reporte_fo(
    fo, sucursal, problema, coord_komatsu, responsable, observacion, file_obj
):
  df_rep = cargar_reportes_fo()
  fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
  fo_limpia = limpiar_fo_str(fo)

  nombre_archivo = "Sin respaldo"
  tipo_res = "Sin archivo"

  if file_obj is not None:
    ext = os.path.splitext(file_obj.name)[1].lower()
    nombre_archivo = (
        f"FO_{fo_limpia}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}"
    )
    path_file = os.path.join(CAPTURAS_DIR, nombre_archivo)
    with open(path_file, "wb") as f:
      f.write(file_obj.getbuffer())

    if ext in [".png", ".jpg", ".jpeg"]:
      tipo_res = "Imagen / Foto"
    elif ext == ".pdf":
      tipo_res = "Documento PDF"
    else:
      tipo_res = "Correo / Mensaje"

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
  df_rep.to_csv(REPORTES_FILE, index=False)
  return True


def editar_reporte_fo(
    fo_target,
    sucursal,
    problema,
    coord_komatsu,
    responsable,
    observacion,
    file_obj=None,
):
  df_rep = cargar_reportes_fo()
  fo_target_limpia = limpiar_fo_str(fo_target)
  idx_matches = df_rep[df_rep["FO"] == fo_target_limpia].index

  if len(idx_matches) > 0:
    idx = idx_matches[-1]
    df_rep.at[idx, "Sucursal"] = sucursal
    df_rep.at[idx, "Tipo_Problema"] = problema
    df_rep.at[idx, "Coordinador_Komatsu"] = (
        coord_komatsu if coord_komatsu else "-"
    )
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
          "Imagen / Foto"
          if ext in [".png", ".jpg", ".jpeg"]
          else ("Documento PDF" if ext == ".pdf" else "Correo / Mensaje")
      )

    df_rep.to_csv(REPORTES_FILE, index=False)
    return True
  return False


def eliminar_reporte_fo(fo_target):
  df_rep = cargar_reportes_fo()
  fo_target_limpia = limpiar_fo_str(fo_target)
  if fo_target_limpia in df_rep["FO"].values:
    df_rep = df_rep[df_rep["FO"] != fo_target_limpia]
    df_rep.to_csv(REPORTES_FILE, index=False)
    return True
  return False


# --- FUNCIÓN AUXILIAR DE TIEMPO / DESVIACIÓN ---
def desviacion_a_minutos(val_str):
  if pd.isna(val_str) or str(val_str).strip() in ["-", "", "nan", "None"]:
    return None
  texto = str(val_str).strip().lower().replace("+", "")
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

  total = horas * 60 + minutos
  return total if total > 0 else None


def minutos_a_formato_horas(total_minutos):
  if pd.isna(total_minutos) or total_minutos is None or total_minutos <= 0:
    return "-"
  total_minutos = int(round(total_minutos))
  hrs = total_minutos // 60
  mins = total_minutos % 60
  return f"+{hrs}h {mins:02d}m"


# --- EXPORTAR EXCEL ---
def generar_excel_descarga(df_cruce, dict_metricas, df_ranking=None):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_met = pd.DataFrame([dict_metricas])
    df_met.to_excel(writer, sheet_name="Resumen_KPIs", index=False)
    df_cruce.to_excel(writer, sheet_name="Vencidas_y_Vigentes", index=False)
    df_retraso = df_cruce[
        df_cruce["Estado de la ruta"] == "Finalizada con Retraso"
    ]
    if not df_retraso.empty:
      df_retraso.to_excel(
          writer, sheet_name="Finalizadas_Con_Retraso", index=False
      )
    if df_ranking is not None and not df_ranking.empty:
      df_ranking.to_excel(writer, sheet_name="Ranking_Conductores", index=False)
  return output.getvalue()


def generar_excel_solo_tabla(df_tabla, nombre_hoja="Datos"):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_tabla.to_excel(writer, sheet_name=nombre_hoja, index=False)
  return output.getvalue()


def color_filas_estado(val):
  val_str = str(val)
  if val_str == "Vencida":
    return "background-color: #fbe9e7; color: #c62828; font-weight: bold;"
  elif val_str == "Vigente":
    return "background-color: #ffffff; color: #212529;"
  elif val_str == "Finalizada con Retraso":
    return "background-color: #fff3e0; color: #ef6c00; font-weight: bold;"
  elif val_str == "Rechazada":
    return "background-color: #e1bee7; color: #4a148c; font-weight: bold;"
  elif "REPORTADA" in val_str:
    return "background-color: #ffe0b2; color: #e65100; font-weight: bold;"
  return ""


# --- LECTORES AUXILIARES ---
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
      text_limpio = re.sub(
          r"^(descripción:|descripcion:)\s*", "", text, flags=re.IGNORECASE
      ).strip()
      return text_limpio if text_limpio else "PENDIENTE EN GO TRUCK"

    df_clean["FO"] = df_clean["FO"].apply(limpiar_fo_str)
    df_clean["ORIGEN_SM"] = df_clean["ORIGEN_SM"].apply(clean_desc)
    df_clean["DESTINO_SM"] = df_clean["DESTINO_SM"].apply(clean_desc)
    return df_clean.dropna(subset=["FO"]).drop_duplicates(subset=["FO"])
  except Exception:
    return pd.DataFrame()


def procesar_drivin_excel(file_excel_drivin):
  totales = {
      "total_fo": 0,
      "integradas": 0,
      "a_tiempo": 0,
      "con_retraso": 0,
      "vigentes": 0,
      "vencidas": 0,
      "rechazadas": 0,
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
      inicio_p = formatear_fecha_str(
          buscar_valor_columna_flexible(
              row, ["inicio plan", "inicio planificado", "inicio"]
          )
      )
      fin_p = formatear_fecha_str(
          buscar_valor_columna_flexible(
              row, ["fin plan", "fin planificado", "fin estimad"]
          )
      )
      val_fin_real = buscar_valor_columna_flexible(
          row,
          [
              "fin real",
              "fin de entrega real",
              "llegada real",
              "cierre real",
              "fecha real",
          ],
      )
      fin_r = formatear_fecha_str(val_fin_real)

      lista_fos.append({
          "FO": limpiar_fo_str(row["FO"]),
          "RESPONSABLE": str(row.get("Responsable KCC", "-")),
          "Fin plan": fin_p,
          "Fin real": fin_r,
          "Inicio Planificado": inicio_p,
          "Fin Planificado": fin_p,
          "Conductor": str(row.get("Conductor", "-")),
          "Estado de la ruta": (
              str(row.get("Estado", "Vigente")).strip().capitalize()
          ),
          "Evento": (
              "-" if pd.isna(row.get("Evento")) else str(row.get("Evento"))
          ),
          "Desviación": "-",
      })

  sheet_ret = None
  for k in sheet_names_map:
    if "retraso" in k or "finalizad" in k:
      sheet_ret = sheet_names_map[k]
      break

  if sheet_ret:
    df_ret = pd.read_excel(xls, sheet_name=sheet_ret)
    df_ret = df_ret.dropna(subset=["FO"])

    for _, row in df_ret.iterrows():
      inicio_p = formatear_fecha_str(
          buscar_valor_columna_flexible(
              row, ["inicio plan", "inicio planificado", "inicio"]
          )
      )
      fin_p = formatear_fecha_str(
          buscar_valor_columna_flexible(
              row, ["fin plan", "fin planificado", "fin estimad"]
          )
      )
      val_fin_real = buscar_valor_columna_flexible(
          row,
          [
              "fin real",
              "fin de entrega real",
              "llegada real",
              "cierre real",
              "fecha real",
          ],
      )
      fin_r = formatear_fecha_str(val_fin_real)

      desviacion_val = "-"
      val_desv = buscar_valor_columna_flexible(
          row, ["desviacion", "desfase", "delay"]
      )
      if val_desv:
        desviacion_val = str(val_desv).strip()

      if fin_r == "-" and fin_p != "-" and desviacion_val != "-":
        fin_r = calcular_fin_real_automatico(fin_p, desviacion_val)

      if desviacion_val == "-" and fin_p != "-" and fin_r != "-":
        desviacion_val = calcular_desviacion_automatica(fin_p, fin_r)

      lista_fos.append({
          "FO": limpiar_fo_str(row["FO"]),
          "RESPONSABLE": str(row.get("Responsable KCC", "-")),
          "Fin plan": fin_p,
          "Fin real": fin_r,
          "Inicio Planificado": inicio_p,
          "Fin Planificado": fin_p,
          "Conductor": str(row.get("Conductor", "-")),
          "Estado de la ruta": "Finalizada con Retraso",
          "Evento": (
              "-" if pd.isna(row.get("Evento")) else str(row.get("Evento"))
          ),
          "Desviación": desviacion_val,
      })

  sheet_rank = sheet_names_map.get("ranking")
  if sheet_rank:
    df_ranking = pd.read_excel(xls, sheet_name=sheet_rank)

  sheet_tot = (
      sheet_names_map.get("hoja4")
      or sheet_names_map.get("hoja 4")
      or sheet_names_map.get("totales")
  )
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
    df_fos["Prioridad"] = (
        df_fos["Estado de la ruta"]
        .map({
            "Vencida": 1,
            "Rechazada": 2,
            "Finalizada con Retraso": 3,
            "Vigente": 4,
        })
        .fillna(5)
    )
    df_fos = (
        df_fos.sort_values(by="Prioridad")
        .drop_duplicates(subset=["FO"], keep="first")
        .drop(columns=["Prioridad"])
    )

  if totales["total_fo"] == 0 and not df_fos.empty:
    totales["integradas"] = len(df_fos)
    totales["vencidas"] = len(
        df_fos[df_fos["Estado de la ruta"] == "Vencida"]
    )
    totales["vigentes"] = len(
        df_fos[df_fos["Estado de la ruta"] == "Vigente"]
    )
    totales["con_retraso"] = len(
        df_fos[df_fos["Estado de la ruta"] == "Finalizada con Retraso"]
    )
    totales["total_fo"] = totales["integradas"]

  return totales, df_fos, df_ranking


def procesar_drivin_pdf(file_pdf_drivin):
  totales = {
      "total_fo": 0,
      "integradas": 0,
      "a_tiempo": 0,
      "con_retraso": 0,
      "vigentes": 0,
      "vencidas": 0,
      "rechazadas": 0,
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
    if (
        "RUTAS QUE REQUIEREN" in line_up
        or "RUTAS VENCIDAS" in line_up
        or "FO VENCIDAS" in line_up
        or ("VENCIDAS" in line_up and "A TIEMPO" not in line_up)
    ):
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
      fin_real = (
          fechas[2]
          if len(fechas) > 2 and estado_linea == "Finalizada con Retraso"
          else "-"
      )

      desfase_txt = "-"
      m_desf = re.search(
          r"\b(\d+\s*d[ií]as?\s*\d+\s*hrs?|\d+\s*hrs?|\+\d+h\s*\d+m)\b",
          bloque_texto,
          re.IGNORECASE,
      )
      if m_desf:
        desfase_txt = m_desf.group(1)

      if fin_real == "-" and fin_plan != "-" and desfase_txt != "-":
        fin_real = calcular_fin_real_automatico(fin_plan, desfase_txt)

      if desfase_txt == "-" and fin_plan != "-" and fin_real != "-":
        desfase_txt = calcular_desviacion_automatica(fin_plan, fin_real)

      resp_kcc = "-"
      m_kcc = re.search(
          r"\b(Danna\s+[A-Za-z\-]+|Nicol\s+[A-Za-z\s]+|Yheri\s+[A-Za-z\s]+|Carolina\s+[A-Za-z\s]+|Karina\s+[A-Za-z\s]+|Patricia\s+[A-Za-z\s]+|Elizabeth\s+[A-Za-z\s]+|Diego\s+[A-Za-z\s]+|Camila\s+[A-Za-z\s]+|Franco\s+[A-Za-z\s]+)\b",
          bloque_texto,
      )
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
      })

  df_fos = pd.DataFrame(lista_fos)
  if not df_fos.empty:
    df_fos["Prioridad"] = (
        df_fos["Estado de la ruta"]
        .map({
            "Vencida": 1,
            "Rechazada": 2,
            "Finalizada con Retraso": 3,
            "Vigente": 4,
        })
        .fillna(5)
    )
    df_fos = (
        df_fos.sort_values(by="Prioridad")
        .drop_duplicates(subset=["FO"], keep="first")
        .drop(columns=["Prioridad"])
    )

  totales["integradas"] = len(df_fos)
  totales["vencidas"] = (
      len(df_fos[df_fos["Estado de la ruta"] == "Vencida"])
      if not df_fos.empty
      else 0
  )
  totales["vigentes"] = (
      len(df_fos[df_fos["Estado de la ruta"] == "Vigente"])
      if not df_fos.empty
      else 0
  )
  totales["con_retraso"] = (
      len(df_fos[df_fos["Estado de la ruta"] == "Finalizada con Retraso"])
      if not df_fos.empty
      else 0
  )
  totales["rechazadas"] = (
      len(df_fos[df_fos["Estado de la ruta"] == "Rechazada"])
      if not df_fos.empty
      else 0
  )
  totales["total_fo"] = totales["integradas"]

  return totales, df_fos, pd.DataFrame()


# --- FUNCIÓN DE CRUCE DE DATOS GENERAL ---
def ejecutar_cruce_completo(file_drivin, file_sae=None, file_smartsheet=None):
  if hasattr(file_drivin, "name") and file_drivin.name.endswith(".pdf"):
    totales_p, df_fos_p, df_rank_p = procesar_drivin_pdf(file_drivin)
  else:
    totales_p, df_fos_p, df_rank_p = procesar_drivin_excel(file_drivin)

  if totales_p["total_fo"] == 0 or totales_p["total_fo"] < len(df_fos_p):
    totales_p["total_fo"] = len(df_fos_p) if not df_fos_p.empty else 500
    totales_p["integradas"] = totales_p["total_fo"]

  df_sae_p = (
      procesar_excel_sae(file_sae)
      if file_sae is not None
      else pd.DataFrame()
  )
  df_sm_p = (
      procesar_smartsheet(file_smartsheet)
      if file_smartsheet is not None
      else pd.DataFrame()
  )

  df_cruce_p = pd.DataFrame()
  if not df_fos_p.empty:
    df_cruce_p = df_fos_p.copy()

    if not df_sae_p.empty:
      df_cruce_p = pd.merge(df_cruce_p, df_sae_p, on="FO", how="left")
    else:
      (
          df_cruce_p["SUCURSAL"],
          df_cruce_p["ORIGEN"],
          df_cruce_p["DESTINO"],
          df_cruce_p["FOLIO"],
      ) = (None, None, None, None)

    if not df_sm_p.empty:
      df_cruce_p = pd.merge(df_cruce_p, df_sm_p, on="FO", how="left")
      if "ORIGEN_SM" in df_cruce_p.columns:
        df_cruce_p["ORIGEN"] = df_cruce_p["ORIGEN"].fillna(
            df_cruce_p["ORIGEN_SM"]
        )
        df_cruce_p["DESTINO"] = df_cruce_p["DESTINO"].fillna(
            df_cruce_p["DESTINO_SM"]
        )
        df_cruce_p = df_cruce_p.drop(
            columns=["ORIGEN_SM", "DESTINO_SM"]
        )

    df_cruce_p["SUCURSAL"] = df_cruce_p["SUCURSAL"].fillna(
        "PENDIENTE EN GO TRUCK"
    )
    df_cruce_p["ORIGEN"] = df_cruce_p["ORIGEN"].fillna(
        "PENDIENTE EN GO TRUCK"
    )
    df_cruce_p["DESTINO"] = df_cruce_p["DESTINO"].fillna(
        "PENDIENTE EN GO TRUCK"
    )
    df_cruce_p["FOLIO"] = df_cruce_p["FOLIO"].apply(limpiar_fo_str)

    # Integración de reportes
    df_reps_p = cargar_reportes_fo()
    if not df_reps_p.empty:
      df_reps_grouped = (
          df_reps_p.groupby("FO")
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
          df_reps_grouped[["FO", "Estado_Reporte", "Detalle_Reporte"]],
          on="FO",
          how="left",
      )
      df_cruce_p["Estado_Reporte"] = df_cruce_p["Estado_Reporte"].fillna(
          "Sin Reporte"
      )
      df_cruce_p["Detalle_Reporte"] = df_cruce_p[
          "Detalle_Reporte"
      ].fillna("-")
    else:
      df_cruce_p["Estado_Reporte"] = "Sin Reporte"
      df_cruce_p["Detalle_Reporte"] = "-"

    cols_order = [
        "FO",
        "Fin plan",
        "Fin real",
        "Desviación",
        "Conductor",
        "Estado_Reporte",
        "Detalle_Reporte",
        "SUCURSAL",
        "RESPONSABLE",
        "Inicio Planificado",
        "Fin Planificado",
        "Estado de la ruta",
        "Evento",
        "ORIGEN",
        "DESTINO",
        "FOLIO",
    ]
    cols_exist = [c for c in cols_order if c in df_cruce_p.columns]
    df_cruce_p = df_cruce_p[cols_exist]

  return totales_p, df_cruce_p, df_rank_p


# --- INICIALIZACIÓN AUTOMÁTICA DESDE ARCHIVOS BASE EN GITHUB O CACHÉ ---
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


# --- VENTANA MODAL FLOTANTE (CONFIRMACIÓN DE GUARDADO) ---
@st.dialog("💾 Guardar Información en Histórico")
def modal_guardar_historico():
  st.write("¿Deseas guardar los datos cargados en el histórico?")
  col_m_si, col_m_no = st.columns(2)

  if "mostrar_confirmar_fecha" not in st.session_state:
    st.session_state.mostrar_confirmar_fecha = False

  with col_m_si:
    if st.button("✅ Sí, Deseo Guardar", use_container_width=True, type="primary"):
      st.session_state.mostrar_confirmar_fecha = True

  with col_m_no:
    if st.button("❌ No", use_container_width=True):
      st.session_state.mostrar_modal_guardar = False
      st.session_state.mostrar_confirmar_fecha = False
      st.rerun()

  if st.session_state.mostrar_confirmar_fecha:
    st.divider()
    fecha_modal = st.date_input(
        "📅 Selecciona el día a registrar:", value=date.today()
    )

    totales_act = st.session_state.get("totales", {})
    df_act = st.session_state.get("df_cruce", pd.DataFrame())

    v_tot = totales_act.get("total_fo", 0)
    v_int = totales_act.get("integradas", 0)
    v_at = totales_act.get("a_tiempo", 0)
    v_ret = totales_act.get("con_retraso", 0)
    v_vig = totales_act.get("vigentes", 0)
    v_ven = totales_act.get("vencidas", 0)
    v_rec = totales_act.get("rechazadas", 0)

    v_rep = 0
    v_evt = 0
    v_ret_rep = 0

    if not df_act.empty:
      v_rep = len(df_act[df_act["Estado_Reporte"] == "🚨 REPORTADA"]) if "Estado_Reporte" in df_act.columns else 0
      v_evt = len(df_act[df_act["Evento"].notna() & (df_act["Evento"].astype(str).str.strip() != "-") & (df_act["Evento"].astype(str).str.strip() != "")]) if "Evento" in df_act.columns else 0
      if "Estado de la ruta" in df_act.columns and "Estado_Reporte" in df_act.columns:
        v_ret_rep = len(df_act[(df_act["Estado de la ruta"] == "Finalizada con Retraso") & (df_act["Estado_Reporte"] == "🚨 REPORTADA")])

    evaluables = v_at + v_ret + v_ven
    p_at = (v_at / evaluables * 100) if evaluables > 0 else 0.0
    p_v = (v_ven / v_int * 100) if v_int > 0 else 0.0

    reg = {
        "Fecha_Registro": str(fecha_modal),
        "FO_Totales": v_tot,
        "Integradas": v_int,
        "A_Tiempo": v_at,
        "Con_Retraso": v_ret,
        "Vigentes": v_vig,
        "Vencidas": v_ven,
        "Rechazadas": v_rec,
        "FO_Reportadas": v_rep,
        "FO_Con_Evento": v_evt,
        "Retrasos_Reportados": v_ret_rep,
        "Pct_A_Tiempo": round(p_at, 2),
        "Pct_Vencidas": round(p_v, 2),
    }

    if st.button("💾 Guardar Día", use_container_width=True, type="primary"):
      if guardar_nuevo_registro(reg):
        st.session_state.mostrar_modal_guardar = False
        st.session_state.mostrar_confirmar_fecha = False
        st.rerun()
      else:
        st.warning(f"⚠️ El día {fecha_modal} ya existe en el histórico.")


# --- BARRA LATERAL (ADMINISTRADOR) ---
with st.sidebar:
  st.title(" Panel Drivin")
  st.caption("Operación Komatsu Cummins - SAE")

  fecha_act_display = st.session_state.get(
      "ultima_fecha_act", ultima_fecha_act
  )
  if fecha_act_display:
    st.info(f"🕒 **Última actualización local:**\n{fecha_act_display}")

  if st.button("🔄 Actualizar / Recargar Sistema", use_container_width=True):
    st.cache_data.clear()
    (
        c_totales,
        c_df_cruce,
        c_df_ranking,
        c_fecha_act,
    ) = cargar_cache_global()
    st.session_state.totales = c_totales
    st.session_state.df_cruce = c_df_cruce
    st.session_state.df_ranking = c_df_ranking
    st.session_state.ultima_fecha_act = c_fecha_act
    st.success("✅ Sistema recargado exitosamente.")
    st.rerun()

  st.divider()

  st.subheader("⚙️ Panel Administrador")
  servicios_mensuales_admin = st.number_input(
      "📦 Total Servicios Mensuales Operativos:",
      min_value=1,
      value=500,
      step=10,
      help="Indica el total estimado de servicios del mes para autocompletar en caso de faltar totales.",
  )

  # --- ADMINISTRACIÓN DE PROBLEMAS FO DINÁMICO ---
  st.markdown("##### ⚙️ Configurar Opciones de Problemas FO")
  opciones_actuales_prob = cargar_opciones_problemas()

  nueva_opcion_prob = st.text_input("Agregar nueva opción de problema:", key="txt_nuevo_prob")
  if st.button("➕ Añadir Problema", use_container_width=True):
    txt_clean = nueva_opcion_prob.strip()
    if txt_clean and txt_clean not in opciones_actuales_prob:
      opciones_actuales_prob.append(txt_clean)
      guardar_opciones_problemas(opciones_actuales_prob)
      st.success(f"✅ Añadido: '{txt_clean}'")
      st.rerun()
    elif txt_clean in opciones_actuales_prob:
      st.warning("⚠️ Esta opción ya existe en la lista.")

  opcion_eliminar_prob = st.selectbox(
      "Eliminar opción de problema:",
      options=["-- Seleccionar para eliminar --"] + opciones_actuales_prob,
      key="sb_elim_prob"
  )
  if st.button("🗑️ Eliminar Problema", use_container_width=True):
    if opcion_eliminar_prob != "-- Seleccionar para eliminar --":
      opciones_actuales_prob.remove(opcion_eliminar_prob)
      guardar_opciones_problemas(opciones_actuales_prob)
      st.success(f"🗑️ Eliminado: '{opcion_eliminar_prob}'")
      st.rerun()

  # --- ADMINISTRACIÓN DE COORDINADORES KOMATSU / KCC DINÁMICO ---
  st.divider()
  st.markdown("##### ⚙️ Configurar Coordinadores Komatsu / KCC")
  opciones_actuales_coord = cargar_opciones_coordinadores()

  nuevo_coord = st.text_input("Agregar nuevo coordinador:", key="txt_nuevo_coord")
  if st.button("➕ Añadir Coordinador", use_container_width=True):
    txt_coord_clean = nuevo_coord.strip()
    if txt_coord_clean and txt_coord_clean not in opciones_actuales_coord:
      opciones_actuales_coord.append(txt_coord_clean)
      guardar_opciones_coordinadores(opciones_actuales_coord)
      st.success(f"✅ Añadido: '{txt_coord_clean}'")
      st.rerun()
    elif txt_coord_clean in opciones_actuales_coord:
      st.warning("⚠️ Este coordinador ya existe en la lista.")

  coord_eliminar = st.selectbox(
      "Eliminar coordinador existente:",
      options=["-- Seleccionar para eliminar --"] + opciones_actuales_coord,
      key="sb_elim_coord"
  )
  if st.button("🗑️ Eliminar Coordinador", use_container_width=True):
    if coord_eliminar != "-- Seleccionar para eliminar --":
      opciones_actuales_coord.remove(coord_eliminar)
      guardar_opciones_coordinadores(opciones_actuales_coord)
      st.success(f"🗑️ Eliminado: '{coord_eliminar}'")
      st.rerun()

  st.divider()

  st.subheader("📂 Sobrescribir Planillas (Opcional)")
  formato_drivin = st.radio(
      "Tipo de Archivo Drivin a Cargar:",
      ["📊 Excel Drivin (Nuevo Formato)", "📄 PDF Drivin (Formato Legacy)"],
      index=0,
  )

  uploaded_drivin_excel = None
  uploaded_drivin_pdf = None

  if "Excel" in formato_drivin:
    uploaded_drivin_excel = st.file_uploader(
        "Subir Drivin Excel (.xlsx)", type=["xlsx", "xls"]
    )
  else:
    uploaded_drivin_pdf = st.file_uploader(
        "Subir Drivin PDF (.pdf)", type=["pdf"]
    )

  st.markdown("##### Tablas de Cruce Auxiliares")
  uploaded_excel_sae = st.file_uploader(
      "2. Planilla Excel SAE (EDP)", type=["xlsx", "xls"]
  )
  uploaded_smartsheet = st.file_uploader(
      "3. Planilla SmartSheet", type=["xlsx", "xls"]
  )

  btn_procesar = st.button(
      "🚀 Procesar Carga de Información", type="primary", use_container_width=True
  )

  if btn_procesar:
    f_drivin_sel = uploaded_drivin_excel if "Excel" in formato_drivin else uploaded_drivin_pdf
    if f_drivin_sel is None and not os.path.exists(DRIVIN_BASE_FILE):
      st.error("⚠️ Debes adjuntar el archivo de Drivin.")
    else:
      try:
        f_d = f_drivin_sel if f_drivin_sel is not None else DRIVIN_BASE_FILE
        f_sae = uploaded_excel_sae if uploaded_excel_sae is not None else (SAE_BASE_FILE if os.path.exists(SAE_BASE_FILE) else None)
        f_sm = uploaded_smartsheet if uploaded_smartsheet is not None else (SMARTSHEET_BASE_FILE if os.path.exists(SMARTSHEET_BASE_FILE) else None)

        totales_p, df_cruce_p, df_rank_p = ejecutar_cruce_completo(f_d, f_sae, f_sm)

        if totales_p["total_fo"] == 0 or totales_p["total_fo"] < len(df_cruce_p):
          totales_p["total_fo"] = (
              len(df_cruce_p) if not df_cruce_p.empty else servicios_mensuales_admin
          )
          totales_p["integradas"] = totales_p["total_fo"]

        guardar_cache_global(totales_p, df_cruce_p, df_rank_p)

        st.session_state.totales = totales_p
        st.session_state.df_cruce = df_cruce_p
        st.session_state.df_ranking = df_rank_p
        nueva_fecha_txt = datetime.now().strftime("%d/%m/%Y a las %H:%M hrs")
        st.session_state.ultima_fecha_act = nueva_fecha_txt

        st.session_state.mostrar_modal_guardar = True
        st.rerun()
      except Exception as e:
        st.error(f"Error al procesar la información: {e}")

  st.divider()
  st.subheader("🗓️ 2. Registrar Día en Histórico")
  fecha_manual = st.date_input("Fecha a guardar:", value=date.today())

  totales_actuales = st.session_state.get("totales", None)
  df_actuales = st.session_state.get("df_cruce", pd.DataFrame())

  val_tot = totales_actuales["total_fo"] if totales_actuales else 0
  val_int = totales_actuales["integradas"] if totales_actuales else 0
  val_at = totales_actuales["a_tiempo"] if totales_actuales else 0
  val_ret = totales_actuales["con_retraso"] if totales_actuales else 0
  val_vig = totales_actuales["vigentes"] if totales_actuales else 0
  val_ven = totales_actuales["vencidas"] if totales_actuales else 0
  val_rec = totales_actuales["rechazadas"] if totales_actuales else 0

  val_rep = 0
  val_evt = 0
  val_ret_rep = 0

  if not df_actuales.empty:
    val_rep = len(df_actuales[df_actuales["Estado_Reporte"] == "🚨 REPORTADA"]) if "Estado_Reporte" in df_actuales.columns else 0
    val_evt = len(df_actuales[df_actuales["Evento"].notna() & (df_actuales["Evento"].astype(str).str.strip() != "-") & (df_actuales["Evento"].astype(str).str.strip() != "")]) if "Evento" in df_actuales.columns else 0
    if "Estado de la ruta" in df_actuales.columns and "Estado_Reporte" in df_actuales.columns:
      val_ret_rep = len(df_actuales[(df_actuales["Estado de la ruta"] == "Finalizada con Retraso") & (df_actuales["Estado_Reporte"] == "🚨 REPORTADA")])

  evaluables_tot = val_at + val_ret + val_ven
  pct_at = (val_at / evaluables_tot * 100) if evaluables_tot > 0 else 0.0
  pct_v = (val_ven / val_int * 100) if val_int > 0 else 0.0

  registro_preparado = {
      "Fecha_Registro": str(fecha_manual),
      "FO_Totales": val_tot,
      "Integradas": val_int,
      "A_Tiempo": val_at,
      "Con_Retraso": val_ret,
      "Vigentes": val_vig,
      "Vencidas": val_ven,
      "Rechazadas": val_rec,
      "FO_Reportadas": val_rep,
      "FO_Con_Evento": val_evt,
      "Retrasos_Reportados": val_ret_rep,
      "Pct_A_Tiempo": round(pct_at, 2),
      "Pct_Vencidas": round(pct_v, 2),
  }

  st.subheader("💾 3. Acciones de Histórico")
  col_b1, col_b2 = st.columns(2)
  with col_b1:
    if st.button("💾 Guardar", use_container_width=True):
      guardar_nuevo_registro(registro_preparado)
  with col_b2:
    if st.button("🔄 Actualizar", use_container_width=True):
      actualizar_registro_existente(registro_preparado)

  st.divider()
  st.subheader("🗑️ 4. Eliminar Día del Histórico")

  df_h_elim = cargar_historico()
  if not df_h_elim.empty:
    fechas_elim_list = [str(f) for f in df_h_elim["Fecha_Registro"].tolist()]
    fecha_a_eliminar = st.selectbox(
        "Seleccionar día a eliminar:", options=fechas_elim_list
    )
    clave_eliminar = st.text_input(
        "Ingresa la clave para eliminar:", type="password", key="clave_elim"
    )

    if st.button("🔥 Eliminar Registro", use_container_width=True):
      if eliminar_registro_historico(fecha_a_eliminar, clave_eliminar):
        st.rerun()
  else:
    st.caption("No hay días guardados en el histórico para eliminar.")


# DISPARADOR DEL MODAL
if st.session_state.get("mostrar_modal_guardar", False):
  modal_guardar_historico()


# --- DASHBOARD ---
st.title("🚚 Drivin SAE")

fecha_encabezado = st.session_state.get("ultima_fecha_act", ultima_fecha_act)
if fecha_encabezado:
  st.caption(f"📅 **Información actualizada al:** {fecha_encabezado}")

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
      fecha_consultada = st.selectbox(
          "🔍 Seleccionar Día:", options=fechas_disponibles
      )
    else:
      fecha_consultada = "📅 Carga Actual Local"
  else:
    fecha_consultada = "Mes Completo Acumulado"
    st.write("📊 Consolidando todos los registros guardados del mes.")

df_cruce = st.session_state.get("df_cruce", pd.DataFrame())
df_ranking = st.session_state.get("df_ranking", pd.DataFrame())

# CALCULOS DE KPIs
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
  fila_dia = df_hist[
      df_hist["Fecha_Registro"].astype(str) == fecha_consultada
  ].iloc[-1]
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
  # CARGA ACTUAL LOCAL REAL
  val_tot = st.session_state.totales.get("total_fo", 0) if st.session_state.get("totales") else len(df_cruce)
  val_int = st.session_state.totales.get("integradas", len(df_cruce)) if st.session_state.get("totales") else len(df_cruce)
  val_at = st.session_state.totales.get("a_tiempo", 0) if st.session_state.get("totales") else 0
  val_ret = len(df_cruce[df_cruce["Estado de la ruta"] == "Finalizada con Retraso"]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
  val_vig = len(df_cruce[df_cruce["Estado de la ruta"] == "Vigente"]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
  val_ven = len(df_cruce[df_cruce["Estado de la ruta"] == "Vencida"]) if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns else 0
  val_rep = len(df_cruce[df_cruce["Estado_Reporte"] == "🚨 REPORTADA"]) if not df_cruce.empty and "Estado_Reporte" in df_cruce.columns else 0
  val_evt = len(df_cruce[df_cruce["Evento"].notna() & (df_cruce["Evento"].astype(str).str.strip() != "-") & (df_cruce["Evento"].astype(str).str.strip() != "")]) if not df_cruce.empty and "Evento" in df_cruce.columns else 0
  
  if not df_cruce.empty and "Estado de la ruta" in df_cruce.columns and "Estado_Reporte" in df_cruce.columns:
    val_ret_rep = len(df_cruce[(df_cruce["Estado de la ruta"] == "Finalizada con Retraso") & (df_cruce["Estado_Reporte"] == "🚨 REPORTADA")])
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
    st.markdown(
        f'<div class="kpi-card kpi-info"><div class="kpi-title">FO'
        f' Integradas</div><div class="kpi-value">{val_int}</div><div'
        ' class="kpi-sub" style="color: #3498db;">Total evaluables:'
        f' {val_tot}</div></div>',
        unsafe_allow_html=True,
    )
  with c2:
    at_color = "#2ecc71" if pct_at >= 90 else "#e74c3c"
    st.markdown(
        '<div class="kpi-card'
        f' {"kpi-success" if pct_at >= 90 else "kpi-danger"}"><div'
        ' class="kpi-title">Finalizadas a Tiempo</div><div'
        f' class="kpi-value">{pct_at:.1f}%</div><div class="kpi-sub"'
        f' style="color: {at_color};">{val_at} FO | Meta >= 90%</div></div>',
        unsafe_allow_html=True,
    )
  with c3:
    st.markdown(
        '<div class="kpi-card kpi-warning"><div class="kpi-title">Con'
        ' Retraso</div><div'
        f' class="kpi-value">{(val_ret/val_int*100):.1f}%</div><div'
        ' class="kpi-sub" style="color: #e67e22;">'
        f"{val_ret} FO entregadas tarde</div></div>",
        unsafe_allow_html=True,
    )
  with c4:
    v_color = "#2ecc71" if pct_v <= 5 else "#e74c3c"
    st.markdown(
        '<div class="kpi-card'
        f' {"kpi-success" if pct_v <= 5 else "kpi-danger"}"><div'
        ' class="kpi-title">FO Vencidas</div><div'
        f' class="kpi-value">{pct_v:.1f}%</div><div class="kpi-sub"'
        f' style="color: {v_color};">{val_ven} FO | Límite <= 5%</div></div>',
        unsafe_allow_html=True,
    )
  with c5:
    st.markdown(
        '<div class="kpi-card kpi-info"><div class="kpi-title">FO'
        f' Vigentes</div><div class="kpi-value">{val_vig}</div><div'
        ' class="kpi-sub" style="color: #3498db;">En proceso dentro del'
        " plazo</div></div>",
        unsafe_allow_html=True,
    )

  # SEGUNDA FILA DE KPIs: REPORTES Y EVENTOS
  st.markdown("##### 🚨 Análisis de Reportes de Inconsistencias y Eventos")
  cr1, cr2, cr3 = st.columns(3)
  with cr1:
    st.markdown(
        '<div class="kpi-card kpi-purple"><div class="kpi-title">FO'
        ' Reportadas con Problema</div><div'
        f' class="kpi-value">{pct_rep_total:.1f}%</div><div'
        ' class="kpi-sub" style="color: #8e44ad;">'
        f"{val_rep} de {val_int} FO integradas</div></div>",
        unsafe_allow_html=True,
    )
  with cr2:
    st.markdown(
        '<div class="kpi-card kpi-warning"><div class="kpi-title">Retrasos'
        ' Reportados / Justificados</div><div'
        f' class="kpi-value">{pct_ret_justificados:.1f}%</div><div'
        ' class="kpi-sub" style="color: #d35400;">'
        f"{val_ret_rep} de {val_ret} retrasos tienen reporte</div></div>",
        unsafe_allow_html=True,
    )
  with cr3:
    st.markdown(
        '<div class="kpi-card kpi-info"><div class="kpi-title">FO con Evento'
        ' Registrado</div><div'
        f' class="kpi-value">{pct_evt_total:.1f}%</div><div'
        ' class="kpi-sub" style="color: #2980b9;">'
        f"{val_evt} de {val_int} FO con evento anotado</div></div>",
        unsafe_allow_html=True,
    )

  st.markdown("<br>", unsafe_allow_html=True)

# PESTAÑAS
(
    tab_vencidas_vigentes,
    tab_rango_fechas,
    tab_metas_graficos,
    tab_ranking,
    tab_sucursales,
    tab_retraso,
    tab_reportar,
    tab_historico,
) = st.tabs([
    "📋 Vencidas y Vigentes",
    "📅 Simulación MES actual",
    "📊 Rendimiento KPIS Komatsu",
    "🏆 Ranking Conductores",
    "🏢 Rendimiento por Sucursales",
    "⏱️ Finalizadas con Retraso",
    "🚨 Reportar FO / Inconsistencias",
    "📈 Evolución Histórica",
])

with tab_vencidas_vigentes:
  if not df_cruce.empty:
    st.subheader("📋 Tabla Unificada de Vencidas y Vigentes")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      est_opts = df_cruce["Estado de la ruta"].unique()
      estado_sel = st.multiselect(
          "Filtrar Estado de la Ruta:", options=est_opts, default=est_opts
      )
    with col_f2:
      suc_opts = df_cruce["SUCURSAL"].unique()
      sucursal_sel = st.multiselect(
          "Filtrar Sucursal:", options=suc_opts, default=suc_opts
      )

    df_vv_filtrado = df_cruce[df_cruce["Estado de la ruta"].isin(estado_sel)]
    if sucursal_sel:
      df_vv_filtrado = df_vv_filtrado[
          df_vv_filtrado["SUCURSAL"].isin(sucursal_sel)
      ]

    st.dataframe(
        df_vv_filtrado.style.map(
            color_filas_estado, subset=["Estado de la ruta"]
        ),
        use_container_width=True,
        height="auto",
    )

    dict_met_export = {
        "Fecha_Consulta": fecha_consultada,
        "Total_FO": val_tot,
        "FO_Integradas": val_int,
        "A_Tiempo": val_at,
        "Con_Retraso": val_ret,
        "Vigentes": val_vig,
        "Vencidas": val_ven,
        "FO_Reportadas": val_rep,
        "FO_Con_Evento": val_evt,
        "Retrasos_Reportados": val_ret_rep,
        "Pct_A_Tiempo": pct_at,
        "Pct_Vencidas": pct_v,
    }
    bytes_excel = generar_excel_descarga(
        df_vv_filtrado, dict_met_export, df_ranking
    )
    st.download_button(
        label="📥 Descargar Excel Vencidas y Vigentes",
        data=bytes_excel,
        file_name=f"Vencidas_y_Vigentes_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  else:
    st.info(
        "👈 Carga y procesa archivos desde el menú lateral para ver la tabla."
    )

with tab_rango_fechas:
  st.subheader(
      "📅 Filtrado de FO por Rango de Fechas (Basado en Servicios Mensuales)"
  )

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
        f_desde = st.date_input(
            "📅 Fecha Desde:", value=min_f_plan, key="f_desde_plan"
        )
      with col_rf2:
        f_hasta = st.date_input(
            "📅 Fecha Hasta:", value=max_f_plan, key="f_hasta_plan"
        )

      if f_desde > f_hasta:
        st.error(
            "⚠️ La 'Fecha Desde' no puede ser posterior a la 'Fecha Hasta'."
        )
      else:
        dias_rango = (f_hasta - f_desde).days + 1
        servicios_proporcionales_rango = max(
            1, int((servicios_mensuales_admin / 30) * dias_rango)
        )

        df_filtrado_desviaciones = df_rf[
            (df_rf["Fecha_Final_Eval"] >= f_desde)
            & (df_rf["Fecha_Final_Eval"] <= f_hasta)
        ]

        ret_r = len(
            df_filtrado_desviaciones[
                df_filtrado_desviaciones["Estado de la ruta"]
                == "Finalizada con Retraso"
            ]
        )
        ven_r = len(
            df_filtrado_desviaciones[
                df_filtrado_desviaciones["Estado de la ruta"] == "Vencida"
            ]
        )
        vig_r = len(
            df_filtrado_desviaciones[
                df_filtrado_desviaciones["Estado de la ruta"] == "Vigente"
            ]
        )
        rec_r = len(
            df_filtrado_desviaciones[
                df_filtrado_desviaciones["Estado de la ruta"] == "Rechazada"
            ]
        )

        total_r = max(
            servicios_proporcionales_rango,
            (ret_r + ven_r + vig_r + rec_r),
        )
        at_r = max(0, total_r - (ret_r + ven_r + vig_r + rec_r))

        pct_at_r = (at_r / total_r * 100) if total_r > 0 else 0.0
        pct_v_r = (ven_r / total_r * 100) if total_r > 0 else 0.0

        st.markdown(
            f"##### 📊 Resultados para los **{dias_rango} días del periodo**"
            f" (**{f_desde.strftime('%d/%m/%Y')}** al"
            f" **{f_hasta.strftime('%d/%m/%Y')}**)"
        )
        st.caption(
            f"ℹ️ Calculado en base a los **{servicios_mensuales_admin}**"
            " servicios mensuales configurados en el Panel Administrador (~"
            f" {servicios_proporcionales_rango} FO en este rango)."
        )

        rk1, rk2, rk3, rk4, rk5 = st.columns(5)
        with rk1:
          st.markdown(
              '<div class="kpi-card kpi-info"><div class="kpi-title">Total'
              f' Estimado</div><div class="kpi-value">{total_r}</div><div'
              ' class="kpi-sub" style="color: #3498db;">Servicios'
              ' Periodo</div></div>',
              unsafe_allow_html=True,
          )
        with rk2:
          color_at_r = "#2ecc71" if pct_at_r >= 90 else "#e74c3c"
          st.markdown(
              '<div class="kpi-card'
              f' {"kpi-success" if pct_at_r >= 90 else "kpi-danger"}"><div'
              ' class="kpi-title">% On-Time Est.</div><div'
              f' class="kpi-value">{pct_at_r:.1f}%</div><div class="kpi-sub"'
              f' style="color: {color_at_r};">{at_r} FO A Tiempo</div></div>',
              unsafe_allow_html=True,
          )
        with rk3:
          st.markdown(
              '<div class="kpi-card kpi-warning"><div class="kpi-title">Con'
              ' Retraso</div><div'
              f' class="kpi-value">{(ret_r/total_r*100):.1f}%</div><div'
              ' class="kpi-sub" style="color: #e67e22;">'
              f"{ret_r} FO Tarde</div></div>",
              unsafe_allow_html=True,
          )
        with rk4:
          color_v_r = "#2ecc71" if pct_v_r <= 5 else "#e74c3c"
          st.markdown(
              '<div class="kpi-card'
              f' {"kpi-success" if pct_v_r <= 5 else "kpi-danger"}"><div'
              ' class="kpi-title">% Vencidas</div><div'
              f' class="kpi-value">{pct_v_r:.1f}%</div><div class="kpi-sub"'
              f' style="color: {color_v_r};">{ven_r} FO Vencidas</div></div>',
              unsafe_allow_html=True,
          )
        with rk5:
          st.markdown(
              '<div class="kpi-card kpi-info"><div class="kpi-title">FO'
              f' Vigentes</div><div class="kpi-value">{vig_r}</div><div'
              ' class="kpi-sub" style="color: #3498db;">En Proceso</div></div>',
              unsafe_allow_html=True,
          )

        st.divider()
        gr1, gr2 = st.columns(2)
        with gr1:
          df_bar_plan = (
              df_filtrado_desviaciones.groupby(
                  ["Fecha_Final_Eval", "Estado de la ruta"]
              )
              .size()
              .reset_index(name="Cantidad")
          )
          fig_r_bar = px.bar(
              df_bar_plan,
              x="Fecha_Final_Eval",
              y="Cantidad",
              color="Estado de la ruta",
              title="Desviaciones por Fecha Planificada",
              barmode="stack",
              color_discrete_map={
                  "Finalizada con Retraso": "#f39c12",
                  "Vencida": "#e74c3c",
                  "Vigente": "#3498db",
              },
          )
          st.plotly_chart(fig_r_bar, use_container_width=True)

        with gr2:
          df_pie_data = pd.DataFrame({
              "Estado": ["A Tiempo (Est.)", "Con Retraso", "Vencida", "Vigente"],
              "Cantidad": [at_r, ret_r, ven_r, vig_r],
          })
          fig_r_pie = px.pie(
              df_pie_data[df_pie_data["Cantidad"] > 0],
              names="Estado",
              values="Cantidad",
              title="Distribución Proporcional en el Rango",
              hole=0.4,
              color_discrete_map={
                  "A Tiempo (Est.)": "#2ecc71",
                  "Con Retraso": "#f39c12",
                  "Vencida": "#e74c3c",
                  "Vigente": "#3498db",
              },
          )
          st.plotly_chart(fig_r_pie, use_container_width=True)

        st.markdown("##### 📋 Listado de Desviaciones Registradas en el Rango")
        cols_mostrar = [
            c
            for c in [
                "FO",
                "Fin plan",
                "Fin real",
                "Desviación",
                "Conductor",
                "RESPONSABLE",
                "Estado de la ruta",
                "SUCURSAL",
            ]
            if c in df_filtrado_desviaciones.columns
        ]
        st.dataframe(
            df_filtrado_desviaciones[cols_mostrar].style.map(
                color_filas_estado, subset=["Estado de la ruta"]
            ),
            use_container_width=True,
            height="auto",
        )
    else:
      st.warning(
          "No se pudieron interpretar fechas válidas en las FO del archivo"
          " cargado."
      )
  else:
    st.info(
        "👈 Carga y procesa un archivo Drivin para filtrar FO por su fecha"
        " planificada."
    )

with tab_metas_graficos:
  st.subheader(
      "📊 Evaluación de Cumplimiento de Metas Operativas (95% y 90%)"
  )
  if not df_cruce.empty:
    cf1, cf2 = st.columns(2)
    with cf1:
      suc_disponibles = sorted(df_cruce["SUCURSAL"].unique().tolist())
      suc_seleccionadas = st.multiselect(
          "Filtrar por Sucursal(es):",
          options=suc_disponibles,
          default=suc_disponibles,
          key="filter_suc_metas",
      )
    with cf2:
      resp_disponibles = sorted(df_cruce["RESPONSABLE"].unique().tolist())
      resp_seleccionados = st.multiselect(
          "Filtrar por Responsable KCC:",
          options=resp_disponibles,
          default=resp_disponibles,
          key="filter_resp_metas",
      )

    df_m_filtrado = df_cruce[
        (df_cruce["SUCURSAL"].isin(suc_seleccionadas))
        & (df_cruce["RESPONSABLE"].isin(resp_seleccionados))
    ]

    tot_f = len(df_m_filtrado)
    if tot_f > 0:
      cnt_at = len(
          df_m_filtrado[
              df_m_filtrado["Estado de la ruta"].isin(
                  ["Finalizada a tiempo", "A Tiempo"]
              )
          ]
      )
      cnt_ret = len(
          df_m_filtrado[
              df_m_filtrado["Estado de la ruta"] == "Finalizada con Retraso"
          ]
      )
      cnt_ven = len(
          df_m_filtrado[df_m_filtrado["Estado de la ruta"] == "Vencida"]
      )
      cnt_vig = len(
          df_m_filtrado[df_m_filtrado["Estado de la ruta"] == "Vigente"]
      )

      if len(suc_seleccionadas) == len(suc_disponibles) and len(
          resp_seleccionados
      ) == len(resp_disponibles):
        calc_tot = val_tot
        calc_at = val_at
        calc_ret = val_ret
        calc_ven = val_ven
        calc_vig = val_vig
      else:
        calc_tot = tot_f
        calc_at = cnt_at
        calc_ret = cnt_ret
        calc_ven = cnt_ven
        calc_vig = cnt_vig

      calc_fin = calc_at + calc_ret
      pct_fin_tot = (calc_fin / calc_tot * 100) if calc_tot > 0 else 0.0
      pct_a_tiempo_tot = (calc_at / calc_tot * 100) if calc_tot > 0 else 0.0

      g1, g2 = st.columns(2)
      with g1:
        st.markdown("#### 1. Meta 95% FO Finalizadas Totales")
        fig_g95 = go.Figure(
            go.Indicator(
                mode="gauge+number+delta",
                value=pct_fin_tot,
                domain={"x": [0, 1], "y": [0, 1]},
                title={"text": "% Finalizadas (A tiempo + Retraso)"},
                delta={
                    "reference": 95.0,
                    "increasing": {"color": "green"},
                    "decreasing": {"color": "red"},
                },
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {"color": "#1E88E5"},
                    "steps": [
                        {"range": [0, 85], "color": "#ffebee"},
                        {"range": [85, 95], "color": "#fff3e0"},
                        {"range": [95, 100], "color": "#e8f5e9"},
                    ],
                    "threshold": {
                        "line": {"color": "red", "width": 4},
                        "thickness": 0.75,
                        "value": 95.0,
                    },
                },
            )
        )
        fig_g95.update_layout(height=320)
        st.plotly_chart(fig_g95, use_container_width=True)

      with g2:
        st.markdown("#### 2. Meta 90% FO Finalizadas a Tiempo")
        fig_g90 = go.Figure(
            go.Indicator(
                mode="gauge+number+delta",
                value=pct_a_tiempo_tot,
                domain={"x": [0, 1]},
                title={"text": "% Finalizadas A Tiempo (On-Time)"},
                delta={
                    "reference": 90.0,
                    "increasing": {"color": "green"},
                    "decreasing": {"color": "red"},
                },
                gauge={
                    "axis": {"range": [0, 100]},
                    "bar": {
                        "color": "#2ecc71"
                        if pct_a_tiempo_tot >= 90
                        else "#e74c3c"
                    },
                    "steps": [
                        {"range": [0, 75], "color": "#ffebee"},
                        {"range": [75, 90], "color": "#fff3e0"},
                        {"range": [90, 100], "color": "#e8f5e9"},
                    ],
                    "threshold": {
                        "line": {"color": "red", "width": 4},
                        "thickness": 0.75,
                        "value": 90.0,
                    },
                },
            )
        )
        fig_g90.update_layout(height=320)
        st.plotly_chart(fig_g90, use_container_width=True)

      st.divider()
      c_pie1, c_pie2 = st.columns(2)
      with c_pie1:
        df_pie_95 = pd.DataFrame({
            "Estado": ["A Tiempo", "Con Retraso", "Vigentes", "Vencidas"],
            "Cantidad": [calc_at, calc_ret, calc_vig, calc_ven],
        })
        fig_pie1 = px.pie(
            df_pie_95,
            names="Estado",
            values="Cantidad",
            color="Estado",
            color_discrete_map={
                "A Tiempo": "#2ecc71",
                "Con Retraso": "#f39c12",
                "Vigentes": "#3498db",
                "Vencidas": "#e74c3c",
            },
            hole=0.4,
        )
        st.plotly_chart(fig_pie1, use_container_width=True)
      with c_pie2:
        df_pie_90 = pd.DataFrame({
            "Condición": ["A Tiempo", "Con Retraso"],
            "Cantidad": [calc_at, calc_ret],
        })
        fig_pie2 = px.pie(
            df_pie_90,
            names="Condición",
            values="Cantidad",
            color="Condición",
            color_discrete_map={"A Tiempo": "#2ecc71", "Con Retraso": "#e67e22"},
            hole=0.4,
        )
        st.plotly_chart(fig_pie2, use_container_width=True)

with tab_ranking:
  st.subheader("🏆 Ranking de Cumplimiento de Conductores")
  if df_ranking is not None and not df_ranking.empty:
    st.dataframe(df_ranking, use_container_width=True, height="auto")
  else:
    st.info("Sin datos de ranking actualmente.")

with tab_sucursales:
  if not df_cruce.empty:
    st.subheader("🏢 Rendimiento por Sucursales")
    fig_suc = px.bar(
        df_cruce,
        x="SUCURSAL",
        color="Estado de la ruta",
        barmode="group",
        text_auto=True,
    )
    st.plotly_chart(fig_suc, use_container_width=True)

# --- PESTAÑA: FINALIZADAS CON RETRASO ---
with tab_retraso:
  col_tit_ret, col_btn_act = st.columns([3, 1])
  with col_tit_ret:
    st.subheader("⏱️ FO Finalizadas con Retraso")
  with col_btn_act:
    if st.button(
        "🔄 Actualizar Planilla",
        key="btn_act_retrasos",
        use_container_width=True,
        type="primary",
    ):
      st.cache_data.clear()
      c_tot, c_cruce, c_rank, c_fecha = cargar_cache_global()

      if not c_cruce.empty:
        df_reps_sync = cargar_reportes_fo()
        if not df_reps_sync.empty:
          df_reps_sync["FO"] = df_reps_sync["FO"].apply(limpiar_fo_str)
          df_reps_grouped_s = (
              df_reps_sync.groupby("FO")
              .agg({
                  "Tipo_Problema": "last",
                  "Coordinador_Komatsu": "last",
                  "Responsable": "last",
                  "Detalle_Observacion": "last",
              })
              .reset_index()
          )
          df_reps_grouped_s["Estado_Reporte"] = "🚨 REPORTADA"
          df_reps_grouped_s["Detalle_Reporte"] = (
              "Motivo: "
              + df_reps_grouped_s["Tipo_Problema"].astype(str)
              + " | Coord. Komatsu: "
              + df_reps_grouped_s["Coordinador_Komatsu"].astype(str)
              + " | Resp: "
              + df_reps_grouped_s["Responsable"].astype(str)
              + " | Obs: "
              + df_reps_grouped_s["Detalle_Observacion"].astype(str)
          )

          for col_drop in ["Estado_Reporte", "Detalle_Reporte"]:
            if col_drop in c_cruce.columns:
              c_cruce = c_cruce.drop(columns=[col_drop])

          c_cruce = pd.merge(
              c_cruce,
              df_reps_grouped_s[["FO", "Estado_Reporte", "Detalle_Reporte"]],
              on="FO",
              how="left",
          )
          c_cruce["Estado_Reporte"] = c_cruce["Estado_Reporte"].fillna(
              "Sin Reporte"
          )
          c_cruce["Detalle_Reporte"] = c_cruce["Detalle_Reporte"].fillna("-")

      guardar_cache_global(c_tot, c_cruce, c_rank)
      st.session_state.totales = c_tot
      st.session_state.df_cruce = c_cruce
      st.session_state.df_ranking = c_rank
      st.session_state.ultima_fecha_act = (
          datetime.now().strftime("%d/%m/%Y a las %H:%M hrs")
      )
      st.success("✅ Planilla de retrasos y reportes actualizada.")
      st.rerun()

  if not df_cruce.empty:
    df_solo_retraso = df_cruce[
        df_cruce["Estado de la ruta"] == "Finalizada con Retraso"
    ].copy()

    if not df_solo_retraso.empty:
      df_solo_retraso["FO"] = df_solo_retraso["FO"].apply(limpiar_fo_str)

      df_reps_info = cargar_reportes_fo()
      if not df_reps_info.empty:
        df_reps_info["FO"] = df_reps_info["FO"].apply(limpiar_fo_str)

        df_reps_grouped = (
            df_reps_info.groupby("FO")
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

        for c_drop in ["Estado_Reporte", "Detalle_Reporte"]:
          if c_drop in df_solo_retraso.columns:
            df_solo_retraso = df_solo_retraso.drop(columns=[c_drop])

        df_solo_retraso = pd.merge(
            df_solo_retraso,
            df_reps_grouped[["FO", "Estado_Reporte", "Detalle_Reporte"]],
            on="FO",
            how="left",
        )
        df_solo_retraso["Estado_Reporte"] = df_solo_retraso[
            "Estado_Reporte"
        ].fillna("Sin Reporte")
        df_solo_retraso["Detalle_Reporte"] = df_solo_retraso[
            "Detalle_Reporte"
        ].fillna("-")
      else:
        df_solo_retraso["Estado_Reporte"] = "Sin Reporte"
        df_solo_retraso["Detalle_Reporte"] = "-"

      sucursales_retraso = sorted(
          df_solo_retraso["SUCURSAL"].dropna().unique().tolist()
      )

      sucursales_sel_ret = st.multiselect(
          "🔍 Filtrar por Sucursal(es):",
          options=sucursales_retraso,
          default=sucursales_retraso,
          key="ms_sucursales_retraso_fijo",
          help="Busca o selecciona sucursales para filtrar la vista.",
      )

      if sucursales_sel_ret:
        df_solo_retraso = df_solo_retraso[
            df_solo_retraso["SUCURSAL"].isin(sucursales_sel_ret)
        ]

      if not df_solo_retraso.empty:
        df_mostrar_retraso = df_solo_retraso.copy()

        if "Fin plan" not in df_mostrar_retraso.columns:
          if "Fin Planificado" in df_mostrar_retraso.columns:
            df_mostrar_retraso["Fin plan"] = df_mostrar_retraso[
                "Fin Planificado"
            ]
          else:
            df_mostrar_retraso["Fin plan"] = "-"

        def resolver_fin_real_y_desviacion(row):
          f_plan = row.get("Fin plan") or row.get("Fin Planificado") or "-"
          f_real = row.get("Fin real") or row.get("Fin Real") or "-"
          desv = row.get("Desviación") or "-"

          f_real_str = str(f_real).strip()
          desv_str = str(desv).strip()

          if f_real_str in ["", "-", "nan", "None"] and desv_str not in [
              "",
              "-",
              "nan",
              "None",
          ]:
            f_real_str = calcular_fin_real_automatico(f_plan, desv_str)

          if desv_str in ["", "-", "nan", "None"] and f_real_str not in [
              "",
              "-",
              "nan",
              "None",
          ]:
            desv_str = calcular_desviacion_automatica(f_plan, f_real_str)

          return pd.Series([f_real_str, desv_str])

        df_mostrar_retraso[["Fin real", "Desviación"]] = (
            df_mostrar_retraso.apply(resolver_fin_real_y_desviacion, axis=1)
        )

        df_mostrar_retraso["Desfase_Minutos"] = df_mostrar_retraso[
            "Desviación"
        ].apply(desviacion_a_minutos)
        minutos_validos = df_mostrar_retraso["Desfase_Minutos"].dropna()

        promedio_desfase_str = "-"
        if len(minutos_validos) > 0:
          prom_minutos = minutos_validos.mean()
          promedio_desfase_str = minutos_a_formato_horas(prom_minutos)

        # RESUMEN DE RETRASOS REPORTADOS VS NO REPORTADOS
        num_ret_totales = len(df_mostrar_retraso)
        num_ret_con_rep = len(df_mostrar_retraso[df_mostrar_retraso["Estado_Reporte"] == "🚨 REPORTADA"])
        pct_ret_con_rep = (num_ret_con_rep / num_ret_totales * 100) if num_ret_totales > 0 else 0.0

        c_ret_kpi1, c_ret_kpi2 = st.columns(2)
        with c_ret_kpi1:
          st.markdown(
              f'<div class="kpi-card kpi-warning"><div class="kpi-title">⏱️'
              ' Promedio General de Desfase</div><div'
              f' class="kpi-value">{promedio_desfase_str}</div><div'
              ' class="kpi-sub" style="color: #e67e22;">Calculado sobre'
              f" {len(minutos_validos)} FO con retraso</div></div>",
              unsafe_allow_html=True,
          )
        with c_ret_kpi2:
          st.markdown(
              f'<div class="kpi-card kpi-purple"><div class="kpi-title">🚨'
              ' Retrasos Justificados / Reportados</div><div'
              f' class="kpi-value">{pct_ret_con_rep:.1f}%</div><div'
              ' class="kpi-sub" style="color: #8e44ad;">'
              f"{num_ret_con_rep} de {num_ret_totales} FO con retraso tienen"
              " reporte</div></div>",
              unsafe_allow_html=True,
          )

        st.markdown("<br>", unsafe_allow_html=True)

        c_ret_g1, c_ret_g2 = st.columns(2)

        with c_ret_g1:
          df_suc_ret = (
              df_mostrar_retraso.groupby("SUCURSAL")
              .size()
              .reset_index(name="Cantidad Retrasos")
              .sort_values(by="Cantidad Retrasos", ascending=False)
          )
          fig_bar_ret = px.bar(
              df_suc_ret,
              x="SUCURSAL",
              y="Cantidad Retrasos",
              text="Cantidad Retrasos",
              title="Cantidad de FO con Retraso por Sucursal",
              color="SUCURSAL",
              color_discrete_sequence=px.colors.qualitative.Set2,
          )
          fig_bar_ret.update_traces(textposition="outside")
          st.plotly_chart(fig_bar_ret, use_container_width=True)

        with c_ret_g2:
          fig_pie_ret = px.pie(
              df_suc_ret,
              names="SUCURSAL",
              values="Cantidad Retrasos",
              title="Distribución % de Retrasos por Sucursal",
              hole=0.4,
              color_discrete_sequence=px.colors.qualitative.Set2,
          )
          st.plotly_chart(fig_pie_ret, use_container_width=True)

        st.divider()
        st.markdown("##### 📋 Listado Detallado de FO Finalizadas con Retraso")

        cols_orden_foto = [
            "FO",
            "Fin plan",
            "Fin real",
            "Desviación",
            "Conductor",
            "Estado_Reporte",
            "Detalle_Reporte",
            "SUCURSAL",
            "RESPONSABLE",
            "ORIGEN",
            "DESTINO",
        ]

        cols_existentes = [
            c for c in cols_orden_foto if c in df_mostrar_retraso.columns
        ]

        st.dataframe(
            df_mostrar_retraso[cols_existentes].style.map(
                color_filas_estado, subset=["Estado_Reporte"]
            ),
            use_container_width=True,
            height="auto",
        )

        excel_bytes_retraso = generar_excel_solo_tabla(
            df_mostrar_retraso[cols_existentes],
            nombre_hoja="Finalizadas_Con_Retraso",
        )
        st.download_button(
            label="📥 Descargar Informe en Excel (Finalizadas con Retraso)",
            data=excel_bytes_retraso,
            file_name=f"Informe_Retrasos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
      else:
        st.warning(
            "⚠️ No existen registros con retraso para las sucursales"
            " seleccionadas."
        )
    else:
      st.success("🎉 No hay retrasos registrados.")
  else:
    st.info("👈 Carga y procesa un archivo Drivin para visualizar los datos.")

# --- PESTAÑA: REPORTAR Y EDITAR FO ---
with tab_reportar:
  st.subheader("🚨 Gestión y Reporte de Inconsistencias en FO")
  col_form, col_hist = st.columns([1, 1])

  with col_form:
    st.markdown("#### 📝 Crear / Editar Reporte de FO")

    df_reps_actuales = cargar_reportes_fo()
    opciones_modo = ["➕ Nuevo Reporte"]
    if not df_reps_actuales.empty:
      opciones_modo.append("✏️ Editar Reporte Existente")

    modo_reporte = st.radio(
        "Acción a realizar:", opciones_modo, horizontal=True
    )

    lista_sucursales = [
        "La Negra",
        "La Chimba",
        "Copiapó",
        "Coquimbo",
        "Santiago",
        "Otra",
    ]

    lista_problemas = cargar_opciones_problemas()
    lista_coordinadores_komatsu = cargar_opciones_coordinadores()

    if modo_reporte == "➕ Nuevo Reporte":
      with st.form(key="form_rep_nuevo", clear_on_submit=True):
        fo_in = st.text_input("Número de FO:")
        suc_in = st.selectbox("Sucursal:", lista_sucursales)
        prob_in = st.selectbox("Problema / Inconsistencia:", lista_problemas)
        coord_sel = st.selectbox("Coordinador Komatsu / KCC:", lista_coordinadores_komatsu)

        resp_in = st.text_input("Reportado por (Usuario SAE):")
        obs_in = st.text_area("Observaciones / Justificación:")
        arch_in = st.file_uploader(
            "Adjuntar Respaldo (Imagen/PDF):",
            type=["png", "jpg", "jpeg", "pdf", "msg", "eml", "txt"],
        )

        if st.form_submit_button("📩 Guardar Nuevo Reporte"):
          if fo_in and resp_in:
            guardar_reporte_fo(
                fo_in,
                suc_in,
                prob_in,
                coord_sel,
                resp_in,
                obs_in,
                arch_in,
            )

            if (
                "df_cruce" in st.session_state
                and not st.session_state.df_cruce.empty
            ):
              guardar_cache_global(
                  st.session_state.totales,
                  st.session_state.df_cruce,
                  st.session_state.df_ranking,
              )

            st.success("✅ Reporte guardado con éxito.")
            st.rerun()
          else:
            st.error("⚠️ Debes completar la FO y la persona que reporta.")

    else:
      fos_reportadas_list = [
          str(f) for f in df_reps_actuales["FO"].unique() if str(f).strip()
      ]
      fo_a_editar = st.selectbox(
          "Seleccionar FO a Editar:", options=fos_reportadas_list
      )

      fo_a_editar_limpia = limpiar_fo_str(fo_a_editar)
      row_edit = df_reps_actuales[
          df_reps_actuales["FO"] == fo_a_editar_limpia
      ].iloc[-1]

      idx_suc = (
          lista_sucursales.index(row_edit["Sucursal"])
          if row_edit["Sucursal"] in lista_sucursales
          else 0
      )
      
      prob_actual = row_edit["Tipo_Problema"]
      idx_prob = (
          lista_problemas.index(prob_actual)
          if prob_actual in lista_problemas
          else 0
      )

      coord_actual = str(row_edit.get("Coordinador_Komatsu", "-"))
      idx_coord = (
          lista_coordinadores_komatsu.index(coord_actual)
          if coord_actual in lista_coordinadores_komatsu
          else 0
      )

      with st.form(key="form_rep_editar"):
        st.info(f"Modificando datos para la **FO {fo_a_editar_limpia}**")
        suc_edit = st.selectbox("Sucursal:", lista_sucursales, index=idx_suc)
        prob_edit = st.selectbox(
            "Problema / Inconsistencia:", lista_problemas, index=idx_prob
        )
        coord_edit_sel = st.selectbox(
            "Coordinador Komatsu / KCC:",
            lista_coordinadores_komatsu,
            index=idx_coord,
        )

        resp_edit = st.text_input(
            "Reportado por (Usuario SAE):",
            value=str(row_edit.get("Responsable", "")),
        )
        obs_edit = st.text_area(
            "Observaciones:",
            value=str(row_edit.get("Detalle_Observacion", "")),
        )
        arch_edit = st.file_uploader(
            "Reemplazar Respaldo (Opcional):",
            type=["png", "jpg", "jpeg", "pdf", "msg", "eml", "txt"],
        )

        c_ed1, c_ed2 = st.columns(2)
        with c_ed1:
          btn_actualizar = st.form_submit_button("🔄 Guardar Cambios")
        with c_ed2:
          btn_eliminar = st.form_submit_button("🔥 Eliminar Reporte")

        if btn_actualizar:
          if editar_reporte_fo(
              fo_a_editar_limpia,
              suc_edit,
              prob_edit,
              coord_edit_sel,
              resp_edit,
              obs_edit,
              arch_edit,
          ):
            st.success("✅ Reporte actualizado con éxito.")
            st.rerun()
          else:
            st.error("❌ Ocurrió un error al intentar editar.")

        if btn_eliminar:
          if eliminar_reporte_fo(fo_a_editar_limpia):
            st.success("🗑️ Reporte eliminado.")
            st.rerun()

  with col_hist:
    st.markdown("#### 📋 Registro de FO Reportadas")
    df_reps = cargar_reportes_fo()
    if not df_reps.empty:
      st.dataframe(
          df_reps[[
              "Fecha_Reporte",
              "FO",
              "Sucursal",
              "Tipo_Problema",
              "Coordinador_Komatsu",
              "Responsable",
              "Detalle_Observacion",
          ]],
          use_container_width=True,
          height="auto",
      )

      excel_bytes_reportes = generar_excel_solo_tabla(
          df_reps, nombre_hoja="FO_Reportadas"
      )
      st.download_button(
          label="📥 Descargar Registro de FO Reportadas en Excel",
          data=excel_bytes_reportes,
          file_name=f"Registro_FO_Reportadas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
          use_container_width=True,
      )

      st.divider()
      st.markdown("##### 🔍 Consultar y Visualizar Respaldo")
      lista_fos_reportadas = [
          str(f) for f in df_reps["FO"].unique() if str(f).strip()
      ]
      fo_sel_ver = st.selectbox(
          "Seleccionar FO para ver la imagen:",
          options=lista_fos_reportadas,
          key="sb_ver_img",
      )

      fo_ver_limpia = limpiar_fo_str(fo_sel_ver)
      row_ver = df_reps[df_reps["FO"] == fo_ver_limpia].iloc[-1]
      st.write(f"**Fecha Reporte:** {row_ver.get('Fecha_Reporte', '-')}")
      st.write(
          "**Coordinador Komatsu / KCC:**"
          f" {row_ver.get('Coordinador_Komatsu', '-')}"
      )
      st.write(f"**Reportado por:** {row_ver.get('Responsable', '-')}")
      st.write(f"**Observación:** {row_ver.get('Detalle_Observacion', '-')}")

      archivo_nombre = str(row_ver.get("Archivo_Pantallazo", "Sin respaldo"))
      if archivo_nombre != "Sin respaldo" and archivo_nombre != "nan":
        path_img_local = os.path.join(CAPTURAS_DIR, archivo_nombre)
        if os.path.exists(path_img_local):
          ext_ver = os.path.splitext(archivo_nombre)[1].lower()
          if ext_ver in [".png", ".jpg", ".jpeg"]:
            st.image(
                path_img_local,
                caption=f"Respaldo FO {fo_ver_limpia}",
                use_container_width=True,
            )
          else:
            with open(path_img_local, "rb") as f_down:
              st.download_button(
                  label=f"📥 Descargar Archivo Respaldo ({ext_ver.upper()})",
                  data=f_down,
                  file_name=archivo_nombre,
              )
        else:
          st.warning(
              "⚠️ No se encontró el archivo de imagen en la carpeta local."
          )
      else:
        st.info("ℹ️ Este reporte no adjuntó ningún archivo de respaldo.")
    else:
      st.caption("Aún no hay reportes de FO registrados.")

# --- PESTAÑA: EVOLUCIÓN HISTÓRICA CON GRÁFICO ---
with tab_historico:
  st.subheader("📈 Evolución Histórica de Cumplimiento")
  df_hist = cargar_historico()
  df_papelera = cargar_papelera()

  if not df_hist.empty:
    df_hist_val = df_hist[df_hist["FO_Totales"] > 0].copy()

    if not df_hist_val.empty:
      fig_hist = go.Figure()

      fig_hist.add_trace(
          go.Scatter(
              x=df_hist_val["Fecha_Registro"],
              y=df_hist_val["Pct_A_Tiempo"],
              mode="lines+markers+text",
              name="% A Tiempo",
              text=[f"{v:.1f}%" for v in df_hist_val["Pct_A_Tiempo"]],
              textposition="top center",
              line=dict(color="#1E88E5", width=3),
              marker=dict(size=8, symbol="circle"),
          )
      )

      fig_hist.add_trace(
          go.Scatter(
              x=df_hist_val["Fecha_Registro"],
              y=df_hist_val["Pct_Vencidas"],
              mode="lines+markers+text",
              name="% Vencidas",
              text=[f"{v:.2f}%" for v in df_hist_val["Pct_Vencidas"]],
              textposition="bottom center",
              line=dict(color="#E53935", width=3),
              marker=dict(size=8, symbol="square"),
          )
      )

      fig_hist.add_shape(
          type="line",
          x0=df_hist_val["Fecha_Registro"].min(),
          x1=df_hist_val["Fecha_Registro"].max(),
          y0=90,
          y1=90,
          line=dict(color="#2ECC71", width=2, dash="dash"),
      )
      fig_hist.add_annotation(
          x=df_hist_val["Fecha_Registro"].max(),
          y=90,
          text="Meta A Tiempo (≥ 90%)",
          showarrow=False,
          yshift=10,
          font=dict(color="#2ECC71", size=12),
      )

      fig_hist.add_shape(
          type="line",
          x0=df_hist_val["Fecha_Registro"].min(),
          x1=df_hist_val["Fecha_Registro"].max(),
          y0=5,
          y1=5,
          line=dict(color="#D32F2F", width=2, dash="dot"),
      )
      fig_hist.add_annotation(
          x=df_hist_val["Fecha_Registro"].max(),
          y=5,
          text="Límite Vencidas (≤ 5%)",
          showarrow=False,
          yshift=10,
          font=dict(color="#D32F2F", size=12),
      )

      fig_hist.update_layout(
          title=(
              "Evolución por Día Registrado: % A Tiempo vs. % Vencidas con Metas"
          ),
          xaxis_title="Fecha de Registro",
          yaxis_title="Porcentaje (%)",
          yaxis=dict(range=[0, 105]),
          legend=dict(
              orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
          ),
          height=450,
          margin=dict(l=40, r=40, t=60, b=40),
      )

      st.plotly_chart(fig_hist, use_container_width=True)

    st.divider()
    st.markdown("##### 📋 Tabla Histórico Activo")
    st.dataframe(df_hist, use_container_width=True, height="auto")
  else:
    st.info("Aún no hay datos guardados en el histórico para graficar.")

  st.divider()
  st.markdown("##### 🗑️ Papelera de Respaldo")
  if not df_papelera.empty:
    st.dataframe(df_papelera, use_container_width=True, height="auto")
  else:
    st.caption("Papelera vacía.")
