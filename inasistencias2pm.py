# -*- coding: utf-8 -*-
"""
Created on Fri Oct  3 15:53:06 2025

@author: carlos.grajales
"""

import requests
import json
import pandas as pd
import io
import os
from datetime import datetime

# ==========================
# CONFIGURACIÓN TABLEAU
# ==========================
TABLEAU_SERVER = "https://prod-useast-b.online.tableau.com"
TOKEN_NAME = "Python"
TOKEN_VALUE = ""
SITE = "superdealimentos40"
WORKBOOK_NAME = "prueba marcaciones"

VIEWS_CONFIG = [
    {"view_name": "Inasistencias"}
]
try:
	os.remove("Inasistencias.xlsx")
except FileNotFoundError:
	print("paila")
#SAVE_PATH = r"\\192.168.110.20\super\UNIDAD PRODUCTIVA\INFORMACIóN GLOBAL\INDICADORES\DESARROLLOS\Alerta_Notificacion_Entrada_Salida_Danger"
#SAVE_PATH = r"C:\Users\sebastian.giraldo\OneDrive - SUPER DE ALIMENTOS S A\Escritorio\Marcaciones"
OUTPUT_FILE = os.path.join("Inasistencias.xlsx")

CORREO = os.path.join("Correos.xlsx")

df_correo=pd.read_excel(CORREO)

# Configuración de reintentos
MAX_RETRIES = 3
RETRY_DELAY = 60  # 60 segundos entre reintentos


# ==========================
# CÓDIGO PRINCIPAL (sin cambios hasta el envío)
# ==========================
print(f"=== INICIO DEL PROCESO - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")

AUTH_URL = f"{TABLEAU_SERVER}/api/3.17/auth/signin"
payload = {
    "credentials": {
        "personalAccessTokenName": TOKEN_NAME,
        "personalAccessTokenSecret": TOKEN_VALUE,
        "site": {"contentUrl": SITE}
    }
}
headers = {"Content-Type": "application/json", "Accept": "application/json"}

print("Autenticando en Tableau...")
response = requests.post(AUTH_URL, data=json.dumps(payload), headers=headers)
if response.status_code != 200:
    print(f"✗ Error en autenticación: {response.text}")
    exit()

auth_data = response.json()
token = auth_data["credentials"]["token"]
site_id = auth_data["credentials"]["site"]["id"]
headers["X-Tableau-Auth"] = token
print("✓ Autenticación exitosa en Tableau.")

# ==========================
# BUSCAR WORKBOOK
# ==========================
print(f"Buscando workbook: '{WORKBOOK_NAME}'...")
WORKBOOKS_URL = f"{TABLEAU_SERVER}/api/3.17/sites/{site_id}/workbooks?filter=name:eq:{WORKBOOK_NAME}"
response = requests.get(WORKBOOKS_URL, headers=headers)
if response.status_code != 200:
    print(f"✗ Error buscando workbook: {response.text}")
    exit()

workbooks = response.json().get("workbooks", {}).get("workbook", [])
if not workbooks:
    print(f"✗ No se encontró el workbook: '{WORKBOOK_NAME}'")
    exit()
workbook_id = workbooks[0]["id"]
print(f"✓ Workbook encontrado. ID: {workbook_id}")

# ==========================
# OBTENER VISTAS
# ==========================
print("Obteniendo vistas del workbook...")
VIEWS_URL = f"{TABLEAU_SERVER}/api/3.17/sites/{site_id}/workbooks/{workbook_id}/views"
response = requests.get(VIEWS_URL, headers=headers)
if response.status_code != 200:
    print(f"✗ Error obteniendo vistas: {response.text}")
    exit()
views = response.json().get("views", {}).get("view", [])
print(f"✓ Se encontraron {len(views)} vistas en el workbook.")

# Crear directorio si no existe
#os.makedirs(SAVE_PATH, exist_ok=True)

# Descargar datos y guardar en Excel
for view_config in VIEWS_CONFIG:
    view_name = view_config["view_name"]
    print(f"\n--- Procesando vista: '{view_name}' ---")

    # Buscar view_id
    view_id = next((v["id"] for v in views if v["name"] == view_name), None)
    if not view_id:
        print(f"✗ No se encontró la vista: '{view_name}'")
        continue

    print(f"✓ Vista encontrada. ID: {view_id}")

    # Descargar datos
    CROSSTAB_CSV_URL = f"{TABLEAU_SERVER}/api/3.17/sites/{site_id}/views/{view_id}/data"
    csv_response = requests.get(CROSSTAB_CSV_URL, headers=headers)

    if csv_response.status_code == 200:
        csv_data = csv_response.content.decode("utf-8")
        df = pd.read_csv(io.StringIO(csv_data))  # fuerza texto desde el inicio
        df = df.drop(columns=["Count of Area"])
        #df = df.astype(str)  # doble seguridad
        #with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        #    df.to_excel(writer, sheet_name=view_name, index=False)

        # =========================================================
        # FILTRADOS

# ---- Caso 2: TURNO3 ----
        filtro_turnos4 = df["Horario_Entrada"] == "Entrada 2:00 P.M."
        filtro_obs4 = df["Observacion2"] == "Inasistencia de colaborador programado"
        df_Diezpm = df[filtro_turnos4 & filtro_obs4]
        # Obtener los valores únicos de la columna PlantaProg
        plantas = df_Diezpm["PlantaProgramada"].unique()
        print(plantas)
        # =========================================================
         # GUARDAR RESULTADOS
        # =========================================================
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            df_Diezpm.to_excel(writer, sheet_name="Entrada_2_00_P.M.", index=False)
            # Crear una hoja por cada planta, Sieteam
            for planta, i in zip(plantas, range(0, len(df_Diezpm["PlantaProgramada"].unique()))):
                # Filtrar las filas para esa planta
                df_planta = df_Diezpm[df_Diezpm["PlantaProgramada"] == planta]
                # Elegir un nombre de hoja (sheet_name). Debe tener máximo 31 caracteres y no usar caracteres inválidos para nombres de hojas de Excel.
                sheet_name = f"{i}_Entrada_2_00_P.M."
                df_planta.to_excel(writer, sheet_name=sheet_name, index=False)
            df_correo.to_excel(writer, sheet_name="Correos_Plantas", index=False)
        print(":marca_de_verificación_blanca: Archivos generados correctamente")
           
    else:
        print("No dió :c")

# Cerrar sesión en Tableau
LOGOUT_URL = f"{TABLEAU_SERVER}/api/3.17/auth/signout"
requests.post(LOGOUT_URL, headers=headers)
print("✓ Sesión cerrada en Tableau.")
print(f"✓ Archivo final generado: {OUTPUT_FILE}")
