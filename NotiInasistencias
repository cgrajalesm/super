# -*- coding: utf-8 -*-
#VERSIÓN UNIFICADA - ENVÍO DE ENLACE A SLACK + EMAIL CON ADJUNTO

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import requests
import json
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import datetime


# Email setup
EMAIL_USER = 'carlos.grajales@ccsuper.co'
EMAIL_PASS = 'SUPcagm25-10*#'
SMTP_SERVER = 'smtp.office365.com'
SMTP_PORT = 587

# Slack setup
SLACK_TOKEN = ""
SLACK_CHANNEL = "#bot_test"

# Excel path
EXCEL_PATH = r"Inasistencias.xlsx"

# Clean up temp file at the beginning
try:
    os.remove("temp_INASISTENCIAS.xlsx")
except FileNotFoundError:
    print("ni modo")

# Read Excel    
try:
    dfs = pd.read_excel(EXCEL_PATH, sheet_name=None)
    print("Excel file read successfully.")
except Exception as e:
    print(f"Error reading Excel: {e}")
    exit(1)

# Guardar con pandas primero (nuevas columnas ya agregadas en df_first)
with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
    for sheet_name, df in dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Ahora reabrimos con openpyxl para aplicar formatos
wb = openpyxl.load_workbook(EXCEL_PATH)

# Guardar nuevamente con los formatos aplicados
wb.save(EXCEL_PATH)


df_correos = dfs['Correos_Plantas']

# Assume columns: Planta, Correo
plantas_emails = df_correos.groupby('Correo')['Planta'].apply(list).to_dict()
print(plantas_emails)

# Invert to plant_emails
plant_emails = {}
for email, plants in plantas_emails.items():
    for plant in plants:
        if plant not in plant_emails:
            plant_emails[plant] = []
        plant_emails[plant].append(email)

print(f"All sheets: {list(dfs.keys())}")
print(f"Plants: {list(plant_emails.keys())}")

# Function to send email
def send_email(to_emails, subject, body, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USER
    msg['To'] = ', '.join(to_emails)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    if attachment_path:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment_path)}')
            msg.attach(part)
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)
        text = msg.as_string()
        server.sendmail(EMAIL_USER, to_emails, text)
        server.quit()
        print(f"Email sent to {to_emails} for {subject}")
    except Exception as e:
        print(f"Error sending email: {e}")

# Send emails for each plant
for planta, emails in plant_emails.items():
    # Find sheets containing the plant name
    relevant_sheets = []
    for name, df in dfs.items():
        if name != 'Correos_Plantas' and 'PlantaProgramada' in df.columns:
            unique_plants = df['PlantaProgramada'].unique()
            if len(unique_plants) == 1 and unique_plants[0] == planta:
                relevant_sheets.append(name)
    print(f"For plant '{planta}', relevant sheets: {relevant_sheets}")
    if not relevant_sheets:
        print(f"No sheets found for plant '{planta}'")
    for sheet in relevant_sheets:
        df = dfs[sheet]
        body = f"Reporte Inasistencias - {planta} - {sheet}\n\n"
        grouped_df = df.groupby('CentroProgramado')
        for centro, group in grouped_df:
            body += f"Centro: {centro}\n"
            body += f"  Planta: {planta}\n"
            turnos_group = group.groupby('Horario_Entrada')
            for turno, gg in turnos_group:
                body += f"    Turno: {turno}\n"
                for _, row in gg.iterrows():
                    id_doc = row['IdDocumento']
                    nombre = row['Nombre']
                    obs2 = str(row.get('Observacion2', ''))
                    obs3 = str(row.get('Observacion_Act', ''))
                    recomendacion = ' - '.join([obs2, obs3]).strip(' - ')
                    body += f"      - {id_doc}: {nombre}\n"
                    if recomendacion:
                        body += f"          Recomendación: {recomendacion}\n"
            body += "\n"
        subject = f"Reporte Inasistencias - {planta} - {sheet}"
        send_email(emails, subject, body)

# Collect all data for Slack
all_data = []
for sheet_name, df in dfs.items():
    if sheet_name not in ['Correos_Plantas', 'Entrada_2_00_P.M.', 'Entrada_6_00_A.M.', 'TURNO3', 'TURNO1','12H NOC']:
        # Find the plant for this sheet
        if 'PlantaProgramada' in df.columns:
            unique_plants = df['PlantaProgramada'].unique()
            plant_for_sheet = unique_plants[0] if len(unique_plants) == 1 else sheet_name
        else:
            plant_for_sheet = next((p for p in plantas_emails.keys() if p in sheet_name), sheet_name)
        df_copy = df.copy()
        df_copy['Planta'] = plant_for_sheet
        all_data.append(df_copy)

if all_data:
    df_all = pd.concat(all_data, ignore_index=True)
    # Count per plant
    plant_counts = df_all.groupby('Planta').size()
    print(plant_counts)
    mensaje = "📊 *REPORTE INASISTENCIAS* 📊\n\n"
    mensaje += "Conteo de personas por planta:\n"
    for plant, count in plant_counts.items():
        mensaje += f"🏢 *{plant}*: {count} personas\n"
    total_casos = plant_counts.sum()
    mensaje += f"\n📈 *Total de casos compañía*: {total_casos} personas\n\n"

    # Create Excel with only specific sheets and convert to tables
    selected_sheets = ['Entrada_2_00_P.M.', 'TURNO2-12H DIA', 'TURNO3', 'TURNO1', '12H NOC']
    temp_filename = "temp_INASISTENCIAS.xlsx"
    
    # Filter only sheets that exist in the workbook
    existing_sheets = [sheet for sheet in selected_sheets if sheet in dfs]
    
    # Only create temp file if there are existing sheets to include
    if existing_sheets:
        # Paso 1: Write Excel with pandas
        with pd.ExcelWriter(temp_filename, engine='openpyxl') as writer:
            for sheet in existing_sheets:
                dfs[sheet].to_excel(writer, sheet_name=sheet, index=False)
        
        # Paso 2: Open with openpyxl to convert to tables
        wb = openpyxl.load_workbook(temp_filename)
        
        for sheet_name in existing_sheets:
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                # Get the dimensions of the data
                max_row = ws.max_row
                max_col = ws.max_column
                
                # Only create table if there is data
                if max_row > 1 and max_col > 0:
                    # Define table range (from A1 to last cell with data)
                    table_ref = f"A1:{openpyxl.utils.get_column_letter(max_col)}{max_row}"
                    
                    # Create table with unique name for each sheet
                    tab = Table(displayName=f"Tabla_{sheet_name.replace(' ', '_').replace('-', '_')}", ref=table_ref)
                    
                    # Add a default table style
                    style = TableStyleInfo(
                        name="TableStyleMedium9",
                        showFirstColumn=False,
                        showLastColumn=False,
                        showRowStripes=True,
                        showColumnStripes=False
                    )
                    tab.tableStyleInfo = style
                    
                    # Add table to worksheet
                    ws.add_table(tab)
        
        # Save the workbook with tables
        wb.save(temp_filename)

        # Send the Excel file via email to carlos.grajales@ccsuper.co
        email_body = "Adjunto el reporte de inasistencias."
        send_email(['carlos.grajales@ccsuper.co'], "Inasistencias - PowerAutomate", email_body, temp_filename)
        
        # Clean up temp file
        if os.path.exists(temp_filename):
            os.remove(temp_filename)
    else:
        print("No se encontraron hojas válidas para el archivo temporal. No se enviará email con adjunto.")
    
    # Agregar el enlace del aplicativo al mensaje de Slack
    mensaje += ":warning::bangbang: *Haga click <https://apps.powerapps.com/play/e/default-a8da24a4-606b-4d75-ae3f-79ed5e74a421/a/d899f343-04c0-4cf2-8bf6-7a30443d5ca3?tenantId=a8da24a4-606b-4d75-ae3f-79ed5e74a421&hint=a941663c-ac85-4342-b939-2951b30c42bd&sourcetime=1761076365896|aquí> para reportar inasistencias* :warning::bangbang:"

    # Send to Slack with link (sin archivo adjunto en Slack)
    slack_payload = {
        "channel": SLACK_CHANNEL,
        "text": mensaje
    }

    headers_slack = {
        "Authorization": f"Bearer {SLACK_TOKEN}",
        "Content-Type": "application/json; charset=utf-8"
    }
    resp = requests.post(
        "https://slack.com/api/chat.postMessage",
        headers=headers_slack,
        json=slack_payload
    )
    if resp.ok and resp.json().get('ok'):
        print("Mensaje enviado a Slack con enlace del aplicativo.")
    else:
        print(f"Error enviando a Slack: {resp.text}")
else:
    print("No data found for Slack report.")
