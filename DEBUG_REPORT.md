# Reporte de Depuración: NotiInasistencias

## 🔍 Problema Reportado
El script `NotiInasistencias.py` no envía el correo con asunto "Inasistencias - PowerAutomate".

## 🕵️ Análisis Realizado

### 1. Flujo de Ejecución
- **inasistencias2pm.py**: Extrae datos de Tableau y genera `Inasistencias.xlsx`
- **NotiInasistencias.py**: Lee el Excel y envía correos, incluyendo uno crítico con asunto "Inasistencias - PowerAutomate"

### 2. Hallazgos Clave

#### ✅ La lógica del código está correcta
La depuración demostró que:
- El código SÍ llega a la sección donde debe enviar el correo PowerAutomate
- El archivo temporal `temp_INASISTENCIAS.xlsx` SÍ se crea correctamente
- La función `send_email()` SÍ se invoca con todos los parámetros correctos
- La hoja `'Entrada_2_00_P.M.'` SÍ es encontrada en el Excel

#### ❌ El problema está en el envío SMTP
Las posibles causas del fallo son:

1. **Error de autenticación SMTP** (más probable)
   - Credenciales incorrectas o expiradas
   - Password cambiado y no actualizado en el script
   - Autenticación multifactor activada en Office365

2. **Problemas de conectividad**
   - Firewall bloqueando puerto 587
   - Servidor SMTP inaccesible

3. **El correo se envía pero no se recibe**
   - Correos van a carpeta de spam
   - Filtros de correo bloqueando el mensaje

4. **Manejo de errores insuficiente**
   - El bloque `try/except` original captura errores pero no muestra detalles

## 🛠️ Mejoras Implementadas

### 1. Logging Detallado en `send_email()`

**Antes:**
```python
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    server.starttls()
    server.login(EMAIL_USER, EMAIL_PASS)
    # ...
    print(f"Email sent to {to_emails} for {subject}")
except Exception as e:
    print(f"Error sending email: {e}")
```

**Ahora:**
```python
print(f"📧 Preparando envío de email")
print(f"Para: {to_emails}")
print(f"Asunto: {subject}")
# Verifica archivo adjunto
if attachment_path:
    if not os.path.exists(attachment_path):
        print(f"❌ ERROR: El archivo adjunto NO existe")
        return
    file_size = os.path.getsize(attachment_path)
    print(f"✓ Archivo adjunto existe ({file_size} bytes)")

# Logging paso a paso
print(f"🔌 Conectando a {SMTP_SERVER}:{SMTP_PORT}...")
print(f"🔐 Iniciando TLS...")
print(f"👤 Autenticando como {EMAIL_USER}...")
print(f"📤 Enviando correo...")
print(f"✅ Email enviado exitosamente!")
```

### 2. Manejo de Excepciones Específico

**Ahora se distingue entre:**
- `SMTPAuthenticationError`: Problemas de credenciales
- `SMTPException`: Otros errores SMTP
- `Exception`: Errores inesperados con traceback completo

### 3. Logging en Sección Crítica

Se agregó depuración en la verificación de hojas:
```python
print(f"🔍 Verificando hojas para correo PowerAutomate")
print(f"Hojas buscadas: {selected_sheets}")
print(f"Hojas disponibles en Excel: {list(dfs.keys())}")
print(f"Hojas encontradas: {existing_sheets}")
```

## 📋 Próximos Pasos para Diagnóstico

### 1. Ejecutar con las mejoras
```bash
python3 NotiInasistencias.py
```

### 2. Analizar el output
El nuevo logging mostrará **exactamente** en qué paso falla:

#### Si el problema es autenticación:
```
❌ ERROR DE AUTENTICACIÓN: (535, b'5.7.3 Authentication unsuccessful')
   Verifica las credenciales: carlos.grajales@ccsuper.co
```
**Solución**: Actualizar contraseña o habilitar "App Password" en Office365

#### Si el problema es conectividad:
```
❌ ERROR INESPERADO: TimeoutError: [Errno 110] Connection timed out
```
**Solución**: Verificar firewall y conectividad al servidor SMTP

#### Si el archivo no existe:
```
❌ ERROR: El archivo adjunto NO existe: temp_INASISTENCIAS.xlsx
```
**Solución**: Revisar por qué no se crea el archivo temporal

### 3. Verificar que inasistencias2pm.py genera las hojas correctas

El script busca estas hojas específicas:
- `'Entrada_2_00_P.M.'` ✓ (generada por inasistencias2pm.py)
- `'TURNO2-12H DIA'` ❓ (NO generada)
- `'TURNO3'` ❓ (NO generada)
- `'TURNO1'` ❓ (NO generada)
- `'12H NOC'` ❓ (NO generada)

**Si solo existe una hoja, el correo DEBERÍA enviarse igual.**

## 🔐 Notas de Seguridad

⚠️ El archivo `NotiInasistencias.py` contiene credenciales en texto plano:
```python
EMAIL_USER = 'carlos.grajales@ccsuper.co'
EMAIL_PASS = 'SUPcagm25-10*#'  # ⚠️ EXPUESTO
```

**Recomendación**: Usar variables de entorno:
```python
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')
```

## 📊 Resumen

| Aspecto | Estado | Notas |
|---------|--------|-------|
| Lógica del código | ✅ Correcta | El flujo llega hasta el envío |
| Creación de archivos | ✅ Correcta | temp_INASISTENCIAS.xlsx se crea |
| Detección de hojas | ✅ Correcta | Encuentra 'Entrada_2_00_P.M.' |
| Envío SMTP | ❌ A verificar | Requiere prueba con credenciales reales |
| Logging | ✅ Mejorado | Ahora muestra cada paso detalladamente |
| Manejo de errores | ✅ Mejorado | Distingue entre tipos de error |

## 🎯 Conclusión

El problema **NO es la lógica del código**, sino muy probablemente:
1. **Credenciales SMTP incorrectas o expiradas** (más probable)
2. **Problemas de conectividad al servidor SMTP**
3. **Correos enviados pero no recibidos (spam)**

Las mejoras implementadas permitirán identificar **exactamente** cuál es el problema al ejecutar el script con credenciales reales.
