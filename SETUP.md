# Setup — Informes mensuales automáticos GAM El Litoral

## Requisitos previos

### 1. Service account con acceso a GAM, Drive y Gmail

El service account que ya usás para el MCP de GAM necesita:

- **GAM**: acceso de lectura como usuario de la red (ya está configurado)
- **Drive**: scope `https://www.googleapis.com/auth/drive`
- **Gmail**: scope `https://www.googleapis.com/auth/gmail.send`

Si `arcadiaconsultora.com` es Google Workspace, podés activar **domain-wide delegation**:
1. Google Admin → Seguridad → Controles de API → Delegación en todo el dominio
2. Agregar el client_id del service account con los scopes de Drive y Gmail
3. Descomentar `IMPERSONATE_USER` en `informe_mensual_gam.py`

Si es Gmail personal (sin Workspace), necesitás credenciales OAuth2 separadas para Gmail
(el script puede adaptarse — consultame).

### 2. Copiar el JSON del service account

```bash
cp /ruta/a/tu/service_account.json automatizacion/service_account.json
```

### 3. Instalar dependencias

```bash
cd automatizacion/
pip install -r requirements.txt
```

---

## Prueba manual

```bash
# Ambos informes (mes anterior)
python informe_mensual_gam.py

# Solo pautas
python informe_mensual_gam.py --pautas

# Solo bloques/CTR
python informe_mensual_gam.py --bloques
```

---

## Automatización con cron (Linux/macOS)

Agregar al crontab (`crontab -e`):

```cron
# El 1ro de cada mes a las 8:00 AM
0 8 1 * * /usr/bin/python3 /ruta/absoluta/automatizacion/informe_mensual_gam.py >> /var/log/gam_informes.log 2>&1
```

---

## Qué hace el script

| Paso | Acción |
|------|--------|
| 1 | Calcula automáticamente el mes anterior (start/end) |
| 2 | Llama a la API de GAM y corre un ReportJob |
| 3 | Descarga el CSV y genera el Excel con openpyxl |
| 4 | Sube (o actualiza) el archivo en la carpeta de Drive |
| 5 | Envía el email HTML con KPIs y top anunciantes/bloques |

Archivos generados:
- `Informe_Pautas_VentaDirecta_MesAño.xlsx`
- `Informe_CTR_Viewability_MesAño.xlsx`
