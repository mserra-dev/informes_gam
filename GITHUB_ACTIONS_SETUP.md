# Setup GitHub Actions — Informes Mensuales GAM

Automatización 100% en la nube. Tu máquina no necesita estar encendida.

---

## Estructura de archivos a subir a GitHub

```
tu-repo/
└── automatizacion/
    ├── informe_mensual_gam.py        ← script principal
    ├── requirements.txt
    └── .github/
        └── workflows/
            └── informes_mensuales.yml  ← workflow de GitHub Actions
```

> **No subas nunca `service_account.json` al repo.** Las credenciales van como Secret.

---

## Paso 1 — Crear el repositorio en GitHub

1. Ir a https://github.com/new
2. Nombre: `gam-informes-ellitoral` (puede ser privado ✅)
3. Crear sin README

---

## Paso 2 — Subir los archivos

Desde tu terminal, en la carpeta `automatizacion/`:

```bash
cd "/ruta/a/Google Ad Manager/automatizacion"
git init
git remote add origin https://github.com/TU_USUARIO/gam-informes-ellitoral.git
git add informe_mensual_gam.py requirements.txt .github/
git commit -m "Informes mensuales GAM El Litoral"
git push -u origin main
```

---

## Paso 3 — Codificar el service account en base64

Necesitás el archivo `service_account.json` del service account que ya usás para el MCP de GAM.

**En macOS / Linux:**
```bash
base64 -i service_account.json | pbcopy
# Ahora el base64 está en tu portapapeles
```

**En Windows (PowerShell):**
```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("service_account.json")) | Set-Clipboard
```

---

## Paso 4 — Agregar el Secret en GitHub

1. Ir al repo en GitHub
2. **Settings → Secrets and variables → Actions → New repository secret**
3. Nombre: `GAM_SA_JSON_B64`
4. Valor: pegar el base64 del paso anterior
5. Click **Add secret**

> El workflow decodifica este secret y lo escribe como archivo temporal solo durante la ejecución. Se elimina automáticamente al terminar.

---

## Paso 5 — Verificar permisos del service account

El service account necesita tener habilitados estos scopes de Google API:

| API | Scope | Para qué |
|-----|-------|---------|
| Google Ad Manager | `https://www.googleapis.com/auth/dfp` | Reportes GAM |
| Google Drive | `https://www.googleapis.com/auth/drive` | Subir Excel |
| Gmail | `https://www.googleapis.com/auth/gmail.send` | Enviar email |

Si `arcadiaconsultora.com` es Google Workspace, activá **domain-wide delegation**:
- Google Admin → Seguridad → Controles de API → Delegación en todo el dominio
- Agregar el Client ID del service account con los 3 scopes
- Descomentar `IMPERSONATE_USER` en el script

---

## Paso 6 — Probar manualmente

Una vez subido el repo y configurado el secret:

1. Ir a **Actions** en el repo
2. Click en **Informes Mensuales GAM — El Litoral**
3. Click en **Run workflow** (botón verde)
4. Elegir qué informe generar: `todos`, `pautas`, `bloques`, o `programatica`
5. Verificar los logs en tiempo real

---

## Ejecución automática

El workflow corre automáticamente **el 1ro de cada mes a las 8:00 AM (hora Argentina)**
sin que hagas nada. GitHub se encarga de todo.

Para cambiar el horario, editá la línea en el workflow:
```yaml
- cron: "0 11 1 * *"   # 11:00 UTC = 8:00 AM Argentina (UTC-3)
```

---

## Monitoreo y alertas

GitHub Actions envía un email automático a tu cuenta de GitHub si el workflow falla.
Podés verlo en: **Actions → (nombre del run) → Ver logs**

---

## Costos

**GitHub Actions es gratuito** para repositorios privados dentro del free tier:
- 2.000 minutos/mes de ejecución gratis
- El script tarda aprox. 2-3 minutos → usás ~3 minutos por mes

Costo estimado: **$0**
