#!/usr/bin/env python3
"""
Previsión de disponibilidad — El Litoral · Google Ad Manager
=============================================================
Consulta el ForecastService de GAM y devuelve cuántas impresiones quedan
disponibles para un tamaño de creatividad en un conjunto de sitios
(ad units de primer nivel), en una única previsión combinada.

Al pedir UNA sola previsión para todos los sitios (en lugar de sumar una
previsión por sitio) GAM nunca cuenta dos veces la misma impresión: cada
impresión disponible se asigna a un único ad unit. El desglose por sitio sale
de esa misma previsión (targetingCriteriaBreakdown).

Además corre una segunda previsión idéntica pero con frequency cap de
1 impresión por usuario por día, para saber cuántas impresiones quedan si la
campaña no debe repetirse al mismo usuario.

Uso:
  python prevision_disponibilidad.py
  python prevision_disponibilidad.py --sitios notife,radio,sur24 --tamano 300x250 --dias 30
  python prevision_disponibilidad.py --inicio 2026-10-01 --dias 31
  python prevision_disponibilidad.py --sin-frecuencia      # omite la variante con cap

Credenciales: mismas que informe_mensual_gam.py (GAM_SA_PATH o
service_account.json junto al script).
"""

import argparse
import csv
import datetime
import json
import os
import re
import sys

from informe_mensual_gam import GAM_API_VERSION, gam_client

SITIOS_DEFAULT = "notife,radio,radiorafaela,puerto,puertonegocios,sur24"
TAMANO_DEFAULT = "300x250"
DIAS_DEFAULT   = 30
OUTPUT_DIR     = os.getenv("PREVISION_OUTPUT_DIR", "/tmp")

# Frequency cap de la variante "sin repetir al usuario".
FREQ_CAP = {"maxImpressions": 1, "numTimeUnits": 1, "timeUnit": "DAY"}


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def _norm(s: str) -> str:
    """Normaliza nombres para comparar: minúsculas, sin espacios ni signos."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def parse_tamano(s: str) -> tuple[int, int]:
    m = re.fullmatch(r"\s*(\d+)\s*[xX×]\s*(\d+)\s*", s)
    if not m:
        raise SystemExit(f"Tamaño inválido: {s!r} (esperado ANCHOxALTO, ej. 300x250)")
    return int(m.group(1)), int(m.group(2))


def gam_datetime(d: datetime.date, tz: str, end_of_day: bool = False) -> dict:
    return {
        "date":       {"year": d.year, "month": d.month, "day": d.day},
        "hour":       23 if end_of_day else 0,
        "minute":     59 if end_of_day else 0,
        "second":     59 if end_of_day else 0,
        "timeZoneId": tz,
    }


def _int(v) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


# ──────────────────────────────────────────────────────────────
# AD UNITS
# ──────────────────────────────────────────────────────────────
def resolver_sitios(client, sitios: list[str]) -> dict[str, dict]:
    """Devuelve {sitio: {"id", "name"}} buscando ad units cuyo nombre
    (normalizado) coincida exactamente con el sitio pedido.

    Primero mira los ad units de primer nivel (hijos del root, como
    "ellitoral"). Si alguno no aparece ahí, busca en todo el árbol por
    nombre o adUnitCode. Falla listando candidatos si no hay match exacto.
    """
    inv = client.GetService("InventoryService", version=GAM_API_VERSION)
    net = client.GetService("NetworkService",   version=GAM_API_VERSION)
    root_id = net.getCurrentNetwork().effectiveRootAdUnitId

    top = list(getattr(inv.getAdUnitsByStatement(
        {"query": f"WHERE parentId = {root_id} LIMIT 500"}), "results", []) or [])
    print(f"   · {len(top)} ad units de primer nivel: "
          + ", ".join(sorted(au.name for au in top)))

    por_nombre = {}
    for au in top:
        por_nombre.setdefault(_norm(au.name), au)
        por_nombre.setdefault(_norm(getattr(au, "adUnitCode", "")), au)

    encontrados, faltantes = {}, []
    for sitio in sitios:
        au = por_nombre.get(_norm(sitio))
        if au is not None:
            encontrados[sitio] = {"id": int(au.id), "name": au.name}
        else:
            faltantes.append(sitio)

    for sitio in faltantes:
        # Búsqueda en todo el árbol. PQL LIKE es case-sensitive, por eso
        # traemos candidatos amplios y filtramos acá con _norm.
        stmt = {
            "query":  "WHERE name LIKE :p OR adUnitCode LIKE :p LIMIT 500",
            "values": [{"key": "p", "value": {"xsi_type": "TextValue",
                                              "value": f"%{sitio}%"}}],
        }
        cands = list(getattr(inv.getAdUnitsByStatement(stmt), "results", []) or [])
        exact = [au for au in cands
                 if _norm(au.name) == _norm(sitio)
                 or _norm(getattr(au, "adUnitCode", "")) == _norm(sitio)]
        if len(exact) == 1:
            au = exact[0]
            encontrados[sitio] = {"id": int(au.id), "name": au.name}
            print(f"   · '{sitio}' resuelto fuera del primer nivel → {au.name} ({au.id})")
            continue
        # Abreviatura de otro sitio ya pedido (ej. 'radio' → 'Radio Rafaela',
        # 'puerto' → 'puertonegocios'): se omite con aviso en vez de fallar.
        ya_pedidos = {au["id"] for au in encontrados.values()}
        if cands and all(int(au.id) in ya_pedidos for au in cands):
            nombres = ", ".join(au.name for au in cands)
            print(f"   ⚠ '{sitio}' no existe como ad unit propio; ya está cubierto por: "
                  f"{nombres}. Se omite.")
        else:
            listado = ", ".join(f"{au.name} ({au.id})" for au in cands[:20]) or "ninguno"
            raise SystemExit(
                f"No se pudo resolver el sitio '{sitio}' a un ad unit único. "
                f"Candidatos: {listado}")

    if not encontrados:
        raise SystemExit("Ningún sitio pudo resolverse a un ad unit.")
    return encontrados


# ──────────────────────────────────────────────────────────────
# FORECAST
# ──────────────────────────────────────────────────────────────
def prospective_line_item(ad_unit_ids, width, height, inicio, fin, tz,
                          freq_cap=None) -> dict:
    li = {
        "lineItemType":         "STANDARD",
        "endDateTime":          gam_datetime(fin, tz, end_of_day=True),
        "costType":             "CPM",
        "creativeRotationType": "EVEN",
        "creativePlaceholders": [{
            "size": {"width": width, "height": height, "isAspectRatio": False},
        }],
        "targeting": {
            "inventoryTargeting": {
                "targetedAdUnits": [
                    {"adUnitId": au_id, "includeDescendants": True}
                    for au_id in ad_unit_ids
                ],
            },
        },
        # El objetivo no afecta availableUnits; es obligatorio para el forecast.
        "primaryGoal": {"goalType": "LIFETIME", "unitType": "IMPRESSIONS",
                        "units": 1_000_000},
    }
    hoy = datetime.date.today()
    if inicio <= hoy:
        li["startDateTimeType"] = "IMMEDIATELY"
    else:
        li["startDateTimeType"] = "USE_START_DATE_TIME"
        li["startDateTime"]     = gam_datetime(inicio, tz)
    if freq_cap:
        li["frequencyCaps"] = [freq_cap]
    return li


def correr_forecast(client, line_item: dict) -> dict:
    svc = client.GetService("ForecastService", version=GAM_API_VERSION)
    fc  = svc.getAvailabilityForecast(
        {"lineItem": line_item},
        {"includeContendingLineItems": False,
         "includeTargetingCriteriaBreakdown": True},
    )
    breakdown = {}
    for b in (getattr(fc, "targetingCriteriaBreakdowns", None) or []):
        if getattr(b, "targetingDimension", "") != "AD_UNIT":
            continue
        breakdown[_int(b.targetingCriteriaId)] = {
            "name":      getattr(b, "targetingCriteriaName", ""),
            "matched":   _int(getattr(b, "matchedUnits", 0)),
            "available": _int(getattr(b, "availableUnits", 0)),
        }
    return {
        "matched":   _int(getattr(fc, "matchedUnits", 0)),
        "available": _int(getattr(fc, "availableUnits", 0)),
        "possible":  _int(getattr(fc, "possibleUnits", 0)),
        "reserved":  _int(getattr(fc, "reservedUnits", 0)),
        "breakdown": breakdown,
    }


# ──────────────────────────────────────────────────────────────
# SALIDA
# ──────────────────────────────────────────────────────────────
def _fmt(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def imprimir(resultado: dict) -> str:
    """Imprime resumen legible por consola y devuelve la versión Markdown."""
    tam   = resultado["tamano"]
    per   = f"{resultado['inicio']} → {resultado['fin']} ({resultado['dias']} días)"
    std   = resultado["previsiones"]["combinada"]
    cap   = resultado["previsiones"].get("frecuencia_1_por_dia")
    sitios = resultado["sitios"]

    md  = [f"## Previsión de disponibilidad {tam}", "",
           f"**Período:** {per}  ",
           f"**Sitios:** {', '.join(sitios)}  ",
           f"**Previsión combinada** (una sola consulta, sin doble conteo entre sitios)", "",
           "| Métrica | Impresiones |", "|---|---:|",
           f"| Disponibles (available) | **{_fmt(std['available'])}** |",
           f"| Que coinciden con el targeting (matched) | {_fmt(std['matched'])} |",
           f"| Reservadas por otras líneas (reserved) | {_fmt(std['reserved'])} |",
           f"| Máximo bookeable (possible) | {_fmt(std['possible'])} |"]
    if cap:
        md += [f"| Disponibles con cap 1 imp/usuario/día | **{_fmt(cap['available'])}** |"]
    md += ["", "### Desglose por sitio (misma previsión)", "",
           "| Sitio | Ad unit | Disponibles | Matched" + (" | Disp. con cap 1/día |" if cap else " |"),
           "|---|---|---:|---:" + ("|---:|" if cap else "|")]
    for s in resultado["por_sitio"]:
        fila = f"| {s['sitio']} | {s['ad_unit']} ({s['ad_unit_id']}) | {_fmt(s['available'])} | {_fmt(s['matched'])} |"
        if cap:
            fila += f" {_fmt(s['available_cap'])} |"
        md.append(fila)
    md += ["", f"_Generado {resultado['generado']} · Red {resultado['network_code']} · "
              f"Zona horaria {resultado['timezone']}_"]
    texto = "\n".join(md)
    print("\n" + texto + "\n")
    return texto


def guardar(resultado: dict, md: str) -> list[str]:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    base   = f"prevision_{resultado['tamano']}_{resultado['inicio']}_{resultado['dias']}d"
    p_json = os.path.join(OUTPUT_DIR, base + ".json")
    p_csv  = os.path.join(OUTPUT_DIR, base + ".csv")
    p_md   = os.path.join(OUTPUT_DIR, base + ".md")
    with open(p_json, "w", encoding="utf-8") as f:
        json.dump(resultado, f, ensure_ascii=False, indent=2)
    with open(p_csv, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["sitio", "ad_unit", "ad_unit_id", "available", "matched", "available_cap_1_dia"])
        for s in resultado["por_sitio"]:
            w.writerow([s["sitio"], s["ad_unit"], s["ad_unit_id"],
                        s["available"], s["matched"], s.get("available_cap", "")])
        std = resultado["previsiones"]["combinada"]
        cap = resultado["previsiones"].get("frecuencia_1_por_dia") or {}
        w.writerow(["TOTAL COMBINADO", "", "", std["available"], std["matched"],
                    cap.get("available", "")])
    with open(p_md, "w", encoding="utf-8") as f:
        f.write(md + "\n")
    summary = os.getenv("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(md + "\n")
    return [p_json, p_csv, p_md]


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="Previsión de disponibilidad GAM por sitios")
    ap.add_argument("--sitios", default=SITIOS_DEFAULT,
                    help=f"Lista separada por comas (default: {SITIOS_DEFAULT})")
    ap.add_argument("--tamano", default=TAMANO_DEFAULT, help="ANCHOxALTO (default 300x250)")
    ap.add_argument("--dias",   type=int, default=DIAS_DEFAULT, help="Duración en días")
    ap.add_argument("--inicio", default="",
                    help="YYYY-MM-DD. Vacío = desde ahora (IMMEDIATELY)")
    ap.add_argument("--sin-frecuencia", action="store_true",
                    help="No correr la variante con frequency cap 1/usuario/día")
    args = ap.parse_args()

    sitios = [s.strip() for s in args.sitios.split(",") if s.strip()]
    if not sitios:
        raise SystemExit("Debe indicar al menos un sitio.")
    if args.dias < 1:
        raise SystemExit("--dias debe ser >= 1")
    width, height = parse_tamano(args.tamano)

    hoy = datetime.date.today()
    if args.inicio:
        inicio = datetime.date.fromisoformat(args.inicio)
    else:
        inicio = hoy
    fin = inicio + datetime.timedelta(days=args.dias - 1)
    if fin < hoy:
        raise SystemExit("El período termina en el pasado; ajustá --inicio/--dias.")

    print(f"📈 Previsión {width}x{height} · {inicio} → {fin} ({args.dias} días)")
    print(f"   · Sitios: {', '.join(sitios)}")

    client = gam_client()
    net    = client.GetService("NetworkService", version=GAM_API_VERSION).getCurrentNetwork()
    tz     = getattr(net, "timeZone", "America/Argentina/Buenos_Aires")

    print("🔎 Resolviendo ad units...")
    ad_units = resolver_sitios(client, sitios)
    for s, au in ad_units.items():
        print(f"   · {s:16s} → {au['name']} ({au['id']})")
    ids    = [au["id"] for au in ad_units.values()]
    sitios = list(ad_units)  # solo los que se resolvieron

    print("⏳ Previsión combinada (sin frequency cap)...")
    std = correr_forecast(client, prospective_line_item(ids, width, height, inicio, fin, tz))
    print(f"   · available={std['available']:,}  matched={std['matched']:,}  "
          f"possible={std['possible']:,}")

    cap = None
    if not args.sin_frecuencia:
        print("⏳ Previsión combinada con frequency cap 1 imp/usuario/día...")
        cap = correr_forecast(client, prospective_line_item(
            ids, width, height, inicio, fin, tz, freq_cap=FREQ_CAP))
        print(f"   · available={cap['available']:,}  matched={cap['matched']:,}")

    por_sitio = []
    for sitio, au in ad_units.items():
        b = std["breakdown"].get(au["id"], {})
        fila = {
            "sitio":      sitio,
            "ad_unit":    au["name"],
            "ad_unit_id": au["id"],
            "available":  b.get("available", 0),
            "matched":    b.get("matched", 0),
        }
        if cap:
            fila["available_cap"] = cap["breakdown"].get(au["id"], {}).get("available", 0)
        por_sitio.append(fila)
    por_sitio.sort(key=lambda r: -r["available"])

    resultado = {
        "generado":     datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "network_code": str(getattr(net, "networkCode", "")),
        "timezone":     tz,
        "tamano":       f"{width}x{height}",
        "inicio":       inicio.isoformat(),
        "fin":          fin.isoformat(),
        "dias":         args.dias,
        "sitios":       sitios,
        "ad_units":     ad_units,
        "previsiones":  {
            "combinada": {k: v for k, v in std.items() if k != "breakdown"},
            **({"frecuencia_1_por_dia": {k: v for k, v in cap.items() if k != "breakdown"}}
               if cap else {}),
        },
        "por_sitio":    por_sitio,
    }

    md    = imprimir(resultado)
    files = guardar(resultado, md)
    print("💾 Archivos: " + ", ".join(files))
    print("\n✅ Listo.")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (0, None):
            print(f"\n❌ {e}", file=sys.stderr)
        raise
