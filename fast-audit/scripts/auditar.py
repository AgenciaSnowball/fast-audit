#!/usr/bin/env python3
"""Fast Audit: lee un CSV de oportunidades exportado del CRM y devuelve alertas.

Uso:
    python auditar.py <fichero.csv|.xlsx> [--hoy AAAA-MM-DD] [--map campo=Columna ...]
                      [--umbral clave=valor ...] [--fechas dia|mes]

Imprime un JSON con el mapeo de columnas, un resumen, las alertas ordenadas por prioridad
y una lista de tareas por comercial. Solo usa la biblioteca estándar (openpyxl, si está
instalado, para .xlsx). No se conecta a nada: todo se calcula en local.
"""
import argparse
import csv
import io
import json
import re
import statistics
import sys
import unicodedata
from datetime import date, datetime, timedelta, timezone

# --------------------------------------------------------------------------- umbrales

UMBRALES = {
    "dias_sin_primer_contacto": 3,      # oportunidad creada hace 3+ días y sin actividad
    "dias_parada": 14,                  # sin actividad registrada en 14+ días
    "dias_propuesta_sin_seguimiento": 7,
    "dias_comercial_sin_actividad": 5,  # comercial cuya actividad más reciente es más antigua
    "dias_en_etapa": 45,                # con fecha de cambio de etapa o días en etapa
    "dias_etapa_acumula": 60,           # sin esa fecha: edad desde la creación
    "pct_motivo_vacio": 20,             # % de pérdidas sin motivo útil a partir del cual avisa
    "pct_sin_importe": 20,
    "factor_cola": 2.0,                 # cola de un comercial frente a la mediana
}

# --------------------------------------------------------------------------- columnas

CAMPOS = {
    "id": ["record id", "id del registro", "deal id", "id del trato", "id del negocio", "opportunity id",
           "id de la oportunidad", "id"],
    "nombre": ["deal name", "nombre del negocio", "title", "deal title", "titulo", "nombre del trato",
               "nombre de trato", "opportunity name", "nombre de la oportunidad", "oportunidad",
               "negocio", "trato", "name", "nombre"],
    "etapa": ["deal stage", "etapa del negocio", "stage", "etapa", "fase", "etapa del trato",
              "pipeline stage", "sales stage"],
    "estado": ["status", "estado", "deal status", "estado del trato", "estado del negocio"],
    "comercial": ["deal owner", "propietario del negocio", "owner", "propietario", "opportunity owner",
                  "propietario de la oportunidad", "propietario de trato", "propietario del trato",
                  "assigned to", "asignado a", "responsable", "comercial", "vendedor", "owner name",
                  "deal owner name", "usuario asignado", "sales rep"],
    "importe": ["amount", "importe", "value", "valor", "deal value", "valor del trato", "cantidad",
                "monto", "valor del negocio", "importe del negocio", "expected revenue"],
    "creado": ["create date", "fecha de creacion", "deal created", "trato creado", "negocio creado",
               "created time", "hora de creacion", "created date", "created", "creado",
               "fecha creacion", "add time", "created on", "creado el"],
    "ult_actividad": ["last activity date", "fecha de la ultima actividad", "fecha de ultima actividad",
                      "last activity", "ultima actividad", "last activity time",
                      "hora de la ultima actividad", "last contacted", "ultimo contacto",
                      "fecha del ultimo contacto", "last engagement date", "fecha de la ultima interaccion"],
    "prox_actividad": ["next activity date", "fecha de la proxima actividad", "fecha de proxima actividad",
                       "proxima actividad", "next activity", "next step date", "fecha del proximo paso",
                       "next activity time"],
    "cierre": ["close date", "fecha de cierre", "expected close date", "fecha prevista de cierre",
               "fecha de cierre prevista", "closing date", "fecha esperada de cierre"],
    "cambio_etapa": ["last stage change", "last stage change date", "ultimo cambio de etapa",
                     "fecha del ultimo cambio de etapa", "stage change date", "fecha de cambio de etapa",
                     "date entered current stage", "fecha de entrada en la etapa actual",
                     "stage last updated"],
    "dias_etapa": ["stage duration", "duracion de la etapa", "days in stage", "dias en la etapa",
                   "dias en etapa", "time in current stage"],
    "motivo_perdida": ["closed lost reason", "motivo de perdida del negocio cerrado", "lost reason",
                       "motivo de perdida", "reason for loss", "razon de perdida", "loss reason",
                       "motivo perdida", "motivo de la perdida", "razon de la perdida"],
}
CAMPOS_FECHA = ["creado", "ult_actividad", "prox_actividad", "cierre", "cambio_etapa"]
PREFIJOS = re.compile(r"^(deal|trato|negocio|oportunidad|opportunity|deals|tratos|negocios)\s*-\s*")


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    s = PREFIJOS.sub("", s.strip())
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return s.strip()


def mapear(cabeceras, forzado):
    """Asigna cada campo a una columna: primero coincidencias exactas, luego parciales."""
    mapeo = {c: None for c in CAMPOS}
    usadas = set()
    for campo, col in forzado.items():
        if col not in cabeceras:
            raise SystemExit(f"--map {campo}={col}: esa columna no existe. Columnas: {cabeceras}")
        mapeo[campo] = col
        usadas.add(col)
    normas = {h: norm(h) for h in cabeceras}
    # Exactas, en el orden de sinónimos (el primero es el más específico)
    for campo, sins in CAMPOS.items():
        if mapeo[campo]:
            continue
        for s in sins:
            hit = next((h for h in cabeceras if h not in usadas and normas[h] == s), None)
            if hit:
                mapeo[campo] = hit
                usadas.add(hit)
                break
    # Parciales: la cabecera contiene un sinónimo de 2+ palabras
    for campo, sins in CAMPOS.items():
        if mapeo[campo]:
            continue
        mejor = None
        for h in cabeceras:
            if h in usadas or "date entered" in normas[h] or "fecha de entrada" in normas[h]:
                continue
            for s in sins:
                if " " in s and re.search(r"\b" + re.escape(s) + r"\b", normas[h]):
                    if not mejor or len(s) > mejor[1]:
                        mejor = (h, len(s))
        if mejor:
            mapeo[campo] = mejor[0]
            usadas.add(mejor[0])
    return mapeo

# --------------------------------------------------------------------------- lectura


def leer(ruta):
    if ruta.lower().endswith((".xlsx", ".xlsm")):
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("Para leer .xlsx hace falta openpyxl. Exporta en CSV o instálalo.")
        wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        ws = wb.active
        filas = [[("" if v is None else v) for v in r] for r in ws.iter_rows(values_only=True)]
        cab = [str(c).strip() for c in filas[0]]
        return cab, [dict(zip(cab, r)) for r in filas[1:] if any(str(v).strip() for v in r)]
    raw = open(ruta, "rb").read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        texto = raw.decode("utf-16")
    else:
        try:
            texto = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            texto = raw.decode("cp1252", errors="replace")
    muestra = texto[:20000]
    try:
        dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t|")
    except csv.Error:
        class dialecto(csv.excel):
            delimiter = ";" if muestra.count(";") > muestra.count(",") else ","
    lector = csv.reader(io.StringIO(texto), dialecto)
    filas = [r for r in lector if any(c.strip() for c in r)]
    cab = [c.strip() for c in filas[0]]
    return cab, [dict(zip(cab, r)) for r in filas[1:]]

# --------------------------------------------------------------------------- fechas e importes


MESES = {"jan": 1, "ene": 1, "feb": 2, "mar": 3, "apr": 4, "abr": 4, "may": 5, "jun": 6, "jul": 7,
         "aug": 8, "ago": 8, "sep": 9, "set": 9, "oct": 10, "nov": 11, "dec": 12, "dic": 12}
RE_ISO = re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})")
RE_NUM = re.compile(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})")
RE_TXT = re.compile(r"(\d{1,2})?\s*([a-zA-Z]{3,})\.?\s*(\d{1,2})?,?\s*(\d{4})")


def pista_orden(valores):
    """Mira si las fechas numéricas van día/mes o mes/día. Devuelve 'dia', 'mes' o None."""
    dia = mes = False
    for v in valores:
        m = RE_NUM.match(str(v).strip())
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            dia |= a > 12
            mes |= b > 12
    if dia and not mes:
        return "dia"
    if mes and not dia:
        return "mes"
    return None


def a_fecha(v, orden):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v or "").strip()
    if not s or s in ("-", "--", "N/A", "n/a"):
        return None
    try:
        m = RE_ISO.match(s)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        m = RE_NUM.match(s)
        if m:
            a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            y = y + 2000 if y < 100 else y
            d, mo = (a, b) if orden != "mes" else (b, a)
            return date(y, mo, d)
        m = RE_TXT.search(s)
        if m and m.group(2)[:3].lower() in MESES:
            d = int(m.group(1) or m.group(3) or 1)
            return date(int(m.group(4)), MESES[m.group(2)[:3].lower()], d)
        if re.fullmatch(r"\d{5}(\.\d+)?", s):          # número de serie de Excel
            return date(1899, 12, 30) + timedelta(days=int(float(s)))
        if re.fullmatch(r"\d{12,13}", s):               # epoch en milisegundos
            return datetime.fromtimestamp(int(s) / 1000, tz=timezone.utc).date()
    except ValueError:
        return None
    return None


def a_importe(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v or ""))
    if not s or s in "-.,":
        return None
    if "," in s and "." in s:
        dec = "," if s.rfind(",") > s.rfind(".") else "."
        s = s.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in s:
        s = s.replace(",", "") if re.search(r",\d{3}$", s) else s.replace(",", ".")
    elif re.search(r"\.\d{3}$", s) and s.count(".") >= 1 and not re.search(r"\.\d{3}\d", s):
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None

# --------------------------------------------------------------------------- estado


def estado_de(fila, mapeo):
    textos = []
    if mapeo["estado"]:
        textos.append(norm(fila.get(mapeo["estado"])))
    if mapeo["etapa"]:
        textos.append(norm(fila.get(mapeo["etapa"])))
    for t in textos:
        if re.search(r"\b(ganad\w*|won|win|gagne\w*)\b", t):
            return "ganada"
        if re.search(r"\b(perdid\w*|lost|lose|perdu\w*|descartad\w*)\b", t):
            return "perdida"
        if re.search(r"\b(deleted|eliminad\w*|borrad\w*)\b", t):
            return "borrada"
    return "abierta"


GENERICOS = {"", "otros", "otro", "other", "others", "varios", "n a", "na", "none", "ninguno",
             "sin motivo", "no especificado", "no definido", "desconocido", "unknown", "otra", "otras"}
RE_AVANZADA = re.compile(r"reunion|meeting|demo|propuesta|proposal|presupuesto|quote|cotizacion|negociacion|"
                         r"negotiation|contrato|contract|cierre")
RE_PROPUESTA = re.compile(r"propuesta|proposal|presupuesto|quote|cotizacion|oferta enviada|offer")

# --------------------------------------------------------------------------- auditoría


def auditar(cab, filas, mapeo, hoy, u, orden):
    avisos = []
    deals = []
    for f in filas:
        d = {
            "id": str(f.get(mapeo["id"], "") or "").strip() if mapeo["id"] else "",
            "nombre": str(f.get(mapeo["nombre"], "") or "").strip() if mapeo["nombre"] else "",
            "etapa": str(f.get(mapeo["etapa"], "") or "").strip() if mapeo["etapa"] else "",
            "comercial": str(f.get(mapeo["comercial"], "") or "").strip() if mapeo["comercial"] else "",
            "importe": a_importe(f.get(mapeo["importe"])) if mapeo["importe"] else None,
            "motivo": str(f.get(mapeo["motivo_perdida"], "") or "").strip() if mapeo["motivo_perdida"] else "",
            "estado": estado_de(f, mapeo),
        }
        for c in CAMPOS_FECHA:
            d[c] = a_fecha(f.get(mapeo[c]), orden) if mapeo[c] else None
        d["dias_etapa"] = None
        if mapeo["dias_etapa"]:
            n = a_importe(f.get(mapeo["dias_etapa"]))
            d["dias_etapa"] = int(n) if n is not None else None
        if not d["nombre"]:
            d["nombre"] = "(sin nombre)"
        deals.append(d)

    abiertas = [d for d in deals if d["estado"] == "abierta"]
    ganadas = [d for d in deals if d["estado"] == "ganada"]
    perdidas = [d for d in deals if d["estado"] == "perdida"]
    alertas = []

    def dias(f):
        return (hoy - f).days if f else None

    def ficha(d, dato):
        return {"id": d["id"], "nombre": d["nombre"], "comercial": d["comercial"] or "(sin asignar)", "etapa": d["etapa"],
                "importe": d["importe"], "dato": dato}

    def alerta(cat, prioridad, titulo, afectadas, extra=None, comercial=None):
        if not afectadas and not extra:
            return
        a = {"categoria": cat, "prioridad": prioridad, "titulo": titulo, "total_afectadas": len(afectadas),
             "importe_afectado": round(sum(x.get("importe") or 0 for x in afectadas), 2),
             "por_comercial": {}, "oportunidades": afectadas[:12]}
        for x in afectadas:
            a["por_comercial"][x["comercial"]] = a["por_comercial"].get(x["comercial"], 0) + 1
        if comercial:
            a["comercial"] = comercial
        if extra:
            a.update(extra)
        alertas.append(a)

    n_ab = len(abiertas) or 1
    tiene = {c: bool(mapeo[c]) and any(d[c] for d in deals) for c in CAMPOS_FECHA}
    tiene["dias_etapa"] = any(d["dias_etapa"] is not None for d in deals)

    # Velocidad: creada hace días y nadie la ha tocado
    if tiene["creado"] and tiene["ult_actividad"]:
        xs = [d for d in abiertas if d["creado"] and dias(d["creado"]) >= u["dias_sin_primer_contacto"]
              and (not d["ult_actividad"] or d["ult_actividad"] < d["creado"])]
        xs.sort(key=lambda d: d["creado"])
        p = "urgente" if len(xs) >= 10 or len(xs) / n_ab > .15 else "alta" if len(xs) >= 3 else "media"
        avanzadas = sum(bool(RE_AVANZADA.search(norm(d["etapa"]))) for d in xs)
        alerta("Velocidad", p, f"{len(xs)} oportunidades sin ninguna actividad registrada desde que se crearon "
               f"(hace más de {u['dias_sin_primer_contacto'] * 24} h)",
               [ficha(d, f"creada hace {dias(d['creado'])} d, sin actividad") for d in xs],
               extra={"en_etapa_avanzada": avanzadas})
    else:
        avisos.append("Sin fecha de creación o de última actividad: no he podido medir el primer contacto.")

    # Seguimiento: sin próxima tarea y tareas vencidas
    if tiene["prox_actividad"]:
        xs = [d for d in abiertas if not d["prox_actividad"]]
        pct = len(xs) / n_ab * 100
        alerta("Seguimiento", "alta" if pct > 30 else "media",
               f"{len(xs)} oportunidades abiertas sin próxima tarea ({pct:.0f} % de las abiertas)",
               [ficha(d, "sin próxima tarea") for d in xs])
        xs = [d for d in abiertas if d["prox_actividad"] and dias(d["prox_actividad"]) >= 1]
        xs.sort(key=lambda d: d["prox_actividad"])
        alerta("Seguimiento", "alta" if any(dias(d["prox_actividad"]) > 7 for d in xs) else "media",
               f"{len(xs)} oportunidades con la próxima tarea vencida",
               [ficha(d, f"vencida hace {dias(d['prox_actividad'])} d") for d in xs])
    else:
        avisos.append("Sin columna de próxima actividad: no he podido revisar tareas vencidas ni oportunidades sin siguiente paso.")

    # Seguimiento: propuestas enviadas que nadie persigue
    if tiene["ult_actividad"] and mapeo["etapa"]:
        xs = [d for d in abiertas if RE_PROPUESTA.search(norm(d["etapa"]))
              and (not d["ult_actividad"] or dias(d["ult_actividad"]) >= u["dias_propuesta_sin_seguimiento"])]
        xs.sort(key=lambda d: d["ult_actividad"] or date.min)
        alerta("Seguimiento", "alta",
               f"{len(xs)} propuestas sin seguimiento en {u['dias_propuesta_sin_seguimiento']} días",
               [ficha(d, f"última actividad hace {dias(d['ult_actividad'])} d" if d["ult_actividad"]
                      else "sin actividad registrada") for d in xs])

    # Actividad: oportunidades paradas y comerciales sin actividad
    if tiene["ult_actividad"]:
        xs = [d for d in abiertas if d["ult_actividad"] and dias(d["ult_actividad"]) >= u["dias_parada"]]
        xs.sort(key=lambda d: d["ult_actividad"])
        pct = len(xs) / n_ab * 100
        alerta("Actividad", "alta" if pct > 25 else "media",
               f"{len(xs)} oportunidades sin actividad en más de {u['dias_parada']} días",
               [ficha(d, f"última actividad hace {dias(d['ult_actividad'])} d") for d in xs])
        por = {}
        for d in abiertas:
            if d["comercial"]:
                por.setdefault(d["comercial"], []).append(d)
        for com, ds in por.items():
            if len(ds) < 3:
                continue
            ult = max((d["ult_actividad"] for d in ds if d["ult_actividad"]), default=None)
            if not ult or dias(ult) >= u["dias_comercial_sin_actividad"]:
                txt = f"sin actividad registrada en {dias(ult)} días" if ult else "sin ninguna actividad registrada"
                ds.sort(key=lambda d: -(d["importe"] or 0))
                alerta("Actividad", "alta", f"{com}: {txt} con {len(ds)} oportunidades abiertas",
                       [ficha(d, f"última actividad hace {dias(d['ult_actividad'])} d" if d["ult_actividad"]
                              else "sin actividad registrada") for d in ds], comercial=com)
    else:
        avisos.append("Sin columna de última actividad: no he podido detectar oportunidades paradas ni comerciales sin actividad.")

    # Carga: colas desequilibradas
    por = {}
    for d in abiertas:
        if d["comercial"]:
            por[d["comercial"]] = por.get(d["comercial"], 0) + 1
    if len(por) >= 3:
        med = statistics.median(por.values())
        for com, n in por.items():
            if med and n >= u["factor_cola"] * med and n - med >= 5:
                alerta("Carga", "media", f"{com} tiene {n} oportunidades abiertas, {n / med:.1f} veces la mediana del equipo ({med:g})",
                       [], extra={"total_afectadas": n, "mediana_equipo": med}, comercial=com)

    # Pipeline: cierre previsto ya pasado
    if tiene["cierre"]:
        xs = [d for d in abiertas if d["cierre"] and d["cierre"] < hoy]
        xs.sort(key=lambda d: d["cierre"])
        alerta("Pipeline", "alta" if len(xs) / n_ab > .2 else "media",
               f"{len(xs)} oportunidades abiertas con la fecha de cierre ya pasada",
               [ficha(d, f"cierre previsto hace {dias(d['cierre'])} d") for d in xs])

    # Pipeline: estancadas en etapa
    if tiene["cambio_etapa"] or tiene["dias_etapa"]:
        def en_etapa(d):
            return d["dias_etapa"] if d["dias_etapa"] is not None else dias(d["cambio_etapa"])
        xs = [d for d in abiertas if en_etapa(d) is not None and en_etapa(d) >= u["dias_en_etapa"]]
        xs.sort(key=lambda d: -en_etapa(d))
        alerta("Pipeline", "media", f"{len(xs)} oportunidades llevan más de {u['dias_en_etapa']} días en la misma etapa",
               [ficha(d, f"{en_etapa(d)} d en {d['etapa']}") for d in xs])
    elif tiene["creado"] and mapeo["etapa"]:
        viejas = {}
        for d in abiertas:
            if d["creado"] and dias(d["creado"]) >= u["dias_etapa_acumula"]:
                viejas.setdefault(d["etapa"] or "(sin etapa)", []).append(d)
        if viejas:
            etapa, ds = max(viejas.items(), key=lambda kv: len(kv[1]))
            if len(ds) >= 3:
                alerta("Pipeline", "media", f"La etapa «{etapa}» acumula {len(ds)} oportunidades creadas hace más de {u['dias_etapa_acumula']} días",
                       [ficha(d, f"creada hace {dias(d['creado'])} d") for d in sorted(ds, key=lambda d: d['creado'])])
        avisos.append("Sin fecha de cambio de etapa: el tiempo en etapa se ha aproximado con la antigüedad de la oportunidad.")

    # Datos: motivos de pérdida que no dicen nada
    if perdidas:
        if mapeo["motivo_perdida"]:
            motivos = {}
            for d in perdidas:
                motivos[d["motivo"] or "(vacío)"] = motivos.get(d["motivo"] or "(vacío)", 0) + 1
            inutiles = [d for d in perdidas if norm(d["motivo"]) in GENERICOS]
            pct = len(inutiles) / len(perdidas) * 100
            top = sorted(motivos.items(), key=lambda kv: -kv[1])[:6]
            if pct >= u["pct_motivo_vacio"]:
                alerta("Datos", "media", f"El {pct:.0f} % de las pérdidas no tiene un motivo útil (vacío u «Otros»)",
                       [ficha(d, d["motivo"] or "vacío") for d in inutiles],
                       extra={"reparto_motivos": [{"motivo": k, "n": v} for k, v in top], "perdidas": len(perdidas)})
        else:
            avisos.append("No hay columna de motivo de pérdida: no sé por qué se pierden las ventas.")

    # Datos: oportunidades sin comercial o sin importe
    xs = [d for d in abiertas if mapeo["comercial"] and not d["comercial"]]
    alerta("Datos", "alta", f"{len(xs)} oportunidades abiertas sin comercial asignado",
           [ficha(d, "sin comercial") for d in xs])
    if mapeo["importe"]:
        xs = [d for d in abiertas if not d["importe"]]
        pct = len(xs) / n_ab * 100
        if pct >= u["pct_sin_importe"]:
            alerta("Datos", "baja", f"{len(xs)} oportunidades abiertas sin importe ({pct:.0f} %): la previsión no es fiable",
                   [ficha(d, "sin importe") for d in xs])

    orden_p = {"urgente": 0, "alta": 1, "media": 2, "baja": 3}
    alertas.sort(key=lambda a: (orden_p[a["prioridad"]], -a["total_afectadas"]))
    for i, a in enumerate(alertas, 1):
        a["id"] = f"SAI-{i:03d}"

    # Resumen por comercial
    comerciales = {}
    for d in abiertas:
        c = d["comercial"] or "(sin asignar)"
        r = comerciales.setdefault(c, {"comercial": c, "abiertas": 0, "importe_abierto": 0.0, "sin_proxima_tarea": 0,
                                       "tareas_vencidas": 0, "paradas": 0, "sin_primer_contacto": 0, "ultima_actividad": None})
        r["abiertas"] += 1
        r["importe_abierto"] += d["importe"] or 0
        if tiene["prox_actividad"]:
            r["sin_proxima_tarea"] += not d["prox_actividad"]
            r["tareas_vencidas"] += bool(d["prox_actividad"] and dias(d["prox_actividad"]) >= 1)
        if d["ult_actividad"]:
            r["paradas"] += dias(d["ult_actividad"]) >= u["dias_parada"]
            if not r["ultima_actividad"] or d["ult_actividad"] > r["ultima_actividad"]:
                r["ultima_actividad"] = d["ult_actividad"]
        if d["creado"] and dias(d["creado"]) >= u["dias_sin_primer_contacto"] and tiene["ult_actividad"] and \
                (not d["ult_actividad"] or d["ult_actividad"] < d["creado"]):
            r["sin_primer_contacto"] += 1
    for r in comerciales.values():
        r["dias_desde_ultima_actividad"] = dias(r["ultima_actividad"]) if r["ultima_actividad"] else None
        r["ultima_actividad"] = r["ultima_actividad"].isoformat() if r["ultima_actividad"] else None
        r["importe_abierto"] = round(r["importe_abierto"], 2)

    # Tareas de hoy: lo que cada comercial debería hacer primero
    tareas = {}
    for d in abiertas:
        c = d["comercial"] or "(sin asignar)"
        t = tareas.setdefault(c, {"llamar_sin_primer_contacto": [], "tareas_vencidas": [], "propuestas_a_perseguir": []})
        if d["creado"] and tiene["ult_actividad"] and dias(d["creado"]) >= u["dias_sin_primer_contacto"] and \
                (not d["ult_actividad"] or d["ult_actividad"] < d["creado"]):
            t["llamar_sin_primer_contacto"].append(d["nombre"])
        if d["prox_actividad"] and dias(d["prox_actividad"]) >= 1:
            t["tareas_vencidas"].append(d["nombre"])
        if tiene["ult_actividad"] and RE_PROPUESTA.search(norm(d["etapa"])) and \
                (not d["ult_actividad"] or dias(d["ult_actividad"]) >= u["dias_propuesta_sin_seguimiento"]):
            t["propuestas_a_perseguir"].append(d["nombre"])
    tareas = {c: {k: {"n": len(v), "ejemplos": v[:5]} for k, v in t.items()} for c, t in tareas.items()
              if any(t.values())}

    fechas_ref = [d[c] for d in deals for c in ("creado", "ult_actividad", "cambio_etapa") if d[c]]
    resumen = {
        "filas": len(deals), "abiertas": len(abiertas), "ganadas": len(ganadas), "perdidas": len(perdidas),
        "borradas": sum(d["estado"] == "borrada" for d in deals),
        "importe_abierto": round(sum(d["importe"] or 0 for d in abiertas), 2),
        "tasa_cierre_pct": round(len(ganadas) / (len(ganadas) + len(perdidas)) * 100, 1) if ganadas or perdidas else None,
        "fecha_mas_reciente_en_datos": max((f for f in fechas_ref if f <= hoy), default=None),
        "alertas": len(alertas),
    }
    if resumen["fecha_mas_reciente_en_datos"]:
        lag = (hoy - resumen["fecha_mas_reciente_en_datos"]).days
        resumen["fecha_mas_reciente_en_datos"] = resumen["fecha_mas_reciente_en_datos"].isoformat()
        if lag > 30:
            avisos.append(f"La fecha más reciente del fichero es de hace {lag} días: puede ser una exportación antigua. "
                          "Si es así, vuelve a lanzar con --hoy igual a la fecha de exportación.")
    vistos = {}
    for d in abiertas:
        vistos[d["nombre"]] = vistos.get(d["nombre"], 0) + 1
    repes = {k: v for k, v in vistos.items() if v > 1 and k != "(sin nombre)"}
    if repes:
        top = ", ".join(f"{k} ({v})" for k, v in sorted(repes.items(), key=lambda kv: -kv[1])[:3])
        avisos.append(f"{len(repes)} nombres se repiten entre las abiertas ({sum(repes.values())} filas), por ejemplo {top}. "
                      "Pueden ser varias oportunidades del mismo cliente o duplicados que inflan el importe abierto"
                      + ("; usa el id para distinguirlas." if mapeo["id"] else "."))
    if not abiertas:
        avisos.append("No he encontrado oportunidades abiertas. Revisa que la columna de etapa o estado esté bien mapeada.")
    return resumen, alertas, sorted(comerciales.values(), key=lambda r: -r["abiertas"]), tareas, avisos


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("fichero")
    ap.add_argument("--hoy", help="fecha de referencia AAAA-MM-DD (por defecto, hoy)")
    ap.add_argument("--map", action="append", default=[], help="campo=Columna exacta, p. ej. comercial='Deal owner'")
    ap.add_argument("--umbral", action="append", default=[], help="clave=valor, p. ej. dias_parada=21")
    ap.add_argument("--fechas", choices=["dia", "mes"], help="fuerza el orden de las fechas numéricas")
    args = ap.parse_args()

    hoy = date.fromisoformat(args.hoy) if args.hoy else date.today()
    u = dict(UMBRALES)
    for kv in args.umbral:
        k, v = kv.split("=", 1)
        if k not in u:
            raise SystemExit(f"Umbral desconocido: {k}. Válidos: {list(u)}")
        u[k] = type(u[k])(float(v)) if isinstance(u[k], float) else int(float(v))
    forzado = {}
    for kv in args.map:
        k, v = kv.split("=", 1)
        if k not in CAMPOS:
            raise SystemExit(f"Campo desconocido: {k}. Válidos: {list(CAMPOS)}")
        forzado[k] = v

    cab, filas = leer(args.fichero)
    mapeo = mapear(cab, forzado)
    valores_fecha = [f.get(mapeo[c]) for c in CAMPOS_FECHA if mapeo[c] for f in filas]
    orden = args.fechas or pista_orden(valores_fecha)
    avisos_orden = []
    if not args.fechas and orden is None and any(RE_NUM.match(str(v).strip()) for v in valores_fecha if v):
        orden = "dia"
        avisos_orden.append("Las fechas numéricas son ambiguas (ningún día pasa de 12): las he leído como día/mes/año. "
                            "Si el CRM está en inglés de EE. UU., relanza con --fechas mes.")

    resumen, alertas, comerciales, tareas, avisos = auditar(cab, filas, mapeo, hoy, u, orden or "dia")
    faltan = [c for c in ("nombre", "etapa", "comercial") if not mapeo[c]]
    if faltan:
        avisos.insert(0, f"No he encontrado columna para: {', '.join(faltan)}. Usa --map campo='Columna' si existe con otro nombre.")
    salida = {
        "fichero": args.fichero, "hoy": hoy.isoformat(), "columnas": cab,
        "mapeo": mapeo, "orden_fechas": orden or "dia", "umbrales": u,
        "resumen": resumen, "alertas": alertas, "comerciales": comerciales,
        "tareas_de_hoy": tareas, "avisos": avisos_orden + avisos,
    }
    json.dump(salida, sys.stdout, ensure_ascii=False, indent=1, default=str)
    print()


if __name__ == "__main__":
    main()
