# Fast Audit

Skill gratuita para Claude que audita la exportación de tu CRM y te dice qué se está quedando por hacer: leads sin primer contacto, tareas vencidas, propuestas que nadie persigue, comerciales sin actividad, cierres ya pasados y pérdidas sin motivo. Devuelve las alertas ordenadas por prioridad, qué harías con cada una y las tareas de hoy de cada comercial.

Funciona con exportaciones de HubSpot, Pipedrive, Zoho, Salesforce o cualquier CRM que exporte a CSV o Excel.

Hecha por [Sailor](https://labs.agenciasnowball.com/sailor).

## Qué hay dentro

Tres ficheros, todos legibles:

| Fichero | Qué es |
|---|---|
| `fast-audit/SKILL.md` | Las instrucciones que sigue Claude: cómo leer el resultado y cómo escribir el informe |
| `fast-audit/scripts/auditar.py` | El script que cuenta. Python con la biblioteca estándar, sin dependencias |
| `fast-audit/references/exportar-csv.md` | Cómo exportar las oportunidades de cada CRM |

**No se conecta a nada.** El script lee el fichero que le das, calcula y devuelve un JSON. No hace peticiones de red ni guarda nada fuera de tu conversación con Claude. Puedes comprobarlo leyendo `auditar.py`: no importa ninguna librería de red.

## Instalación en Claude

1. Descarga **`fast-audit.zip`** desde [Releases](../../releases/latest). No uses el botón "Code › Download ZIP": mete la carpeta con otro nombre y Claude la rechaza.
2. En Claude, ve a **Settings › Capabilities** y activa **Code execution and file creation**.
3. Ve a **Customize › Skills**, pulsa **+**, luego **Create skill** y **Upload a skill**, y sube el `.zip` tal cual.

Funciona en todos los planes de Claude. En Team y Enterprise, la ejecución de código la activa quien administra la cuenta.

## Uso

1. Exporta las oportunidades de tu CRM (abiertas, ganadas y perdidas de los últimos 12 meses). Los pasos de cada CRM están en `fast-audit/references/exportar-csv.md`.
2. Abre un chat nuevo en Claude, arrastra el fichero y pide algo como: *«Audita este pipeline. ¿Qué se nos está escapando?»*

Si tu ciclo de venta es largo, pide a Claude que cambie los límites (por ejemplo, «considera parada a partir de 30 días»).

## Uso del script sin Claude

```bash
python3 fast-audit/scripts/auditar.py oportunidades.csv --hoy 2026-10-22
```

Opciones: `--map campo='Columna'` para asignar columnas a mano, `--fechas mes` para fechas de EE. UU., `--umbral clave=valor` para cambiar límites. `python3 fast-audit/scripts/auditar.py --help` las lista todas.

## Licencia

MIT. Úsala, cámbiala y compártela citando la fuente.
