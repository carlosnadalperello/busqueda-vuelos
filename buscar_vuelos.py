#!/usr/bin/env python3
"""
Buscador diario de vuelos con aviso por email.

Lee una lista de búsquedas (busquedas.csv o busquedas.xlsx), consulta Google Flights
a través de SerpApi y envía un email cuando encuentra una combinación de ida y vuelta
por debajo del precio máximo por persona indicado.

Variables de entorno:
  SERPAPI_KEY          Clave de SerpApi (obligatoria salvo en --simular)
  SMTP_USER            Cuenta que envía el correo (p. ej. tu Gmail)
  SMTP_PASSWORD        Contraseña de aplicación de esa cuenta
  EMAIL_TO             Destinatario(s), separados por coma (por defecto SMTP_USER)
  SMTP_HOST / SMTP_PORT  Por defecto smtp.gmail.com / 465
  MAX_BUSQUEDAS_DIA    Consultas máximas por ejecución (por defecto 8)
  MAX_BUSQUEDAS_MES    Consultas máximas al mes (por defecto 240; el plan gratuito da 250)
  BAJADA_MIN_AVISO     % de bajada necesario para volver a avisar de la misma opción (5)
  AVISAR_ERRORES       1 para recibir email si la búsqueda falla (por defecto 1)
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import html
import json
import os
import smtplib
import sys
import time
from dataclasses import dataclass, field
from email.message import EmailMessage
from pathlib import Path

import requests

BASE = Path(__file__).resolve().parent
ESTADO = BASE / "estado.json"
HISTORIAL = BASE / "historial.csv"
API_URL = "https://serpapi.com/search"
DIAS = ["lun", "mar", "mié", "jue", "vie", "sáb", "dom"]


def fd(d: dt.date) -> str:
    return f"{DIAS[d.weekday()]} {d:%d/%m/%Y}"


# --------------------------------------------------------------------------- config
@dataclass
class Busqueda:
    nombre: str
    origen: str
    destino: str
    ida_desde: dt.date
    ida_hasta: dt.date
    vuelta_desde: dt.date | None
    vuelta_hasta: dt.date | None
    noches_min: int | None
    noches_max: int | None
    solo_directo: bool
    precio_max: float
    pasajeros: int = 1


def _fecha(v) -> dt.date | None:
    if v is None or str(v).strip() == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for f in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s, f).date()
        except ValueError:
            pass
    raise ValueError(f"Fecha no válida: {s!r} (usa DD/MM/AAAA)")


def _int(v) -> int | None:
    if v is None or str(v).strip() == "":
        return None
    return int(float(str(v).replace(",", ".")))


def _si(v) -> bool:
    return str(v).strip().lower() in ("si", "sí", "s", "yes", "y", "true", "1", "x")


def _codigos(v) -> str:
    cods = [c.strip().upper() for c in str(v).replace(";", ",").replace(" ", ",").split(",") if c.strip()]
    if not cods or any(len(c) != 3 or not c.isalpha() for c in cods):
        raise ValueError(f"Códigos de aeropuerto no válidos: {v!r} (usa p. ej. PMI o OTP,CLJ)")
    return ",".join(cods)


def leer_filas(ruta: Path) -> list[dict]:
    if ruta.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        ws = openpyxl.load_workbook(ruta, data_only=True).active
        filas = list(ws.iter_rows(values_only=True))
        cab = [str(c).strip().lower() if c is not None else "" for c in filas[0]]
        return [dict(zip(cab, f)) for f in filas[1:] if any(c not in (None, "") for c in f)]
    texto = ruta.read_text(encoding="utf-8-sig")
    sep = ";" if texto.splitlines()[0].count(";") >= texto.splitlines()[0].count(",") else ","
    lector = csv.DictReader(texto.splitlines(), delimiter=sep)
    return [{(k or "").strip().lower(): v for k, v in f.items()} for f in lector if any((v or "").strip() for v in f.values())]


def cargar_busquedas(ruta: Path) -> list[Busqueda]:
    res = []
    for i, f in enumerate(leer_filas(ruta), start=2):
        if "activa" in f and str(f["activa"]).strip() and not _si(f["activa"]):
            continue
        try:
            b = Busqueda(
                nombre=str(f.get("nombre") or f"Búsqueda fila {i}").strip(),
                origen=_codigos(f["origen"]),
                destino=_codigos(f["destino"]),
                ida_desde=_fecha(f["ida_desde"]),
                ida_hasta=_fecha(f.get("ida_hasta")) or _fecha(f["ida_desde"]),
                vuelta_desde=_fecha(f.get("vuelta_desde")),
                vuelta_hasta=_fecha(f.get("vuelta_hasta")),
                noches_min=_int(f.get("noches_min")),
                noches_max=_int(f.get("noches_max")),
                solo_directo=_si(f.get("solo_directo", "no")),
                precio_max=float(str(f["precio_max"]).replace("€", "").replace(",", ".").strip()),
                pasajeros=_int(f.get("pasajeros")) or 1,
            )
        except (KeyError, ValueError) as e:
            raise SystemExit(f"Error en la fila {i} de {ruta.name}: {e}")
        if b.noches_min is None and b.noches_max is not None:
            b.noches_min = 1
        if b.noches_max is None and b.noches_min is not None:
            b.noches_max = b.noches_min
        if b.vuelta_desde is None and b.noches_min is None:
            raise SystemExit(f"Fila {i}: indica el rango de vuelta o el de noches.")
        res.append(b)
    return res


def combinaciones(b: Busqueda, hoy: dt.date) -> list[tuple[dt.date, dt.date]]:
    """Pares (ida, vuelta) que cumplen rangos de fechas y de noches, sin fechas pasadas."""
    pares = []
    d = max(b.ida_desde, hoy)
    while d <= b.ida_hasta:
        if b.noches_min is not None:
            vueltas = [d + dt.timedelta(days=n) for n in range(b.noches_min, b.noches_max + 1)]
        else:
            v, vueltas = max(b.vuelta_desde, d + dt.timedelta(days=1)), []
            while v <= b.vuelta_hasta:
                vueltas.append(v)
                v += dt.timedelta(days=1)
        for v in vueltas:
            if b.vuelta_desde and v < b.vuelta_desde:
                continue
            if b.vuelta_hasta and v > b.vuelta_hasta:
                continue
            pares.append((d, v))
        d += dt.timedelta(days=1)
    return pares


# --------------------------------------------------------------------------- estado
def cargar_estado() -> dict:
    if ESTADO.exists():
        return json.loads(ESTADO.read_text(encoding="utf-8"))
    return {"cursor": 0, "uso": {}, "avisados": {}, "error_avisado": None}


def guardar_estado(e: dict) -> None:
    ESTADO.write_text(json.dumps(e, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")


def anotar_historial(filas: list[list]) -> None:
    nuevo = not HISTORIAL.exists()
    with HISTORIAL.open("a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh, delimiter=";")
        if nuevo:
            w.writerow(["consulta", "busqueda", "ida", "vuelta", "noches", "precio_persona",
                        "aerolinea", "aeropuertos", "directo", "bajo_umbral"])
        w.writerows(filas)


# --------------------------------------------------------------------------- API
class Cuota(Exception):
    pass


class ClienteSerpApi:
    def __init__(self, clave: str, estado: dict, max_dia: int, max_mes: int, simulador=None):
        self.clave, self.estado, self.max_dia, self.max_mes = clave, estado, max_dia, max_mes
        self.hechas = 0
        self.simulador = simulador
        self.mes = dt.date.today().strftime("%Y-%m")

    def uso_mes(self) -> int:
        return self.estado["uso"].get(self.mes, 0)

    def queda(self) -> int:
        return min(self.max_dia - self.hechas, self.max_mes - self.uso_mes())

    def buscar(self, params: dict) -> dict:
        if self.queda() <= 0:
            raise Cuota()
        p = {"engine": "google_flights", "currency": "EUR", "hl": "es", "gl": "es", **params}
        if self.simulador:
            datos = self.simulador(p)
        else:
            p["api_key"] = self.clave
            r = requests.get(API_URL, params=p, timeout=90)
            datos = r.json() if r.headers.get("content-type", "").startswith("application/json") else {}
            if r.status_code >= 400 and "error" not in datos:
                datos["error"] = f"HTTP {r.status_code}"
        self.hechas += 1
        err = datos.get("error", "")
        if err and "hasn't returned any results" not in err and "no results" not in err.lower():
            raise RuntimeError(f"SerpApi: {err}")  # SerpApi no cobra las búsquedas con error
        self.estado["uso"][self.mes] = self.uso_mes() + 1
        return datos


# --------------------------------------------------------------------------- lógica
@dataclass
class Opcion:
    busqueda: Busqueda
    ida: dt.date
    vuelta: dt.date
    precio_persona: float
    ida_tramos: list
    vuelta_tramos: list = field(default_factory=list)
    enlace: str = ""
    nivel_precio: str = ""
    rango_tipico: list | None = None

    @property
    def noches(self) -> int:
        return (self.vuelta - self.ida).days

    @property
    def aerolineas(self) -> str:
        nombres = []
        for t in self.ida_tramos + self.vuelta_tramos:
            if t.get("airline") and t["airline"] not in nombres:
                nombres.append(t["airline"])
        return " / ".join(nombres)

    @property
    def ruta(self) -> str:
        ida = f"{self.ida_tramos[0]['departure_airport']['id']}→{self.ida_tramos[-1]['arrival_airport']['id']}"
        if self.vuelta_tramos:
            return ida + f", {self.vuelta_tramos[0]['departure_airport']['id']}→{self.vuelta_tramos[-1]['arrival_airport']['id']}"
        return ida

    @property
    def clave(self) -> str:
        return f"{self.busqueda.nombre}|{self.ida}|{self.vuelta}|{self.ruta}|{self.aerolineas}"


def _opciones(datos: dict) -> list[dict]:
    return (datos.get("best_flights") or []) + (datos.get("other_flights") or [])


def _tramo_txt(tramos: list) -> str:
    if not tramos:
        return "(detalle de vuelta no disponible; míralo en el enlace)"
    a, z = tramos[0], tramos[-1]
    num = ", ".join(t.get("flight_number", "") for t in tramos)
    esc = "directo" if len(tramos) == 1 else f"{len(tramos) - 1} escala(s)"
    return (f"{a['departure_airport']['id']} {a['departure_airport']['time'][-5:]} → "
            f"{z['arrival_airport']['id']} {z['arrival_airport']['time'][-5:]} "
            f"({a.get('airline', '')} {num}, {esc})")


def evaluar(b: Busqueda, ida: dt.date, vuelta: dt.date, api: ClienteSerpApi, consulta: str,
            historial: list) -> Opcion | None:
    params = {"type": "1", "departure_id": b.origen, "arrival_id": b.destino,
              "outbound_date": ida.isoformat(), "return_date": vuelta.isoformat(),
              "adults": str(b.pasajeros), "sort_by": "2"}
    if b.solo_directo:
        params["stops"] = "1"
    datos = api.buscar(params)
    candidatas = [o for o in _opciones(datos) if o.get("price")]
    if b.solo_directo:
        candidatas = [o for o in candidatas if len(o.get("flights", [])) == 1]
    if not candidatas:
        historial.append([consulta, b.nombre, ida, vuelta, (vuelta - ida).days, "", "", "", "", "sin vuelos"])
        return None
    mejor = min(candidatas, key=lambda o: o["price"])
    pp = round(mejor["price"] / b.pasajeros, 2)
    ins = datos.get("price_insights") or {}
    op = Opcion(b, ida, vuelta, pp, mejor["flights"],
                enlace=(datos.get("search_metadata") or {}).get("google_flights_url", ""),
                nivel_precio=ins.get("price_level", ""), rango_tipico=ins.get("typical_price_range"))
    bajo = pp <= b.precio_max
    if bajo and mejor.get("departure_token") and api.queda() > 0:
        try:
            vuelta_datos = api.buscar({**params, "departure_token": mejor["departure_token"]})
            regresos = [o for o in _opciones(vuelta_datos) if o.get("flights")]
            if b.solo_directo:
                regresos = [o for o in regresos if len(o["flights"]) == 1]
            if regresos:
                r = min(regresos, key=lambda o: o.get("price") or 10**9)
                op.vuelta_tramos = r["flights"]
                if r.get("price"):
                    op.precio_persona = round(r["price"] / b.pasajeros, 2)
            elif b.solo_directo:
                bajo = False  # no hay vuelta directa compatible
        except Cuota:
            pass
    bajo = bajo and op.precio_persona <= b.precio_max
    historial.append([consulta, b.nombre, ida, vuelta, op.noches, op.precio_persona, op.aerolineas,
                      op.ruta, "sí" if b.solo_directo else "no", "SÍ" if bajo else ""])
    return op if bajo else None


def debe_avisar(op: Opcion, estado: dict, bajada_pct: float) -> bool:
    prev = estado["avisados"].get(op.clave)
    if prev is None:
        return True
    return op.precio_persona <= prev["precio"] * (1 - bajada_pct / 100)


# --------------------------------------------------------------------------- email
def enviar_email(asunto: str, texto: str, html_cuerpo: str, simular: bool = False) -> None:
    usuario = os.environ.get("SMTP_USER", "")
    destino = os.environ.get("EMAIL_TO") or usuario
    if simular or not usuario:
        print(f"\n===== EMAIL (no enviado{' - simulación' if simular else ': falta SMTP_USER'}) =====")
        print("Asunto:", asunto)
        print(texto)
        return
    msg = EmailMessage()
    msg["Subject"], msg["From"], msg["To"] = asunto, usuario, destino
    msg.set_content(texto)
    msg.add_alternative(html_cuerpo, subtype="html")
    host = os.environ.get("SMTP_HOST", "smtp.gmail.com")
    puerto = int(os.environ.get("SMTP_PORT", "465"))
    if puerto == 465:
        with smtplib.SMTP_SSL(host, puerto, timeout=60) as s:
            s.login(usuario, os.environ["SMTP_PASSWORD"])
            s.send_message(msg)
    else:
        with smtplib.SMTP(host, puerto, timeout=60) as s:
            s.starttls()
            s.login(usuario, os.environ["SMTP_PASSWORD"])
            s.send_message(msg)
    print(f"Email enviado a {destino}: {asunto}")


def email_ofertas(ops: list[Opcion]) -> tuple[str, str, str]:
    ops = sorted(ops, key=lambda o: o.precio_persona)
    top = ops[0]
    asunto = (f"✈️ {top.busqueda.nombre}: {top.precio_persona:.0f} €/pers "
              f"({top.ida:%d/%m}-{top.vuelta:%d/%m})" + (f" y {len(ops) - 1} más" if len(ops) > 1 else ""))
    lineas, bloques = [], []
    for o in ops:
        tipico = (f"; rango típico {o.rango_tipico[0]}-{o.rango_tipico[1]} €" if o.rango_tipico else "")
        nivel = f"Nivel de precio según Google: {o.nivel_precio}{tipico}" if o.nivel_precio else ""
        lineas += [f"• {o.busqueda.nombre} — {o.precio_persona:.0f} € por persona ida y vuelta "
                   f"(máx. {o.busqueda.precio_max:.0f} €)",
                   f"  Ida {fd(o.ida)}: {_tramo_txt(o.ida_tramos)}",
                   f"  Vuelta {fd(o.vuelta)}: {_tramo_txt(o.vuelta_tramos)}",
                   f"  {o.noches} noches · {o.aerolineas}" + (f" · {nivel}" if nivel else ""),
                   f"  {o.enlace}", ""]
        bloques.append(
            f"<li><b>{html.escape(o.busqueda.nombre)}</b> — <b>{o.precio_persona:.0f} €</b> por persona ida y vuelta "
            f"(máx. {o.busqueda.precio_max:.0f} €)<br>Ida {fd(o.ida)}: {html.escape(_tramo_txt(o.ida_tramos))}"
            f"<br>Vuelta {fd(o.vuelta)}: {html.escape(_tramo_txt(o.vuelta_tramos))}"
            f"<br>{o.noches} noches · {html.escape(o.aerolineas)}"
            + (f"<br><small>{html.escape(nivel)}</small>" if nivel else "")
            + (f"<br><a href=\"{html.escape(o.enlace)}\">Ver en Google Flights</a>" if o.enlace else "") + "</li>")
    pie = ("Precio de Google Flights (vía SerpApi) en el momento de la consulta, tarifa básica; "
           "la maleta de cabina grande o facturada suele costar aparte. Compruébalo antes de comprar.")
    texto = "\n".join(lineas) + pie
    cuerpo = f"<ul>{''.join(bloques)}</ul><p><small>{pie}</small></p>"
    return asunto, texto, cuerpo


# --------------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", help="Ruta a busquedas.csv o busquedas.xlsx")
    ap.add_argument("--listar", action="store_true", help="Solo muestra las combinaciones, sin consultar")
    ap.add_argument("--simular", action="store_true", help="Usa respuestas falsas y no envía email")
    args = ap.parse_args()

    ruta = Path(args.config) if args.config else next(
        (BASE / n for n in ("busquedas.xlsx", "busquedas.csv") if (BASE / n).exists()), BASE / "busquedas.csv")
    hoy = dt.date.today()
    busquedas = cargar_busquedas(ruta)
    trabajo = [(b, i, v) for b in busquedas for (i, v) in combinaciones(b, hoy)]
    print(f"{len(busquedas)} búsqueda(s) activas, {len(trabajo)} combinaciones de fechas pendientes.")
    if args.listar:
        for b, i, v in trabajo:
            print(f"  {b.nombre}: {b.origen}→{b.destino} {fd(i)} - {fd(v)} ({(v - i).days} n)")
        return 0
    if not trabajo:
        print("Nada que buscar (fechas pasadas o búsquedas inactivas).")
        return 0

    estado = cargar_estado()
    simulador = None
    if args.simular:
        from tests.simulador import responder as simulador  # noqa: E402
    clave = os.environ.get("SERPAPI_KEY", "")
    if not clave and not simulador:
        raise SystemExit("Falta la variable SERPAPI_KEY.")
    api = ClienteSerpApi(clave, estado, int(os.environ.get("MAX_BUSQUEDAS_DIA", "8")),
                         int(os.environ.get("MAX_BUSQUEDAS_MES", "240")), simulador)
    bajada = float(os.environ.get("BAJADA_MIN_AVISO", "5"))

    # Recorre las combinaciones en rueda: cada día sigue donde lo dejó el anterior.
    cursor = estado.get("cursor", 0) % len(trabajo)
    consulta = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    historial, ofertas, error = [], [], None
    revisadas = 0
    try:
        # Cada combinación puede necesitar 2 llamadas (ida + detalle de la vuelta).
        while revisadas < len(trabajo) and api.queda() >= 2:
            b, i, v = trabajo[(cursor + revisadas) % len(trabajo)]
            op = evaluar(b, i, v, api, consulta, historial)
            revisadas += 1
            if op and debe_avisar(op, estado, bajada):
                ofertas.append(op)
            time.sleep(0 if simulador else 1)
    except Cuota:
        pass
    except Exception as e:  # noqa: BLE001
        error = str(e)
        print("ERROR:", error, file=sys.stderr)
    estado["cursor"] = (cursor + revisadas) % len(trabajo)

    anotar_historial(historial)
    print(f"Consultadas {revisadas} combinaciones ({api.hechas} llamadas; uso del mes {api.uso_mes()}).")

    if ofertas:
        enviar_email(*email_ofertas(ofertas), simular=args.simular)
        for o in ofertas:
            estado["avisados"][o.clave] = {"precio": o.precio_persona, "fecha": consulta}
    else:
        print("Sin ofertas nuevas por debajo del umbral.")

    if error and os.environ.get("AVISAR_ERRORES", "1") == "1" and estado.get("error_avisado") != hoy.isoformat():
        enviar_email("⚠️ El buscador de vuelos ha fallado",
                     f"La búsqueda diaria ha fallado con este error:\n\n{error}\n\n"
                     "Revisa la clave de SerpApi y tu cupo mensual.",
                     f"<p>La búsqueda diaria ha fallado:</p><pre>{html.escape(error)}</pre>",
                     simular=args.simular)
        estado["error_avisado"] = hoy.isoformat()
    guardar_estado(estado)
    return 1 if error else 0


if __name__ == "__main__":
    sys.exit(main())
