"""Respuestas falsas de SerpApi para probar el buscador sin gastar consultas."""
import datetime as dt

LLAMADAS = []


def _vuelo(origen, destino, fecha, hora_s, hora_l, aerolinea, numero):
    return {"departure_airport": {"id": origen, "name": origen, "time": f"{fecha} {hora_s}"},
            "arrival_airport": {"id": destino, "name": destino, "time": f"{fecha} {hora_l}"},
            "airline": aerolinea, "flight_number": numero, "duration": 185}


def responder(p: dict) -> dict:
    LLAMADAS.append(p)
    ida, vuelta = p["outbound_date"], p["return_date"]
    meta = {"search_metadata": {"google_flights_url": f"https://www.google.com/travel/flights?test={ida}"}}
    if "departure_token" in p:  # vuelos de vuelta
        precio = int(p["departure_token"].split(":")[1])
        return {**meta, "best_flights": [
            {"flights": [_vuelo("OTP", "PMI", vuelta, "19:50", "22:10", "Wizz Air Malta", "W4 3025")],
             "price": precio}]}
    d = dt.date.fromisoformat(ida)
    # Precio barato solo saliendo el miércoles 24/03; resto caro; el 23/03 solo con escala.
    if d.day == 23:
        return {**meta, "other_flights": [{"flights": [
            _vuelo("PMI", "BCN", ida, "07:00", "08:00", "Vueling", "VY 1"),
            _vuelo("BCN", "OTP", ida, "10:00", "14:00", "Wizz Air", "W6 2")], "price": 90,
            "departure_token": "t:90"}]}
    precio = 189 if d.day == 24 else 310
    precio *= int(p.get("adults", 1))
    return {**meta, "price_insights": {"price_level": "low", "typical_price_range": [220, 340]},
            "best_flights": [{"flights": [_vuelo("PMI", "CLJ" if precio < 300 else "OTP", ida, "22:30", "02:30",
                                                  "Wizz Air Malta", "W4 3420")],
                              "price": precio, "departure_token": f"t:{precio}"}]}
