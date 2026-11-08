# Buscador diario de vuelos

Busca cada día las combinaciones de ida y vuelta de `busquedas.csv` en Google Flights (a través de SerpApi) y te manda un email cuando alguna baja del precio máximo por persona.

## 1. El archivo de búsquedas

Se abre con Excel. Va una fila por búsqueda, separada por `;`. Si lo prefieres, guárdalo como `busquedas.xlsx` con las mismas columnas: si existe, se usa ese en lugar del CSV.

| Columna | Qué poner | Ejemplo |
|---|---|---|
| nombre | Nombre que verás en el email | Rumanía Semana Santa 2027 |
| activa | Sí / No | Sí |
| origen | Código IATA, o varios separados por coma | PMI |
| destino | Código IATA, o varios separados por coma | OTP,CLJ,SBZ |
| ida_desde / ida_hasta | Rango de fechas de salida (DD/MM/AAAA) | 22/03/2027 – 27/03/2027 |
| vuelta_desde / vuelta_hasta | Rango de fechas de vuelta (opcional si indicas las noches) | 28/03/2027 – 02/04/2027 |
| noches_min / noches_max | Duración de la estancia (opcional si indicas el rango de vuelta) | 5 – 6 |
| solo_directo | Sí / No | Sí |
| precio_max | Máximo en € por persona, ida + vuelta | 250 |
| pasajeros | Adultos (el precio se compara por persona) | 1 |
| notas | Libre, no se usa | |

Si pones varios destinos, en cada consulta se busca en todos a la vez. La ida y la vuelta salen siempre del mismo aeropuerto: no se cubren combinaciones como llegar a Cluj y volver desde Bucarest.

Para comprobar qué fechas se van a buscar, sin gastar consultas: `python buscar_vuelos.py --listar`

## 2. Cuentas necesarias (una sola vez)

1. **SerpApi**: regístrate en serpapi.com y copia tu *API key*. El plan gratuito da 250 búsquedas al mes. El script se limita a 8 consultas al día y 240 al mes, y cada día continúa por donde lo dejó el anterior. Una combinación que baja del umbral gasta 2 consultas, porque también se pide el detalle de la vuelta.
2. **Gmail**: activa la verificación en dos pasos y crea una *contraseña de aplicación* (Cuenta de Google → Seguridad → Contraseñas de aplicaciones). Úsala en lugar de tu contraseña normal.

## 3. Ponerlo en marcha en GitHub (gratis, sin dejar el ordenador encendido)

1. Crea un repositorio **privado** en GitHub y sube todos estos archivos, incluida la carpeta `.github`.
2. En el repositorio, ve a *Settings → Secrets and variables → Actions → New repository secret* y crea:
   - `SERPAPI_KEY`: tu clave de SerpApi
   - `SMTP_USER`: tu Gmail
   - `SMTP_PASSWORD`: la contraseña de aplicación
   - `EMAIL_TO`: a quién avisar (puede ser el mismo Gmail; para varios, sepáralos con comas)
3. En la pestaña *Actions*, abre «Buscar vuelos» y pulsa **Run workflow** para probarlo. Después se ejecuta solo cada día a las 06:17 UTC.
4. Para cambiar las búsquedas, edita `busquedas.csv` en GitHub (o súbelo de nuevo desde Excel).

## Cómo avisa

- Envía un email con las opciones que bajan del umbral: aerolínea, aeropuertos, horas, noches, precio por persona, el nivel de precio según Google y un enlace.
- No repite el aviso de la misma opción salvo que el precio baje al menos un 5 % más (`BAJADA_MIN_AVISO`).
- Si la búsqueda falla (clave incorrecta, cupo agotado…) envía un aviso, como mucho uno al día.
- `historial.csv` guarda el precio más bajo de cada consulta, para ver cómo evoluciona.

Los precios son de tarifa básica. La maleta de cabina grande o facturada suele ir aparte, así que confirma siempre en la web de la aerolínea antes de comprar.

## Ejecutarlo en tu ordenador (opcional)

```
pip install -r requirements.txt
export SERPAPI_KEY=... SMTP_USER=... SMTP_PASSWORD=... EMAIL_TO=...
python buscar_vuelos.py            # búsqueda real
python buscar_vuelos.py --simular  # prueba con datos falsos, sin email
```
