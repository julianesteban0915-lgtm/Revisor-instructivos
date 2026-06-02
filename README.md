# Revisor de Instructivos — Garces Fruit

App web que revisa los kits de un instructivo de embalaje (PDF, formato
GF-IND-PL-003) contra el kit base maestro. Ignora las columnas
**Observación** y **Pallet**.

El usuario final solo abre el link, sube el PDF y ve el resultado
(OK / Error / Advertencia) con opción de descargar a Excel.

## Archivos

- `app.py` — la aplicación.
- `requirements.txt` — librerías necesarias.
- `KIT_EMBALAJE_TODOS.xlsx` — kit base maestro (debes subir el tuyo aquí).

El kit base se incluye en el repo y se carga automáticamente. Cuando
cambie, basta con reemplazar ese archivo (o subir uno nuevo desde la
propia app, en "⚙️ Kit base maestro").

## Cómo publicarla (gratis, sin servidores)

1. Crea una cuenta en GitHub (si no tienes) y un repositorio nuevo.
2. Sube estos tres archivos al repositorio:
   `app.py`, `requirements.txt` y tu `KIT_EMBALAJE_TODOS.xlsx`.
3. Entra a https://share.streamlit.io e inicia sesión con GitHub.
4. "Create app" → elige tu repositorio, la rama, y `app.py` como archivo principal.
5. "Deploy". En 1–2 minutos tendrás un link público para compartir.

## Probarla en tu PC primero (opcional)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abre en el navegador en http://localhost:8501

## Lógica de revisión

- Compara cada kit por **Envase + Embalaje + Etiqueta**.
- Normaliza los códigos (toma solo el código: `NAGR NACIONAL` → `NAGR`).
- **Error**: la combinación no existe en el kit base (sugiere etiquetas válidas).
- **Advertencia**: el kit existe pero las cajas/pallet no coinciden.
- **OK**: todo cuadra.
