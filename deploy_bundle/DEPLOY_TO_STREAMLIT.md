# Subir la app a Streamlit Community Cloud

## Antes de empezar

Necesitas:

- una cuenta de GitHub
- una cuenta de Streamlit Community Cloud

## Opcion simple sin Git instalado

1. En GitHub, crea un repositorio nuevo.
2. Entra al repositorio y usa `Add file` -> `Upload files`.
3. Sube estos archivos y carpetas:
   - `generate_plan.py`
   - `streamlit_app.py`
   - `requirements.txt`
   - `README.md`
   - `templates/` si la quieres guardar en el repo
   - `static/` si la quieres guardar en el repo
4. No subas:
   - `output/`
   - `web_uploads/`
   - `__pycache__/`
   - `drive_config.json`
5. Crea el commit desde GitHub.

## Publicarlo en Streamlit

1. Entra a [share.streamlit.io](https://share.streamlit.io/).
2. Conecta tu cuenta de GitHub.
3. Haz clic en `Create app`.
4. Elige tu repositorio.
5. Selecciona la rama principal.
6. En `Main file path` escribe:
   `streamlit_app.py`
7. Opcional: elige un subdominio facil de recordar.
8. En `Advanced settings`, si quieres, define la version de Python.
9. Haz clic en `Deploy`.

## Despues

- Cada cambio nuevo que subas a GitHub actualiza la app.
- Si una libreria falla en la nube, actualiza `requirements.txt` y vuelve a desplegar.
