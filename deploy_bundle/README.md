# Generador de Planes Estrategicos en PPTX

Esta carpeta convierte un documento de briefing en una presentacion `.pptx` y tambien incluye una interfaz web para que el equipo lo use desde navegador.

## Uso simple local

1. Ejecuta `install_dependencies.bat` una sola vez.
2. Deja el documento de entrada en la carpeta `input`.
3. Ejecuta `run_generator.bat`.
4. El archivo final aparecera en `output`.

## Uso como app web local

1. Ejecuta `install_dependencies.bat`.
2. Ejecuta `start_web.bat`.
3. Abre `http://127.0.0.1:8000` en el navegador.
4. Sube el briefing y descarga el resultado.

## Uso como app simple tipo "sube y descarga"

Esta es la opcion recomendada para una experiencia mas parecida a Kimi:

1. Ejecuta `install_dependencies.bat`.
2. En una terminal, entra a la carpeta del proyecto.
3. Ejecuta:
   `C:\Users\dsala\AppData\Local\Python\bin\python.exe -m streamlit run streamlit_app.py`
4. Se abrira una app con subida de archivo y descarga directa del PPTX.

## Uso con Google Drive compartido

Esta es la forma mas simple para empezar sin pagar hosting.

1. Crea una carpeta compartida en Drive, por ejemplo `GeneradorPPT`.
2. Dentro de esa carpeta crea:
   - `incoming`
   - `processed`
   - `output`
   - `error`
3. Copia `drive_config.example.json` como `drive_config.json`.
4. Edita `drive_config.json` con las rutas locales sincronizadas por Google Drive Desktop.
5. Ejecuta `start_drive_mode.bat` en un computador que quede encendido durante el horario de uso.
6. El equipo solo tiene que subir briefings a `incoming`.
7. El sistema mueve el briefing a `processed` y deja el `.pptx` y el resumen en `output`.

### Que hace cada carpeta

- `incoming`: donde el equipo deja el briefing
- `processed`: archivos ya procesados
- `output`: presentaciones y resumenes generados
- `error`: archivos que fallaron y requieren revision

## Formatos soportados

- `.txt`
- `.docx`
- `.pdf`

## Salidas generadas

- Presentacion `.pptx`
- Resumen `.txt` con secciones completas y campos pendientes

## Para subirla online

- Ya esta incluido `render.yaml` para desplegar en Render.
- El backend principal es `web_app.py`.
- Si mas adelante quieres, se puede conectar con Google Drive para guardar automaticamente briefing y entregables.
- Para una experiencia gratis mas simple, recomiendo subir `streamlit_app.py` a Streamlit Community Cloud.
