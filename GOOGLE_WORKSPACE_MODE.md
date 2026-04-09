# Modo Pro: Google Workspace

Este archivo define la base para el flujo multiusuario con Google.

## Objetivo

Permitir que cada persona del equipo conecte su propia cuenta de Google y use:

- Google Drive
- Google Sheets
- Google Slides

sin depender de descargar un `.pptx` para corregir datos o editar el reporte final.

## Enfoque recomendado

1. El usuario inicia sesion con Google.
2. La app obtiene permisos para Drive, Sheets y Slides.
3. El usuario elige una Sheet como fuente de datos.
4. La app crea o actualiza una presentacion en Google Slides.
5. La presentacion queda guardada y editable en el Drive del usuario.

## Estado actual del proyecto

Ya existe una base local en [google_workspace.py](./google_workspace.py) con:

- carga de configuracion
- manejo de tokens por usuario
- autorizacion OAuth
- cliente para Drive / Sheets / Slides
- lectura de filas desde Google Sheets
- creacion de una presentacion en Google Slides
- creacion de una slide de titulo basica

## Credenciales necesarias

1. Crear un proyecto en Google Cloud.
2. Habilitar estas APIs:
   - Google Drive API
   - Google Sheets API
   - Google Slides API
3. Crear credenciales OAuth para la app.
4. Descargar el archivo `client_secret.json`.
5. Crear `google_workspace_config.json` a partir de `google_workspace_config.example.json`.

## Alcance inicial de la primera version

La primera version del modo pro deberia cubrir:

- login por usuario
- seleccion de spreadsheet
- lectura de una hoja estructurada
- creacion de una presentacion en Google Slides
- exportacion opcional a PPTX o PDF

## Estructura recomendada de la Sheet

Hoja `Datos`:

- `fecha`
- `cliente`
- `medio`
- `tipo_medio`
- `tier`
- `tipo_comunicado`
- `valor_estimado`
- `alcance_estimado`

Hoja `Resumen`:

- `comentario_ejecutivo`
- `pasos_siguientes`
- `mes_reporte`
- `titulo_reporte`

## Siguiente implementacion sugerida

1. Agregar un modo `Google Workspace (beta)` en la app.
2. Incorporar el flujo de autorizacion por usuario.
3. Leer una Sheet y transformarla a la estructura del reporte.
4. Crear la presentacion directamente en Google Slides.
