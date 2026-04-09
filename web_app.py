from __future__ import annotations

import shutil
from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from generate_plan import OUTPUT_DIR, generate_from_file
from report_generator import generate_report_from_pdf


BASE_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = BASE_DIR / "templates"
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR = BASE_DIR / "web_uploads"

UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Generador de Planes Estrategicos")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "message": None,
            "results": [],
        },
    )


@app.post("/generate", response_class=HTMLResponse)
async def generate(
    request: Request,
    briefing: UploadFile = File(...),
    generator_type: str = Form("plan"),
    mode: str = Form("auto"),
    report_title: str = Form("Reporte Automatico Mensual"),
    client_name: str = Form(""),
    report_month: str = Form(""),
    executive_comment: str = Form(""),
    next_steps: str = Form(""),
    background: UploadFile | None = File(None),
    logo: UploadFile | None = File(None),
):
    suffix = Path(briefing.filename or "").suffix.lower()
    if suffix not in {".txt", ".docx", ".pdf"}:
        raise HTTPException(status_code=400, detail="Formato no soportado. Usa .txt, .docx o .pdf.")

    token = uuid4().hex
    upload_path = UPLOADS_DIR / f"{token}{suffix}"
    with upload_path.open("wb") as buffer:
        shutil.copyfileobj(briefing.file, buffer)

    if generator_type == "report" and suffix != ".pdf":
        raise HTTPException(status_code=400, detail="El modo reporte necesita un PDF como documento base.")

    background_path = None
    if background and background.filename:
        background_suffix = Path(background.filename).suffix.lower()
        background_path = UPLOADS_DIR / f"{token}_background{background_suffix}"
        with background_path.open("wb") as buffer:
            shutil.copyfileobj(background.file, buffer)

    logo_path = None
    if logo and logo.filename:
        logo_suffix = Path(logo.filename).suffix.lower()
        logo_path = UPLOADS_DIR / f"{token}_logo{logo_suffix}"
        with logo_path.open("wb") as buffer:
            shutil.copyfileobj(logo.file, buffer)

    if generator_type == "report":
        outputs, report_data = generate_report_from_pdf(
            pdf_path=upload_path,
            executive_comment=executive_comment,
            next_steps=next_steps,
            report_title=report_title,
            client_name=client_name,
            report_month=report_month,
            background_path=background_path,
            logo_path=logo_path,
        )
        detected = "report"
        pending = []
        notes = report_data.extraction_notes
    else:
        selected_mode = None if mode == "auto" else mode
        outputs, data, detected = generate_from_file(upload_path, mode=selected_mode, prompt_on_combined=False)
        pending = data.pending
        notes = []

    results = []
    for output in outputs:
        summary = output.with_suffix(".txt")
        results.append(
            {
                "name": output.name,
                "download_url": f"/download/{output.name}",
                "summary_name": summary.name,
                "summary_url": f"/download/{summary.name}",
            }
        )

    message = {
        "detected": detected,
        "pending": pending,
        "notes": notes,
    }
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "message": message,
            "results": results,
        },
    )


@app.get("/download/{filename}")
async def download(filename: str):
    target = OUTPUT_DIR / filename
    if not target.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado.")
    media_type = "application/octet-stream"
    if target.suffix.lower() == ".pptx":
        media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    elif target.suffix.lower() == ".txt":
        media_type = "text/plain; charset=utf-8"
    return FileResponse(target, media_type=media_type, filename=target.name)
