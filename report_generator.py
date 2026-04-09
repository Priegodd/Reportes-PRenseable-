from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from generate_plan import extract_text, get_output_dir

BASE_DIR = Path(__file__).resolve().parent
REPORT_ASSETS_DIR = BASE_DIR / "assets" / "report-kit"
DEFAULT_COVER_BACKGROUND = REPORT_ASSETS_DIR / "cover-background.png"
DEFAULT_SLIDE_BACKGROUND = REPORT_ASSETS_DIR / "slide-background.png"
DEFAULT_REPORT_LOGO = REPORT_ASSETS_DIR / "logo-prenseable.png"
FONT_FAMILY = "Open Sans"
TITLE_SIZE = 16
BODY_SIZE = 12

COLOR_PRIMARY = RGBColor(0xFF, 0x40, 0xB4)
COLOR_SECONDARY = RGBColor(0xEC, 0xEC, 0xEC)
COLOR_ACCENT = RGBColor(0xDE, 0x0A, 0x98)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_DARK = RGBColor(0x23, 0x22, 0x27)
COLOR_MUTED = RGBColor(0x6D, 0x6A, 0x73)
COLOR_BORDER = RGBColor(0xD9, 0xD5, 0xDD)

MONTH_ALIASES = {
    "jan": "Ene",
    "january": "Ene",
    "ene": "Ene",
    "enero": "Ene",
    "feb": "Feb",
    "february": "Feb",
    "febrero": "Feb",
    "mar": "Mar",
    "march": "Mar",
    "marzo": "Mar",
    "apr": "Abr",
    "april": "Abr",
    "abr": "Abr",
    "abril": "Abr",
    "may": "May",
    "mayo": "May",
    "jun": "Jun",
    "june": "Jun",
    "junio": "Jun",
    "jul": "Jul",
    "july": "Jul",
    "julio": "Jul",
    "aug": "Ago",
    "august": "Ago",
    "ago": "Ago",
    "agosto": "Ago",
    "sep": "Sep",
    "sept": "Sep",
    "september": "Sep",
    "septiembre": "Sep",
    "oct": "Oct",
    "october": "Oct",
    "octubre": "Oct",
    "nov": "Nov",
    "november": "Nov",
    "noviembre": "Nov",
    "dec": "Dic",
    "december": "Dic",
    "dic": "Dic",
    "diciembre": "Dic",
}


@dataclass
class MetricPoint:
    label: str
    value: float
    unit: str = ""
    raw_value: str = ""
    context: str = ""


@dataclass
class ReportData:
    source_text: str
    title: str
    source_name: str
    client_name: str = ""
    report_month: str = ""
    executive_comment: str = ""
    next_steps: str = ""
    metrics: list[MetricPoint] = field(default_factory=list)
    bullets: list[str] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)
    extraction_notes: list[str] = field(default_factory=list)


def clean_line(line: str) -> str:
    return re.sub(r"\s+", " ", line.replace("\uf0b7", "-")).strip()


def safe_metric_label(label: str) -> str:
    label = re.sub(r"^[\-\*\u2022]+\s*", "", label).strip(" :-")
    label = re.sub(r"\s+", " ", label)
    return label[:40] or "Metrica"


def parse_numeric_token(token: str) -> float | None:
    current = token.strip().replace(" ", "")
    if not re.search(r"\d", current):
        return None
    if "," in current and "." in current:
        current = current.replace(".", "").replace(",", ".")
    elif "," in current:
        current = current.replace(".", "").replace(",", ".")
    try:
        return float(current)
    except ValueError:
        return None


def detect_unit(raw: str) -> str:
    if "%" in raw:
        return "%"
    if re.search(r"\b(?:usd|clp|uf|\$)\b", raw.lower()) or "$" in raw:
        return "$"
    return ""


def metric_sort_key(metric: MetricPoint) -> tuple[int, float]:
    weight = 1 if metric.unit == "%" else 2 if metric.unit == "$" else 0
    return (weight, abs(metric.value))


def normalize_month(label: str) -> str | None:
    key = re.sub(r"[^a-zA-Z]", "", label).lower()
    return MONTH_ALIASES.get(key)


def extract_candidate_metrics(text: str) -> list[MetricPoint]:
    metrics: list[MetricPoint] = []
    seen: set[tuple[str, str]] = set()
    lines = [clean_line(line) for line in text.splitlines()]
    numeric_pattern = re.compile(r"([-+]?\d[\d\.,]*\s*%?)")

    for line in lines:
        if len(line) < 3 or len(line) > 130:
            continue
        matches = list(numeric_pattern.finditer(line))
        if not matches:
            continue
        if len(matches) > 3:
            continue

        first = matches[-1]
        raw_value = first.group(1).strip()
        value = parse_numeric_token(raw_value.replace("%", ""))
        if value is None:
            continue

        prefix = line[: first.start()].strip(" :-|")
        suffix = line[first.end() :].strip(" :-|")
        label = safe_metric_label(prefix or suffix or "Metrica")
        if len(label) < 3:
            continue
        if re.fullmatch(r"\d{1,4}", label):
            continue

        candidate = MetricPoint(
            label=label,
            value=value,
            unit=detect_unit(raw_value),
            raw_value=raw_value,
            context=line,
        )
        fingerprint = (candidate.label.lower(), candidate.raw_value)
        if fingerprint in seen:
            continue
        seen.add(fingerprint)
        metrics.append(candidate)

    return sorted(metrics, key=metric_sort_key, reverse=True)


def build_quant_bullets(metrics: list[MetricPoint]) -> list[str]:
    bullets: list[str] = []
    if not metrics:
        return [
            "No se encontraron metricas numericas suficientemente limpias en el PDF para construir lectura automatica.",
            "El flujo actual funciona mejor cuando el PDF trae texto seleccionable, tablas o etiquetas numericas visibles.",
        ]

    top = metrics[0]
    bullets.append(f"La metrica mas fuerte detectada es {top.label} con {top.raw_value}.")

    percent_metrics = [metric for metric in metrics if metric.unit == "%"]
    if percent_metrics:
        best_percent = max(percent_metrics, key=lambda metric: metric.value)
        bullets.append(f"Entre los porcentajes extraidos, destaca {best_percent.label} con {best_percent.raw_value}.")

    comparable = [metric for metric in metrics if metric.unit == top.unit][:4]
    if len(comparable) >= 2:
        highest = max(comparable, key=lambda metric: metric.value)
        lowest = min(comparable, key=lambda metric: metric.value)
        if highest.label != lowest.label:
            gap = highest.value - lowest.value
            gap_text = f"{gap:.1f}".rstrip("0").rstrip(".")
            suffix = highest.unit if highest.unit else ""
            bullets.append(
                f"La brecha entre la mejor y la menor lectura comparable es {gap_text}{suffix}, entre {highest.label} y {lowest.label}."
            )

    if len(metrics) >= 4:
        bullets.append(
            f"El PDF entrego al menos {len(metrics)} puntos numericos reutilizables, suficiente para una lectura cuantitativa inicial."
        )
    else:
        bullets.append("La lectura cuantitativa es preliminar; conviene complementar con mas datos estructurados si quieren mayor precision.")

    return bullets[:4]


def build_evidence(text: str, metrics: list[MetricPoint]) -> list[str]:
    evidence = [metric.context for metric in metrics[:6] if metric.context]
    if evidence:
        return evidence
    lines = [clean_line(line) for line in text.splitlines() if clean_line(line)]
    return lines[:6]


def parse_report_pdf(
    pdf_path: Path,
    executive_comment: str = "",
    next_steps: str = "",
    report_title: str | None = None,
    client_name: str = "",
    report_month: str = "",
) -> ReportData:
    text = extract_text(pdf_path)
    metrics = extract_candidate_metrics(text)
    extraction_notes = []
    if metrics:
        extraction_notes.append(
            f"Se detectaron {len(metrics)} metricas candidatas desde texto legible del PDF."
        )
    else:
        extraction_notes.append(
            "No se detectaron metricas legibles; si el grafico esta embebido como imagen, sera necesario OCR o carga manual."
        )

    title = report_title or f"Reporte Automatico {datetime.now():%B %Y}"
    return ReportData(
        source_text=text,
        title=title,
        source_name=pdf_path.name,
        client_name=client_name.strip(),
        report_month=report_month.strip(),
        executive_comment=executive_comment.strip(),
        next_steps=next_steps.strip(),
        metrics=metrics,
        bullets=build_quant_bullets(metrics),
        evidence=build_evidence(text, metrics),
        extraction_notes=extraction_notes,
    )


def add_picture_safe(slide, path: Path, left, top, width=None, height=None):
    if not path or not path.exists():
        return None
    kwargs = {}
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    return slide.shapes.add_picture(str(path), left, top, **kwargs)


def resolve_asset_path(candidate: Path | None, default: Path | None) -> Path | None:
    if candidate and candidate.exists():
        return candidate
    if default and default.exists():
        return default
    return None


def apply_report_background(slide, background_path: Path | None) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_SECONDARY

    resolved_background = resolve_asset_path(background_path, DEFAULT_SLIDE_BACKGROUND)
    if resolved_background:
        add_picture_safe(slide, resolved_background, 0, 0, width=Inches(13.333), height=Inches(7.5))
    else:
        top_band = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.0))
        top_band.fill.solid()
        top_band.fill.fore_color.rgb = COLOR_PRIMARY
        top_band.line.fill.background()
        accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.35), Inches(0.35), Inches(12.6), Inches(0.03))
        accent.fill.solid()
        accent.fill.fore_color.rgb = COLOR_ACCENT
        accent.line.fill.background()


def add_logo_box(slide, logo_path: Path | None) -> None:
    resolved_logo = resolve_asset_path(logo_path, DEFAULT_REPORT_LOGO)
    box = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(11.58), Inches(0.2), Inches(1.4), Inches(1.0))
    box.fill.solid()
    box.fill.fore_color.rgb = COLOR_WHITE
    box.line.color.rgb = COLOR_BORDER
    if resolved_logo:
        add_picture_safe(slide, resolved_logo, Inches(11.72), Inches(0.33), width=Inches(1.12))
    else:
        tf = box.text_frame
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = "LOGO"
        run.font.name = FONT_FAMILY
        run.font.size = Pt(BODY_SIZE)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY


def add_slide_base(prs: Presentation, title: str, background_path: Path | None, logo_path: Path | None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    apply_report_background(slide, background_path)
    add_logo_box(slide, logo_path)

    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.28), Inches(9.8), Inches(0.6))
    p = title_box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.name = FONT_FAMILY
    run.font.bold = True
    run.font.size = Pt(TITLE_SIZE)
    run.font.color.rgb = COLOR_DARK
    return slide


def add_text_block(slide, title: str, body: str, left, top, width, height, accent_color: RGBColor = COLOR_PRIMARY) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_WHITE
    shape.line.color.rgb = COLOR_BORDER
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(TITLE_SIZE)
    p.font.color.rgb = COLOR_DARK
    p2 = tf.add_paragraph()
    p2.text = body
    p2.font.name = FONT_FAMILY
    p2.font.size = Pt(BODY_SIZE)
    p2.font.color.rgb = COLOR_MUTED
    accent = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, Inches(0.14), height)
    accent.fill.solid()
    accent.fill.fore_color.rgb = accent_color
    accent.line.fill.background()


def add_bullet_list(slide, title: str, bullets: list[str], left, top, width, height) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = COLOR_WHITE
    shape.line.color.rgb = COLOR_BORDER
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = FONT_FAMILY
    p.font.bold = True
    p.font.size = Pt(TITLE_SIZE)
    p.font.color.rgb = COLOR_DARK
    for bullet in bullets:
        item = tf.add_paragraph()
        item.text = f"- {bullet}"
        item.level = 0
        item.font.name = FONT_FAMILY
        item.font.size = Pt(BODY_SIZE)
        item.font.color.rgb = COLOR_MUTED


def add_metric_cards(slide, metrics: list[MetricPoint]) -> None:
    cards = metrics[:4]
    while len(cards) < 4:
        cards.append(MetricPoint(label="Dato pendiente", value=0, raw_value="[sin dato]"))

    for index, metric in enumerate(cards):
        left = Inches(0.85 + (index % 2) * 6.0)
        top = Inches(1.45 + (index // 2) * 1.7)
        shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, left, top, Inches(5.5), Inches(1.35))
        shape.fill.solid()
        shape.fill.fore_color.rgb = COLOR_WHITE
        shape.line.color.rgb = COLOR_BORDER
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = metric.label
        p.font.name = FONT_FAMILY
        p.font.size = Pt(BODY_SIZE)
        p.font.color.rgb = COLOR_MUTED
        p2 = tf.add_paragraph()
        p2.text = metric.raw_value or "[sin dato]"
        p2.font.name = FONT_FAMILY
        p2.font.bold = True
        p2.font.size = Pt(TITLE_SIZE)
        p2.font.color.rgb = COLOR_PRIMARY if index % 2 == 0 else COLOR_ACCENT


def add_metric_chart(slide, metrics: list[MetricPoint]) -> None:
    chart_metrics = [metric for metric in metrics if metric.value >= 0][:6]
    if not chart_metrics:
        add_text_block(
            slide,
            "Datos cuantitativos",
            "No hubo suficientes valores consistentes para graficar automaticamente desde el PDF.",
            Inches(0.85),
            Inches(1.55),
            Inches(11.7),
            Inches(1.2),
        )
        return

    chart_data = CategoryChartData()
    chart_data.categories = [metric.label[:18] for metric in chart_metrics]
    chart_data.add_series("Valor", [metric.value for metric in chart_metrics])

    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(0.9),
        Inches(1.7),
        Inches(7.1),
        Inches(4.5),
        chart_data,
    ).chart
    chart.has_legend = False
    chart.value_axis.has_major_gridlines = True
    chart.category_axis.tick_labels.font.size = Pt(BODY_SIZE)
    chart.value_axis.tick_labels.font.size = Pt(BODY_SIZE)
    series = chart.series[0]
    series.format.fill.solid()
    series.format.fill.fore_color.rgb = COLOR_PRIMARY
    series.format.line.color.rgb = COLOR_ACCENT
    plot = chart.plots[0]
    plot.has_data_labels = True
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END

    add_bullet_list(
        slide,
        "Lectura cuantitativa",
        build_quant_bullets(metrics),
        Inches(8.25),
        Inches(1.7),
        Inches(4.2),
        Inches(4.5),
    )


def add_cover_slide(prs: Presentation, data: ReportData, background_path: Path | None, logo_path: Path | None) -> None:
    cover_background = resolve_asset_path(background_path, DEFAULT_COVER_BACKGROUND)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = COLOR_PRIMARY
    if cover_background:
        add_picture_safe(slide, cover_background, 0, 0, width=Inches(13.333), height=Inches(7.5))

    client_box = slide.shapes.add_textbox(Inches(0.75), Inches(5.45), Inches(6.3), Inches(0.55))
    p = client_box.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = data.client_name or data.title or "[Nombre del cliente]"
    run.font.name = FONT_FAMILY
    run.font.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = COLOR_WHITE

    month_box = slide.shapes.add_textbox(Inches(0.75), Inches(6.02), Inches(4.0), Inches(0.35))
    p2 = month_box.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    run2 = p2.add_run()
    run2.text = data.report_month or datetime.now().strftime("%B %Y")
    run2.font.name = FONT_FAMILY
    run2.font.bold = False
    run2.font.size = Pt(TITLE_SIZE)
    run2.font.color.rgb = COLOR_WHITE


def add_extraction_slide(prs: Presentation, data: ReportData, background_path: Path | None, logo_path: Path | None) -> None:
    slide = add_slide_base(prs, "Lectura del PDF", background_path, logo_path)
    add_text_block(
        slide,
        "Estado de extraccion",
        "\n".join(data.extraction_notes),
        Inches(0.85),
        Inches(1.45),
        Inches(5.45),
        Inches(1.25),
    )
    add_text_block(
        slide,
        "Capacidad actual",
        "El sistema reutiliza texto, tablas y etiquetas numericas presentes en el PDF. Si un grafico viene solo como imagen, la lectura automatica puede quedar parcial.",
        Inches(6.45),
        Inches(1.45),
        Inches(5.95),
        Inches(1.25),
        accent_color=COLOR_ACCENT,
    )
    add_bullet_list(
        slide,
        "Evidencia detectada",
        data.evidence[:5] or ["Sin evidencia legible para mostrar."],
        Inches(0.85),
        Inches(3.0),
        Inches(11.55),
        Inches(3.2),
    )


def add_exec_slide(prs: Presentation, data: ReportData, background_path: Path | None, logo_path: Path | None) -> None:
    slide = add_slide_base(prs, "Comentario Ejecutivo", background_path, logo_path)
    add_text_block(
        slide,
        "Resumen de gestion",
        data.executive_comment or "[Completar comentario ejecutivo de la gestion del mes]",
        Inches(0.85),
        Inches(1.5),
        Inches(11.55),
        Inches(4.8),
    )


def add_next_steps_slide(prs: Presentation, data: ReportData, background_path: Path | None, logo_path: Path | None) -> None:
    slide = add_slide_base(prs, "Pasos a Seguir", background_path, logo_path)
    bullets = [clean_line(line) for line in data.next_steps.splitlines() if clean_line(line)]
    if not bullets and data.next_steps.strip():
        bullets = [item.strip() for item in re.split(r"[;\n]", data.next_steps) if item.strip()]
    add_bullet_list(
        slide,
        "Mes siguiente",
        bullets or ["[Completar siguientes pasos para el proximo mes]"],
        Inches(0.85),
        Inches(1.45),
        Inches(11.55),
        Inches(4.9),
    )


def write_report_summary(
    output_path: Path,
    data: ReportData,
    background_path: Path | None,
    logo_path: Path | None,
) -> Path:
    summary_path = output_path.with_suffix(".txt")
    lines = [
        f"Archivo generado: {output_path.name}",
        f"Fuente: {data.source_name}",
        f"Fecha: {datetime.now().isoformat(timespec='seconds')}",
        "",
        "Branding aplicado:",
        f"- Fondo personalizado: {'si' if background_path and background_path.exists() else 'no'}",
        f"- Logo personalizado: {'si' if logo_path and logo_path.exists() else 'no'}",
        "",
        "Lectura cuantitativa:",
    ]
    lines.extend(f"- {bullet}" for bullet in data.bullets)
    lines.append("")
    lines.append("Metricas detectadas:")
    if data.metrics:
        lines.extend(f"- {metric.label}: {metric.raw_value}" for metric in data.metrics[:10])
    else:
        lines.append("- No se detectaron metricas legibles.")
    summary_path.write_text("\n".join(lines), encoding="utf-8")
    return summary_path


def generate_report_from_pdf(
    pdf_path: Path,
    executive_comment: str = "",
    next_steps: str = "",
    report_title: str | None = None,
    client_name: str = "",
    report_month: str = "",
    background_path: Path | None = None,
    logo_path: Path | None = None,
    output_dir: Path | None = None,
) -> tuple[list[Path], ReportData]:
    output_dir = output_dir or get_output_dir()
    output_dir.mkdir(parents=True, exist_ok=True)

    data = parse_report_pdf(
        pdf_path=pdf_path,
        executive_comment=executive_comment,
        next_steps=next_steps,
        report_title=report_title,
        client_name=client_name,
        report_month=report_month,
    )

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    add_cover_slide(prs, data, background_path, logo_path)

    slide_metrics = add_slide_base(prs, "Resumen de KPIs", background_path, logo_path)
    add_metric_cards(slide_metrics, data.metrics)

    slide_chart = add_slide_base(prs, "Graficos y Opinion Cuantitativa", background_path, logo_path)
    add_metric_chart(slide_chart, data.metrics)

    add_extraction_slide(prs, data, background_path, logo_path)
    add_exec_slide(prs, data, background_path, logo_path)
    add_next_steps_slide(prs, data, background_path, logo_path)

    slug = re.sub(r"[^a-z0-9_]+", "_", pdf_path.stem.lower()).strip("_") or "reporte"
    output_path = output_dir / f"Reporte_Automatico_{slug}_{datetime.now().year}.pptx"
    prs.save(output_path)
    write_report_summary(output_path, data, background_path, logo_path)
    return [output_path], data
