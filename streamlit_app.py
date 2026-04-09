from __future__ import annotations

import tempfile
import traceback
from pathlib import Path

import streamlit as st

from generate_plan import generate_from_file
from report_generator import generate_report_from_manual_fields, generate_report_from_pdf


st.set_page_config(
    page_title="Generador de Planes de Medios",
    page_icon="P",
    layout="wide",
)


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
          --pink: #FF40B4;
          --light: #ECECEC;
          --muted: #727276;
          --dark: #1A1A1A;
        }

        .stApp {
          background:
            radial-gradient(circle at top right, rgba(255, 64, 180, 0.14), transparent 26%),
            linear-gradient(180deg, #fff8fc 0%, #f7f6f8 100%);
        }

        .brand-bar {
          display: flex;
          align-items: center;
          gap: 16px;
          padding: 8px 0 24px 0;
        }

        .brand-badge {
          width: 56px;
          height: 56px;
          border-radius: 14px;
          display: flex;
          align-items: center;
          justify-content: center;
          background: var(--pink);
          color: white;
          font-size: 30px;
          font-weight: 800;
        }

        .brand-copy h1 {
          margin: 0;
          color: #5d6472;
          font-size: 2.2rem;
          line-height: 1;
        }

        .brand-copy p {
          margin: 6px 0 0;
          color: #9096a3;
          font-size: 1.1rem;
        }

        .hero-card {
          background: rgba(255, 255, 255, 0.88);
          border: 1px solid rgba(114, 114, 118, 0.14);
          border-radius: 24px;
          padding: 36px;
          box-shadow: 0 24px 80px rgba(26, 26, 26, 0.08);
          margin-top: 10px;
        }

        .hero-card h2 {
          text-align: center;
          margin: 0 0 12px;
          color: #6b7280;
          font-size: 3rem;
        }

        .hero-card .subtitle {
          text-align: center;
          margin: 0 0 28px;
          color: #9ca3af;
          font-size: 1.35rem;
        }

        .helper-box {
          background: #f7f7f8;
          border-radius: 18px;
          padding: 18px 22px;
          margin-top: 16px;
        }

        .helper-box h3 {
          margin: 0 0 12px;
          color: #6b7280;
        }

        .helper-box p {
          margin: 6px 0;
          color: #8e95a3;
        }

        .pending-box {
          background: #fff2fa;
          border: 1px solid rgba(255, 64, 180, 0.18);
          border-radius: 16px;
          padding: 14px 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
        <div class="brand-bar">
          <div class="brand-badge">P</div>
          <div class="brand-copy">
            <h1>Planes de Medios Builder</h1>
            <p>Generador de planes estrategicos en PowerPoint</p>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_intro() -> None:
    st.markdown(
        """
        <div class="hero-card">
          <h2>Genera tu presentacion de medios</h2>
          <p class="subtitle">Ahora puedes crear planes estrategicos o reportes automaticos con branding configurable</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def save_uploaded_file(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix.lower()
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        return Path(tmp.name)


def main() -> None:
    inject_styles()
    render_header()
    render_intro()

    with st.container(border=False):
        col1, col2 = st.columns([1.3, 0.9], gap="large")

        with col1:
            generator_type = st.selectbox(
                "Tipo de generador",
                options=[
                    ("plan", "Plan estrategico"),
                    ("report", "Reporte automatico"),
                ],
                format_func=lambda item: item[1],
            )
            uploaded_file = st.file_uploader(
                "Documento de entrada",
                type=["txt", "docx", "pdf"],
                help="Planes: TXT, DOCX o PDF. Reportes: PDF opcional como referencia.",
            )
            if generator_type[0] == "plan":
                mode = st.selectbox(
                    "Tipo de salida",
                    options=[
                        ("auto", "Detectar automaticamente"),
                        ("combined", "Presentacion combinada"),
                        ("separate", "Dos presentaciones separadas"),
                        ("press", "Solo plan de prensa"),
                        ("content", "Solo marketing de contenidos"),
                    ],
                    format_func=lambda item: item[1],
                )
            else:
                mode = None
                client_name = st.text_input("Nombre del cliente")
                report_month = st.text_input("Mes del reporte", value="")
                report_title = st.text_input("Titulo interno del reporte", value="Reporte de Resultados")
                monthly_summary = st.text_area(
                    "Breve resumen del mes",
                    placeholder="Resumen corto de resultados, hitos y aprendizajes del mes.",
                    height=110,
                )
                st.markdown("**Distribucion de tiers**")
                tier_cols = st.columns(4)
                tier_1 = tier_cols[0].number_input("Tier 1", min_value=0, value=0, step=1)
                tier_2 = tier_cols[1].number_input("Tier 2", min_value=0, value=0, step=1)
                tier_3 = tier_cols[2].number_input("Tier 3", min_value=0, value=0, step=1)
                tier_4 = tier_cols[3].number_input("Tier 4", min_value=0, value=0, step=1)

                st.markdown("**Distribucion de medios**")
                media_cols = st.columns(4)
                media_digital = media_cols[0].number_input("Digital", min_value=0, value=0, step=1)
                media_tv = media_cols[1].number_input("TV", min_value=0, value=0, step=1)
                media_radio = media_cols[2].number_input("Radio", min_value=0, value=0, step=1)
                media_written = media_cols[3].number_input("Escrito", min_value=0, value=0, step=1)

                st.markdown("**Cantidad de apariciones por mes**")
                monthly_values: list[int] = []
                month_grid_rows = [st.columns(4), st.columns(4), st.columns(4)]
                month_labels = [f"Mes {index}" for index in range(1, 13)]
                month_index = 0
                for row in month_grid_rows:
                    for column in row:
                        with column:
                            monthly_values.append(
                                st.number_input(month_labels[month_index], min_value=0, value=0, step=1, key=f"month_{month_index+1}")
                            )
                        month_index += 1

                reach_value = st.text_input("Alcance de la gestion", value="")
                valuation_value = st.text_input("Valorizacion", value="")
                executive_comment = st.text_area(
                    "Comentario ejecutivo",
                    placeholder="Resumen de la gestion, hitos, aprendizajes y contexto del mes.",
                    height=140,
                )
                next_steps = st.text_area(
                    "Pasos a seguir",
                    placeholder="Una accion por linea o separadas por punto y coma.",
                    height=120,
                )
                table_rows_text = st.text_area(
                    "Tabla final de resultados",
                    placeholder="Una fila por linea con este formato: fecha | cliente | medio | tipo medio | tier | tipo comunicado",
                    height=180,
                )
                table_file = st.file_uploader(
                    "Excel o CSV para la tabla final",
                    type=["xlsx", "csv"],
                    help="Si subes este archivo, reemplaza el texto manual de la tabla final.",
                    key="report_table_file",
                )
                background_file = st.file_uploader(
                    "Fondo personalizado opcional",
                    type=["png", "jpg", "jpeg"],
                    help="Si no subes uno, se usa el fondo PRenseable integrado por defecto.",
                    key="report_background",
                )
                logo_file = st.file_uploader(
                    "Logo opcional",
                    type=["png", "jpg", "jpeg"],
                    help="Si no subes uno, se usa el logo PRenseable integrado por defecto.",
                    key="report_logo",
                )
            generate_clicked = st.button("Generar PowerPoint", type="primary", use_container_width=True)

        with col2:
            if generator_type[0] == "plan":
                st.markdown(
                    """
                    <div class="helper-box">
                      <h3>Formato esperado del documento</h3>
                      <p>DATOS GENERALES</p>
                      <p>ANALISIS DE COMPETIDORES</p>
                      <p>BUYER PERSONA</p>
                      <p>PILARES DE COMUNICACION</p>
                      <p>PROPUESTAS DE TEMATICAS</p>
                      <p>METRICAS DE EXITO</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    """
                    <div class="helper-box">
                      <h3>Como funciona el reporte</h3>
                      <p>1. Completa los campos clave del reporte dentro de la app.</p>
                      <p>2. La portada usa el fondo de portada integrado y permite montar cliente + mes.</p>
                      <p>3. El resto del deck usa el fondo PRenseable y el logo en la esquina superior derecha.</p>
                      <p>4. El PDF puede usarse solo como apoyo para copiar la tabla del locker.</p>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        if generate_clicked:
            if generator_type[0] == "plan" and uploaded_file is None:
                st.error("Primero sube un archivo para generar la presentacion.")
                return

            temp_path = save_uploaded_file(uploaded_file) if uploaded_file else None
            temp_assets: list[Path] = [temp_path] if temp_path else []
            try:
                with st.spinner("Generando presentacion..."):
                    if generator_type[0] == "plan":
                        assert temp_path is not None
                        selected_mode = None if mode[0] == "auto" else mode[0]
                        outputs, data, detected = generate_from_file(
                            temp_path,
                            mode=selected_mode,
                            prompt_on_combined=False,
                        )
                        st.success(f"Listo. Modo aplicado: {detected}")

                        if data.pending:
                            st.markdown('<div class="pending-box">', unsafe_allow_html=True)
                            st.markdown("**Informacion pendiente de completar**")
                            for item in data.pending:
                                st.write(f"- {item}")
                            st.markdown("</div>", unsafe_allow_html=True)
                    else:
                        background_path = save_uploaded_file(background_file) if background_file else None
                        logo_path = save_uploaded_file(logo_file) if logo_file else None
                        if background_path:
                            temp_assets.append(background_path)
                        if logo_path:
                            temp_assets.append(logo_path)
                        table_file_path = save_uploaded_file(table_file) if table_file else None
                        if table_file_path:
                            temp_assets.append(table_file_path)

                        outputs, manual_report = generate_report_from_manual_fields(
                            report_title=report_title,
                            client_name=client_name or "Cliente",
                            report_month=report_month or "",
                            monthly_summary=monthly_summary,
                            executive_comment=executive_comment,
                            next_steps=next_steps,
                            tier_values={
                                "Tier 1": int(tier_1),
                                "Tier 2": int(tier_2),
                                "Tier 3": int(tier_3),
                                "Tier 4": int(tier_4),
                            },
                            media_values={
                                "Digital": int(media_digital),
                                "TV": int(media_tv),
                                "Radio": int(media_radio),
                                "Escrito": int(media_written),
                            },
                            monthly_values=[int(value) for value in monthly_values],
                            reach_value=reach_value,
                            valuation_value=valuation_value,
                            table_rows_text=table_rows_text,
                            table_rows_file=table_file_path,
                            source_path=temp_path,
                            background_path=background_path,
                            logo_path=logo_path,
                        )
                        st.success("Listo. Se genero el reporte desde campos manuales.")
                        if temp_path and temp_path.suffix.lower() == ".pdf":
                            st.info("El PDF se uso solo como referencia opcional. Los graficos se construyeron desde los campos del formulario.")
                        if manual_report.table_rows:
                            st.info(f"Se cargaron {len(manual_report.table_rows)} filas en la tabla final.")

                for output in outputs:
                    ppt_bytes = output.read_bytes()
                    summary = output.with_suffix(".txt")
                    st.download_button(
                        label=f"Descargar {output.name}",
                        data=ppt_bytes,
                        file_name=output.name,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True,
                    )
                    if summary.exists():
                        st.download_button(
                            label=f"Descargar resumen {summary.name}",
                            data=summary.read_bytes(),
                            file_name=summary.name,
                            mime="text/plain",
                            use_container_width=True,
                        )
            except ValueError as exc:
                st.error(
                    "No pudimos procesar ese archivo. "
                    "Sube un TXT, DOCX o PDF valido y, si es posible, usa la plantilla esperada."
                )
                st.info(str(exc))
            except Exception as exc:
                st.error(
                    "La app encontro un problema inesperado al procesar el documento. "
                    "Prueba con otro archivo o avisanos para revisarlo."
                )
                with st.expander("Ver detalle tecnico"):
                    st.code(str(exc) or repr(exc))
                    st.code(traceback.format_exc())
            finally:
                for asset in temp_assets:
                    asset.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
