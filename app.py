import base64
from io import BytesIO
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import qrcode
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, LongTable, TableStyle, Image

st.set_page_config(page_title="Entrega y Recepción de Equipos", page_icon="📦", layout="wide", initial_sidebar_state="expanded")

PRIMARY = "#102A56"
GREEN = "#D6E53F"
EXCEL_FILE = Path("registro_entregas.xlsx")
PDF_DIR = Path("pdf_generados")
PDF_DIR.mkdir(exist_ok=True)

st.markdown("""
<style>
[data-testid="stSidebar"]{background:linear-gradient(180deg,#061B38 0%,#102A56 100%)}
[data-testid="stSidebar"] *{color:white}
[data-testid="stSidebar"] img{background:transparent;padding:6px}
.block-container{padding-top:1rem;padding-left:1.5rem;padding-right:1.5rem;max-width:100%}
.main-title{font-size:34px;line-height:1.2;font-weight:900;color:#102A56;margin:0}
.subtitle{font-size:15px;color:#667085;margin-top:8px;margin-bottom:18px}
.section-title{color:#169B62;font-weight:800;font-size:17px;margin-bottom:12px}
.stButton>button{border-radius:10px;font-weight:700;min-height:48px}
[data-testid="stSidebar"] .stButton>button{background:#1A365F;border:0;color:white;text-align:left;justify-content:flex-start;width:100%;padding-left:18px}
[data-testid="stSidebar"] .stButton>button:hover{background:#169B62;color:white}
div[data-testid="stDownloadButton"] button{background:#169B62;color:white;border:none;font-weight:800}
</style>
""", unsafe_allow_html=True)

def obtener_siguiente_documento():
    if not EXCEL_FILE.exists():
        return "ER-ST-001"
    try:
        wb = load_workbook(EXCEL_FILE)
        if "Registro General" not in wb.sheetnames:
            return "ER-ST-001"
        nums = []
        for row in wb["Registro General"].iter_rows(min_row=2, values_only=True):
            d = row[0]
            if d and str(d).startswith("ER-ST-"):
                try:
                    nums.append(int(str(d).replace("ER-ST-", "")))
                except Exception:
                    pass
        return f"ER-ST-{(max(nums)+1 if nums else 1):03d}"
    except Exception:
        return "ER-ST-001"

def crear_excel():
    if EXCEL_FILE.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro General"
    ws.append(["N° Documento","Fecha","Hora","Destino","Proyecto","Transportista","Entregado por","Recibido por","Equipo responsable","Observaciones","Archivo PDF"])
    wd = wb.create_sheet("Detalle Materiales")
    wd.append(["N° Documento","Material / Equipo","Cantidad","Unidad","Entregado","Recibido"])
    wb.save(EXCEL_FILE)

def guardar_excel(data, materiales, pdf_name):
    crear_excel()
    wb = load_workbook(EXCEL_FILE)
    wb["Registro General"].append([
        data["numero_documento"], data["fecha"], data["hora"], data["destino"],
        data["proyecto"], data["transportista"], data["entregado_por"],
        data["recibido_por"], data["equipo_responsable"], data["observaciones"], pdf_name
    ])
    ws = wb["Detalle Materiales"]
    for m in materiales:
        ws.append([
            data["numero_documento"], m.get("Material / Equipo",""), m.get("Cantidad",""),
            m.get("Unidad",""), "Sí" if m.get("Entregado",False) else "No",
            "Sí" if m.get("Recibido",False) else "No"
        ])
    wb.save(EXCEL_FILE)

def cargar_historial():
    if not EXCEL_FILE.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name="Registro General")
    except Exception:
        return pd.DataFrame()

def normalizar_materiales(materiales):
    return [m for m in materiales if str(m.get("Material / Equipo","")).strip() not in ("", "None")]

def generar_qr(data):
    texto = f"N° Documento: {data['numero_documento']}\nProyecto: {data['proyecto']}\nDestino: {data['destino']}\nFecha: {data['fecha']}"
    qr = qrcode.QRCode(box_size=3, border=2)
    qr.add_data(texto)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    b = BytesIO()
    img.save(b, format="PNG")
    b.seek(0)
    return b

def generar_pdf(data, materiales):
    b = BytesIO()
    doc = SimpleDocTemplate(b, pagesize=A4, rightMargin=1.3*cm, leftMargin=1.3*cm, topMargin=1.1*cm, bottomMargin=1.1*cm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle("TituloST", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=14, alignment=1, leading=18)
    txt = ParagraphStyle("TextoST", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10)
    sec = ParagraphStyle("SeccionST", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, textColor=colors.white)
    story = []
    logo = Image("logo.png", width=4.1*cm, height=2*cm) if Path("logo.png").exists() else Paragraph("SOLARTEAM", title)
    head = Table([[logo, Paragraph("FORMULARIO DE ENTREGA Y<br/>RECEPCIÓN DE EQUIPOS", title), Paragraph(f"<b>N° DOCUMENTO</b><br/>{data['numero_documento']}", txt)]], colWidths=[5*cm,8.4*cm,4.2*cm])
    head.setStyle(TableStyle([("BOX",(0,0),(-1,-1),.8,colors.HexColor("#9AA4B2")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),7)]))
    story += [head, Spacer(1,8)]
    qr = Image(generar_qr(data), width=2.2*cm, height=2.2*cm)
    info = Table([[Paragraph(f"<b>Fecha:</b> {data['fecha']}",txt), qr]], colWidths=[7*cm,7*cm,3.4*cm])
    story += [info, Spacer(1,8)]
    s1 = Table([[Paragraph("DATOS GENERALES",sec)]], colWidths=[17.6*cm])
    s1.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor(PRIMARY)),("PADDING",(0,0),(-1,-1),5)]))
    story.append(s1)
    datos = [
        ["Destino / Área:",data["destino"],"Proyecto asignado:",data["proyecto"]],
        ["Transportista:",data["transportista"],"",""],
        ["Entregado por:",data["entregado_por"],"",""],
        ["Recibido por:",data["recibido_por"],"Equipo responsable:",data["equipo_responsable"]],
        ["Observaciones:",data["observaciones"] or "Sin observaciones.","",""]
    ]
    td = Table(datos,colWidths=[3.2*cm,6*cm,3*cm,5.4*cm])
    td.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#F1F5F9")),("BACKGROUND",(2,0),(2,-1),colors.HexColor("#F1F5F9")),("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("FONTNAME",(2,0),(2,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("SPAN",(1,1),(3,1)),("SPAN",(1,2),(3,2)),("SPAN",(1,4),(3,4)),("PADDING",(0,0),(-1,-1),5)]))
    story += [td, Spacer(1,10)]
    s2 = Table([[Paragraph("DETALLE DE MATERIALES / EQUIPOS",sec)]], colWidths=[17.6*cm])
    s2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor(PRIMARY)),("PADDING",(0,0),(-1,-1),5)]))
    story.append(s2)
    detalle = [["No.","Material / Equipo","Cantidad","Unidad","Entregado","Recibido"]]
    for i,m in enumerate(materiales,1):
        detalle.append([str(i),Paragraph(str(m.get("Material / Equipo","")),txt),str(m.get("Cantidad","")),str(m.get("Unidad","")),"X" if m.get("Entregado",False) else "","X" if m.get("Recibido",False) else ""])
    tm = LongTable(detalle,colWidths=[1*cm,8.1*cm,2.2*cm,2.3*cm,2*cm,2*cm],repeatRows=1)
    tm.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1F5F9")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(0,0),(0,-1),"CENTER"),("ALIGN",(2,1),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("PADDING",(0,0),(-1,-1),5)]))
    story += [tm, Spacer(1,18)]
    firmas = Table([["ENTREGA","RECEPCIÓN"],["",""],[data["entregado_por"],data["recibido_por"]],["",data["equipo_responsable"]],["Nombre y firma","Nombre y firma"]],colWidths=[8.8*cm,8.8*cm],rowHeights=[.6*cm,2*cm,.5*cm,.5*cm,.5*cm])
    firmas.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(-1,0),colors.HexColor(PRIMARY)),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("FONTSIZE",(0,0),(-1,-1),8)]))
    story.append(firmas)
    doc.build(story)
    b.seek(0)
    return b

def preview_html(data, materiales):
    logo = ""
    if Path("logo.png").exists():
        logo = base64.b64encode(Path("logo.png").read_bytes()).decode()
        logo = f'<img src="data:image/png;base64,{logo}" style="width:145px;height:auto">'
    rows = "".join(f"<tr><td>{i}</td><td>{m.get('Material / Equipo','')}</td><td>{m.get('Cantidad','')}</td><td>{m.get('Unidad','')}</td><td>{'✓' if m.get('Entregado',False) else ''}</td><td>{'✓' if m.get('Recibido',False) else ''}</td></tr>" for i,m in enumerate(materiales,1))
    if not rows: rows="<tr><td colspan='6'>Sin materiales registrados</td></tr>"
    return f"""<!doctype html><html><head><style>*{{box-sizing:border-box}}body{{margin:0;font-family:Arial;background:white}}.page{{border:1px solid #CBD5E1;padding:20px;min-height:760px}}.head{{display:grid;grid-template-columns:150px 1fr 170px;gap:12px;align-items:center;border-bottom:1px solid #94A3B8;padding-bottom:14px}}.title{{font-size:19px;line-height:1.5;font-weight:800;text-align:center;color:#071B38}}.doc{{border:1px solid #94A3B8;padding:15px;text-align:center;font-weight:800}}.meta{{display:flex;justify-content:space-between;padding:14px 0;font-size:13px}}.sec{{background:#102A56;color:white;font-weight:800;padding:7px 10px;font-size:13px}}table{{width:100%;border-collapse:collapse;font-size:12px}}th,td{{border:1px solid #D0D5DD;padding:7px}}th{{background:#F2F4F7}}.sign{{margin-top:18px}}</style></head><body><div class="page"><div class="head"><div>{logo}</div><div class="title">FORMULARIO DE ENTREGA<br>Y<br>RECEPCIÓN DE EQUIPOS</div><div class="doc">N° DOCUMENTO<br>{data['numero_documento']}</div></div><div class="meta"><div><b>Fecha:</b> {data['fecha']}</div><div><b>Hora:</b> {data['hora']}</div></div><div class="sec">DATOS GENERALES</div><table><tr><td><b>Destino / Área:</b></td><td>{data['destino']}</td><td><b>Proyecto:</b></td><td>{data['proyecto']}</td></tr><tr><td><b>Transportista:</b></td><td colspan="3">{data['transportista']}</td></tr><tr><td><b>Entregado por:</b></td><td colspan="3">{data['entregado_por']}</td></tr><tr><td><b>Recibido por:</b></td><td>{data['recibido_por']}</td><td><b>Equipo responsable:</b></td><td>{data['equipo_responsable']}</td></tr><tr><td><b>Observaciones:</b></td><td colspan="3">{data['observaciones']}</td></tr></table><div class="sec">DETALLE DE MATERIALES / EQUIPOS</div><table><tr><th>No.</th><th>Material / Equipo</th><th>Cantidad</th><th>Unidad</th><th>Entregado</th><th>Recibido</th></tr>{rows}</table><table class="sign"><tr><th>ENTREGA</th><th>RECEPCIÓN</th></tr><tr><td style="height:75px;text-align:center">____________________<br>{data['entregado_por']}</td><td style="height:75px;text-align:center">____________________<br>{data['recibido_por']}<br>{data['equipo_responsable']}</td></tr></table></div></body></html>"""

if "pagina" not in st.session_state:
    st.session_state.pagina = "Nuevo Formulario"
if "materiales" not in st.session_state:
    st.session_state.materiales = [{"Material / Equipo":"","Cantidad":1,"Unidad":"Unidad","Entregado":True,"Recibido":True}]

with st.sidebar:
    if Path("logo.png").exists():
        st.image("logo.png", use_container_width=True)
    if st.button("📄  Nuevo Formulario", use_container_width=True): st.session_state.pagina="Nuevo Formulario"; st.rerun()
    if st.button("🕘  Historial de Entregas", use_container_width=True): st.session_state.pagina="Historial"; st.rerun()
    if st.button("📦  Catálogo de Materiales", use_container_width=True): st.session_state.pagina="Catálogo"; st.rerun()
    if st.button("📊  Reportes", use_container_width=True): st.session_state.pagina="Reportes"; st.rerun()
    if st.button("⚙️  Configuración", use_container_width=True): st.session_state.pagina="Configuración"; st.rerun()

pagina = st.session_state.pagina

if pagina == "Nuevo Formulario":
    numero = obtener_siguiente_documento()
    ahora = datetime.now()
    st.markdown('<h1 class="main-title">Entrega y Recepción de Equipos</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Complete la información para generar el documento.</p>', unsafe_allow_html=True)
    left,right = st.columns([1.05,1.15],gap="large")
    with left:
        st.markdown('<div class="section-title">👷 1. DATOS GENERALES</div>',unsafe_allow_html=True)
        c1,c2=st.columns(2)
        with c1:
            destino=st.text_input("Destino / Área *")
            proyecto=st.text_input("Proyecto asignado *")
            transportista=st.text_input("Transportista *")
            entregado_por=st.text_input("Entregado por *",value="SOLARTEAM")
        with c2:
            recibido_por=st.text_input("Recibido por *")
            equipo_responsable=st.text_input("Equipo responsable")
            fecha=st.date_input("Fecha *",value=date.today())
            observaciones=st.text_area("Observaciones",height=125)
        st.divider()
        st.markdown('<div class="section-title">📋 2. MATERIALES / EQUIPOS</div>',unsafe_allow_html=True)
        st.caption("Puedes copiar desde Excel y pegar Material / Equipo, Cantidad y Unidad.")
        materiales=st.data_editor(st.session_state.materiales,num_rows="dynamic",use_container_width=True,hide_index=True,column_config={
            "Material / Equipo":st.column_config.TextColumn("Material / Equipo"),
            "Cantidad":st.column_config.NumberColumn("Cantidad",min_value=0.0,step=1.0),
            "Unidad":st.column_config.TextColumn("Unidad"),
            "Entregado":st.column_config.CheckboxColumn("Entregado"),
            "Recibido":st.column_config.CheckboxColumn("Recibido")
        })
        st.session_state.materiales=materiales
        data={"numero_documento":numero,"fecha":fecha.strftime("%d/%m/%Y"),"hora":ahora.strftime("%H:%M"),"destino":destino,"proyecto":proyecto,"transportista":transportista,"entregado_por":entregado_por,"recibido_por":recibido_por,"equipo_responsable":equipo_responsable,"observaciones":observaciones}
        mats=normalizar_materiales(materiales)
        if st.button("⬇️ Generar PDF y Guardar",type="primary",use_container_width=True):
            if not all(str(x).strip() for x in [destino,proyecto,transportista,entregado_por,recibido_por]):
                st.warning("Completa los campos obligatorios.")
            elif not mats:
                st.warning("Agrega al menos un material o equipo.")
            else:
                pdf=generar_pdf(data,mats)
                name=f"{numero}.pdf"
                path=PDF_DIR/name
                path.write_bytes(pdf.getvalue())
                guardar_excel(data,mats,str(path))
                st.success(f"Formulario {numero} generado correctamente.")
                st.download_button("Descargar PDF",pdf.getvalue(),name,"application/pdf",use_container_width=True)
    with right:
        st.markdown('<div class="section-title">VISTA PREVIA DEL DOCUMENTO</div>',unsafe_allow_html=True)
        components.html(preview_html(data,mats),height=900,scrolling=True)

elif pagina == "Historial":
    st.markdown('<h1 class="main-title">Historial de Entregas</h1>',unsafe_allow_html=True)
    hist=cargar_historial()
    if hist.empty: st.info("Todavía no existen entregas registradas.")
    else:
        q=st.text_input("Buscar por proyecto, destino o N° de documento")
        if q.strip():
            hist=hist[hist.astype(str).apply(lambda r:r.str.contains(q,case=False,na=False).any(),axis=1)]
        st.dataframe(hist,use_container_width=True,hide_index=True)
        if EXCEL_FILE.exists():
            st.download_button("📗 Exportar Excel",EXCEL_FILE.read_bytes(),"registro_entregas.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif pagina == "Catálogo":
    st.markdown('<h1 class="main-title">Catálogo de Materiales</h1>',unsafe_allow_html=True)
    st.info("Módulo preparado para una siguiente etapa. Por ahora los materiales y unidades son de texto libre.")

elif pagina == "Reportes":
    st.markdown('<h1 class="main-title">Reportes</h1>',unsafe_allow_html=True)
    hist=cargar_historial()
    if hist.empty: st.info("Aún no hay datos para reportes.")
    else:
        c1,c2=st.columns(2)
        c1.metric("Entregas registradas",len(hist))
        c2.metric("Proyectos registrados",hist["Proyecto"].nunique())

else:
    st.markdown('<h1 class="main-title">Configuración</h1>',unsafe_allow_html=True)
    st.info("Configuración general del formulario. La numeración automática usa el formato ER-ST-001.")
