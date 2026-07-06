
import streamlit as st
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
import pandas as pd
import qrcode
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

st.set_page_config(page_title="Entrega y Recepción de Equipos", page_icon="📦", layout="wide", initial_sidebar_state="expanded")
PRIMARY = "#102A56"
GREEN = "#169B62"
EXCEL_FILE = Path("registro_entregas.xlsx")
PDF_DIR = Path("pdf_generados")
PDF_DIR.mkdir(exist_ok=True)
CATALOGO_MATERIALES = ["Panel Solar Jinko 625W","Panel Solar Jinko 620W","Inversor Huawei 100KTL-M1","Inversor Huawei 50KTL","Cable Solar 6mm²","Cable Solar 4mm²","Estructura Grace Solar","Conector MC4","Breaker DC","Breaker AC","Canaleta","Tubería EMT","Tablero eléctrico","Transformador","Otro"]
UNIDADES = ["Unidad","Metro","Rollo","Caja","Kit","Juego","Par","Global","Otro"]

st.markdown("""
<style>
[data-testid="stSidebar"] {background: linear-gradient(180deg, #061B38 0%, #102A56 100%);} 
[data-testid="stSidebar"] * {color: white;}
[data-testid="stSidebar"] img {background: white; padding: 8px;}
.block-container {padding-top: 1.3rem; padding-left: 1.3rem; padding-right: 1.3rem;}
.main-title {font-size: 30px; font-weight: 800; color: #102A56; margin-bottom: 0;}
.subtitle {font-size: 14px; color: #667085; margin-top: 2px; margin-bottom: 18px;}
.card {background: white; border: 1px solid #E5E7EB; border-radius: 14px; padding: 18px; box-shadow: 0 1px 4px rgba(16,42,86,.06);}
.section-title {color: #169B62; font-weight: 800; font-size: 17px; margin-bottom: 12px;}
.small-muted {color: #667085; font-size: 12px;}
.stButton>button {border-radius: 9px; font-weight: 700; height: 42px;}
div[data-testid="stDownloadButton"] button {background-color: #169B62; color: white; border: none; border-radius: 9px; font-weight: 800; height: 42px;}
div[data-testid="stDownloadButton"] button:hover {background-color: #11814F; color: white;}
.sidebar-btn {display:block; padding:13px 14px; margin:10px 0; border-radius:9px; font-weight:700; background:rgba(255,255,255,.06);}
.sidebar-btn-active {background:#169B62;}
.help-box {border:1px solid rgba(255,255,255,.35); border-radius:10px; padding:12px; margin-top:100px; font-size:13px;}
.preview-box {background:white; border:1px solid #E5E7EB; border-radius:13px; padding:18px; box-shadow:0 8px 24px rgba(16,42,86,.10);}
.pdf-page {background:white; border:1px solid #D0D5DD; padding:16px; min-height:680px;}
.pdf-header {display:flex; align-items:center; justify-content:space-between; border-bottom:1px solid #98A2B3; padding-bottom:12px; margin-bottom:12px;}
.pdf-title {font-size:18px; font-weight:800; color:#111827; text-align:center;}
.doc-number {border:1px solid #98A2B3; padding:10px 16px; text-align:center; font-weight:800;}
.pdf-section {background:#102A56; color:white; font-weight:800; padding:5px 8px; font-size:12px; margin-top:12px;}
.pdf-table {width:100%; border-collapse:collapse; font-size:12px;}
.pdf-table th,.pdf-table td {border:1px solid #D0D5DD; padding:5px;}
.pdf-table th {background:#F2F4F7;}
.check {color:#169B62; font-weight:900;}
.history-card {background:white; border:1px solid #E5E7EB; border-radius:14px; padding:16px; margin-top:8px;}
</style>
""", unsafe_allow_html=True)

def obtener_siguiente_documento():
    if not EXCEL_FILE.exists(): return "ER-ST-001"
    try:
        wb = load_workbook(EXCEL_FILE)
        if "Registro General" not in wb.sheetnames: return "ER-ST-001"
        numeros=[]
        for row in wb["Registro General"].iter_rows(min_row=2, values_only=True):
            doc=row[0]
            if doc and str(doc).startswith("ER-ST-"):
                try: numeros.append(int(str(doc).replace("ER-ST-","")))
                except Exception: pass
        return f"ER-ST-{(max(numeros)+1 if numeros else 1):03d}"
    except Exception:
        return "ER-ST-001"

def crear_excel_si_no_existe():
    if EXCEL_FILE.exists(): return
    wb=Workbook(); ws1=wb.active; ws1.title="Registro General"
    ws1.append(["N° Documento","Fecha","Hora","Destino","Proyecto","Transportista","Entregado por","Cargo entrega","Recibido por","Cargo recibe","Observaciones","Archivo PDF"])
    ws2=wb.create_sheet("Detalle Materiales")
    ws2.append(["N° Documento","Material / Equipo","Cantidad","Unidad","Entregado","Recibido"])
    wb.save(EXCEL_FILE)

def guardar_excel(data, materiales, pdf_name):
    crear_excel_si_no_existe(); wb=load_workbook(EXCEL_FILE)
    wb["Registro General"].append([data["numero_documento"],data["fecha"],data["hora"],data["destino"],data["proyecto"],data["transportista"],data["entregado_por"],data["cargo_entrega"],data["recibido_por"],data["cargo_recibe"],data["observaciones"],pdf_name])
    ws2=wb["Detalle Materiales"]
    for m in materiales:
        ws2.append([data["numero_documento"],m.get("Material / Equipo",""),m.get("Cantidad",""),m.get("Unidad",""),"Sí" if m.get("Entregado",False) else "No","Sí" if m.get("Recibido",False) else "No"])
    wb.save(EXCEL_FILE)

def cargar_historial():
    if not EXCEL_FILE.exists(): return pd.DataFrame()
    try: return pd.read_excel(EXCEL_FILE, sheet_name="Registro General")
    except Exception: return pd.DataFrame()

def generar_qr(data):
    texto=f"N° Documento: {data['numero_documento']}\nProyecto: {data['proyecto']}\nDestino: {data['destino']}\nTransportista: {data['transportista']}\nFecha: {data['fecha']}"
    qr=qrcode.QRCode(box_size=3,border=2); qr.add_data(texto); qr.make(fit=True)
    img=qr.make_image(fill_color="black", back_color="white"); buf=BytesIO(); img.save(buf, format="PNG"); buf.seek(0); return buf

def logo_b64():
    import base64
    p=Path("logo.png")
    return base64.b64encode(p.read_bytes()).decode("utf-8") if p.exists() else ""

def generar_pdf(data, materiales):
    buf=BytesIO(); doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=1.3*cm,leftMargin=1.3*cm,topMargin=1.1*cm,bottomMargin=1.1*cm)
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloPDF",parent=styles["Title"],fontName="Helvetica-Bold",fontSize=14,textColor=colors.black,alignment=1,leading=18))
    styles.add(ParagraphStyle(name="TextoPDF",parent=styles["Normal"],fontName="Helvetica",fontSize=8,leading=10))
    styles.add(ParagraphStyle(name="SeccionPDF",parent=styles["Normal"],fontName="Helvetica-Bold",fontSize=9,textColor=colors.white))
    story=[]; logo=Image("logo.png",width=4.1*cm,height=2.0*cm) if Path("logo.png").exists() else Paragraph("SOLARTEAM", styles["TituloPDF"])
    enc=Table([[logo,Paragraph("FORMULARIO DE ENTREGA Y<br/>RECEPCIÓN DE EQUIPOS", styles["TituloPDF"]),Paragraph(f"<b>N° DOCUMENTO</b><br/>{data['numero_documento']}",styles["TextoPDF"])]],colWidths=[5*cm,8.4*cm,4.2*cm])
    enc.setStyle(TableStyle([("BOX",(0,0),(-1,-1),.8,colors.HexColor("#9AA4B2")),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(0,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),7)])); story += [enc, Spacer(1,8)]
    info=Table([[Paragraph(f"<b>Fecha:</b> {data['fecha']}",styles["TextoPDF"]),Paragraph(f"<b>Hora:</b> {data['hora']}",styles["TextoPDF"]),Image(generar_qr(data),width=2.4*cm,height=2.4*cm)]],colWidths=[7*cm,7*cm,3.4*cm])
    info.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"MIDDLE"),("ALIGN",(2,0),(2,0),"RIGHT")]))
    story += [info, Spacer(1,8), Table([[Paragraph("DATOS GENERALES",styles["SeccionPDF"])]],colWidths=[17.6*cm],style=[("BACKGROUND",(0,0),(-1,-1),colors.HexColor(PRIMARY)),("PADDING",(0,0),(-1,-1),5)])]
    datos=[["Destino / Área:",data["destino"],"Proyecto asignado:",data["proyecto"]],["Transportista:",data["transportista"],"",""] ,["Entregado por:",data["entregado_por"],"Cargo:",data["cargo_entrega"]],["Recibido por:",data["recibido_por"],"Cargo:",data["cargo_recibe"]],["Observaciones:",data["observaciones"] or "Sin observaciones.","",""]]
    td=Table(datos,colWidths=[3.2*cm,6*cm,3*cm,5.4*cm]); td.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#F1F5F9")),("BACKGROUND",(2,0),(2,-1),colors.HexColor("#F1F5F9")),("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),("FONTNAME",(2,0),(2,-1),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("SPAN",(1,1),(3,1)),("SPAN",(1,4),(3,4)),("PADDING",(0,0),(-1,-1),5)])); story += [td, Spacer(1,10)]
    story.append(Table([[Paragraph("DETALLE DE MATERIALES / EQUIPOS",styles["SeccionPDF"])]],colWidths=[17.6*cm],style=[("BACKGROUND",(0,0),(-1,-1),colors.HexColor(PRIMARY)),("PADDING",(0,0),(-1,-1),5)]))
    det=[["No.","Material / Equipo","Cantidad","Unidad","Entregado","Recibido"]]
    for i,m in enumerate(materiales,1): det.append([str(i),Paragraph(str(m.get("Material / Equipo","")),styles["TextoPDF"]),str(m.get("Cantidad","")),str(m.get("Unidad","")),"✓" if m.get("Entregado",False) else "","✓" if m.get("Recibido",False) else ""])
    tm=Table(det,colWidths=[1*cm,8.1*cm,2.2*cm,2.3*cm,2*cm,2*cm],repeatRows=1); tm.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F1F5F9")),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),8),("ALIGN",(0,0),(0,-1),"CENTER"),("ALIGN",(2,1),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("PADDING",(0,0),(-1,-1),5)])); story += [tm, Spacer(1,18)]
    firmas=Table([["ENTREGA","RECEPCIÓN"],["",""] ,[data["entregado_por"],data["recibido_por"]],[data["cargo_entrega"],data["cargo_recibe"]],["Nombre y firma","Nombre y firma"]],colWidths=[8.8*cm,8.8*cm],rowHeights=[.6*cm,2*cm,.5*cm,.5*cm,.5*cm])
    firmas.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.45,colors.HexColor("#BFC7D5")),("BACKGROUND",(0,0),(-1,0),colors.HexColor(PRIMARY)),("TEXTCOLOR",(0,0),(-1,0),colors.white),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("LINEABOVE",(0,2),(0,2),.7,colors.black),("LINEABOVE",(1,2),(1,2),.7,colors.black),("FONTSIZE",(0,0),(-1,-1),8)])); story += [firmas, Spacer(1,12)]
    pie=Table([[Paragraph("Este documento certifica la entrega y recepción de los equipos y/o materiales descritos.",styles["TextoPDF"])]],colWidths=[17.6*cm]); pie.setStyle(TableStyle([("LINEABOVE",(0,0),(-1,-1),1,colors.HexColor(GREEN)),("ALIGN",(0,0),(-1,-1),"CENTER"),("PADDING",(0,0),(-1,-1),8)])); story.append(pie)
    doc.build(story); buf.seek(0); return buf

def pdf_preview_html(data, materiales):
    rows="".join([f"<tr><td>{i}</td><td>{m.get('Material / Equipo','')}</td><td>{m.get('Cantidad','')}</td><td>{m.get('Unidad','')}</td><td class='check'>{'✓' if m.get('Entregado',False) else ''}</td><td class='check'>{'✓' if m.get('Recibido',False) else ''}</td></tr>" for i,m in enumerate(materiales[:6],1)]) or "<tr><td colspan='6'>Sin materiales registrados</td></tr>"
    return f"""<div class='preview-box'><div class='pdf-page'><div class='pdf-header'><img src='data:image/png;base64,{logo_b64()}' style='width:125px;'><div class='pdf-title'>FORMULARIO DE ENTREGA Y<br>RECEPCIÓN DE EQUIPOS</div><div class='doc-number'>N° DOCUMENTO<br>{data['numero_documento']}</div></div><div style='display:flex;justify-content:space-between;align-items:start;'><div style='font-size:12px;'><b>Fecha:</b> {data['fecha']}</div><div style='font-size:12px;'><b>Hora:</b> {data['hora']}</div><div style='width:74px;height:74px;border:1px solid #D0D5DD;text-align:center;font-size:11px;display:flex;align-items:center;justify-content:center;'>QR</div></div><div class='pdf-section'>DATOS GENERALES</div><table class='pdf-table'><tr><td><b>Destino / Área:</b></td><td>{data['destino']}</td><td><b>Proyecto:</b></td><td>{data['proyecto']}</td></tr><tr><td><b>Transportista:</b></td><td colspan='3'>{data['transportista']}</td></tr><tr><td><b>Entregado por:</b></td><td>{data['entregado_por']}</td><td><b>Cargo:</b></td><td>{data['cargo_entrega']}</td></tr><tr><td><b>Recibido por:</b></td><td>{data['recibido_por']}</td><td><b>Cargo:</b></td><td>{data['cargo_recibe']}</td></tr><tr><td><b>Observaciones:</b></td><td colspan='3'>{data['observaciones']}</td></tr></table><div class='pdf-section'>DETALLE DE MATERIALES / EQUIPOS</div><table class='pdf-table'><tr><th>No.</th><th>Material / Equipo</th><th>Cantidad</th><th>Unidad</th><th>Entregado</th><th>Recibido</th></tr>{rows}</table><br><table class='pdf-table'><tr style='background:#102A56;color:white;'><th>ENTREGA</th><th>RECEPCIÓN</th></tr><tr><td style='height:70px;text-align:center;'>____________________<br>{data['entregado_por']}<br>{data['cargo_entrega']}</td><td style='height:70px;text-align:center;'>____________________<br>{data['recibido_por']}<br>{data['cargo_recibe']}</td></tr></table><p style='border-top:2px solid #169B62;text-align:center;font-size:12px;padding-top:8px;'>Este documento certifica la entrega y recepción de los equipos descritos.</p></div></div>"""

with st.sidebar:
    if Path("logo.png").exists(): st.image("logo.png", use_container_width=True)
    st.markdown('<div class="sidebar-btn sidebar-btn-active">📄 &nbsp; Nuevo Formulario</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-btn">🕘 &nbsp; Historial de Entregas</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-btn">📦 &nbsp; Catálogo de Materiales</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-btn">📊 &nbsp; Reportes</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-btn">⚙️ &nbsp; Configuración</div>', unsafe_allow_html=True)
    st.markdown('<div class="help-box"><b>¿Necesitas ayuda?</b><br>Contacta al administrador del sistema.</div>', unsafe_allow_html=True)

numero_documento=obtener_siguiente_documento(); ahora=datetime.now()
top1,top2=st.columns([5,1])
with top1:
    st.markdown('<div class="main-title">Entrega y Recepción de Equipos</div>', unsafe_allow_html=True)
    st.markdown('<div class="subtitle">Complete la información para generar el documento.</div>', unsafe_allow_html=True)
with top2:
    st.markdown(f"<div class='small-muted' style='text-align:right;'>📅 {date.today().strftime('%d/%m/%Y')} &nbsp; {ahora.strftime('%H:%M')}</div>", unsafe_allow_html=True)

if "materiales" not in st.session_state:
    st.session_state.materiales=[{"Material / Equipo":"Panel Solar Jinko 625W","Cantidad":50,"Unidad":"Unidad","Entregado":True,"Recibido":True},{"Material / Equipo":"Inversor Huawei 100KTL-M1","Cantidad":2,"Unidad":"Unidad","Entregado":True,"Recibido":True},{"Material / Equipo":"Cable Solar 6mm²","Cantidad":200,"Unidad":"Metro","Entregado":True,"Recibido":True},{"Material / Equipo":"Estructura Grace Solar","Cantidad":10,"Unidad":"Unidad","Entregado":True,"Recibido":True}]

left,right=st.columns([1.05,1.15], gap="large")
with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">👷 1. DATOS GENERALES</div>', unsafe_allow_html=True)
    c1,c2=st.columns(2)
    with c1:
        destino=st.text_input("Destino / Área *", value="Obra - Zona Norte")
        proyecto=st.text_input("Proyecto asignado *", value="Fibroacero")
        transportista=st.text_input("Transportista *", value="Transportes XYZ S.A.")
        entregado_por=st.text_input("Entregado por *", value="Juan Pérez")
        cargo_entrega=st.text_input("Cargo", value="Almacén")
    with c2:
        recibido_por=st.text_input("Recibido por *", value="Carlos Ruiz")
        cargo_recibe=st.text_input("Cargo ", value="Residente de Obra")
        fecha=st.date_input("Fecha *", value=date.today())
        observaciones=st.text_area("Observaciones", value="Entrega conforme a lo solicitado para avance de obra.", height=125)
    st.divider(); st.markdown('<div class="section-title">📋 2. MATERIALES / EQUIPOS</div>', unsafe_allow_html=True)
    materiales=st.data_editor(st.session_state.materiales,num_rows="dynamic",use_container_width=True,hide_index=True,column_config={"Material / Equipo":st.column_config.SelectboxColumn("Material / Equipo",options=CATALOGO_MATERIALES),"Cantidad":st.column_config.NumberColumn("Cantidad",min_value=0.0,step=1.0),"Unidad":st.column_config.SelectboxColumn("Unidad",options=UNIDADES),"Entregado":st.column_config.CheckboxColumn("Entregado"),"Recibido":st.column_config.CheckboxColumn("Recibido")})
    material_otro=st.text_input("Si seleccionaste 'Otro', escribe el producto nuevo")
    if material_otro.strip():
        for m in materiales:
            if m.get("Material / Equipo")=="Otro": m["Material / Equipo"]=material_otro.strip()
    st.session_state.materiales=materiales
    data={"numero_documento":numero_documento,"fecha":fecha.strftime("%d/%m/%Y"),"hora":ahora.strftime("%H:%M"),"destino":destino,"proyecto":proyecto,"transportista":transportista,"entregado_por":entregado_por,"cargo_entrega":cargo_entrega,"recibido_por":recibido_por,"cargo_recibe":cargo_recibe,"observaciones":observaciones}
    materiales_validos=[m for m in materiales if str(m.get("Material / Equipo","")).strip() not in ["","None"]]
    generar=st.button("⬇️ Generar PDF y Guardar", type="primary", use_container_width=True)
    if generar:
        if not all(str(c).strip() for c in [destino,proyecto,transportista,entregado_por,recibido_por]): st.warning("Completa destino, proyecto, transportista, entregado por y recibido por.")
        elif not materiales_validos: st.warning("Agrega al menos un material o equipo.")
        else:
            pdf=generar_pdf(data, materiales_validos); pdf_name=f"{numero_documento}.pdf"; pdf_path=PDF_DIR/pdf_name; pdf_path.write_bytes(pdf.getvalue()); guardar_excel(data,materiales_validos,str(pdf_path)); st.success(f"Formulario {numero_documento} generado y guardado correctamente."); st.download_button("Descargar PDF", pdf.getvalue(), pdf_name, "application/pdf", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">VISTA PREVIA DEL DOCUMENTO</div>', unsafe_allow_html=True)
    st.markdown(pdf_preview_html(data, materiales_validos), unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="history-card">', unsafe_allow_html=True)
hcol1,hcol2=st.columns([4,1])
with hcol1: st.markdown('<div class="section-title">🧾 HISTORIAL DE ENTREGAS</div>', unsafe_allow_html=True)
with hcol2:
    if EXCEL_FILE.exists():
        with open(EXCEL_FILE,"rb") as f: st.download_button("📗 Exportar Excel", f, "registro_entregas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
historial=cargar_historial()
if historial.empty:
    ejemplo=pd.DataFrame([["ER-ST-001","02/07/2026","Fibroacero","Obra - Zona Norte","Transportes XYZ S.A.","Juan Pérez","Carlos Ruiz"],["ER-ST-002","01/07/2026","Pronaca - Quito","Obra - Quito","Transporte Andino","María López","Diego Salazar"],["ER-ST-003","30/06/2026","Holcim Ecuador","Bodega - Guayaquil","Logística Express","Pedro García","Kevin Mera"]],columns=["N° Documento","Fecha","Proyecto","Destino","Transportista","Entregado por","Recibido por"])
    st.dataframe(ejemplo,use_container_width=True,hide_index=True)
else:
    busqueda=st.text_input("Buscar por proyecto, destino o N° de documento")
    if busqueda.strip(): historial=historial[historial.astype(str).apply(lambda row: row.str.contains(busqueda,case=False,na=False).any(), axis=1)]
    st.dataframe(historial,use_container_width=True,hide_index=True)
st.markdown('</div>', unsafe_allow_html=True)
