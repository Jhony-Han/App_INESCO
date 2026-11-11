import streamlit as st
import pandas as pd
import pdfplumber
import io
import re
import json
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import urllib.parse

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Inesco | Gestión y Extracción",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    .cocacola-header {
        background: linear-gradient(135deg, #E41E2B 0%, #B3000C 100%);
        padding: 20px;
        border-radius: 15px;
        text-align: center;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0px 6px 15px rgba(228, 30, 43, 0.4);
    }
    .cocacola-header h1 { 
        color: white !important; 
        margin: 0; 
        font-weight: 800; 
        font-size: 1.8rem;
        letter-spacing: 1px;
    }
    .cocacola-header p { 
        color: #FFEBEE !important; 
        margin: 5px 0 0 0; 
        font-size: 0.95rem; 
    }
    .sub-title {
        color: #E41E2B; 
        font-weight: 700; 
        font-size: 1.2rem;
        margin-top: 15px; 
        margin-bottom: 10px;
        border-bottom: 2px solid #E41E2B; 
        padding-bottom: 5px;
    }
    .stButton>button {
        background-color: #E41E2B !important; 
        color: white !important;
        font-weight: bold !important; 
        border-radius: 8px !important; 
        width: 100%;
        padding: 10px 0px;
    }
    </style>
""", unsafe_allow_html=True)

# --- CREDENCIALES DE ACCESO ---
ADMIN_PASSWORD = "inesco2026admin"
DRIVER_PASSWORD = "ruta2026inesco"

col_h1, col_h2, col_h3 = st.columns([1, 6, 1])
with col_h1:
    st.markdown("<h1 style='text-align: center; font-size: 2.5rem; margin-top: 20px;'>🥤</h1>", unsafe_allow_html=True)
with col_h2:
    st.markdown("""
        <div class="cocacola-header">
            <h1>DISTRIBUCIONES INESCO</h1>
            <p>Gestión de Inventario, Vencimientos y Extracción por Rutas</p>
        </div>
    """, unsafe_allow_html=True)
with col_h3:
    st.markdown("<h1 style='text-align: center; font-size: 2.5rem; margin-top: 20px;'>🚚</h1>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. GESTIÓN DE DATOS EN SESIÓN Y CATÁLOGO
# ---------------------------------------------------------
if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = []

if "modo_captura" not in st.session_state:
    st.session_state.modo_captura = "Tomar datos manual"

if "voz_temp_input" not in st.session_state:
    st.session_state.voz_temp_input = ""

if "pedidos_autoventa" not in st.session_state:
    st.session_state.pedidos_autoventa = {}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "rol_usuario" not in st.session_state:
    st.session_state.rol_usuario = None

CATALOGO_INICIAL = [
    {"sku": "120180", "descripcion": "MONSTER VIOLET 473ML LAT(12) MQ"},
    {"sku": "120208", "descripcion": "POWERADE CONTRATAQUE 500ML PET(6)"},
    {"sku": "136665", "descripcion": "6PK CCSO 269ML LAT NVO"},
    {"sku": "136705", "descripcion": "TRIDENT 5SX18U MENTA 153G KIT30(1)"},
    {"sku": "136707", "descripcion": "TRIDENT 1SX60U MENTA 90G KIT40(1)"},
    {"sku": "136710", "descripcion": "CLUB SOCIAL 9U ORIG 216G KIT24(1)"},
    {"sku": "120215", "descripcion": "FURY LAT 473ML(12)"},
    {"sku": "120226", "descripcion": "6PK(3MG+3MF)188ML"},
    {"sku": "136783", "descripcion": "OREO REGULAR 540G KIT12(1)"},
    {"sku": "120214", "descripcion": "MONSTER VICE GUAVA LAT 473ML(12)"},
    {"sku": "92892", "descripcion": "BRISA LIMA LIMON 600ML PET (6)NVO"},
    {"sku": "56452", "descripcion": "SCHWEPPES SODA 400ML PET(12)"},
    {"sku": "56497", "descripcion": "SCHWEPPES GINGER 10OZ VNR(12)NvaB"},
    {"sku": "56498", "descripcion": "SCHWEPPES SODA 10OZ VNR(12)NvaB"},
    {"sku": "56521", "descripcion": "SPRITE SO 8OZ VIR (30)"},
    {"sku": "56557", "descripcion": "SCHWEPPES SODA 8OZ RGB (30)"},
    {"sku": "56622", "descripcion": "PREMIO ROJO FNT 350ML VIR(30)"},
    {"sku": "56624", "descripcion": "PREMIO ROJO FNT 8OZ VIR(30)"},
    {"sku": "56625", "descripcion": "PREMIO ROJO FNT 3LTS PET (6)"},
    {"sku": "56627", "descripcion": "PREMIO ROJO 400ML PET"},
    {"sku": "56651", "descripcion": "PREMIO ROJO FNT 1.5LT PET"},
    {"sku": "56703", "descripcion": "QUATRO CHOICE 250ML PET (12)"},
    {"sku": "56704", "descripcion": "QUATRO CHOICE 8OZ VIR (30)"},
    {"sku": "56705", "descripcion": "QUATRO CHOICE 350ML VIR(30)"},
    {"sku": "56706", "descripcion": "QUATRO CHOICE 400ML NR (12)"},
    {"sku": "56709", "descripcion": "QUATRO 1.5 L"},
    {"sku": "56711", "descripcion": "QUATRO CHOICE 3LTS PET (6)"},
    {"sku": "56820", "descripcion": "QUATRO CHOICE 2.5L PET (8)"},
    {"sku": "56841", "descripcion": "FLASHLYTE UVA 625ML PET(6)"},
    {"sku": "56842", "descripcion": "FLASHLYTE COCO LIM 625ML PET(6)"},
    {"sku": "92764", "descripcion": "MANANTIAL SIN GAS 500ML VNR(12)"},
    {"sku": "92768", "descripcion": "BRISA MANZANA 288ML"},
    {"sku": "92769", "descripcion": "BRISA MANZANA 600ML"},
    {"sku": "92770", "descripcion": "BRISA MANZANA 1.5 L"},
    {"sku": "92793", "descripcion": "BRISA MARACUYA 280ML PET(24)"},
    {"sku": "92794", "descripcion": "BRISA MARACUYA 600ML PET(6)"},
    {"sku": "92795", "descripcion": "BRISA MARACUYA 1.5LT PET(12)"},
    {"sku": "92893", "descripcion": "AGUA BRISA GAS LIMA LIMON PET 280 ML (24)"},
    {"sku": "92911", "descripcion": "BRISA LIMA LIMON 1.5L PET(12)NV"},
    {"sku": "120116", "descripcion": "FRESH CITRUS 188ML TRP(24)NVO"},
    {"sku": "120117", "descripcion": "FRESH CITRUS 400ML PET (12)NVO"},
    {"sku": "120118", "descripcion": "DEL VALLE CITRUS 1.5 LTS PET (12)"},
    {"sku": "120119", "descripcion": "FRESH CITRUS 2,5LT"},
    {"sku": "120120", "descripcion": "2PK FRESH CITRUS 2.5L PET NVO"},
    {"sku": "120128", "descripcion": "FRUTAL MANGOFRS 500ML PET(6)NVO"},
    {"sku": "120129", "descripcion": "FRUTAL MORA 500ML PET(6)NVO"},
    {"sku": "120130", "descripcion": "FRUTAL SALPICON 500ML PET(6)NVO"},
    {"sku": "120131", "descripcion": "FRUTAL MANGO 500ML PET(6)NVO"},
    {"sku": "120134", "descripcion": "FRUTAL MORA 188ML TRP(24)NVO"},
    {"sku": "120135", "descripcion": "FRUTAL SALPICON 188ML TRP(24)NV"},
    {"sku": "120136", "descripcion": "FRUTAL MANGO 188ML TRP(24)"},
    {"sku": "120156", "descripcion": "MONSTER GREEN 473ML LAT(12) MQ"},
    {"sku": "120158", "descripcion": "FRESH MANDARINA 400ML(12)NVO"},
    {"sku": "120159", "descripcion": "FRESH MANDARINA 1.5LT PET(12)NVO"},
    {"sku": "120175", "descripcion": "MONSTER PIPELINE PUNCH 473ML LA"},
    {"sku": "120181", "descripcion": "MONSTER MGLOCO 473ML LAT(12) MQ"},
    {"sku": "120182", "descripcion": "MONSTER PARADISE 473ML LAT(12) MQ"},
    {"sku": "120183", "descripcion": "MONSTER ULTRA 473ML LAT(12) MQ"},
    {"sku": "120184", "descripcion": "FRESH PONCHE FRUTAS 400ML PET(12)"},
    {"sku": "120188", "descripcion": "FRUTAL MANGO 1.2L PET (6)"},
    {"sku": "120190", "descripcion": "FRUTAL MANGOFRS 1.2LT PET(6)"},
    {"sku": "120191", "descripcion": "FRUTAL MORA 1.2LT PET(6)"},
    {"sku": "120194", "descripcion": "6PK(2MG+2MR+2MF)188ML"},
    {"sku": "120205", "descripcion": "MONSTER GREEN ZERO LAT 473ML(12)"},
    {"sku": "135662", "descripcion": "COCA COLA 3 LTS PET (6)"},
    {"sku": "135663", "descripcion": "CC PET 2.5LTNR 8 PK  Nvo"},
    {"sku": "135664", "descripcion": "COCA-COLA 1.5 LT NR (12) Nvo"},
    {"sku": "135675", "descripcion": "COMBO AHORRO CCSO 3LT"},
    {"sku": "135718", "descripcion": "COCA-COLA  8OZ VIR(30)"},
    {"sku": "135755", "descripcion": "COCA COLA ZERO 2.5 LTS (8)"},
    {"sku": "135757", "descripcion": "COCA COLA ZERO 1,5 NR (12)"},
    {"sku": "135760", "descripcion": "COCA-COLA SA 400ml PET(12)"},
    {"sku": "135761", "descripcion": "CC ZERO 250ML (12) CON PRECIO"},
    {"sku": "135763", "descripcion": "COCA-COLA ZERO-SA 350ML VIR(30)"},
    {"sku": "135764", "descripcion": "COCA-COLA SA 8OZ VIR(30)"},
    {"sku": "135766", "descripcion": "CC SA 10onz"},
    {"sku": "135795", "descripcion": "6PK CCZR 250 ML PET KZ"},
    {"sku": "136020", "descripcion": "BOCAD ATUN ACEIT 140GR KIT48(1)"},
    {"sku": "136024", "descripcion": "ATUN RALL ACEITE 150GR KIT48(1)"},
    {"sku": "136027", "descripcion": "BOCADO ATUN AGUA 140GR KIT18(1)"},
    {"sku": "136113", "descripcion": "ESPEJO CCSO 1.5LT PET(12)"},
    {"sku": "136133", "descripcion": "COCA COLA 2.25L PET(8)EINSTEIN"},
    {"sku": "136159", "descripcion": "COCA COLA SO 1L PET (12)"},
    {"sku": "136247", "descripcion": "2 PK (1CCSO 3LT + 1 QT3LT) Nvo"},
    {"sku": "136399", "descripcion": "COCA COLA ORIGINAL 269ML LAT(12)"},
    {"sku": "136400", "descripcion": "COCA COLA ZERO 269ML LAT(12)KZ"},
    {"sku": "136401", "descripcion": "CCZR CAFE 269ML LAT(12)"},
    {"sku": "136403", "descripcion": "6PK LAT CCZR KZ 269ML"},
    {"sku": "136461", "descripcion": "ZUCARITAS MEGA 115G KIT54(1)"},
    {"sku": "136462", "descripcion": "CHOKRISPIS MEGA 115G KIT54(1)"},
    {"sku": "136465", "descripcion": "FROOT LOOPS MEGA 90GR KIT54(1)"},
    {"sku": "136466", "descripcion": "OREO ROLLO 109GR KIT30(1)"},
    {"sku": "136467", "descripcion": "OREO REGULAR COL 324GR KIT18(1)"},
    {"sku": "136468", "descripcion": "HALLS 100S EXT STRONG COL 310GR KIT30(1)"},
    {"sku": "136474", "descripcion": "ATUN ISABEL ACEI 142G KIT48(1)"},
    {"sku": "136475", "descripcion": "BARRILETE 440GR PAIS KIT18(1)"},
    {"sku": "136478", "descripcion": "CHAO MENTA 350GR KIT24(1)"},
    {"sku": "136479", "descripcion": "CHAO FRESA 350GR KIT24(1)"},
    {"sku": "136559", "descripcion": "ZUCARITAS 240GR KIT30(1)"},
    {"sku": "136560", "descripcion": "CHOCO KRISPIS 240GR KIT30(1)"},
    {"sku": "136561", "descripcion": "FROOT LOOPS 200GR KIT30(1)"},
    {"sku": "136630", "descripcion": "TRULU MINIPACK 24U 960GR KIT1(1)"},
    {"sku": "136632", "descripcion": "TRUL MAXIPACK 12U 1.2K KIT1(1)"},
    {"sku": "136670", "descripcion": "TRUL MXPK 12U 1.3K TOPEX KIT6(1)"},
    {"sku": "136671", "descripcion": "TRUL MAXI+VEND 12U 1.2K KIT6(1)"},
    {"sku": "136675", "descripcion": "TRUL SABOR SURT 12U 360G KIT12(1)"},
    {"sku": "136691", "descripcion": "EXH TRUL 72UX70G 6.7K KIT1(1)"},
    {"sku": "136700", "descripcion": "BIANCHI CHOCOLA 600G KIT18(1)"},
    {"sku": "136701", "descripcion": "BIANCHI CHOCOBLA 600G KIT18(1)"},
    {"sku": "136712", "descripcion": "TRULU 70U MAS VENDID 1.6K KIT6(1)"},
    {"sku": "136747", "descripcion": "CHOCOLORES SURT 540G KIT12(1)"},
    {"sku": "160008", "descripcion": "BRISA ECOFLEX 1LT"},
    {"sku": "160029", "descripcion": "COCA-COLA 330ML LATA (12)"},
    {"sku": "160033", "descripcion": "SCHWEPPES SODA 350 ML VIR (30)"},
    {"sku": "160046", "descripcion": "AGUA BRISA GAS PET 600ML (24)"},
    {"sku": "160047", "descripcion": "AGUA BRISA PET 600ML (24)"},
    {"sku": "160048", "descripcion": "AGUA BRISA BOTELLON 5 GL TV"},
    {"sku": "160053", "descripcion": "COCA-COLA 350ML VIR(30)"},
    {"sku": "160060", "descripcion": "MANANTIAL 1.5 PET(9)"},
    {"sku": "160070", "descripcion": "SPRITE LL 350ML VIR(30)"},
    {"sku": "160099", "descripcion": "KOLA ROMAN 1.5L"},
    {"sku": "160101", "descripcion": "SCHWEPPES SODA 1.5 L (12)"},
    {"sku": "160102", "descripcion": "SPRITE 1.5L (12)"},
    {"sku": "160121", "descripcion": "POWERADE ROJO FT 500ML PET(6)"},
    {"sku": "160122", "descripcion": "POWERADE AZUL MB 500ML PET(6)"},
    {"sku": "160124", "descripcion": "SPRITE 3 LTS PET (6)"},
    {"sku": "160168", "descripcion": "AGUA BRISA PET 280 ML (24)"},
    {"sku": "160186", "descripcion": "FUZE TEA DURAZNO PET 400ML  (6)"},
    {"sku": "160187", "descripcion": "FUZE TEA LIMON PET 400 (6)"},
    {"sku": "160192", "descripcion": "COCACOLA 10OZ VNR (12) VIDRIO"},
    {"sku": "160200", "descripcion": "COCA COLA 250 ML (12)"},
    {"sku": "160201", "descripcion": "6PACK CCR 250ML MINI PET"},
    {"sku": "160240", "descripcion": "MANANTIAL GAS 600 ML (12)"},
    {"sku": "160241", "descripcion": "MANANTIAL 600 ML (12)"},
    {"sku": "160297", "descripcion": "SPRITE 400PET"},
    {"sku": "160318", "descripcion": "COCA-COLA 400ML PET  (12)"},
    {"sku": "160479", "descripcion": "KOLA ROMAN 400 ML NR (12)"},
    {"sku": "92899", "descripcion": "REFAJO KOLA ROMAN 330M LAT(6)"},
    {"sku": "160151", "descripcion": "AGUA BRISA BIDON 6 LT (2)"},
    {"sku": "92740", "descripcion": "BRISA LIMA LIMÓN 1.5LT PET (12)"},
    {"sku": "160170", "descripcion": "MANANTIAL 1.5 LTS PET (9)"},
    {"sku": "120189", "descripcion": "FRUTAL SALPICON 1.2L PET (6)"},
    {"sku": "136475", "descripcion": "LOKINO CARAMELO 400GR KIT24(1)"},
    {"sku": "136720", "descripcion": "TRUL MINIPACK 24U 840GR KIT12(1)"},
    {"sku": "160061", "descripcion": "MANANTIAL 600ML (12)"},
    {"sku": "160242", "descripcion": "MANANTIAL BIDON 6 LT (2)"},
    {"sku": "56445", "descripcion": "QUATRO CHOI 400NR (12)"},
    {"sku": "56450", "descripcion": "SPRITE 400NR (12)"},
    {"sku": "56455", "descripcion": "COCA-COLA 400 NR (12)"},
    {"sku": "56456", "descripcion": "COCA-COLA ZERO 400NR (12)"},
    {"sku": "56645", "descripcion": "KOLA ROMAN 400 NR (12)"},
    {"sku": "56650", "descripcion": "PREMIO ROJO 400 NR (12)"},
    {"sku": "120121", "descripcion": "2PK FRESH CITRUS 2.5LT (4)"},
    {"sku": "120157", "descripcion": "FRESH CITRUS 1.5L PET (12)"},
    {"sku": "120186", "descripcion": "FRESH MANDARINA 1.2 LT (6)"},
    {"sku": "120187", "descripcion": "FRESH MANDARINA 1.2L (6)"},
    {"sku": "135676", "descripcion": "COMBO AHORRO CCSO 3L (6)"},
    {"sku": "135756", "descripcion": "COCOA-COLA ZERO 2.25L (8)"},
    {"sku": "135762", "descripcion": "COCA COLA 250ML (12)"},
    {"sku": "136114", "descripcion": "COCA-COLA 1.5L (12)"},
    {"sku": "136134", "descripcion": "COCA-COLA 2.25L (8)"},
    {"sku": "136160", "descripcion": "COCA COLA S/AZUCAR 1L (12)"},
    {"sku": "136402", "descripcion": "COCA-COLA ZERO CAFE 269ML (12)"},
    {"sku": "160009", "descripcion": "BRISA ECOFLEX 600ML (24)"},
    {"sku": "160049", "descripcion": "BRISA BOTELLON 5 GL"},
    {"sku": "160059", "descripcion": "MANANTIAL 600ML (12)"},
    {"sku": "160062", "descripcion": "MANANTIAL 3L (6)"},
    {"sku": "160167", "descripcion": "BRISA 280ML (24)"},
    {"sku": "160169", "descripcion": "BRISA CON GAS 280ML (24)"},
    {"sku": "160190", "descripcion": "COCA-COLA 10OZ (12)"},
    {"sku": "160191", "descripcion": "COCA COLA ZERO 10OZ (12)"}
]

if "catalogo" not in st.session_state:
    st.session_state.catalogo = CATALOGO_INICIAL

# ---------------------------------------------------------
# 3. EXPORTADOR EXCEL PROFESIONAL (RUTAS Y VENCIMIENTOS)
# ---------------------------------------------------------
def exportar_excel_multiruta(rutas_dict, fecha_str, es_reporte_rutas=False):
    wb = Workbook()
    default_sheet = wb.active
    
    red_title_font = Font(color="FFFFFF", bold=True, size=13, name="Calibri")
    title_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    sub_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    sub_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(bold=True, size=11, name="Calibri", color="000000")
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
    )
    thick_bottom = Border(
        left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'), bottom=Side(style='double', color='000000')
    )
    
    first_sheet = True
    for nombre_ruta, r_info in rutas_dict.items():
        safe_title = re.sub(r'[\\/*?:[\]]', '_', nombre_ruta)
        ws = default_sheet if first_sheet else wb.create_sheet(title=safe_title[:30])
        first_sheet = False
            
        df_r = r_info["df"]
        header_text = r_info["header"]
        num_cols = len(df_r.columns)
        titulo_reporte = "DISTRIBUCIONES INESCO - REPORTE TOTAL DE RUTA" if es_reporte_rutas else "DISTRIBUCIONES INESCO - REPORTE DE VENCIMIENTOS"
        
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(num_cols, 3))
        cell_t = ws.cell(row=1, column=1, value=titulo_reporte)
        cell_t.font = red_title_font
        cell_t.fill = title_fill
        cell_t.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, max(num_cols, 3) + 1):
            ws.cell(row=1, column=col).border = thin_border
            ws.cell(row=1, column=col).fill = title_fill
        
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(num_cols, 3))
        cell_s = ws.cell(row=2, column=1, value=f"📌 {header_text}    (Fecha de Exportación: {fecha_str})")
        cell_s.font = sub_font
        cell_s.fill = sub_fill
        cell_s.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(1, max(num_cols, 3) + 1):
            ws.cell(row=2, column=col).border = thin_border
            ws.cell(row=2, column=col).fill = sub_fill
        
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 24
        ws.row_dimensions[3].height = 10
        ws.row_dimensions[4].height = 24
        
        for col_idx, col_name in enumerate(df_r.columns, start=1):
            c = ws.cell(row=4, column=col_idx, value=col_name)
            c.fill = header_fill
            c.font = header_font
            c.border = thin_border
            c.alignment = Alignment(horizontal="center", vertical="center")
            
        for row_idx, row_data in enumerate(df_r.values, start=5):
            ws.row_dimensions[row_idx].height = 20
            is_even = (row_idx % 2 == 0)
            row_fill = zebra_fill if is_even else white_fill
            
            for col_idx, val in enumerate(row_data, start=1):
                c = ws.cell(row=row_idx, column=col_idx)
                c.border = thin_border
                c.fill = row_fill
                c.font = Font(name="Calibri", size=11)
                
                col_header_name = str(df_r.columns[col_idx-1]).lower()
                if "cajas" in col_header_name or "unidades" in col_header_name or "cantidad" in col_header_name:
                    try:
                        c.value = int(val) if val != "" and val is not None else 0
                    except:
                        c.value = 0
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.number_format = '#,##0'
                else:
                    c.value = val
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    
        if es_reporte_rutas and not df_r.empty:
            last_row = 4 + len(df_r)
            total_row_idx = last_row + 1
            ws.row_dimensions[total_row_idx].height = 22
            
            for col_idx in range(1, num_cols + 1):
                c = ws.cell(row=total_row_idx, column=col_idx)
                c.fill = total_fill
                c.font = total_font
                c.border = thick_bottom
                
                col_header_name = str(df_r.columns[col_idx-1]).lower()
                if col_idx == 1:
                    c.value = "TOTALES"
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif "cajas" in col_header_name or "unidades" in col_header_name or "cantidad" in col_header_name:
                    col_letter = get_column_letter(col_idx)
                    c.value = f"=SUM({col_letter}5:{col_letter}{last_row})"
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.number_format = '#,##0'
                else:
                    c.value = ""

        for col_idx in range(1, num_cols + 1):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for row in range(4, ws.max_row + 1):
                cell_val = ws.cell(row=row, column=col_idx).value
                if cell_val:
                    max_len = max(max_len, len(str(cell_val)))
            ws.column_dimensions[col_letter].width = max(max_len + 8, 25)
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 4. FUNCIÓN DE EXTRACCIÓN DE PDF DE RUTAS
# ---------------------------------------------------------
def procesar_pdf_rutas(pdf_file):
    routes_data = {}
    mapa_catalogo = {item["sku"]: item["descripcion"] for item in st.session_state.catalogo}
    
    with pdfplumber.open(pdf_file) as pdf:
        current_route = "ML3E51"
        current_header_info = "Ruta / No.de Carga: ML3E51"
        
        for page_idx, page in enumerate(pdf.pages, start=1):
            texto_pagina = page.extract_text()
            if not texto_pagina:
                continue
                
            lineas = texto_pagina.split('\n')
            for linea in lineas:
                l_str = linea.strip()
                if not l_str:
                    continue
                
                m_ruta = re.search(r'ML3E5[1-3]', l_str)
                if m_ruta:
                    current_route = m_ruta.group(0)
                
                if "Ruta" in l_str or "No.de Carga" in l_str or "Fecha de Entrega" in l_str:
                    current_header_info = l_str
                    continue 
                
                m_sku = re.search(r'\b(\d{5,6})\b', l_str)
                if m_sku:
                    sku = m_sku.group(1)
                    cajas = 0
                    unidades = 0
                    
                    m_slash = re.search(r'(?:^|\s)(\d+)?\s*/\s*(\d+)', l_str)
                    if m_slash:
                        c_str = m_slash.group(1)
                        u_str = m_slash.group(2)
                        cajas = int(c_str) if c_str else 0
                        unidades = int(u_str) if u_str else 0
                    else:
                        m_qty = re.search(r'[\$\s]*(\d+)\s*[/\\-]\s*(\d+)', l_str)
                        if m_qty:
                            cajas = int(m_qty.group(1))
                            unidades = int(m_qty.group(2))
                    
                    desc_limpia = l_str.replace(sku, "")
                    if m_slash:
                        desc_limpia = desc_limpia.replace(m_slash.group(0), "")
                    
                    desc_limpia = re.sub(r'\s+\d+\s*[/\\-]\s*\d+\s*$', '', desc_limpia)
                    desc_limpia = re.sub(r'\s+[/\\-]\s*\d+\s*$', '', desc_limpia)
                    desc_limpia = re.sub(r'\b\d{5,7}\b\s*_+', '', desc_limpia)
                    desc_limpia = re.sub(r'\b\d{5,7}\b', '', desc_limpia)
                    desc_limpia = re.sub(r'_+', '', desc_limpia)
                    desc_limpia = re.sub(r'[\$\|\(\)]', '', desc_limpia).strip()
                    
                    descripcion_final = desc_limpia if len(desc_limpia) > 3 else mapa_catalogo.get(sku, "PRODUCTO FEMSA")
                    
                    if current_route not in routes_data:
                        routes_data[current_route] = {
                            "header": current_header_info,
                            "items": []
                        }
                    
                    routes_data[current_route]["items"].append({
                        "SKU": sku,
                        "Descripción del Producto": descripcion_final,
                        "Cajas": cajas,
                        "Unidades": unidades
                    })
                    
    return routes_data

# ---------------------------------------------------------
# 5. FUNCIONES PARA VENCIMIENTOS
# ---------------------------------------------------------
def calcular_estado_vencimiento(fecha_venc_str):
    try:
        f_venc = datetime.strptime(fecha_venc_str, "%d/%m/%Y").date()
        dias_restantes = (f_venc - date.today()).days
        if dias_restantes < 0:
            return "🔴 VENCIDO", dias_restantes
        elif dias_restantes <= 30:
            return "🚨 URGENTE (< 30 días)", dias_restantes
        elif dias_restantes <= 60:
            return "⚠️ PRÓXIMO (< 60 días)", dias_restantes
        else:
            return "✅ SEGURO", dias_restantes
    except:
        return "⚪ DESCONOCIDO", 9999

def exportar_excel_vencimientos_organizado(vencimientos_list, fecha_str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Vencimientos"
    
    red_title_font = Font(color="FFFFFF", bold=True, size=13, name="Calibri")
    title_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    
    black_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_rojo = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_rojo = Font(name="Calibri", size=11, color="9C0006", bold=True)
    fill_amarillo = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_amarillo = Font(name="Calibri", size=11, color="9C6500", bold=True)
    fill_verde = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_verde = Font(name="Calibri", size=11, color="006100")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    cell_t = ws.cell(row=1, column=1, value="DISTRIBUCIONES INESCO - CONTROL DE VENCIMIENTOS POR SECCIÓN")
    cell_t.font = red_title_font
    cell_t.fill = title_fill
    cell_t.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, 7):
        ws.cell(row=1, column=col).border = black_border
        ws.cell(row=1, column=col).fill = title_fill

    sub_fill = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
    cell_s = ws.cell(row=2, column=1, value=f"📌 Reporte Agrupado por Secciones (Fecha: {fecha_str})")
    cell_s.font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    cell_s.fill = sub_fill
    cell_s.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, 7):
        ws.cell(row=2, column=col).border = black_border
        ws.cell(row=2, column=col).fill = sub_fill

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 24

    headers = ["SKU", "Descripción del Producto", "Fecha Vencimiento", "Cantidad", "Estado de Alerta", "Días Restantes"]
    for col_idx, h_name in enumerate(headers, start=1):
        c = ws.cell(row=4, column=col_idx, value=h_name)
        c.fill = header_fill
        c.font = header_font
        c.border = black_border
        c.alignment = Alignment(horizontal="center", vertical="center")

    catalogo_actual = st.session_state.get("catalogo", CATALOGO_INICIAL)
    orden_catalogo = {item["sku"]: idx for idx, item in enumerate(catalogo_actual)}

    datos_procesados = []
    for item in vencimientos_list:
        estado, dias = calcular_estado_vencimiento(item["Fecha Vencimiento"])
        pos_catalogo = orden_catalogo.get(item["SKU"], 999)
        datos_procesados.append({
            "SKU": item["SKU"],
            "Descripción del Producto": item["Descripción del Producto"],
            "Fecha Vencimiento": item["Fecha Vencimiento"],
            "Cantidad": item.get("Cantidad", "-"),
            "Estado": estado,
            "Dias": dias,
            "PosCat": pos_catalogo
        })

    datos_ordenados = sorted(datos_procesados, key=lambda x: x["PosCat"])

    for row_idx, r_data in enumerate(datos_ordenados, start=5):
        ws.row_dimensions[row_idx].height = 20
        if r_data["Dias"] < 0 or r_data["Dias"] <= 30:
            f_fill, f_font = fill_rojo, font_rojo
        elif r_data["Dias"] <= 60:
            f_fill, f_font = fill_amarillo, font_amarillo
        else:
            f_fill, f_font = fill_verde, font_verde

        vals = [
            r_data["SKU"],
            r_data["Descripción del Producto"],
            r_data["Fecha Vencimiento"],
            r_data["Cantidad"],
            r_data["Estado"],
            r_data["Dias"] if r_data["Dias"] != 9999 else "N/A"
        ]

        for col_idx, val in enumerate(vals, start=1):
            c = ws.cell(row=row_idx, column=col_idx, value=val)
            c.border = black_border
            if col_idx in [1, 2, 3, 4]:
                c.fill = white_fill
                c.font = Font(name="Calibri", size=11, color="000000")
                if col_idx in [1, 3, 4]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            elif col_idx == 5:
                c.fill = f_fill
                c.font = f_font
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.fill = white_fill
                c.font = Font(name="Calibri", size=11, color="000000")
                c.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, 7):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or '')) for r in range(4, ws.max_row + 1))
        ws.column_dimensions[col_letter].width = max(max_len + 6, 20)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 6. PANTALLA DE ACCESO (LOGIN) CON CONTRASEÑA
# ---------------------------------------------------------
if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; font-size: 1.8rem; font-weight: bold; color: #E41E2B;">Distribuciones Inesco</p>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #666;">Sistema de Control Logístico, Bodega y Autoventa</p><br>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.container():
            st.markdown("### 🔐 Iniciar Sesión")
            pwd_input = st.text_input("Ingrese la contraseña de acceso:", type="password", key="input_pwd")
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("Ingresar", use_container_width=True):
                    if pwd_input == ADMIN_PASSWORD:
                        st.session_state.autenticado = True
                        st.session_state.rol_usuario = "admin"
                        st.rerun()
                    elif pwd_input == DRIVER_PASSWORD:
                        st.session_state.autenticado = True
                        st.session_state.rol_usuario = "driver"
                        st.rerun()
                    else:
                        st.error("Contraseña incorrecta.")
            with col_b2:
                if st.button("Limpiar", use_container_width=True):
                    st.rerun()
    st.stop()

# ---------------------------------------------------------
# 7. VISTA EXCLUSIVA PARA CONDUCTORES (AUTOVENTA Y CARGUES)
# ---------------------------------------------------------
if st.session_state.rol_usuario == "driver":
    st.markdown('<p class="sub-title">📱 Portal de Autoventa / Solicitud de Carga - Rutas Inesco</p>', unsafe_allow_html=True)
    st.write("Formulario móvil para que los conductores soliciten su cargue sin pasar por oficina.")
    
    if st.button("🚪 Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.rol_usuario = None
        st.rerun()
        
    st.markdown("---")
    ruta_sel = st.selectbox("Seleccione su Ruta:", ["Ruta 71", "Ruta 74", "Ruta 75 (Especial)", "Ruta 76", "Ruta 77", "ML3E51", "ML3E52", "ML3E53"])
    placa_vehiculo = st.text_input("Placa / Cupo del Vehículo:", placeholder="Ej: ABC-123")
    
    st.markdown("#### Ingrese las cantidades solicitadas por producto:")
    pedido_ruta = {}
    for item in st.session_state.catalogo:
        cant = st.number_input(f"{item['sku']} - {item['descripcion']}", min_value=0, step=1, key=f"drv_{item['sku']}")
        if cant > 0:
            pedido_ruta[item['sku']] = {"Descripción del Producto": item['descripcion'], "Cantidad": cant}
            
    if st.button("🚀 Enviar Solicitud de Carga a Bodega"):
        if placa_vehiculo and pedido_ruta:
            st.session_state.pedidos_autoventa[ruta_sel] = {
                "placa": placa_vehiculo,
                "fecha": str(date.today()),
                "productos": pedido_ruta
            }
            st.success(f"¡Pedido de la {ruta_sel} enviado exitosamente a la bodega!")
        else:
            st.warning("⚠️ Ingrese la placa del vehículo y al menos un producto con cantidad mayor a 0.")
    st.stop()

# ---------------------------------------------------------
# 8. VISTA DE ADMINISTRADOR (TODO INTEGRADO)
# ---------------------------------------------------------
col_head1, col_head2 = st.columns([6, 1])
with col_head1:
    st.markdown("### Panel de Administración - Inesco")
with col_head2:
    if st.button("🚪 Salir"):
        st.session_state.autenticado = False
        st.session_state.rol_usuario = None
        st.rerun()

tab1, tab2, tab3, tab4 = st.tabs([
    "📅 Fechas de Vencimiento", 
    "📦 Administrar SKUs", 
    "📄 Extracción PDF (Rutas)",
    "🚚 Recepción de Cargues (Autoventa)"
])

# --- TAB 1: FECHAS DE VENCIMIENTO ---
with tab1:
    st.markdown('<p class="sub-title">➕ Agregar o Actualizar Fecha de Vencimiento (Fecha Más Próxima)</p>', unsafe_allow_html=True)
    
    with st.expander("📂 Opciones de Respaldo, Sincronización y Limpieza (JSON)"):
        datos_respaldo = {
            "vencimientos": st.session_state.vencimientos,
            "catalogo": st.session_state.catalogo
        }
        json_str = json.dumps(datos_respaldo, ensure_ascii=False, indent=4)
        st.download_button(
            label="📥 Descargar Archivo Respaldo Completo (.json)",
            data=json_str,
            file_name=f"respaldo_inesco_{date.today()}.json",
            mime="application/json"
        )
        
        uploaded_backup = st.file_uploader("📤 Subir Respaldo Previo (.json)", type=["json"], key="uploader_backup")
        if uploaded_backup is not None:
            if "last_uploaded_file" not in st.session_state or st.session_state.last_uploaded_file != uploaded_backup.name:
                try:
                    stringio = io.StringIO(uploaded_backup.getvalue().decode("utf-8"))
                    data_recuperada = json.load(stringio)
                    if isinstance(data_recuperada, dict) and "vencimientos" in data_recuperada:
                        st.session_state.vencimientos = data_recuperada["vencimientos"]
                        if "catalogo" in data_recuperada:
                            st.session_state.catalogo = data_recuperada["catalogo"]
                        st.session_state.last_uploaded_file = uploaded_backup.name
                        st.success("✅ ¡Datos y catálogo restaurados exitosamente!")
                        st.rerun()
                except Exception as e:
                    st.error(f"⚠️ Error al leer el respaldo: {e}")
        
        if st.button("🗑️ Limpiar Todos los Vencimientos Registrados", key="btn_limpiar_todo"):
            st.session_state.vencimientos = []
            if "last_uploaded_file" in st.session_state:
                del st.session_state.last_uploaded_file
            st.toast("🧹 ¡Se han borrado todos los registros!", icon="✅")
            st.rerun()

    st.session_state.modo_captura = st.selectbox(
        "🎛️ Método de entrada de datos:",
        ["Tomar datos manual", "Búsqueda rápida de SKU"],
        index=0,
        key="select_modo_captura"
    )

    skus_opt = [f"{item['sku']} - {item['descripcion']}" for item in st.session_state.catalogo]

    if st.session_state.modo_captura == "Búsqueda rápida de SKU":
        filtro_voz = st.text_input("🔎 Escribe o busca el producto:", value=st.session_state.voz_temp_input, placeholder="Ej: Coca-Cola 350, Brisa, Quatro...", key="input_voz_busqueda")
        st.session_state.voz_temp_input = filtro_voz
        skus_filtrados = [s for s in skus_opt if filtro_voz.lower() in s.lower()] if filtro_voz else skus_opt
    else:
        skus_filtrados = skus_opt

    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        sel_sku = st.selectbox("Seleccionar Producto:", options=skus_filtrados, index=0 if len(skus_filtrados) == 1 else None, placeholder="🔎 Buscar SKU...", key="select_sku_venc")
    with c2:
        cantidad_input = st.text_input("Cant. (Opcional):", value="", placeholder="Ej: 50 o vacío", key="input_cantidad_lote")
    with c3:
        fecha_seleccionada = st.date_input("Fecha Vencimiento:", value=date.today(), key="input_calendar_venc")
        
    if st.button("💾 Guardar Fecha de Vencimiento", key="btn_save_venc"):
        if sel_sku:
            s_code, s_desc = sel_sku.split(" - ", 1)
            nueva_fecha_dt = fecha_seleccionada
            nueva_fecha_str = nueva_fecha_dt.strftime("%d/%m/%Y")
            cant_final = cantidad_input.strip() if cantidad_input.strip().isdigit() else "-"
            
            sku_existente_idx = next(
                (i for i, r in enumerate(st.session_state.vencimientos) if r["SKU"] == s_code), 
                None
            )
            
            if sku_existente_idx is not None:
                fecha_actual_str = st.session_state.vencimientos[sku_existente_idx]["Fecha Vencimiento"]
                try:
                    fecha_actual_dt = datetime.strptime(fecha_actual_str, "%d/%m/%Y").date()
                except:
                    fecha_actual_dt = nueva_fecha_dt
                
                if nueva_fecha_dt < fecha_actual_dt:
                    st.session_state.vencimientos[sku_existente_idx]["Fecha Vencimiento"] = nueva_fecha_str
                    st.toast(f"🔄 Se actualizó a la fecha más próxima ({nueva_fecha_str}) para {s_code}.", icon="✅")
                else:
                    st.toast(f"ℹ️ Ya existe una fecha más próxima ({fecha_actual_str}) para este producto.", icon="ℹ️")
                
                if cant_final != "-":
                    st.session_state.vencimientos[sku_existente_idx]["Cantidad"] = cant_final
            else:
                nuevo_registro = {
                    "SKU": s_code,
                    "Descripción del Producto": s_desc,
                    "Cantidad": cant_final,
                    "Fecha Vencimiento": nueva_fecha_str
                }
                st.session_state.vencimientos.append(nuevo_registro)
                st.toast(f"¡Guardado correctamente ({s_code} - {nueva_fecha_str})!", icon="✅")
            
            st.rerun()
        else:
            st.warning("⚠️ Debes seleccionar un SKU primero.")

    st.markdown('<p class="sub-title">📋 Productos Registrados en Bodega (Fecha Única más Próxima)</p>', unsafe_allow_html=True)
    
    if st.session_state.vencimientos:
        catalogo_actual = st.session_state.get("catalogo", CATALOGO_INICIAL)
        orden_catalogo = {item["sku"]: idx for idx, item in enumerate(catalogo_actual)}

        lotes_con_orden = []
        for idx, item in enumerate(st.session_state.vencimientos):
            estado, dias = calcular_estado_vencimiento(item["Fecha Vencimiento"])
            pos_cat = orden_catalogo.get(item["SKU"], 999)
            lote_data = item.copy()
            lote_data["_idx_original"] = idx
            lote_data["_dias"] = dias
            lote_data["_estado"] = estado
            lote_data["_pos"] = pos_cat
            lotes_con_orden.append(lote_data)

        lotes_ordenados = sorted(lotes_con_orden, key=lambda x: x["_pos"])

        to_delete = None
        for row in lotes_ordenados:
            orig_idx = row["_idx_original"]
            dias_rest = row["_dias"]
            estado_txt = row["_estado"]
            
            col_a, col_b, col_c, col_d, col_e, col_f = st.columns([1.5, 3.0, 2.2, 1.2, 2.2, 1.2])
            col_a.write(f"**{row['SKU']}**")
            col_b.write(row['Descripción del Producto'])
            fecha_real_str = st.session_state.vencimientos[orig_idx]['Fecha Vencimiento']
            col_c.write(f"📅 **{fecha_real_str}**")
            cant_mostrada = row.get('Cantidad', '-')
            col_d.write(f"📦 {cant_mostrada}")
            
            if dias_rest <= 30:
                col_e.error(f"{estado_txt} ({dias_rest}d)")
            elif dias_rest <= 60:
                col_e.warning(f"{estado_txt} ({dias_rest}d)")
            else:
                col_e.success(f"{estado_txt} ({dias_rest}d)")
            
            if col_f.button("❌ Borrar", key=f"del_lote_{orig_idx}"):
                to_delete = orig_idx

        if to_delete is not None:
            st.session_state.vencimientos.pop(to_delete)
            st.rerun()

        st.divider()
        excel_bytes = exportar_excel_vencimientos_organizado(st.session_state.vencimientos, fecha_str=date.today().strftime('%d/%m/%Y'))
        
        col_dl1, col_dl2 = st.columns([1, 1])
        with col_dl1:
            st.download_button(
                label="📥 Descargar Reporte Agrupado (.xlsx)",
                data=excel_bytes,
                file_name=f"Vencimientos_Inesco_Agrupado_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_venc_excel_org"
            )
        with col_dl2:
            texto_wa = urllib.parse.quote(f"Hola, comparto el reporte de control de vencimientos agrupado de Distribuciones Inesco del {date.today().strftime('%d/%m/%Y')}.")
            st.markdown(f"""
                <a href="https://api.whatsapp.com/send?text={texto_wa}" target="_blank" style="text-decoration: none;">
                    <div style="background-color: #25D366; color: white; padding: 10px 15px; border-radius: 8px; text-align: center; font-weight: bold; font-size: 0.95rem;">
                        💬 Abrir WhatsApp con Aviso
                    </div>
                </a>
            """, unsafe_allow_html=True)

# --- TAB 2: ADMINISTRAR SKUS ---
with tab2:
    st.markdown('<p class="sub-title">⚙️ Administrar Catálogo de SKUs</p>', unsafe_allow_html=True)
    col_add1, col_add2 = st.columns([1, 2])
    with col_add1:
        nuevo_sku = st.text_input("Código SKU:", key="input_new_sku")
    with col_add2:
        nueva_desc = st.text_input("Descripción:", key="input_new_desc")
        
    if st.button("➕ Agregar SKU", key="btn_add_sku"):
        if nuevo_sku and nueva_desc:
            if not any(item['sku'] == nuevo_sku.strip() for item in st.session_state.catalogo):
                st.session_state.catalogo.append({"sku": nuevo_sku.strip(), "descripcion": nueva_desc.strip()})
                st.success("✅ SKU agregado permanentemente.")
                st.rerun()
            else:
                st.warning("⚠️ El SKU ya existe.")
        else:
            st.error("Completa ambos campos.")

    st.divider()
    for idx, item in enumerate(st.session_state.catalogo):
        c_k, c_d, c_b = st.columns([2, 5, 2])
        c_k.write(f"**{item['sku']}**")
        c_d.write(item['descripcion'])
        if c_b.button("🗑️ Eliminar", key=f"cat_del_{item['sku']}_{idx}"):
            st.session_state.catalogo.pop(idx)
            st.rerun()

# --- TAB 3: EXTRACCIÓN PDF (RUTAS) ---
with tab3:
    st.markdown('<p class="sub-title">📄 Extracción Total de Rutas y Cargues de FEMSA</p>', unsafe_allow_html=True)
    st.info("ℹ️ Sube tu PDF de cargue para extraer la información por rutas")
    
    archivo_pdf = st.file_uploader("📂 Seleccionar archivo PDF de Rutas", type=["pdf"], key="uploader_pdf_rutas")
    
    if archivo_pdf is not None:
        if st.button("🚀 Procesar PDF y Generar Excel", key="btn_procesar_pdf"):
            with st.spinner("🔄 Procesando PDF y ajustando separación de cajas/unidades..."):
                try:
                    datos_rutas = procesar_pdf_rutas(archivo_pdf)
                    
                    if datos_rutas:
                        dfs_para_excel = {}
                        for r_name, r_info in datos_rutas.items():
                            dfs_para_excel[r_name] = {
                                "df": pd.DataFrame(r_info["items"]),
                                "header": r_info["header"]
                            }
                            
                        excel_rutas_bytes = exportar_excel_multiruta(dfs_para_excel, fecha_str=date.today().strftime('%d/%m/%Y'), es_reporte_rutas=True)
                        
                        st.success(f"✅ Se procesaron exitosamente {len(datos_rutas)} ruta(s).")
                        
                        for r_name, r_info in datos_rutas.items():
                            with st.expander(f"🚚 Ruta: {r_name} ({len(r_info['items'])} registros)"):
                                st.write(f"**Encabezado detectado:** {r_info['header']}")
                                st.dataframe(pd.DataFrame(r_info["items"]), use_container_width=True)
                        
                        st.download_button(
                            label="📥 Descargar Reporte Completo de Rutas (.xlsx)",
                            data=excel_rutas_bytes,
                            file_name=f"Reporte_Rutas_Inesco_{date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_rutas_excel"
                        )
                    else:
                        st.warning("⚠️ No se detectaron productos estructurados en el PDF.")
                except Exception as e:
                    st.error(f"⚠️ Ocurrió un error al procesar el PDF: {e}")

# --- TAB 4: RECEPCIÓN DE CARGUES (AUTOVENTA) ---
with tab4:
    st.markdown('<p class="sub-title">🚚 Bandeja de Pedidos y Cargues en Vivo (Enviados por Conductores)</p>', unsafe_allow_html=True)
    
    if st.session_state.pedidos_autoventa:
        st.info(f"Hay pedidos activos de {len(st.session_state.pedidos_autoventa)} ruta(s).")
        
        # Diccionario para armar el consolidado para Excel
        dict_pedidos_excel = {}
        
        for r_nombre, info in st.session_state.pedidos_autoventa.items():
            with st.expander(f"📌 {r_nombre} (Placa/Cupo: {info['placa']} - Fecha: {info['fecha']})"):
                df_pedido_temp = pd.DataFrame.from_dict(info['productos'], orient='index')
                df_pedido_temp.reset_index(inplace=True)
                df_pedido_temp.rename(columns={'index': 'SKU'}, inplace=True)
                
                st.dataframe(df_pedido_temp, use_container_width=True)
                
                dict_pedidos_excel[r_nombre] = {
                    "df": df_pedido_temp,
                    "header": f"Ruta: {r_nombre} | Placa: {info['placa']} | Fecha: {info['fecha']}"
                }
                
                if st.button(f"🗑️ Borrar Pedido de {r_nombre}", key=f"del_ped_{r_nombre}"):
                    del st.session_state.pedidos_autoventa[r_nombre]
                    st.rerun()
                    
        st.divider()
        if dict_pedidos_excel:
            excel_pedidos_bytes = exportar_excel_multiruta(dict_pedidos_excel, fecha_str=date.today().strftime('%d/%m/%Y'), es_reporte_rutas=True)
            st.download_button(
                label="📥 Descargar Consolidado de Pedidos en Excel (.xlsx)",
                data=excel_pedidos_bytes,
                file_name=f"Consolidado_Pedidos_Autoventa_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_pedidos_excel"
            )
    else:
        st.write("ℹ️ Aún no hay pedidos enviados por las rutas desde el portal móvil de autoventa.")
