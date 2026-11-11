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

ADMIN_PASSWORD = "inesco2026admin"
DRIVER_PASSWORD = "ruta2026inesco"

PLACAS_POR_RUTA = {
    "ML3E71": ["JYY 174"],
    "ML3E74": ["JYY 262"],
    "ML3E75": ["JYY 180"],
    "ML3E76": ["JYY 177"],
    "ML3E77": ["JYY 175"]
}

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
# 2. CATÁLOGO OFICIAL Y ESTRUCTURA DE BODEGA
# ---------------------------------------------------------
if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = []

if "pedidos_autoventa" not in st.session_state:
    st.session_state.pedidos_autoventa = {}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "rol_usuario" not in st.session_state:
    st.session_state.rol_usuario = None

CATALOGO_SECCIONES = {
    "1. PRODUCTOS 8 ONZAS (VIR / RGB)": [
        {"sku": "135718", "descripcion": "COCA-COLA 8OZ VIR(30)"},
        {"sku": "135764", "descripcion": "COCA-COLA SA 8OZ VIR(30)"},
        {"sku": "56557", "descripcion": "SCHWEPPES SODA 8OZ RGB (30)"},
        {"sku": "56521", "descripcion": "SPRITE 8OZ VIR (30)"},
        {"sku": "56704", "descripcion": "QUATRO TJA 8OZ VIR(30)"},
        {"sku": "56624", "descripcion": "PREMIO ROJO 8 OZ VIR (30)"}
    ],
    "2. PRODUCTOS 350 ML (VIR / RGB)": [
        {"sku": "160053", "descripcion": "COCA-COLA 350ML VIR(30)"},
        {"sku": "135763", "descripcion": "COCA COLA ZERO 350 ML VIR"},
        {"sku": "56622", "descripcion": "PREMIO ROJO 350 ML RGB"},
        {"sku": "56705", "descripcion": "QUATRO TJA 350ML VIR(30)"},
        {"sku": "160070", "descripcion": "SPRITE LL 350ML VIR(30)"},
        {"sku": "160033", "descripcion": "SCHWEPPES SODA 350 ML"}
    ],
    "3. PRODUCTOS 10 ONZAS VNR / GINGER / SODAS": [
        {"sku": "160192", "descripcion": "COCA COLA 10 OZ VNR (12)"},
        {"sku": "135766", "descripcion": "COCA COLA ZERO 10Oz VNR(12)KZ"},
        {"sku": "56497", "descripcion": "SCHWEPPES GINGER 10OZ"},
        {"sku": "56498", "descripcion": "SCHWEPPES SODA 10OZ"}
    ],
    "4. BIDONES Y BOTELLONES DE AGUA": [
        {"sku": "160151", "descripcion": "AGUA BRISA BIDON 6 LT (2)"},
        {"sku": "160048", "descripcion": "AGUA BOTELLON 5 GL TV"}
    ],
    "5. FORMATOS 2.5 LITROS": [
        {"sku": "135755", "descripcion": "COCA COLA ZERO 2.5 LT (8)"},
        {"sku": "135663", "descripcion": "COCA COLA 2.5LT PET(8) Nvo"},
        {"sku": "120119", "descripcion": "DEL VALLE CITRUS 2.5LT PET (8)"},
        {"sku": "56820", "descripcion": "QTRO 2.5LT PET(8) Nvo"}
    ],
    "6. FORMATOS 3 LITROS Y 1.5 LITROS": [
        {"sku": "135664", "descripcion": "COCA-COLA 1.5 LT NR (12)"},
        {"sku": "135757", "descripcion": "COCA-COLA ZERO 1.5 LT (12)"},
        {"sku": "136159", "descripcion": "COCA-COLA 1 LT NR (12)"},
        {"sku": "160099", "descripcion": "KOLA ROMAN 1.5 LT PET (12)"},
        {"sku": "56709", "descripcion": "QUATRO CHOICE 1.5 LT PET (12)"},
        {"sku": "160102", "descripcion": "SPRITE LIMA LIMON 1.5 LT NR (12)"},
        {"sku": "56651", "descripcion": "PREMIO ROJO FNT 1.5LT PET(12)"},
        {"sku": "160101", "descripcion": "SCHWEPPES SODA 1.5LT PET(12)"},
        {"sku": "120118", "descripcion": "DEL VALLE CITRUS 1.5L NAJ"},
        {"sku": "120185", "descripcion": "FRESH PONCHE FRUTAS 1.5LT PET(12)"},
        {"sku": "120159", "descripcion": "FRESH MAND 1,5 LTS PET (12)"},
        {"sku": "92770", "descripcion": "BRISA MANZANA 1.5LT PET(12)"},
        {"sku": "92740", "descripcion": "BRISA LIMA LIMÓN 1.5LT PET (12)"},
        {"sku": "92795", "descripcion": "BRISA MARACUYA BOT 1.5LT PET 12"}
    ],
    "7. FORMATO 3 LITROS": [
        {"sku": "135662", "descripcion": "COCA COLA 3 LTS PET (6)"},
        {"sku": "160124", "descripcion": "SPRITE 3 LTS PET (6)"},
        {"sku": "56711", "descripcion": "QUATRO 3 LTS PET (6)"},
        {"sku": "56625", "descripcion": "PREMIO ROJO FNT 3LT PET(6)"}
    ],
    "8. AGUAS PET 600 ML": [
        {"sku": "160046", "descripcion": "BRISA GAS PET 600ML (24)"},
        {"sku": "160047", "descripcion": "AGUA BRISA PET 600ML (24)"},
        {"sku": "92892", "descripcion": "BRISA LIMALIM BOT 600ML NR PET"},
        {"sku": "92794", "descripcion": "BRISA MARACUYA 600ML NR PET 6X6"},
        {"sku": "92769", "descripcion": "BRISA MANZANA 600ML PET(6)"}
    ],
    "9. MANANTIAL Y OTRAS AGUAS": [
        {"sku": "92764", "descripcion": "MANANTIAL SIN GAS 500ML VNR(12)"},
        {"sku": "160240", "descripcion": "MANANTIAL GAS 600 ML (12)"},
        {"sku": "160060", "descripcion": "MANANTIAL 1.5 PET(9)"},
        {"sku": "160241", "descripcion": "MANANTIAL 25g PET 600 ML (12)"}
    ],
    "10. AGUAS ECOFLEX Y PEQUEÑAS (280 ML)": [
        {"sku": "160008", "descripcion": "BRISA ECOFLEX 1 LITRO (6)"},
        {"sku": "160168", "descripcion": "AGUA BRISA 280 PET ML (24)"},
        {"sku": "160170", "descripcion": "BRISA GAS LIMA LI PET 280 ML (24)"},
        {"sku": "92768", "descripcion": "BRISA MANZANA 280ML PET(24)"},
        {"sku": "92793", "descripcion": "BRISA MARACUYA BOT 280ML NR PET"}
    ],
    "11. COMBOS AHORRO": [
        {"sku": "135675", "descripcion": "COMBO AHORRO CCSO 3LT"},
        {"sku": "136247", "descripcion": "2 PK CC +QT 3L"},
        {"sku": "120120", "descripcion": "2PK FRESH CITRUS 2.5L"}
    ],
    "12. PRODUCTOS 400 ML": [
        {"sku": "160297", "descripcion": "SPRITE 400 PET"},
        {"sku": "56706", "descripcion": "QUATRO CHOICE 400ML NR (12)"},
        {"sku": "56627", "descripcion": "PREMIO ROJO FNT 400 ML PET(12)"},
        {"sku": "160318", "descripcion": "COCA-COLA 400ML PET (12)"},
        {"sku": "135760", "descripcion": "COCA COLA SA 400ML PET(12)"},
        {"sku": "160479", "descripcion": "KOLA ROMAN 400 ML NR (12)"},
        {"sku": "120184", "descripcion": "FRESH PONCHE FRUTAS 400ML PET(12)"},
        {"sku": "120117", "descripcion": "DV CITRUS 400ML PET(12)"},
        {"sku": "120158", "descripcion": "DV MANDARINA 400 ML PET(12)"},
        {"sku": "56452", "descripcion": "SCHWEPPES SODA 400ML PET(12)"},
        {"sku": "56453", "descripcion": "SCHWEPPES GINGER 400ML PET(12)"}
    ],
    "13. POWERADE Y FLASHLYTE": [
        {"sku": "160121", "descripcion": "POWERADE (ROJO) 500ML PET(6)"},
        {"sku": "120208", "descripcion": "POWERADE CONTRATAQUE 500ML PET(6)"},
        {"sku": "160122", "descripcion": "POWERADE MB (AZUL) 500ML PET(6)"},
        {"sku": "56842", "descripcion": "FLASHLYTE COCO LIM 625ML PET(6)"},
        {"sku": "56841", "descripcion": "FLASHLYTE UVA 625ML PET(6)"}
    ],
    "14. FORMATOS 250 ML / MINI PET": [
        {"sku": "160200", "descripcion": "COCA COLA 250 ML PET (12)"},
        {"sku": "135761", "descripcion": "COCA COLA ZERO 250 ML PET (12)"},
        {"sku": "56703", "descripcion": "QUATRO 250 ML PET (12)"},
        {"sku": "120187", "descripcion": "FRESH CITRUS 250ML PET(12)"},
        {"sku": "160201", "descripcion": "6 PK COCA COLA MINI"},
        {"sku": "135795", "descripcion": "6PK CCZR 250 ML PET KZ"}
    ],
    "15. LATAS 269 ML / 330 ML": [
        {"sku": "135803", "descripcion": "6PK COCA COLA SO 269ML LAT"},
        {"sku": "135767", "descripcion": "COCA COLA ZERO 330ML LAT(12)KZ"},
        {"sku": "160029", "descripcion": "COCA-COLA 330ML LATA (12)"},
        {"sku": "136402", "descripcion": "6PK CCSO 269ML LAT"},
        {"sku": "136403", "descripcion": "6PK CCZR 269ML LAT KZ"},
        {"sku": "136400", "descripcion": "COCA COLA ZERO 269ML LAT(12)KZ"},
        {"sku": "136399", "descripcion": "COCA COLA ORIGINAL 269ML LAT(12)"},
        {"sku": "136401", "descripcion": "CCZERO CAFE 269ML LAT(12)"},
        {"sku": "136402", "descripcion": "COCA-COLA ZERO CAFE 269ML (12)"},
        {"sku": "136409", "descripcion": "COCA-COLA ORIGINAL CAFE 269ML LAT(12)"}
    ],
    "16. ENERGIZANTES MONSTER Y FURY (473 ML)": [
        {"sku": "120175", "descripcion": "MONSTER PIPELINE PUNCH 473ML LAT(24)"},
        {"sku": "120112", "descripcion": "MONSTER PARADISE 473ML LAT(24)"},
        {"sku": "120027", "descripcion": "MONSTER VIOLET 473ML LAT(24)"},
        {"sku": "160378", "descripcion": "MONSTER ULTRA LATA 473ML (24)"},
        {"sku": "120156", "descripcion": "MONSTER GREEN 473ML LAT(12) MQ"},
        {"sku": "160205", "descripcion": "MONSTER ZERO GREEN 473ML LAT(12) MQ"},
        {"sku": "120214", "descripcion": "MONSTER VICE GUAVA LAT 473ML(12)"},
        {"sku": "120215", "descripcion": "FURY LAT 473ML(12)"},
        {"sku": "120181", "descripcion": "MONSTER MGLOCO 473ML LAT(12) MQ"}
    ],
    "17. TÉS FUZE TEA Y PACKS": [
        {"sku": "160187", "descripcion": "FUZE TEA NEGRO LIMON 400(6)"},
        {"sku": "160223", "descripcion": "FUZE MANGO M 400ML PET (6)"},
        {"sku": "160186", "descripcion": "FUZE TEA NEGRO DURAZNO(6)"},
        {"sku": "160732", "descripcion": "3 PACK POWER SURT  (MB + FT)"},
        {"sku": "160247", "descripcion": "3 PACK FUZE 400 DUR MANZ LIM"}
    ],
    "18. JUGOS FRUTAL (500 ML / 946 ML / 1.2 L / 188 ML)": [
        {"sku": "120131", "descripcion": "DV FRUTAL MANGO 500ML PET(6)"},
        {"sku": "120130", "descripcion": "FRUTAL SALPICON 500ML PET(6)"},
        {"sku": "120132", "descripcion": "DV FRUTAL MDRINA-PIÑA 500ML PET"},
        {"sku": "120128", "descripcion": "FRUTAL MANGOFRS 500ML PET(6)"},
        {"sku": "120129", "descripcion": "FRUTAL MORA 500ML PET(6)"},
        {"sku": "120031", "descripcion": "FRUTAL MORA 946 ML T"},
        {"sku": "120141", "descripcion": "FRUTAL MANGO 946 ML"},
        {"sku": "120139", "descripcion": "FRUTAL MORA 946ML TRP(6)NVO"},
        {"sku": "120138", "descripcion": "FRUTAL MANG FRESA 946"},
        {"sku": "120084", "descripcion": "FRUTAL SALPICON 946ML TRP(6)"},
        {"sku": "120190", "descripcion": "FRUTAL MANGOFRS 1.2LT PET(6)"},
        {"sku": "120188", "descripcion": "FRUTAL MANGO 1.2L PET (6)"},
        {"sku": "120191", "descripcion": "FRUTAL MORA 1.2L PET (6)"},
        {"sku": "120085", "descripcion": "FRUTAL SALPICON 188ML TRP(24)"},
        {"sku": "120136", "descripcion": "FRUTAL MANGO 188ML TRP(24)"},
        {"sku": "120134", "descripcion": "FRUTAL MORA 188ML TRP(24)"},
        {"sku": "120133", "descripcion": "FRUTAL MANGOFRS 188ML TRP(24)"},
        {"sku": "120137", "descripcion": "FRUTAL MDR-PIÑA 188ML TRP(24)"},
        {"sku": "120064", "descripcion": "DVFRESH MANDARINA 188ML TRP(24)"},
        {"sku": "120194", "descripcion": "6PK(2MG+2MR+1MP+1MF)188ML NVO"},
        {"sku": "120135", "descripcion": "FRUTAL SALPICON 188ML TRP(24)NVO"},
        {"sku": "120116", "descripcion": "DVFRESH CITRUS 188ML TRP(24)"}
    ],
    "19. SNACKS, GALLETAS, TRIDENT Y GOLOSINAS": [
        {"sku": "136691", "descripcion": "EXH TRUL 72UX70G 6.7K KIT1(1)"},
        {"sku": "136468", "descripcion": "HALLS 100S EXT STRONG COL 310GR KIT30(1)"},
        {"sku": "136487", "descripcion": "LOKINO CARAMELO 400GR KIT24(1)"},
        {"sku": "136747", "descripcion": "CHOCOLORES SURT 540G KIT12(1)"},
        {"sku": "136475", "descripcion": "BARRILETE 440GR PAIS KIT18(1)"},
        {"sku": "136701", "descripcion": "BIANCHI CHOCOBLA 600G KIT18(1)"},
        {"sku": "136700", "descripcion": "BIANCHI CHOCOLA 600G KIT18(1)"},
        {"sku": "136479", "descripcion": "CHAO FRESA 350GR KIT24(1)"},
        {"sku": "136478", "descripcion": "CHAO MENTA 350GR KIT24(1)"},
        {"sku": "136466", "descripcion": "OREO ROLLO 109GR KIT30(1)"},
        {"sku": "136712", "descripcion": "TRULU 70U MAS VENDID 1.6K KIT6(1)"},
        {"sku": "136467", "descripcion": "OREO REGULAR COL 324GR KIT18(1)"},
        {"sku": "136671", "descripcion": "TRUL MAXI+VEND 12U 1.2K KIT6(1)"},
        {"sku": "136630", "descripcion": "TRULULU MINIPACK 24U 960GR KIT1(1)"},
        {"sku": "136720", "descripcion": "TRUL MINIPACK 24U 840GR KIT12(1)"},
        {"sku": "136632", "descripcion": "TRUL MAXIPACK 12U 1.2K KIT1(1)"},
        {"sku": "136675", "descripcion": "TRUL SABOR SURT 12U 360G KIT12(1)"},
        {"sku": "136670", "descripcion": "TRUL MXPK 12U 1.3K TOPEX KIT6(1)"},
        {"sku": "136020", "descripcion": "BOCAD ATUN ACEIT 140GR KIT48(1)"},
        {"sku": "136024", "descripcion": "ATUN RALL ACEITE 150GR KIT48(1)"},
        {"sku": "136027", "descripcion": "BOCADO ATUN AGUA 140GR KIT18(1)"},
        {"sku": "136474", "descripcion": "ATUN ISABEL ACEI 142G KIT48(1)"},
        {"sku": "136461", "descripcion": "ZUCARITAS MEGA 115G KIT54(1)"},
        {"sku": "136465", "descripcion": "FROOT LOOPS MEGA 90GR KIT54(1)"},
        {"sku": "136462", "descripcion": "CHOKRISPIS MEGA 115G KIT54(1)"},
        {"sku": "136560", "descripcion": "CHOCO KRISPIS 240GR KIT30(1)"},
        {"sku": "136559", "descripcion": "ZUCARITAS 240GR KIT30(1)"},
        {"sku": "136561", "descripcion": "FROOT LOOPS 200GR KIT30(1)"},
        {"sku": "136705", "descripcion": "TRIDENT 5SX18U MENTA 153G KIT30(1)"},
        {"sku": "136707", "descripcion": "TRIDENT 1SX60U MENTA 90G KIT40(1)"},
        {"sku": "136710", "descripcion": "CLUB SOCIAL 9U ORIG 216G KIT24(1)"},
        {"sku": "136783", "descripcion": "OREO REGULAR 540G KIT12(1)"}
    ]
}

if "catalogo" not in st.session_state:
    st.session_state.catalogo = []
    for sec, productos in CATALOGO_SECCIONES.items():
        for p in productos:
            if not any(item["sku"] == p["sku"] for item in st.session_state.catalogo):
                st.session_state.catalogo.append(p)

# ---------------------------------------------------------
# 3. EXPORTADOR EXCEL CON FILTROS AUTOMÁTICOS Y FECHA EDITABLE
# ---------------------------------------------------------
def exportar_excel_estilo_bodega(rutas_dict, fecha_str):
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
    section_total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    total_font = Font(bold=True, size=11, name="Calibri", color="000000")
    
    black_border = Border(
        left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
    )
    
    filas_totales_secciones = []

    for nombre_ruta, r_info in rutas_dict.items():
        safe_title = re.sub(r'[\\/*?:[\]]', '_', nombre_ruta)
        ws = wb.create_sheet(title=safe_title[:30])
        
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        
        # Encabezado visual profesional y centrado
        ws.merge_cells('A1:G1')
        title_cell = ws.cell(row=1, column=1, value="DISTRIBUCIONES INESCO - HOJA DE CARGUE")
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        ws.merge_cells('A2:C2')
        r_cell = ws.cell(row=2, column=1, value=f"Ruta: {nombre_ruta}")
        r_cell.font = Font(bold=True, size=12, color="1F4E78")
        r_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells('D2:G2')
        p_cell = ws.cell(row=2, column=4, value=f"Placa: {r_info.get('placa', 'JYY 175')}")
        p_cell.font = Font(bold=True, size=12, color="1F4E78")
        p_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 24

        # Fecha editable en Excel (Celda de texto limpia)
        ws.cell(row=3, column=1, value="Fecha:").font = Font(bold=True, size=10)
        date_cell = ws.cell(row=3, column=2, value=fecha_str)
        date_cell.font = Font(size=10, bold=True)
        date_cell.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[3].height = 18

        # Cabeceras y activación de AutoFiltro nativo de Excel
        headers = ["SKU", "ELEMENTO", "CANTIDAD", "DEVOLUCIÓN 1", "DEVOLUCIÓN 2", "DEVOLUCIÓN 3", "PISO / UBICACIÓN"]
        for col_idx, h_text in enumerate(headers, start=1):
            c = ws.cell(row=4, column=col_idx, value=h_text)
            c.fill = header_fill
            c.font = header_font
            c.border = black_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[4].height = 24
        
        # Habilitar filtros automáticos en toda la tabla
        ws.auto_filter.ref = f"A4:G1000"
            
        current_row = 5
        productos_dict_pedido = r_info.get("productos", {})
        
        for nombre_seccion, lista_prods in CATALOGO_SECCIONES.items():
            start_sec_row = current_row
            
            for item in lista_prods:
                sku = item["sku"]
                desc = item["descripcion"]
                match_key = next((k for k in productos_dict_pedido if k.startswith(sku)), None)
                qty = productos_dict_pedido[match_key]["Cantidad"] if match_key else 0
                
                ws.cell(row=current_row, column=1, value=int(sku) if sku.isdigit() else sku).alignment = Alignment(horizontal="center", vertical="center")
                
                cell_desc = ws.cell(row=current_row, column=2, value=desc)
                cell_desc.font = Font(name="Calibri", size=10, bold=True)
                cell_desc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                
                cell_qty = ws.cell(row=current_row, column=3, value=qty) # Mantiene el cero visible para poder editar en Excel si es necesario
                cell_qty.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.cell(row=current_row, column=4, value="")
                ws.cell(row=current_row, column=5, value="")
                ws.cell(row=current_row, column=6, value="")
                ws.cell(row=current_row, column=7, value="")
                
                for c_i in range(1, 8):
                    cell = ws.cell(row=current_row, column=c_i)
                    cell.border = black_border
                ws.row_dimensions[current_row].height = 20
                current_row += 1
                
            end_sec_row = current_row - 1
            ws.cell(row=current_row, column=1, value="")
            
            cell_t_label = ws.cell(row=current_row, column=2, value="TOTAL")
            cell_t_label.font = total_font
            cell_t_label.alignment = Alignment(horizontal="center", vertical="center")
            
            ws.cell(row=current_row, column=3, value=f"=SUM(C{start_sec_row}:C{end_sec_row})")
            ws.cell(row=current_row, column=4, value="")
            ws.cell(row=current_row, column=5, value="")
            ws.cell(row=current_row, column=6, value="")
            ws.cell(row=current_row, column=7, value="")
            
            filas_totales_secciones.append(f"C{current_row}")
            
            for c_i in range(1, 8):
                cell = ws.cell(row=current_row, column=c_i)
                cell.fill = section_total_fill
                cell.font = total_font
                cell.border = black_border
                if c_i == 3:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.number_format = '#,##0'
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 22
            current_row += 1
            
        extras = [
            ("BOT COCA-COLA 8 OZ (CANASTA)", 183303),
            ("BOT COCA-COLA 350ML", 184017),
            ("CANASTA PLASTICA 30UN", 184008),
            ("CANASTA PLASTICA 9 UN", 184010),
            ("BOTELLON BRISA 5 GLS", 184002),
            ("RACK O ESTIBA METALICA", 197794),
            ("ESTIBA DE MADERA NORMAL", 197826),
            ("TARIMA PLASTICA ISO OW 1.0 cm X 1.2", 197815),
            ("CANASTA PLASTICA BRISA KOF", 184003)
        ]
        
        current_row += 1
        for item_extra, cod_ext in extras:
            ws.cell(row=current_row, column=2, value=item_extra).font = Font(name="Calibri", size=10, bold=True)
            ws.cell(row=current_row, column=4, value=cod_ext).alignment = Alignment(horizontal="center", vertical="center")
            for c_i in range(1, 8):
                cell = ws.cell(row=current_row, column=c_i)
                cell.border = black_border
                if c_i not in [2]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        gt_label = ws.cell(row=current_row, column=1, value="TOTAL CAJAS DEL CARGUE:")
        gt_label.font = Font(name="Calibri", size=12, bold=True, color="1F4E78")
        gt_label.alignment = Alignment(horizontal="right", vertical="center")
        
        formula_gran_total = "=SUM(" + ",".join(filas_totales_secciones) + ")" if filas_totales_secciones else "=0"
        cell_gt = ws.cell(row=current_row, column=3, value=formula_gran_total)
        cell_gt.font = Font(name="Calibri", size=13, bold=True, color="C00000")
        cell_gt.alignment = Alignment(horizontal="center", vertical="center")
        cell_gt.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        for c_i in range(1, 8):
            ws.cell(row=current_row, column=c_i).border = black_border
        ws.row_dimensions[current_row].height = 26

        ws.column_dimensions['A'].width = 14
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 16
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['F'].width = 16
        ws.column_dimensions['G'].width = 18
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ---------------------------------------------------------
# 4. EXTRACCIÓN DE PDF Y VENCIMIENTOS
# ---------------------------------------------------------
def procesar_pdf_rutas(pdf_file):
    routes_data = {}
    mapa_catalogo = {item["sku"]: item["descripcion"] for item in st.session_state.catalogo}
    with pdfplumber.open(pdf_file) as pdf:
        current_route = "ML3E71"
        current_header_info = "Ruta / No.de Carga: ML3E71"
        for page in pdf.pages:
            texto_pagina = page.extract_text()
            if not texto_pagina:
                continue
            for linea in texto_pagina.split('\n'):
                l_str = linea.strip()
                if not l_str:
                    continue
                m_ruta = re.search(r'ML3[E5]\d+', l_str)
                if m_ruta:
                    current_route = m_ruta.group(0)
                m_sku = re.search(r'\b(\d{5,6})\b', l_str)
                if m_sku:
                    sku = m_sku.group(1)
                    cajas, unidades = 0, 0
                    m_slash = re.search(r'(?:^|\s)(\d+)?\s*/\s*(\d+)', l_str)
                    if m_slash:
                        cajas = int(m_slash.group(1)) if m_slash.group(1) else 0
                        unidades = int(m_slash.group(2)) if m_slash.group(2) else 0
                    if current_route not in routes_data:
                        routes_data[current_route] = {"header": current_header_info, "items": []}
                    routes_data[current_route]["items"].append({
                        "SKU": sku,
                        "Descripción del Producto": mapa_catalogo.get(sku, "PRODUCTO FEMSA"),
                        "Cajas": cajas,
                        "Unidades": unidades
                    })
    return routes_data

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

# ---------------------------------------------------------
# 5. PANTALLA DE ACCESO (LOGIN)
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
# 6. PORTAL DE CONDUCTORES
# ---------------------------------------------------------
if st.session_state.rol_usuario == "driver":
    st.markdown('<p class="sub-title">📱 Portal de Autoventa / Solicitud de Carga - Rutas Inesco</p>', unsafe_allow_html=True)
    st.write("Selecciona tu sección, digita las cantidades y envía tu pedido a bodega.")
    
    if st.button("🚪 Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.rol_usuario = None
        st.rerun()
        
    st.markdown("---")
    ruta_sel = st.selectbox("Seleccione su Ruta:", list(PLACAS_POR_RUTA.keys()))
    placa_vehiculo = st.selectbox("Seleccione la Placa / Cupo del Vehículo:", PLACAS_POR_RUTA.get(ruta_sel, ["N/A"]))
    
    st.markdown("---")
    st.markdown("#### 📦 Ingrese las cantidades solicitadas por sección:")
    
    pedido_ruta = {}
    for nombre_seccion, lista_productos in CATALOGO_SECCIONES.items():
        items_con_qty = sum(1 for idx, item in enumerate(lista_productos) if st.session_state.get(f"drv_{item['sku']}_{idx}", 0) > 0)
        badge_txt = f" ({items_con_qty} seleccionados)" if items_con_qty > 0 else ""
        
        with st.expander(f"📌 {nombre_seccion}{badge_txt}"):
            for idx, item in enumerate(lista_productos):
                cant = st.number_input(f"{item['sku']} - {item['descripcion']}", min_value=0, step=1, key=f"drv_{item['sku']}_{idx}")
                if cant > 0:
                    pedido_ruta[f"{item['sku']}_{idx}"] = {"SKU": item['sku'], "Descripción del Producto": item['descripcion'], "Cantidad": cant}
            
    col_env1, col_env2 = st.columns(2)
    with col_env1:
        if st.button("🚀 Enviar Solicitud de Carga a Bodega"):
            if placa_vehiculo and pedido_ruta:
                st.session_state.pedidos_autoventa[ruta_sel] = {
                    "placa": placa_vehiculo,
                    "fecha": str(date.today()),
                    "productos": pedido_ruta
                }
                st.success(f"¡Pedido de la {ruta_sel} enviado exitosamente a la bodega!")
            else:
                st.warning("⚠️ Seleccione la placa y al menos un producto con cantidad mayor a 0.")
    with col_env2:
        if pedido_ruta:
            excel_driver_bytes = exportar_excel_estilo_bodega({ruta_sel: {"placa": placa_vehiculo, "productos": pedido_ruta}}, fecha_str=str(date.today()))
            st.download_button(
                label="📥 Descargar mi Solicitud en Excel (.xlsx)",
                data=excel_driver_bytes,
                file_name=f"Solicitud_{ruta_sel}_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.stop()

# ---------------------------------------------------------
# 7. PANEL DE ADMINISTRACIÓN
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

with tab1:
    st.markdown('<p class="sub-title">➕ Agregar o Actualizar Fecha de Vencimiento</p>', unsafe_allow_html=True)
    skus_opt = [f"{item['sku']} - {item['descripcion']}" for item in st.session_state.catalogo]
    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        sel_sku = st.selectbox("Seleccionar Producto:", options=skus_opt, key="select_sku_venc")
    with c2:
        cantidad_input = st.text_input("Cant. (Opcional):", value="", key="input_cantidad_lote")
    with c3:
        fecha_seleccionada = st.date_input("Fecha Vencimiento:", value=date.today(), key="input_calendar_venc")
        
    if st.button("💾 Guardar Fecha", key="btn_save_venc"):
        if sel_sku:
            s_code, s_desc = sel_sku.split(" - ", 1)
            nueva_fecha_str = fecha_seleccionada.strftime("%d/%m/%Y")
            cant_final = cantidad_input.strip() if cantidad_input.strip().isdigit() else "-"
            st.session_state.vencimientos.append({
                "SKU": s_code, "Descripción del Producto": s_desc, "Cantidad": cant_final, "Fecha Vencimiento": nueva_fecha_str
            })
            st.success("¡Guardado correctamente!")
            st.rerun()

    if st.session_state.vencimientos:
        if st.button("🗑️ Limpiar Todos los Vencimientos"):
            st.session_state.vencimientos = []
            st.rerun()

with tab2:
    st.markdown('<p class="sub-title">⚙️ Administrar Catálogo de SKUs</p>', unsafe_allow_html=True)
    col_add1, col_add2 = st.columns([1, 2])
    with col_add1:
        nuevo_sku = st.text_input("Código SKU:", key="input_new_sku")
    with col_add2:
        nueva_desc = st.text_input("Descripción:", key="input_new_desc")
    if st.button("➕ Agregar SKU", key="btn_add_sku"):
        if nuevo_sku and nueva_desc:
            st.session_state.catalogo.append({"sku": nuevo_sku.strip(), "descripcion": nueva_desc.strip()})
            st.success("✅ SKU agregado.")
            st.rerun()

with tab3:
    st.markdown('<p class="sub-title">📄 Extracción Total de Rutas y Cargues de FEMSA</p>', unsafe_allow_html=True)
    archivo_pdf = st.file_uploader("📂 Seleccionar archivo PDF de Rutas", type=["pdf"], key="uploader_pdf_rutas")
    if archivo_pdf is not None and st.button("🚀 Procesar PDF"):
        with st.spinner("🔄 Procesando PDF..."):
            datos_rutas = procesar_pdf_rutas(archivo_pdf)
            if datos_rutas:
                st.success(f"✅ Se procesaron {len(datos_rutas)} ruta(s).")

with tab4:
    st.markdown('<p class="sub-title">🚚 Bandeja de Pedidos y Cargues en Vivo (Enviados por Conductores)</p>', unsafe_allow_html=True)
    
    col_ctrl1, col_ctrl2 = st.columns([2, 2])
    with col_ctrl1:
        if st.button("🧹 Limpiar / Archivar Pedidos del Día Anterior"):
            st.session_state.pedidos_autoventa = {}
            st.success("¡Bandeja restablecida para los pedidos de hoy!")
            st.rerun()
    with col_ctrl2:
        # Checkbox en app web: permite ocultar o mostrar ceros en pantalla para revisión rápida
        ocultar_ceros_web = st.checkbox("🔍 Ocultar productos en cero en pantalla", value=False)

    if st.session_state.pedidos_autoventa:
        st.info(f"Hay pedidos activos de {len(st.session_state.pedidos_autoventa)} ruta(s) para hoy ({date.today()}).")
        
        dict_pedidos_excel = {}
        for r_nombre, info in st.session_state.pedidos_autoventa.items():
            with st.expander(f"📌 Ruta: {r_nombre} | Placa: {info['placa']} | Fecha: {info['fecha']}"):
                df_pedido_temp = pd.DataFrame.from_dict(info['productos'], orient='index')
                if 'SKU' in df_pedido_temp.columns:
                    df_pedido_temp = df_pedido_temp[['SKU', 'Descripción del Producto', 'Cantidad']]
                
                if ocultar_ceros_web and not df_pedido_temp.empty:
                    df_pedido_temp = df_pedido_temp[df_pedido_temp['Cantidad'] > 0]
                
                st.dataframe(df_pedido_temp, use_container_width=True)
                
                dict_pedidos_excel[r_nombre] = {
                    "placa": info['placa'],
                    "productos": info['productos']
                }
                
                if st.button(f"🗑️ Borrar Pedido de {r_nombre}", key=f"del_ped_{r_nombre}"):
                    del st.session_state.pedidos_autoventa[r_nombre]
                    st.rerun()
                    
        st.divider()
        if dict_pedidos_excel:
            excel_bodega_bytes = exportar_excel_estilo_bodega(dict_pedidos_excel, fecha_str=date.today().strftime('%Y-%m-%d'))
            st.download_button(
                label=f"📥 Descargar Cargue Formato Bodega ({date.today()})",
                data=excel_bodega_bytes,
                file_name=f"Cargue_Inesco_Consolidado_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_pedidos_excel_bodega"
            )
    else:
        st.info("ℹ️ No hay pedidos activos en este momento en la bandeja.")
