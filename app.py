import streamlit as st
import pandas as pd
import pdfplumber
import io
import os
import re
import json
from datetime import datetime, date, timezone, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Zona horaria de Colombia (UTC-5)
COLOMBIA_TZ = timezone(timedelta(hours=-5))

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

# ---------------------------------------------------------
# 2. CREDENCIALES SIMPLIFICADAS PARA CONDUCTORES Y ADMIN
# ---------------------------------------------------------
USUARIOS_SISTEMA = {
    "inesco2026admin": {"rol": "admin",  "ruta": None,     "placa": None},
    "71":              {"rol": "driver", "ruta": "ML3E71", "placa": "JYY 174"},
    "74":              {"rol": "driver", "ruta": "ML3E74", "placa": "JYY 262"},
    "75":              {"rol": "driver", "ruta": "ML3E75", "placa": "JYY 180"},
    "76":              {"rol": "driver", "ruta": "ML3E76", "placa": "JYY 177"},
    "77":              {"rol": "driver", "ruta": "ML3E77", "placa": "JYY 175"},
}

ARCHIVO_PEDIDOS = "pedidos_autoventa_db.json"
ARCHIVO_VENCIMIENTOS = "vencimientos_db.json"

def cargar_json(filepath, default_val):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default_val
    return default_val

def guardar_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

if "vencimientos" not in st.session_state:
    st.session_state.vencimientos = cargar_json(ARCHIVO_VENCIMIENTOS, [])

if "pedidos_autoventa" not in st.session_state:
    st.session_state.pedidos_autoventa = cargar_json(ARCHIVO_PEDIDOS, {})

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if "rol_usuario" not in st.session_state:
    st.session_state.rol_usuario = None

if "ruta_asignada" not in st.session_state:
    st.session_state.ruta_asignada = None

if "placa_asignada" not in st.session_state:
    st.session_state.placa_asignada = None

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
# 3. CATÁLOGO OFICIAL Y ESTRUCTURA ORDENADA POR SECCIONES
# ---------------------------------------------------------
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
    for sec, prods in CATALOGO_SECCIONES.items():
        for p in prods:
            if not any(item["sku"] == p["sku"] for item in st.session_state.catalogo):
                st.session_state.catalogo.append(p)

# ---------------------------------------------------------
# 4. EXPORTADOR EXCEL PROFESIONAL (BODEGA, TAMAÑO CARTA, CÁLCULOS Y OCULTO DE CEROS)
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
    
    for nombre_ruta, r_info in rutas_dict.items():
        safe_title = re.sub(r'[\\/*?:[\]]', '_', nombre_ruta)
        ws = wb.create_sheet(title=safe_title[:30])
        
        # Configuración de página tamaño Carta, horizontal ajustada a 1 página de ancho
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.horizontalCentered = True
        
        # Título superior
        ws.merge_cells('A1:G1')
        title_cell = ws.cell(row=1, column=1, value="DISTRIBUCIONES INESCO - HOJA DE CARGUE")
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Ruta y Placa centrados
        ws.merge_cells('A2:C2')
        r_cell = ws.cell(row=2, column=1, value=f"Ruta: {nombre_ruta}")
        r_cell.font = Font(bold=True, size=12, color="1F4E78")
        r_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells('D2:G2')
        p_cell = ws.cell(row=2, column=4, value=f"Placa: {r_info.get('placa', 'JYY 175')}")
        p_cell.font = Font(bold=True, size=12, color="1F4E78")
        p_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 24

        # Fecha centrada, resaltada y en negrita
        ws.merge_cells('A3:B3')
        f_lbl = ws.cell(row=3, column=1, value="FECHA:")
        f_lbl.font = Font(bold=True, size=11, color="000000")
        f_lbl.alignment = Alignment(horizontal="right", vertical="center")
        
        ws.merge_cells('C3:D3')
        date_cell = ws.cell(row=3, column=3, value=fecha_str)
        date_cell.font = Font(size=12, bold=True, color="C00000")
        date_cell.alignment = Alignment(horizontal="center", vertical="center")
        date_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        ws.row_dimensions[3].height = 22

        # Cabeceras
        headers = ["SKU", "ELEMENTO", "CANTIDAD", "DEVOLUCIÓN 1", "DEVOLUCIÓN 2", "DEVOLUCIÓN 3", "PISO / UBICACIÓN"]
        for col_idx, h_text in enumerate(headers, start=1):
            c = ws.cell(row=4, column=col_idx, value=h_text)
            c.fill = header_fill
            c.font = header_font
            c.border = black_border
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[4].height = 24
        
        ws.auto_filter.ref = f"A4:G1000"
            
        current_row = 5
        productos_dict_pedido = r_info.get("productos", {})
        filas_totales_secciones = []
        filas_secciones_retornables_cajas = [] # Para sumar botellas retornables (8oz y 350ml)
        
        for nombre_seccion, lista_prods in CATALOGO_SECCIONES.items():
            start_sec_row = current_row
            rows_in_section = 0
            
            for item in lista_prods:
                sku = item["sku"]
                desc = item["descripcion"]
                match_key = next((k for k in productos_dict_pedido if k.startswith(sku)), None)
                qty = productos_dict_pedido[match_key]["Cantidad"] if match_key else 0
                
                ws.cell(row=current_row, column=1, value=int(sku) if sku.isdigit() else sku).alignment = Alignment(horizontal="center", vertical="center")
                
                cell_desc = ws.cell(row=current_row, column=2, value=desc)
                cell_desc.font = Font(name="Calibri", size=10, bold=True)
                cell_desc.alignment = Alignment(horizontal="center", vertical="center")
                
                cell_qty = ws.cell(row=current_row, column=3, value=qty)
                cell_qty.alignment = Alignment(horizontal="center", vertical="center")
                
                ws.cell(row=current_row, column=4, value="")
                ws.cell(row=current_row, column=5, value="")
                ws.cell(row=current_row, column=6, value="")
                ws.cell(row=current_row, column=7, value="")
                
                for c_i in range(1, 8):
                    cell = ws.cell(row=current_row, column=c_i)
                    cell.border = black_border
                ws.row_dimensions[current_row].height = 20
                
                # Ocultar automáticamente filas con cantidad 0
                if qty == 0:
                    ws.row_dimensions[current_row].hidden = True
                    
                current_row += 1
                rows_in_section += 1
                
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
            
            # Si es la sección 1 o 2 (Productos retornables 8oz y 350ml), guardamos la celda del total de cajas
            if "1. PRODUCTOS 8 ONZAS" in nombre_seccion or "2. PRODUCTOS 350 ML" in nombre_seccion:
                filas_secciones_retornables_cajas.append(f"C{current_row}")
            
            for c_i in range(1, 8):
                cell = ws.cell(row=current_row, column=c_i)
                cell.fill = section_total_fill
                cell.font = total_font
                cell.border = black_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 22
            current_row += 1
            
        # Bloque final de Canastas, Estibas y Códigos de Botellas
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
        
        current_row += 2
        fila_inicio_extras = current_row
        
        for item_extra, cod_ext in extras:
            c_ext = ws.cell(row=current_row, column=2, value=item_extra)
            c_ext.font = Font(name="Calibri", size=10, bold=True)
            c_ext.alignment = Alignment(horizontal="center", vertical="center")
            
            ws.cell(row=current_row, column=1, value=cod_ext).alignment = Alignment(horizontal="center", vertical="center")
            
            # Si es canasta de botellas retornables, calcular automáticamente con fórmula el total de cajas de 8oz y 350ml
            if "CANASTA" in item_extra and ("8 OZ" in item_extra or "350ML" in item_extra) and filas_secciones_retornables_cajas:
                formula_cajas = "=" + "+".join(filas_secciones_retornables_cajas)
                ws.cell(row=current_row, column=3, value=formula_cajas)
            else:
                ws.cell(row=current_row, column=3, value="")
            
            for c_i in range(1, 8):
                cell = ws.cell(row=current_row, column=c_i)
                cell.border = black_border
                if c_i not in [2]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[current_row].height = 20
            current_row += 1

        # TOTAL CAJAS DEL CARGUE FINAL
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
# 5. EXTRACCIÓN DE PDF Y VENCIMIENTOS (ORIGINAL RESTAURADO)
# ---------------------------------------------------------
def procesar_pdf_rutas(pdf_file):
    routes_data = {}
    mapa_catalogo = {item["sku"]: item["descripcion"] for item in st.session_state.catalogo}
    with pdfplumber.open(pdf_file) as pdf:
        current_route = "ML3E51"
        current_header_info = "Ruta / No.de Carga: ML3E51"
        for page in pdf.pages:
            texto_pagina = page.extract_text()
            if not texto_pagina:
                continue
            for linea in texto_pagina.split('\n'):
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
                    cajas, unidades = 0, 0
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
                        routes_data[current_route] = {"header": current_header_info, "items": []}
                    routes_data[current_route]["items"].append({
                        "SKU": sku,
                        "Descripción del Producto": descripcion_final,
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
    except Exception:
        return "⚪ DESCONOCIDO", 9999

def convertir_df_a_excel(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in df_dict.items():
            safe_name = re.sub(r'[\\/*?:[\]]', '_', str(sheet_name))[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return output.getvalue()

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
    fill_rojo_celda = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_rojo_celda = Font(name="Calibri", size=11, color="9C0006", bold=True)
    fill_amarillo_celda = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_amarillo_celda = Font(name="Calibri", size=11, color="9C6500", bold=True)
    fill_verde_celda = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_verde_celda = Font(name="Calibri", size=11, color="006100")

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
    cell_s = ws.cell(row=2, column=1, value=f"📌 Reporte Agrupado por Secciones del Catálogo (Fecha: {fecha_str})")
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

    catalogo_actual = st.session_state.get("catalogo", [])
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
            f_fill, f_font = fill_rojo_celda, font_rojo_celda
        elif r_data["Dias"] <= 60:
            f_fill, f_font = fill_amarillo_celda, font_amarillo_celda
        else:
            f_fill, f_font = fill_verde_celda, font_verde_celda

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
# 6. PANTALLA DE ACCESO (LOGIN SIMPLIFICADO CON CONTRASEÑA)
# ---------------------------------------------------------
def limpiar_campos_login():
    st.session_state["input_pwd"] = ""

if not st.session_state.autenticado:
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; font-size: 1.8rem; font-weight: bold; color: #E41E2B;">Distribuciones Inesco</p>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; color: #666;">Sistema de Control Logístico, Bodega y Autoventa</p><br>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.container():
            st.markdown("### 🔐 Iniciar Sesión")
            pwd_input = st.text_input("Ingrese su Contraseña (o número de ruta 71, 74...):", type="password", key="input_pwd").strip()
            
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                if st.button("Ingresar", use_container_width=True):
                    if pwd_input in USUARIOS_SISTEMA:
                        datos_usr = USUARIOS_SISTEMA[pwd_input]
                        st.session_state.autenticado = True
                        st.session_state.rol_usuario = datos_usr["rol"]
                        st.session_state.ruta_asignada = datos_usr["ruta"]
                        st.session_state.placa_asignada = datos_usr["placa"]
                        st.rerun()
                    else:
                        st.error("❌ Contraseña incorrecta.")
            with col_b2:
                st.button("Limpiar", on_click=limpiar_campos_login, use_container_width=True)
    st.stop()

# ---------------------------------------------------------
# 7. PORTAL DE CONDUCTORES (CON RECARGA DE PEDIDO PREVIO Y HORA COLOMBIA)
# ---------------------------------------------------------
if st.session_state.rol_usuario == "driver":
    ruta_fija = st.session_state.ruta_asignada
    placa_fija = st.session_state.placa_asignada

    st.markdown(f'<p class="sub-title">📱 Portal de Autoventa | Ruta Asignada: {ruta_fija} (Placa: {placa_fija})</p>', unsafe_allow_html=True)
    st.info(f"Has iniciado sesión exclusivamente para la **{ruta_fija}** — Vehículo **{placa_fija}**.")
    
    if st.button("🚪 Cerrar Sesión"):
        st.session_state.autenticado = False
        st.session_state.rol_usuario = None
        st.session_state.ruta_asignada = None
        st.session_state.placa_asignada = None
        st.rerun()
        
    st.markdown("---")
    st.markdown("#### 📦 Digita las cantidades solicitadas por sección:")
    
    # Cargar pedidos existentes para precargar si ya había enviado antes hoy
    hoy_str = datetime.now(COLOMBIA_TZ).strftime("%Y-%m-%d")
    pedidos_actuales = cargar_json(ARCHIVO_PEDIDOS, {})
    pedido_existente_ruta = pedidos_actuales.get(ruta_fija, {})
    
    # Si es del mismo día, precargamos los productos ya pedidos en el session_state si no están cargados
    if pedido_existente_ruta.get("fecha") == hoy_str:
        prod_previos = pedido_existente_ruta.get("productos", {})
        for k, v in prod_previos.items():
            key_st = f"drv_{ruta_fija}_{v['SKU']}"
            if key_st not in st.session_state:
                st.session_state[key_st] = v['Cantidad']

    pedido_ruta = {}
    for nombre_seccion, lista_productos in CATALOGO_SECCIONES.items():
        with st.expander(f"📌 {nombre_seccion}", expanded=False):
            cant_seccion_actual = 0
            for idx, item in enumerate(lista_productos):
                key_input = f"drv_{ruta_fija}_{item['sku']}"
                val_inicial = st.session_state.get(key_input, 0)
                
                cant = st.number_input(
                    f"{item['sku']} - {item['descripcion']}",
                    min_value=0,
                    step=1,
                    value=val_inicial,
                    key=key_input
                )
                if cant > 0:
                    cant_seccion_actual += 1
                    pedido_ruta[f"{item['sku']}_{idx}"] = {
                        "SKU": item['sku'],
                        "Descripción del Producto": item['descripcion'],
                        "Cantidad": cant
                    }
            if cant_seccion_actual > 0:
                st.caption(f"✅ Llevas **{cant_seccion_actual}** producto(s) seleccionados en esta sección.")
            
    st.markdown("---")
    col_env1, col_env2 = st.columns(2)
    with col_env1:
        if st.button("🚀 Enviar Solicitud de Carga a Bodega"):
            if pedido_ruta:
                hora_colombia = datetime.now(COLOMBIA_TZ).strftime("%H:%M:%S")
                pedidos_actuales[ruta_fija] = {
                    "placa": placa_fija,
                    "fecha": hoy_str,
                    "hora_envio": hora_colombia,
                    "productos": pedido_ruta
                }
                guardar_json(ARCHIVO_PEDIDOS, pedidos_actuales)
                st.session_state.pedidos_autoventa = pedidos_actuales
                st.success(f"✅ ¡Pedido de la {ruta_fija} ({placa_fija}) enviado y guardado exitosamente a las {hora_colombia}!")
            else:
                st.warning("⚠️ Digita al menos un producto con cantidad mayor a 0.")
    with col_env2:
        if pedido_ruta:
            excel_driver_bytes = exportar_excel_estilo_bodega({ruta_fija: {"placa": placa_fija, "productos": pedido_ruta}}, fecha_str=hoy_str)
            st.download_button(
                label="📥 Descargar mi Solicitud en Excel (.xlsx)",
                data=excel_driver_bytes,
                file_name=f"Solicitud_{ruta_fija}_{hoy_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.stop()

# ---------------------------------------------------------
# 8. PANEL DE ADMINISTRACIÓN (PESTAÑAS ORIGINALES RESTAURADAS)
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

# --- PESTAÑA 1: FECHAS DE VENCIMIENTO ---
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
                "SKU": s_code,
                "Descripción del Producto": s_desc,
                "Cantidad": cant_final,
                "Fecha Vencimiento": nueva_fecha_str
            })
            guardar_json(ARCHIVO_VENCIMIENTOS, st.session_state.vencimientos)
            st.success("✅ ¡Guardado correctamente!")
            st.rerun()

    if st.session_state.vencimientos:
        if st.button("🗑️ Limpiar Todos los Vencimientos"):
            st.session_state.vencimientos = []
            guardar_json(ARCHIVO_VENCIMIENTOS, [])
            st.rerun()

        filas_venc = []
        for reg in st.session_state.vencimientos:
            estado_txt, dias_rest = calcular_estado_vencimiento(reg["Fecha Vencimiento"])
            filas_venc.append({
                "SKU": reg["SKU"],
                "Descripción del Producto": reg["Descripción del Producto"],
                "Cantidad": reg["Cantidad"],
                "Fecha Vencimiento": reg["Fecha Vencimiento"],
                "Días Restantes": dias_rest if dias_rest != 9999 else "-",
                "Estado": estado_txt
            })
        df_venc = pd.DataFrame(filas_venc)
        st.dataframe(df_venc, use_container_width=True)
        
        excel_venc_bytes = exportar_excel_vencimientos_organizado(st.session_state.vencimientos, fecha_str=date.today().strftime('%d/%m/%Y'))
        st.download_button(
            label="📥 Descargar Reporte de Vencimientos (.xlsx)",
            data=excel_venc_bytes,
            file_name=f"Vencimientos_Inesco_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("ℹ️ No hay fechas de vencimiento registradas.")

# --- PESTAÑA 2: ADMINISTRAR SKUS ---
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
    
    st.markdown("#### Catálogo Actual:")
    st.dataframe(pd.DataFrame(st.session_state.catalogo), use_container_width=True)

# --- PESTAÑA 3: EXTRACCIÓN PDF (RUTAS) ---
with tab3:
    st.markdown('<p class="sub-title">📄 Extracción Total de Rutas y Cargues de FEMSA</p>', unsafe_allow_html=True)
    archivo_pdf = st.file_uploader("📂 Seleccionar archivo PDF de Rutas", type=["pdf"], key="uploader_pdf_rutas")
    if archivo_pdf is not None:
        if st.button("🚀 Procesar PDF"):
            with st.spinner("🔄 Procesando PDF..."):
                datos_rutas = procesar_pdf_rutas(archivo_pdf)
                if datos_rutas:
                    st.session_state["datos_pdf_extraidos"] = datos_rutas
                    st.success(f"✅ Se procesaron {len(datos_rutas)} ruta(s).")
                else:
                    st.warning("⚠️ No se detectaron rutas.")

    if "datos_pdf_extraidos" in st.session_state and st.session_state["datos_pdf_extraidos"]:
        datos_rutas = st.session_state["datos_pdf_extraidos"]
        dfs_para_excel = {}
        for r_code, r_content in datos_rutas.items():
            with st.expander(f"🚚 {r_content['header']} ({len(r_content['items'])} registros)", expanded=True):
                df_r = pd.DataFrame(r_content["items"])
                st.dataframe(df_r, use_container_width=True)
                dfs_para_excel[r_code] = {"df": df_r, "header": r_content["header"]}
        
        excel_pdf_bytes = exportar_excel_multiruta(dfs_para_excel, fecha_str=date.today().strftime('%d/%m/%Y'), es_reporte_rutas=True)
        st.download_button(
            label="📥 Descargar Extracción de Rutas en Excel (.xlsx)",
            data=excel_pdf_bytes,
            file_name=f"Extraccion_Rutas_FEMSA_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- PESTAÑA 4: RECEPCIÓN DE CARGUES (AUTOVENTA) ---
with tab4:
    st.markdown('<p class="sub-title">🚚 Bandeja de Pedidos y Cargues en Vivo (Enviados por Conductores)</p>', unsafe_allow_html=True)
    
    # Control automático de hora (Corte a las 12:30 PM hora colombiana)
    ahora_col = datetime.now(COLOMBIA_TZ)
    limite_corte = ahora_col.replace(hour=12, minute=30, second=0, microsecond=0)
    
    pedidos_guardados = cargar_json(ARCHIVO_PEDIDOS, {})
    st.session_state.pedidos_autoventa = pedidos_guardados

    col_ctrl1, col_ctrl2 = st.columns([2, 2])
    with col_ctrl1:
        if st.button("🔄 Actualizar Bandeja"):
            st.rerun()
    with col_ctrl2:
        if st.button("🧹 Limpiar / Archivar Pedidos"):
            guardar_json(ARCHIVO_PEDIDOS, {})
            st.session_state.pedidos_autoventa = {}
            st.success("¡Bandeja restablecida!")
            st.rerun()

    if pedidos_guardados:
        st.info(f"Hay pedidos activos de {len(pedidos_guardados)} ruta(s) guardados en el sistema.")
        
        dict_pedidos_excel = {}
        for r_nombre, info in pedidos_guardados.items():
            hora_txt = f" | Hora: {info['hora_envio']}" if 'hora_envio' in info else ""
            with st.expander(f"📌 Ruta: {r_nombre} | Placa: {info['placa']} | Fecha: {info['fecha']}{hora_txt}", expanded=True):
                df_pedido_temp = pd.DataFrame.from_dict(info['productos'], orient='index')
                if 'SKU' in df_pedido_temp.columns:
                    df_pedido_temp = df_pedido_temp[['SKU', 'Descripción del Producto', 'Cantidad']]
                
                st.dataframe(df_pedido_temp, use_container_width=True)
                
                dict_pedidos_excel[r_nombre] = {
                    "placa": info['placa'],
                    "productos": info['productos']
                }
                
                if st.button(f"🗑️ Borrar Pedido de {r_nombre}", key=f"del_ped_{r_nombre}"):
                    del pedidos_guardados[r_nombre]
                    guardar_json(ARCHIVO_PEDIDOS, pedidos_guardados)
                    st.session_state.pedidos_autoventa = pedidos_guardados
                    st.rerun()
                    
        st.divider()
        if dict_pedidos_excel:
            excel_bodega_bytes = exportar_excel_estilo_bodega(dict_pedidos_excel, fecha_str=datetime.now(COLOMBIA_TZ).strftime('%Y-%m-%d'))
            st.download_button(
                label=f"📥 Descargar Cargue Formato Bodega ({datetime.now(COLOMBIA_TZ).strftime('%Y-%m-%d')})",
                data=excel_bodega_bytes,
                file_name=f"Cargue_Inesco_Consolidado_{datetime.now(COLOMBIA_TZ).strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_pedidos_excel_bodega"
            )
    else:
        st.info("ℹ️ No hay pedidos activos en este momento en la bandeja.")
