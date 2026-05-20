import streamlit as st
import pandas as pd
import pdfplumber
import io
import xlsxwriter
import chromadb
from openai import OpenAI
import time

# ==========================================
# CLASES DE PROCESAMIENTO
# ==========================================
class ProcesadorBC3:
    def procesar(self, archivo):
        datos_ejemplo = {
            "Código": ["TUB-050", "ROC-102", "DET-01", "BMB-DIE", "VAL-01"],
            "Descripción": [
                "Tubería acero ranurado DN50 para red de rociadores", 
                "Rociador pendent oculto 68ºC acabado blanco", 
                "Detector óptico de humo analógico", 
                "Grupo de presión contra incendios diésel 1000 l/m",
                "Válvula de mariposa ranurada DN50"
            ],
            "Unidad": ["m", "ud", "ud", "ud", "ud"],
            "Precio": [22.10, 18.50, 35.00, 18500.00, 145.00]
        }
        return pd.DataFrame(datos_ejemplo)

class ProcesadorCliente:
    def procesar_excel(self, archivo):
        return pd.read_excel(archivo)
        
    def procesar_pdf(self, archivo):
        filas = []
        with pdfplumber.open(archivo) as pdf:
            for pagina in pdf.pages:
                tablas = pagina.extract_tables()
                for tabla in tablas:
                    for fila in tabla:
                        if any(fila):
                            filas.append(fila)
        if filas:
            return pd.DataFrame(filas[1:], columns=filas[0])
        return pd.DataFrame()

# ==========================================
# MOTOR SEMÁNTICO IA
# ==========================================
class MotorSemantico:
    def __init__(self, api_key):
        self.cliente = OpenAI(api_key=api_key)
        self.chroma_client = chromadb.EphemeralClient()
        self.collection = self.chroma_client.get_or_create_collection(name="catalogo_pci")
        
    def get_embedding(self, texto):
        response = self.cliente.embeddings.create(
            input=texto,
            model="text-embedding-3-small"
        )
        return response.data[0].embedding
        
    def indexar_catalogo(self, df_bc3):
        for index, row in df_bc3.iterrows():
            embedding = self.get_embedding(row['Descripción'])
            self.collection.upsert(
                documents=[row['Descripción']],
                metadatas=[{"codigo": row['Código'], "precio": row['Precio'], "unidad": row['Unidad']}],
                ids=[row['Código']],
                embeddings=[embedding]
            )
            
    def buscar_similitud(self, texto_cliente):
        try:
            embedding_busqueda = self.get_embedding(texto_cliente)
            resultados = self.collection.query(
                query_embeddings=[embedding_busqueda],
                n_results=1
            )
            if resultados['documents'] and resultados['documents'][0]:
                meta = resultados['metadatas'][0][0]
                distancia = resultados['distances'][0][0] if 'distances' in resultados else 0.5
                confianza = max(0, min(100, int((1 - distancia) * 100)))
                return {
                    "Sugerencia BC3": resultados['documents'][0][0],
                    "Código": meta['codigo'],
                    "Precio Unit.": meta['precio'],
                    "Confianza": f"{confianza}%"
                }
        except Exception as e:
            return {"Sugerencia BC3": "Error", "Código": "-", "Precio Unit.": 0.0, "Confianza": "0%"}
        return {"Sugerencia BC3": "Sin coincidencia", "Código": "-", "Precio Unit.": 0.0, "Confianza": "0%"}

# ==========================================
# EXPORTADOR EXCEL
# ==========================================
class GeneradorExcel:
    def generar(self, df_datos, margen):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Presupuesto_Detallado')
        
        formato_cabecera = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#b30000', 'border': 1, 'align': 'center'})
        formato_moneda = workbook.add_format({'num_format': '#,##0.00 €', 'border': 1})
        formato_normal = workbook.add_format({'border': 1})
        
        cabeceras = ['Código', 'Descripción', 'Unidad', 'Cantidad', 'Precio Unitario', 'Total Fila']
        for col_num, cabecera in enumerate(cabeceras):
            worksheet.write(0, col_num, cabecera, formato_cabecera)
            
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 50)
        worksheet.set_column('C:D', 12)
        worksheet.set_column('E:F', 18)

        fila_excel = 1
        for index, row in df_datos.iterrows():
            precio_con_margen = float(row['Precio Unit.']) * (1 + (margen/100))
            cantidad = float(row['Cantidad'])
            
            worksheet.write(fila_excel, 0, row['Código'], formato_normal)
            worksheet.write(fila_excel, 1, row['Sugerencia BC3'], formato_normal)
            worksheet.write(fila_excel, 2, "ud", formato_normal)
            worksheet.write(fila_excel, 3, cantidad, formato_normal)
            worksheet.write(fila_excel, 4, precio_con_margen, formato_moneda)
            worksheet.write_formula(fila_excel, 5, f'=D{fila_excel+1}*E{fila_excel+1}', formato_moneda)
            fila_excel += 1
            
        worksheet.write(fila_excel + 1, 4, "TOTAL:", formato_cabecera)
        worksheet.write_formula(fila_excel + 1, 5, f'=SUM(F2:F{fila_excel})', formato_moneda)

        workbook.close()
        return output.getvalue()

# ==========================================
# INTERFAZ STREAMLIT
# ==========================================
st.set_page_config(page_title="Presupuestos PCI", layout="wide", page_icon="🔥")
st.title("🔥 Automatizador de Presupuestos PCI")

if 'df_catalogo' not in st.session_state: st.session_state.df_catalogo = None
if 'df_cliente' not in st.session_state: st.session_state.df_cliente = None
if 'df_emparejado' not in st.session_state: st.session_state.df_emparejado = None

with st.sidebar:
    st.header("⚙️ Configuración")
    api_key = st.text_input("OpenAI API Key", type="password")
    margen = st.number_input("Margen (%)", min_value=0, max_value=100, value=15)
    bc3_file = st.file_uploader("Sube tu .bc3", type=['bc3'])
    if bc3_file:
        st.session_state.df_catalogo = ProcesadorBC3().procesar(bc3_file)
        st.success("✅ BC3 cargado.")

tab1, tab2, tab3 = st.tabs(["📄 1. Medición", "🧠 2. IA", "📊 3. Excel"])

with tab1:
    cliente_file = st.file_uploader("Sube Excel o PDF del cliente", type=['pdf', 'xlsx'])
    if cliente_file:
        datos_simulados = {
            "Concepto Solicitado": ["Tuberia ranurada", "Rociador oculto", "Bomba contra incendios", "Detector humos"],
            "Cantidad": [150, 45, 1, 12]
        }
        st.session_state.df_cliente = pd.DataFrame(datos_simulados)
        st.dataframe(st.session_state.df_cliente)

with tab2:
    if st.button("Ejecutar IA") and api_key and st.session_state.df_catalogo is not None:
        motor = MotorSemantico(api_key)
        motor.indexar_catalogo(st.session_state.df_catalogo)
        resultados = []
        for index, row in st.session_state.df_cliente.iterrows():
            match = motor.buscar_similitud(row['Concepto Solicitado'])
            resultados.append({
                "Concepto Original": row['Concepto Solicitado'], "Cantidad": row['Cantidad'],
                "Sugerencia BC3": match['Sugerencia BC3'], "Código": match['Código'],
                "Precio Unit.": match['Precio Unit.'], "Confianza": match['Confianza']
            })
        st.session_state.df_emparejado = pd.DataFrame(resultados)
        st.dataframe(st.session_state.df_emparejado)

with tab3:
    if st.session_state.df_emparejado is not None:
        excel_data = GeneradorExcel().generar(st.session_state.df_emparejado, margen)
        st.download_button("📥 DESCARGAR EXCEL", data=excel_data, file_name="Presupuesto.xlsx")