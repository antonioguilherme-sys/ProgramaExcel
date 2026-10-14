import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import os

# Configuração da página do Streamlit
st.set_page_config(
    page_title="Gestão de Absenteísmo - Impacto Financeiro",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Calculadora do Impacto Financeiro do Absenteísmo")
st.markdown("""
Esta aplicação consolida os registros de ausência por motivo de doença (Agosto e Setembro de 2026), 
cruza com a base cadastral de funcionários e calcula os custos diretos e de cobertura/substituição.
""")

# Sidebar para upload de arquivos
st.sidebar.header("📁 Upload de Arquivos")

file_func = st.sidebar.file_uploader("1. Cadastro de Funcionários (.xlsx)", type=["xlsx"])
file_aug = st.sidebar.file_uploader("2. Relatório de Agosto/2026 (.xls, .xlsx)", type=["xls", "xlsx"])
file_sep = st.sidebar.file_uploader("3. Relatório de Setembro/2026 (.xls, .xlsx)", type=["xls", "xlsx"])

# Função de busca local caso o usuário não faça upload
def obter_fonte(file_uploaded, default_name):
    if file_uploaded is not None:
        return file_uploaded
    elif os.path.exists(default_name):
        return default_name
    return None

source_func = obter_fonte(file_func, 'informções_funcioanrios.xlsx')
source_aug = obter_fonte(file_aug, 'ABS GERAL AGOSTO 2026.xls')
source_sep = obter_fonte(file_sep, 'ABS GERAL SETEMBRO 2026.xls')

def extrair_ausencias_doenca(source, mes_nome):
    xls = pd.ExcelFile(source)
    registros = []
    
    # Processar Aba ATESTADOS -ABONOS
    if 'ATESTADOS -ABONOS' in xls.sheet_names:
        df_atest = pd.read_excel(xls, sheet_name='ATESTADOS -ABONOS')
        col_dias = 'Dias Abs' if 'Dias Abs' in df_atest.columns else 'Dias Afastamento'
        col_descr = 'Descr. Historico Ocorrência' if 'Descr. Historico Ocorrência' in df_atest.columns else 'Descrição.1'
        
        for _, row in df_atest.iterrows():
            descr = str(row.get(col_descr, ''))
            # Filtrar licenças que não são doença/médicas
            if any(termo in descr.upper() for termo in ['PATERNIDADE', 'ÓBITO', 'OBITO', 'CASAMENTO', 'DOAÇÃO', 'SANGUE', 'JUÍZO', 'JUIZO']):
                continue
            
            dias = row.get(col_dias, 1)
            if pd.isna(dias) or dias <= 0:
                dias = 1
            
            registros.append({
                'Matrícula': str(row['Matricula']),
                'Mes': mes_nome,
                'Motivo': 'Atestado Médico / Doença',
                'Dias_Ausencia': float(dias)
            })

    # Processar Aba INSS
    if 'INSS' in xls.sheet_names:
        df_inss = pd.read_excel(xls, sheet_name='INSS')
        for _, row in df_inss.iterrows():
            dias = row.get('Dias Afastamento', row.get('Dias Previstos', 1))
            if pd.isna(dias) or dias <= 0 or dias > 180:
                dias = 15 # Padrão dos primeiros 15 dias pagos pela empresa
            
            registros.append({
                'Matrícula': str(row['Matricula']),
                'Mes': mes_nome,
                'Motivo': 'Auxílio Doença / INSS',
                'Dias_Ausencia': float(dias)
            })
            
    return pd.DataFrame(registros)

def processar_dados(s_func, s_aug, s_sep):
    # 1. Cadastro de Funcionários
    df_func = pd.read_excel(s_func, sheet_name=0)
    df_func_clean = df_func[['Matrícula', 'Nome', 'Descrição (Cargos)', 'Descrição (Departamentos)', 'Salário Base']].copy()
    df_func_clean.columns = ['Matrícula', 'Nome', 'Cargo', 'Departamento', 'Salario_Mensal']
    df_func_clean['Matrícula'] = df_func_clean['Matrícula'].astype(str)
    df_func_clean['Custo_Diario'] = df_func_clean['Salario_Mensal'] / 22

    # 2. Ausências
    df_aug = extrair_ausencias_doenca(s_aug, 'Agosto/2026')
    df_sep = extrair_ausencias_doenca(s_sep, 'Setembro/2026')
    
    df_consolidado = pd.concat([df_aug, df_sep], ignore_index=True)
    df_consolidado = df_consolidado.groupby(['Matrícula', 'Mes', 'Motivo'], as_index=False)['Dias_Ausencia'].sum()

    # 3. Cálculos
    df_calculo = pd.merge(df_consolidado, df_func_clean, on='Matrícula', how='inner')
    df_calculo['Custo_Direto'] = df_calculo['Dias_Ausencia'] * df_calculo['Custo_Diario']
    df_calculo['Custo_Cobertura'] = df_calculo['Custo_Direto'] * 0.50 # Cobertura/Hora Extra (50%)
    df_calculo['Custo_Total_Perda'] = df_calculo['Custo_Direto'] + df_calculo['Custo_Cobertura']

    # 4. Gerar Excel na memória (BytesIO)
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # Remover aba padrão
    wb.remove(wb.active)

    # Estilos
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                         right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'),
                         bottom=Side(style='thin', color='D9D9D9'))

    # Aba 1: Dashboard_Resumo
    ws_dash = wb.create_sheet(title="Dashboard_Resumo")
    ws_dash.views.sheetView[0].showGridLines = True
    
    ws_dash['B2'] = "DASHBOARD EXECUTIVO - IMPACTO FINANCEIRO DO ABSENTEÍSMO"
    ws_dash['B2'].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")

    total_dias = df_calculo['Dias_Ausencia'].sum()
    total_custo = df_calculo['Custo_Total_Perda'].sum()
    dias_aug = df_calculo[df_calculo['Mes'] == 'Agosto/2026']['Dias_Ausencia'].sum()
    dias_sep = df_calculo[df_calculo['Mes'] == 'Setembro/2026']['Dias_Ausencia'].sum()

    kpis = [
        ("Total Dias Perdidos (Bimestre)", total_dias,
