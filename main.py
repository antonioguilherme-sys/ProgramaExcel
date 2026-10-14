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
Esta aplicação consolida os registos da aba **ATESTADOS -ABONOS** (Agosto e Setembro de 2026) 
considerando as nomenclaturas específicas de ausências, cruza com a base cadastral e calcula o **Custo Total da Perda**.
""")

# Sidebar para upload de ficheiros
st.sidebar.header("📁 Upload de Ficheiros")

file_func = st.sidebar.file_uploader("1. Cadastro de Funcionários (.xlsx)", type=["xlsx"])
file_aug = st.sidebar.file_uploader("2. Relatório de Agosto/2026 (.xls, .xlsx)", type=["xls", "xlsx"])
file_sep = st.sidebar.file_uploader("3. Relatório de Setembro/2026 (.xls, .xlsx)", type=["xls", "xlsx"])

# Função de busca local caso não exista upload
def obter_fonte(file_uploaded, default_name):
    if file_uploaded is not None:
        return file_uploaded
    elif os.path.exists(default_name):
        return default_name
    return None

source_func = obter_fonte(file_func, 'informções_funcioanrios.xlsx')
source_aug = obter_fonte(file_aug, 'ABS GERAL AGOSTO 2026.xls')
source_sep = obter_fonte(file_sep, 'ABS GERAL SETEMBRO 2026.xls')

def extrair_ausencias_atestados(source, mes_nome):
    xls = pd.ExcelFile(source)
    registros = []
    
    # Processar EXCLUSIVAMENTE a Aba "ATESTADOS -ABONOS"
    if 'ATESTADOS -ABONOS' in xls.sheet_names:
        df_atest = pd.read_excel(xls, sheet_name='ATESTADOS -ABONOS')
        col_dias = 'Dias Abs' if 'Dias Abs' in df_atest.columns else 'Dias Afastamento'
        col_descr = 'Descr. Historico Ocorrência' if 'Descr. Historico Ocorrência' in df_atest.columns else 'Descrição.1'
        
        for _, row in df_atest.iterrows():
            descr_raw = str(row.get(col_descr, '')).strip().upper()
            
            # Mapeamento e categorização de acordo com as 7 nomenclaturas solicitadas
            motivo = None
            if 'HORA' in descr_raw:
                motivo = 'ATESTADO MÉDICO (HORAS)'
            elif 'FILHO' in descr_raw or 'ACOMP' in descr_raw:
                motivo = 'JUSTIFICATIVA DE ACOMP FILHO MENOR'
            elif 'JUÍZO' in descr_raw or 'JUIZO' in descr_raw:
                motivo = 'LICENÇA COMP. EM JUÍZO'
            elif 'SANGUE' in descr_raw or 'DOAÇÃO' in descr_raw or 'DOACAO' in descr_raw:
                motivo = 'LICENÇA DOAÇÃO DE SANGUE'
            elif 'ÓBITO' in descr_raw or 'OBITO' in descr_raw:
                motivo = 'LICENÇA ÓBITO'
            elif 'PATERNIDADE' in descr_raw:
                motivo = 'LICENÇA PATERNIDADE'
            elif 'ATESTADO' in descr_raw or 'MÉDICO' in descr_raw or 'MEDICO' in descr_raw:
                motivo = 'ATESTADO MÉDICO GERAL (DIAS)'

            # Apenas incluir se corresponder a uma das nomenclaturas
            if motivo is not None:
                dias = row.get(col_dias, 1)
                if pd.isna(dias) or dias <= 0:
                    dias = 1
                
                registros.append({
                    'Matrícula': str(row['Matricula']),
                    'Mes': mes_nome,
                    'Motivo': motivo,
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
    df_aug = extrair_ausencias_atestados(s_aug, 'Agosto/2026')
    df_sep = extrair_ausencias_atestados(s_sep, 'Setembro/2026')
    
    df_consolidado = pd.concat([df_aug, df_sep], ignore_index=True)
    df_consolidado = df_consolidado.groupby(['Matrícula', 'Mes', 'Motivo'], as_index=False)['Dias_Ausencia'].sum()

    # 3. Cálculo do Custo Total da Perda
    df_calculo = pd.merge(df_consolidado, df_func_clean, on='Matrícula', how='inner')
    df_calculo['Custo_Total_Perda'] = df_calculo['Dias_Ausencia'] * df_calculo['Custo_Diario']

    # 4. Gerar Excel na memória (BytesIO)
    output = BytesIO()
    wb = openpyxl.Workbook()
    
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
        ("Total Dias Perdidos (Bimestre)", total_dias, "#,##0"),
        ("Dias Perdidos em Agosto/2026", dias_aug, "#,##0"),
        ("Dias Perdidos em Setembro/2026", dias_sep, "#,##0"),
        ("Impacto Financeiro Total (Custo_Total_Perda)", total_custo, "R$ #,##0.00")
    ]

    r_idx = 4
    for label, val, fmt in kpis:
        ws_dash.cell(row=r_idx, column=2, value=label).font = Font(bold=True)
        cv = ws_dash.cell(row=r_idx, column=4, value=val)
        cv.font = Font(bold=True, color="1F4E78")
        cv.number_format = fmt
        r_idx += 1

    ws_dash.cell(row=10, column=2, value="Top Departamentos com Maior Custo de Absenteísmo").font = Font(size=12, bold=True, color="1F4E78")
    top_depto = df_calculo.groupby('Departamento')['Custo_Total_Perda'].sum().reset_index().sort_values(by='Custo_Total_Perda', ascending=False)
    
    ws_dash.cell(row=11, column=2, value="Departamento").fill = header_fill
    ws_dash.cell(row=11, column=2).font = header_font
    ws_dash.cell(row=11, column=3, value="Custo Total Perda (R$)").fill = header_fill
    ws_dash.cell(row=11, column=3).font = header_font

    r = 12
    for _, row in top_depto.head(10).iterrows():
        ws_dash.cell(row=r, column=2, value=row['Departamento'])
        c = ws_dash.cell(row=r, column=3, value=row['Custo_Total_Perda'])
        c.number_format = 'R$ #,##0.00'
        r += 1

    ws_dash.column_dimensions['B'].width = 42
    ws_dash.column_dimensions['C'].width = 25
    ws_dash.column_dimensions['D'].width = 22

    def criar_aba_df(title, df, fmt_moeda_cols=[], fmt_int_cols=[]):
        ws = wb.create_sheet(title=title)
        ws.views.sheetView[0].showGridLines = True
        
        ws.append(list(df.columns))
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            
        for row in df.itertuples(index=False):
            ws.append(list(row))
            
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                col_name = df.columns[cell.column - 1]
                if col_name in fmt_moeda_cols:
                    cell.number_format = 'R$ #,##0.00'
                elif col_name in fmt_int_cols:
                    cell.number_format = '#,##0'
                    
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Abas da Planilha Final
    criar_aba_df('Dados_Funcionarios', df_func_clean, fmt_moeda_cols=['Salario_Mensal', 'Custo_Diario'])
    criar_aba_df('Consolidado_Absenteismo', df_consolidado, fmt_int_cols=['Dias_Ausencia'])
    
    df_calc_out = df_calculo[['Matrícula', 'Nome', 'Departamento', 'Cargo', 'Mes', 'Motivo', 'Dias_Ausencia', 'Custo_Diario', 'Custo_Total_Perda']]
    criar_aba_df('Calculo_Custo_Perda', df_calc_out, 
                 fmt_moeda_cols=['Custo_Diario', 'Custo_Total_Perda'],
                 fmt_int_cols=['Dias_Ausencia'])

    wb.save(output)
    output.seek(0)
    return df_calculo, output

# Verificação dos ficheiros para execução
if source_func and source_aug and source_sep:
    st.success("✅ Todos os 3 ficheiros de entrada foram identificados com sucesso!")
    
    if st.button("🚀 Processar e Gerar Relatório"):
        with st.spinner("A processar bases de dados..."):
            df_calculo, excel_bytes = processar_dados(source_func, source_aug, source_sep)
            
            st.markdown("---")
            st.subheader("📌 Resumo Executivo (Bimestre - Aba ATESTADOS -ABONOS)")
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Dias Perdidos (Total)", f"{int(df_calculo['Dias_Ausencia'].sum())} dias")
            c2.metric("Agosto/2026", f"{int(df_calculo[df_calculo['Mes']=='Agosto/2026']['Dias_Ausencia'].sum())} dias")
            c3.metric("Setembro/2026", f"{int(df_calculo[df_calculo['Mes']=='Setembro/2026']['Dias_Ausencia'].sum())} dias")
            c4.metric("Custo Total da Perda", f"R$ {df_calculo['Custo_Total_Perda'].sum():,.2f}")
            
            st.markdown("---")
            st.subheader("📊 Distribuição por Motivo de Ausência")
            st.dataframe(df_calculo.groupby('Motivo')[['Dias_Ausencia', 'Custo_Total_Perda']].sum().style.format({'Custo_Total_Perda': 'R$ {:,.2f}', 'Dias_Ausencia': '{:.0f}'}))

            st.markdown("---")
            st.subheader("🏢 Top 5 Departamentos mais Afetados")
            top5_deptos = df_calculo.groupby('Departamento')[['Dias_Ausencia', 'Custo_Total_Perda']].sum().sort_values(by='Custo_Total_Perda', ascending=False).head(5)
            st.dataframe(top5_deptos.style.format({'Custo_Total_Perda': 'R$ {:,.2f}', 'Dias_Ausencia': '{:.0f}'}))
            
            st.download_button(
                label="📥 Baixar Planilha Consolidada Formatada (.xlsx)",
                data=excel_bytes,
                file_name="Relatorio_Impacto_Financeiro_Absenteismo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.warning("⚠️ Aguardando o upload dos 3 ficheiros para prosseguir.")
    st.info("Utilize a barra lateral à esquerda para carregar os ficheiros caso não estejam na pasta do projeto.")
