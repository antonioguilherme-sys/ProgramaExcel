import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def processar_absenteismo():
    print("1. Lendo arquivos de entrada...")
    
    # 1. Carregar cadastro de funcionários
    df_func = pd.read_excel('informções_funcioanrios.xlsx', sheet_name='Planilha1')
    
    # Selecionar e renomear colunas relevantes
    df_func_clean = df_func[['Matrícula', 'Nome', 'Descrição (Cargos)', 'Descrição (Departamentos)', 'Salário Base']].copy()
    df_func_clean.columns = ['Matrícula', 'Nome', 'Cargo', 'Departamento', 'Salario_Mensal']
    df_func_clean['Matrícula'] = df_func_clean['Matrícula'].astype(str)
    df_func_clean['Custo_Diario'] = df_func_clean['Salario_Mensal'] / 22

    # 2. Carregar e unificar ausências por doença (Atestados + INSS)
    def extrair_ausencias_doenca(file_path, mes_nome):
        xls = pd.ExcelFile(file_path)
        registros = []
        
        # Leitura da aba ATESTADOS -ABONOS
        if 'ATESTADOS -ABONOS' in xls.sheet_names:
            df_atest = pd.read_excel(xls, sheet_name='ATESTADOS -ABONOS')
            col_dias = 'Dias Abs' if 'Dias Abs' in df_atest.columns else 'Dias Afastamento'
            col_descr = 'Descr. Historico Ocorrência' if 'Descr. Historico Ocorrência' in df_atest.columns else 'Descrição.1'
            
            for _, row in df_atest.iterrows():
                descr = str(row.get(col_descr, ''))
                # Filtrar licenças não médicas (paternidade, óbito, casamento, doação de sangue)
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

        # Leitura da aba INSS (Auxílio Doença)
        if 'INSS' in xls.sheet_names:
            df_inss = pd.read_excel(xls, sheet_name='INSS')
            for _, row in df_inss.iterrows():
                dias = row.get('Dias Afastamento', row.get('Dias Previstos', 1))
                if pd.isna(dias) or dias <= 0 or dias > 180: # Ajuste de limite razoável
                    dias = 15 # Padrão dos primeiros 15 dias pagos pela empresa
                
                registros.append({
                    'Matrícula': str(row['Matricula']),
                    'Mes': mes_nome,
                    'Motivo': 'Auxílio Doença / INSS',
                    'Dias_Ausencia': float(dias)
                })
                
        return pd.DataFrame(registros)

    df_aug = extrair_ausencias_doenca('ABS GERAL AGOSTO 2026.xls', 'Agosto/2026')
    df_sep = extrair_ausencias_doenca('ABS GERAL SETEMBRO 2026.xls', 'Setembro/2026')
    
    # Consolidação
    df_consolidado = pd.concat([df_aug, df_sep], ignore_index=True)
    
    # Agrupar ausências por funcionário e mês para evitar duplicidades
    df_consolidado = df_consolidado.groupby(['Matrícula', 'Mes', 'Motivo'], as_index=False)['Dias_Ausencia'].sum()

    print("2. Calculando custos de perda...")
    # 3. Cruzamento de dados (Calculo_Custo_Perda)
    df_calculo = pd.merge(df_consolidado, df_func_clean, on='Matrícula', how='inner')
    
    df_calculo['Custo_Direto'] = df_calculo['Dias_Ausencia'] * df_calculo['Custo_Diario']
    df_calculo['Custo_Cobertura'] = df_calculo['Custo_Direto'] * 0.50 # Premissa: 50% de custo de substituição/HE
    df_calculo['Custo_Total_Perda'] = df_calculo['Custo_Direto'] + df_calculo['Custo_Cobertura']

    # 4. Gerar arquivo Excel com openpyxl e formatação visual
    output_filename = "Relatorio_Impacto_Financeiro_Absenteismo.xlsx"
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')
    
    # Salvar abas
    df_func_clean.to_excel(writer, sheet_name='Dados_Funcionarios', index=False)
    df_consolidado.to_excel(writer, sheet_name='Consolidado_Absenteismo', index=False)
    
    df_calculo_out = df_calculo[['Matrícula', 'Nome', 'Departamento', 'Cargo', 'Mes', 'Dias_Ausencia', 'Custo_Diario', 'Custo_Direto', 'Custo_Cobertura', 'Custo_Total_Perda']]
    df_calculo_out.to_excel(writer, sheet_name='Calculo_Custo_Perda', index=False)

    writer.close()

    print("3. Aplicando formatação visual avançada...")
    # Estilização do Excel
    wb = openpyxl.load_workbook(output_filename)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Azul escuro corporativo
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                         right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'),
                         bottom=Side(style='thin', color='D9D9D9'))

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True
        
        # Formatar cabeçalho
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Formatar células
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = thin_border
                # Formatação numérica
                if sheetname == 'Dados_Funcionarios' and cell.column in [5, 6]:
                    cell.number_format = 'R$ #,##0.00'
                elif sheetname == 'Calculo_Custo_Perda' and cell.column in [7, 8, 9, 10]:
                    cell.number_format = 'R$ #,##0.00'
                elif cell.column_letter in ['E', 'F'] and sheetname == 'Calculo_Custo_Perda':
                    cell.number_format = '#,##0'

        # Ajustar largura das colunas
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # 5. Criar Aba Dashboard_Resumo
    ws_dash = wb.create_sheet(title="Dashboard_Resumo", index=0)
    ws_dash.views.sheetView[0].showGridLines = True

    # Título do Dashboard
    ws_dash['B2'] = "DASHBOARD EXECUTIVO - IMPACTO FINANCEIRO DO ABSENTEÍSMO"
    ws_dash['B2'].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")

    # KPIs Principais
    total_dias = df_calculo['Dias_Ausencia'].sum()
    total_custo = df_calculo['Custo_Total_Perda'].sum()
    dias_aug = df_calculo[df_calculo['Mes'] == 'Agosto/2026']['Dias_Ausencia'].sum()
    dias_sep = df_calculo[df_calculo['Mes'] == 'Setembro/2026']['Dias_Ausencia'].sum()

    kpis = [
        ("Total Dias Perdidos (Bimestre)", total_dias, "#,##0"),
        ("Dias Perdidos (Agosto/2026)", dias_aug, "#,##0"),
        ("Dias Perdidos (Setembro/2026)", dias_sep, "#,##0"),
        ("Impacto Financeiro Total (R$)", total_custo, "R$ #,##0.00")
    ]

    row_idx = 4
    for label, val, fmt in kpis:
        ws_dash.cell(row=row_idx, column=2, value=label).font = Font(bold=True)
        cell_v = ws_dash.cell(row=row_idx, column=4, value=val)
        cell_v.font = Font(bold=True, color="1F4E78")
        cell_v.number_format = fmt
        row_idx += 1

    # Tabela Top Departamentos
    ws_dash.cell(row=10, column=2, value="Departamentos com Maior Custo de Absenteísmo").font = Font(size=12, bold=True, color="1F4E78")
    top_depto = df_calculo.groupby('Departamento')['Custo_Total_Perda'].sum().reset_index().sort_values(by='Custo_Total_Perda', ascending=False)
    
    ws_dash.cell(row=11, column=2, value="Departamento").fill = header_fill
    ws_dash.cell(row=11, column=2).font = header_font
    ws_dash.cell(row=11, column=3, value="Custo Total (R$)").fill = header_fill
    ws_dash.cell(row=11, column=3).font = header_font

    r = 12
    for _, row in top_depto.head(10).iterrows():
        ws_dash.cell(row=r, column=2, value=row['Departamento'])
        c = ws_dash.cell(row=r, column=3, value=row['Custo_Total_Perda'])
        c.number_format = 'R$ #,##0.00'
        r += 1

    # Ajustar colunas do Dashboard
    ws_dash.column_dimensions['B'].width = 45
    ws_dash.column_dimensions['C'].width = 22
    ws_dash.column_dimensions['D'].width = 22

    wb.save(output_filename)
    print(f"Sucesso! Planilha final gerada: {output_filename}")

if __name__ == "__main__":
    processar_absenteismo()
