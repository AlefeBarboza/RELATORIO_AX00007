import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import re
import io

# Função para validar nomes de abas no Excel
def sanitize_sheet_name(name):
    invalid_chars = r'[\/:*?[\]]'  # Caracteres inválidos no Excel
    name = re.sub(invalid_chars, '_', str(name))
    return name[:31]  # Limitar a 31 caracteres

# Função para processar o arquivo e criar o Excel
def parse_almoxarifado(file_content):
    consolidated_data = []
    current_almoxarifado = None
    reading_table = False

    # Regex para identificar almoxarifado e linhas de dados
    almoxarifado_pattern = re.compile(r'Almoxarifado:§*(\d+)\s*-\s*([^-§]+)\s*-\s*([^-§]+)\s*-\s*([^-§]+)\s*-\s*([^-§]+)§+')
    data_pattern = re.compile(r'(\d+)\s*-\s*([^§]+)\§+(\w+)\§+([^§]+)\§+(\w+)\§+([\d,.]+)\§+([\d,.]+)\§+([\d,.]+)\§+([\d,.]+)\§+([\d,.]+)\§+([\d,.]+)\§+')

    # Ler o conteúdo do arquivo como texto
    for line in file_content.decode('utf-8').splitlines():
        line = line.strip()
        if not line:
            reading_table = False
            continue

        # Identificar almoxarifado
        almoxarifado_match = almoxarifado_pattern.match(line)
        if almoxarifado_match:
            current_almoxarifado_name = f"{almoxarifado_match.group(1).strip()[5:]} - {almoxarifado_match.group(5).strip()}"  # Extrair número do almoxarifado
            reading_table = True
            continue

        # Processar linhas de dados
        if reading_table:
            data_match = data_pattern.match(line)
            if data_match:
                cod_sigbp, material, unidade_medida, finalidade_compra, endereco, qtd_indisponivel, valor_indisponivel, qtd_disponivel, valor_disponivel, qtd_total, valor_total = data_match.groups()
                consolidated_data.append({
                    "Almoxarifado": current_almoxarifado_name,
                    "Código": cod_sigbp,                    
                    "Material": material.strip(),
                    "U.M.": unidade_medida.strip(),
                    "Qtd total": qtd_total.strip(),
                    "Valor total": valor_total.strip(),
                    "valor unitário": None,
                    "Incorporação/Baixa": None,
                    "Quantidade e Levantamento": 0.0
                })

    # Criar DataFrame
    df = pd.DataFrame(consolidated_data)

    # Converter colunas numéricas
    numeric_columns = ["Qtd total", "Valor total"]
    for col in numeric_columns:
        df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Calcular valor unitário
    df['valor unitário'] = df['Valor total'] / df["Qtd total"]

    # Criar workbook
    wb = Workbook()
    wb.remove(wb.active)  # Remover aba padrão

    # Definir cores para estilização
    header_fill = PatternFill(start_color="006400", end_color="006400", fill_type="solid")  # Verde escuro para cabeçalho
    green_fill = PatternFill(start_color="DAF2D0", end_color="DAF2D0", fill_type="solid")  # Verde claro
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # Branco
    gray_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")   # Cinza claro
    header_font = Font(bold=True, color="FFFFFF")  # Fonte branca e negrito para cabeçalho
    center_alignment = Alignment(horizontal="center", vertical="center")  # Alinhamento central

    # Definir bordas finas
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
       
    # Agrupar por almoxarifado
    grouped = df.groupby('Almoxarifado')

    # Criar uma aba para cada almoxarifado
    for almoxarifado, group_df in grouped:
        # Create a valid sheet title by replacing invalid characters (like '/')
        valid_sheet_title = str(almoxarifado).replace('/', '-')

        # Criar uma nova aba com o nome do almoxarifado
        ws = wb.create_sheet(title=valid_sheet_title)

        # Escrever o DataFrame na planilha
        for r_idx, r in enumerate(dataframe_to_rows(group_df, index=False, header=True), 1):
            ws.append(r)

       # Aplicar estilo ao cabeçalho (primeira linha)
            if r_idx == 1:
                for c_idx, cell in enumerate(ws[r_idx], 1):
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_alignment
                    cell.border = thin_border  # Adicionar bordas ao cabeçalho
            else:
                # Aplicar cores alternadas nas colunas A até G (1 a 7)
                for c_idx in range(1, 8):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.fill = green_fill if (r_idx % 2 == 0) else white_fill
                    cell.alignment = center_alignment
                    cell.border = thin_border  # Adicionar bordas às colunas A-G

                # Aplicar cinza claro nas colunas H e I (8 e 9)
                for c_idx in range(8, 10):
                    cell = ws.cell(row=r_idx, column=c_idx)
                    cell.fill = gray_fill
                    cell.alignment = center_alignment
                    cell.border = thin_border  # Adicionar bordas às colunas H-I
        
        # Identificar colunas no Excel
        col_qtd_total = 5  # Coluna E
        col_incorporacao_baixa = 8  # Coluna H
        col_qtd_levantamento = 9  # Coluna I

        # Inserir fórmula na coluna H
        for row in range(2, len(group_df) + 2):
            ws.cell(row=row, column=col_incorporacao_baixa).value = f'=E{row}-I{row}'
            ws.cell(row=row, column=col_incorporacao_baixa).border = thin_border  # Garantir borda na célula com fórmula

          # Ajustar largura das colunas para melhor visualização
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Limitar largura máxima
            ws.column_dimensions[column].width = adjusted_width
   
    # Salvar o workbook em um buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return df, buffer

# Interface do Streamlit
st.title("Processador de Estoque Analítico")
st.write("Faça upload do arquivo .txt para gerar a tabela consolidada em Excel com abas por almoxarifado.")

uploaded_file = st.file_uploader("Escolha um arquivo .txt", type="txt")

if uploaded_file is not None:
    try:
        df, excel_buffer = parse_almoxarifado(uploaded_file.read())
        st.write("### Tabela Consolidada (Visão Geral)")
        st.dataframe(df)  # Exibir o DataFrame completo
        st.download_button(
            label="Baixar Tabela Consolidada (Excel)",
            data=excel_buffer,
            file_name="tabela_consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Tabela consolidada gerada com sucesso! Cada almoxarifado está em uma aba separada.")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
