import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
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
                    "Código": cod_sigbp,
                    "Almoxarifado": current_almoxarifado_name,
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

    # Agrupar por almoxarifado
    grouped = df.groupby('Almoxarifado')

    # Criar uma aba para cada almoxarifado
    for almoxarifado, group_df in grouped:
        # Create a valid sheet title by replacing invalid characters (like '/')
        valid_sheet_title = str(almoxarifado).replace('/', '-')

        # Criar uma nova aba com o nome do almoxarifado
        ws = wb.create_sheet(title=valid_sheet_title)

        # Escrever o DataFrame na planilha
        for r in dataframe_to_rows(group_df, index=False, header=True):
            ws.append(r)

        # Identificar colunas no Excel
        col_qtd_total = 5  # Coluna E
        col_incorporacao_baixa = 8  # Coluna H
        col_qtd_levantamento = 9  # Coluna I

        # Inserir fórmula na coluna H
        for row in range(2, len(group_df) + 2):
            ws.cell(row=row, column=col_incorporacao_baixa).value = f'=E{row}-I{row}'

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
