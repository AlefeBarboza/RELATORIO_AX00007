import streamlit as st
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import io

# Função parse_almoxarifado (mantida igual, mas com ajustes para Streamlit)
def parse_almoxarifado(file_content):
    consolidated_data = []
    current_almoxarifado = None
    reading_table = False

    almoxarifado_pattern = re.compile(r'^Almoxarifado:(.*?)\s*-\s*PBSAUDE')
    data_pattern = re.compile(r'^(\d+)\s*-\s*([^*]+)\*+(\w+)\*+([^*]+)\*+(\w+)\*+([\d,.]+)\*([\d,.]+)\*+([\d,.]+)\*+([\d,.]+)\*+([\d,.]+)\*+([\d,.]+)\*+')

    # Ler o conteúdo do arquivo como texto
    for line in file_content.decode('utf-8').splitlines():
        line = line.strip()
        if not line:
            reading_table = False
            continue

        almoxarifado_match = almoxarifado_pattern.match(line)
        if almoxarifado_match:
            current_almoxarifado = almoxarifado_match.group(1).strip()[8:]
            reading_table = True
            continue

        if reading_table:
            data_match = data_pattern.match(line)
            if data_match:
                cod_sigbp, material, unidade_medida, finalidade_compra, endereco, qtd_indisponivel, valor_indisponivel, qtd_disponivel, valor_disponivel, qtd_total, valor_total = data_match.groups()
                consolidated_data.append({
                    "Código": cod_sigbp,
                    "Almoxarifado": current_almoxarifado,
                    "Material": material.strip(),
                    "U.M.": unidade_medida.strip(),
                    "Qtd total": qtd_total.strip(),
                    "Valor total": valor_total.strip(),
                    "valor unitário": None,
                    "Incorporação/Baixa": None,
                    "Quantidade e Levantamento": 0.0
                })

    df = pd.DataFrame(consolidated_data)
    numeric_columns = ["Qtd total", "Valor total"]
    for col in numeric_columns:
        df[col] = df[col].str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df['valor unitário'] = df['Valor total'] / df["Qtd total"]

    wb = Workbook()
    ws = wb.active
    ws.title = "Tabela Consolidada"
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    col_qtd_total = 5
    col_incorporacao_baixa = 8
    col_qtd_levantamento = 9
    for row in range(2, len(df) + 2):
        ws.cell(row=row, column=col_incorporacao_baixa).value = f'=E{row}-I{row}'

    # Salvar o workbook em um buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return df, buffer

# Interface do Streamlit
st.title("Processador de Estoque Analítico")
st.write("Faça upload do arquivo .txt para gerar a tabela consolidada em Excel.")

uploaded_file = st.file_uploader("Escolha um arquivo .txt", type="txt")

if uploaded_file is not None:
    try:
        df, excel_buffer = parse_almoxarifado(uploaded_file.read())
        st.write("### Tabela Consolidada")
        st.dataframe(df)  # Exibir o DataFrame na interface
        st.download_button(
            label="Baixar Tabela Consolidada (Excel)",
            data=excel_buffer,
            file_name="tabela_consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.success("Tabela consolidada gerada com sucesso!")
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
