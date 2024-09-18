import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import base64
import openpyxl

# Configuração da página
st.set_page_config(
    page_title="Dashboard de Receita da Mesa de Renda Variável",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Título do dashboard
st.title("Dashboard de Receita da Mesa de Renda Variável")

# Definir tamanho da página para paginação
PAGE_SIZE = 50

# Carregamento dos arquivos Excel
st.sidebar.header("Carregar Planilhas Excel")
uploaded_file_produtos = st.sidebar.file_uploader("Carregar Planilha de Produtos Estruturados", type=["xlsx"])
uploaded_file_corretagem = st.sidebar.file_uploader("Carregar Planilha de Corretagem", type=["xlsx"])

# Variáveis para armazenar os dataframes
df_produtos = None
df_corretagem = None

# Leitura da planilha de Produtos Estruturados
if uploaded_file_produtos is not None:
    try:
        df_produtos = pd.read_excel(uploaded_file_produtos)
        expected_columns_produtos = ['Código Cliente', 'Data da Operação', 'Ação da Estrutura', 'Comissão Gerada', 'Assessor', 'Status']
        missing_columns = [col for col in expected_columns_produtos if col not in df_produtos.columns]
        if missing_columns:
            st.error("Colunas Faltando na Planilha de Produtos Estruturados: " + ', '.join(missing_columns))
        else:
            # Filtrar 'Status' == 'Totalmente Executado'
            df_produtos = df_produtos[df_produtos['Status'] == 'Totalmente Executado']
    except Exception as e:
        st.error("Erro ao ler a planilha de Produtos Estruturados: " + str(e))

# Leitura da planilha de Corretagem
if uploaded_file_corretagem is not None:
    try:
        df_corretagem = pd.read_excel(uploaded_file_corretagem)
        expected_columns_corretagem = ['Código Cliente', 'Data da Operação', 'Comissão BMF', 'Comissão BOV', 'Receita Total', 'Código Assessor', 'Canal']
        missing_columns = [col for col in expected_columns_corretagem if col not in df_corretagem.columns]
        if missing_columns:
            st.error("Colunas Faltando na Planilha de Corretagem: " + ', '.join(missing_columns))
        else:
            # Realizar cálculos
            df_corretagem['Corretagem Bruta'] = df_corretagem['Receita Total']
            df_corretagem['Corretagem Líquida'] = df_corretagem['Corretagem Bruta'] * 0.95
            df_corretagem['Receita Escritório'] = df_corretagem['Corretagem Líquida'] * 0.75
            df_corretagem['Receita da Mesa'] = df_corretagem['Receita Escritório'] * 0.2
    except Exception as e:
        st.error("Erro ao ler a planilha de Corretagem: " + str(e))

# Se pelo menos um dos arquivos foi carregado
if (df_produtos is not None) or (df_corretagem is not None):
    # Criar abas para as planilhas
    tabs = st.tabs(["Produtos Estruturados", "Corretagem"])

    # Função para paginação
    def paginate_dataframe(df, page_size):
        total_rows = df.shape[0]
        total_pages = total_rows // page_size + (total_rows % page_size > 0)
        page = st.number_input('Página', min_value=1, max_value=total_pages, step=1)
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size
        return df.iloc[start_idx:end_idx]

    # Aba Produtos Estruturados
    if df_produtos is not None:
        with tabs[0]:
            st.header("Produtos Estruturados")

            # Filtros
            unique_assessores_prod = df_produtos['Assessor'].drop_duplicates().sort_values()
            selected_assessor_prod = st.multiselect("Selecionar Assessor", unique_assessores_prod)

            unique_clientes_prod = df_produtos['Código Cliente'].drop_duplicates().sort_values()
            selected_cliente_prod = st.multiselect("Selecionar Cliente", unique_clientes_prod)

            # Aplicar filtros
            df_filtered_produtos = df_produtos.copy()
            if selected_assessor_prod:
                df_filtered_produtos = df_filtered_produtos[df_filtered_produtos['Assessor'].isin(selected_assessor_prod)]
            if selected_cliente_prod:
                df_filtered_produtos = df_filtered_produtos[df_filtered_produtos['Código Cliente'].isin(selected_cliente_prod)]

            # Formatação de datas e valores monetários
            df_filtered_produtos['Data da Operação'] = pd.to_datetime(df_filtered_produtos['Data da Operação']).dt.strftime('%d/%m/%Y')
            df_filtered_produtos['Comissão Gerada'] = df_filtered_produtos['Comissão Gerada'].apply(lambda x: f"R$ {x:,.2f}")

            # Somatório para o card
            total_comissao_produtos = df_filtered_produtos['Comissão Gerada'].replace('[R\$\s,]', '', regex=True).astype(float).sum()
            st.metric(label="Total Comissão Gerada", value=f"R$ {total_comissao_produtos:,.2f}")

            # Paginação
            df_produtos_paginated = paginate_dataframe(df_filtered_produtos, PAGE_SIZE)

            # Exibir tabela
            st.dataframe(df_produtos_paginated)

            # Opção de exportar dados
            export_option = st.selectbox("Exportar Dados", ["Nenhum", "Excel", "PDF"], key='export_produtos')
            if export_option == "Excel":
                towrite = io.BytesIO()
                df_filtered_produtos.to_excel(towrite, index=False)
                towrite.seek(0)
                b64 = base64.b64encode(towrite.read()).decode()
                linko = f'<a href="data:application/octet-stream;base64,{b64}" download="produtos_filtrados.xlsx">Baixar Excel</a>'
                st.markdown(linko, unsafe_allow_html=True)
            elif export_option == "PDF":
                st.warning("Exportação para PDF não implementada")

    # Aba Corretagem
    if df_corretagem is not None:
        with tabs[1]:
            st.header("Corretagem")

            # Filtros
            unique_assessores = df_corretagem['Código Assessor'].drop_duplicates().sort_values()
            selected_assessor = st.multiselect("Selecionar Assessor", unique_assessores)

            unique_clientes = df_corretagem['Código Cliente'].drop_duplicates().sort_values()
            selected_cliente = st.multiselect("Selecionar Cliente", unique_clientes)

            unique_canais = df_corretagem['Canal'].drop_duplicates().sort_values()
            selected_canal = st.multiselect("Selecionar Canal", unique_canais)

            # Aplicar filtros
            df_filtered_corretagem = df_corretagem.copy()
            if selected_assessor:
                df_filtered_corretagem = df_filtered_corretagem[df_filtered_corretagem['Código Assessor'].isin(selected_assessor)]
            if selected_cliente:
                df_filtered_corretagem = df_filtered_corretagem[df_filtered_corretagem['Código Cliente'].isin(selected_cliente)]
            if selected_canal:
                df_filtered_corretagem = df_filtered_corretagem[df_filtered_corretagem['Canal'].isin(selected_canal)]

            # Formatação de datas e valores monetários
            df_filtered_corretagem['Data da Operação'] = pd.to_datetime(df_filtered_corretagem['Data da Operação']).dt.strftime('%d/%m/%Y')
            currency_columns = ['Comissão BMF', 'Comissão BOV', 'Receita Total', 'Corretagem Bruta', 'Corretagem Líquida', 'Receita Escritório', 'Receita da Mesa']
            for col in currency_columns:
                df_filtered_corretagem[col] = df_filtered_corretagem[col].apply(lambda x: f"R$ {x:,.2f}")

            # Somatório para o card
            total_receita_mesa = df_filtered_corretagem['Receita da Mesa'].replace('[R\$\s,]', '', regex=True).astype(float).sum()
            st.metric(label="Total Receita da Mesa", value=f"R$ {total_receita_mesa:,.2f}")

            # Paginação
            df_corretagem_paginated = paginate_dataframe(df_filtered_corretagem, PAGE_SIZE)

            # Exibir tabela
            st.dataframe(df_corretagem_paginated)

            # Opção de exportar dados
            export_option_corretagem = st.selectbox("Exportar Dados", ["Nenhum", "Excel", "PDF"], key='export_corretagem')
            if export_option_corretagem == "Excel":
                towrite = io.BytesIO()
                df_filtered_corretagem.to_excel(towrite, index=False)
                towrite.seek(0)
                b64 = base64.b64encode(towrite.read()).decode()
                linko = f'<a href="data:application/octet-stream;base64,{b64}" download="corretagem_filtrada.xlsx">Baixar Excel</a>'
                st.markdown(linko, unsafe_allow_html=True)
            elif export_option_corretagem == "PDF":
                st.warning("Exportação para PDF não implementada")
