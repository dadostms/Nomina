import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.markdown(
    """
    <style>
    div[data-testid="stDecoration"] {
        background-color: #000000;

</style>
""",
    unsafe_allow_html=True
)

# Todos os Markdown são únicos para uma área somente
# markdown somente referente e tratante ao background em geral
st.markdown(
    """
    <style>
    .stApp {
        background-color: #000000;

background-color: #020202;
opacity: 1;
background-image:  linear-gradient(135deg, #0c0c0c 25%, transparent 25%), linear-gradient(225deg, #0c0c0c 25%, transparent 25%), linear-gradient(45deg, #0c0c0c 25%, transparent 25%), linear-gradient(315deg, #0c0c0c 25%, #020202 25%);
background-position:  25px 0, 25px 0, 0 0, 0 0;
background-size: 50px 50px;
background-repeat: repeat;


    </style>
    """,
    unsafe_allow_html=True
)

# markdown somente unicamente referente ao Sidebar
st.markdown(
    """
    <style>
    div[data-testid="stSidebarContent"] {

background-color: #e61c22;
opacity: 1;
background-image:  linear-gradient(135deg, #f60008 25%, transparent 25%), linear-gradient(225deg, #f60008 25%, transparent 25%), linear-gradient(45deg, #f60008 25%, transparent 25%), linear-gradient(315deg, #f60008 25%, #e61c22 25%);
background-position:  25px 0, 25px 0, 0 0, 0 0;
background-size: 50px 50px;
background-repeat: repeat;


    }

</style>
""",
    unsafe_allow_html=True
)

# Dicionário para o preenchimento da carga horária na lista de presença
curso_e_carga_horaria = {
    "NR 30 - Segurança no Trabalho Aquaviário - Capacitação Inicial": "4h",
    "NR 06 - Equipamentos de Proteção Individual (EPI)": "4h",
    "NR 10 - Segurança em Instalações e Serviços em Eletricidade (Básico)": "40h",
    "NR-12 Capacitação para Operação Segura de Máquinas e Equipamentos": "8h",
    "NR 33 - Resgate em espaço confinado - Nível Operacional ": "24h",
    "NR 33 - Espaços Confinados (Treinamento Supervisor) - Reciclagem": "8h",
    "NR 33 - Espaços Confinados (Treinamento Supervisor) - INICIAL": "40h",
    "NR 33 - Espaços Confinados (Treinamento Trabalhador e Vigia)": "16h",
    "NR 34 - Condições de Trabalho na Indústria da Construção e Reparação Naval (ADMISSIONAL/PERIÓDICO) ": "6h",
    "NR 34 - Condições de Trabalho na Indústria da Construção e Reparação Naval (Operadores de Equipamento de Guindar)": "20h",
    "NR 34 - Curso básico de segurança em operações de Movimentação de Cargas": "20h",
    "NR-34 OBSERVADOR DE TRABALHO A QUENTE": "8h",
    "NR 34 - Segurança nas atividades de Pintura": "4h",
    "NR 34 - Condições de Trabalho na Indústria da Construção e Reparação Naval (Básico Trabalho a Quente)": "20h",
    "NR 35 - Resgate em Altura": "24h",
    "NR 35 - Trabalho em altura": "8h",
}

# Menu no sidebar
st.sidebar.markdown("<h1 style='font-size: 62px;'>"
                    " </h1>"
                    "", unsafe_allow_html=True)
st.sidebar.markdown("<h1 style='font-size: 62px;'>"
                    " </h1>"
                    "", unsafe_allow_html=True)
st.sidebar.markdown("<h1 style='font-size: 62px;'>"
                    " </h1>"
                    "", unsafe_allow_html=True)
st.sidebar.markdown("<h1 style='font-size: 62px;'>"
                    " </h1>"
                    "", unsafe_allow_html=True)
logo = st.sidebar.image('logofinal2.png')
menu = st.sidebar.selectbox("Selecione as opções abaixo", ["Lista de Presença", "Grade"])

if menu == "Lista de Presença":

    st.title("Digite os Dados Abaixo 🌐")

    # input de informações para preenchimento

    periodo_treinamento = st.text_input("Digite o Período de Treinamento")
    nome_instrutor = st.text_input("Digite o Nome do Instrutor")
    nome_embarcacao = st.text_input("Digite o Nome da Embarcação")

    st.title("")
    st.title("")

    # ---------------------------------------

    st.title("Insira as Planilhas 📋")

    uploaded_modelo = st.file_uploader("Inserir modelo da Lista de Presença", type=["xlsx"])
    uploaded_matriz = st.file_uploader("Inserir a Crewlist x Matriz", type=["csv", "xlsx"])
    uploaded_agenda = st.file_uploader("Inserir Horários", type=["csv", "xlsx"])

    if uploaded_modelo and uploaded_matriz and uploaded_agenda:
        # Leitura dos dados
        modelo_bytes = uploaded_modelo.read()

        if uploaded_matriz.name.endswith('.csv'):
            df_matriz = pd.read_csv(uploaded_matriz)
        else:
            df_matriz = pd.read_excel(uploaded_matriz)

        if uploaded_agenda.name.endswith('.csv'):
            df_agenda = pd.read_csv(uploaded_agenda)
        else:
            df_agenda = pd.read_excel(uploaded_agenda)
        st.write("")
        st.write("")
        st.title("Seguem os Resultados Abaixo ✅")
        st.write("")
        st.write("")

        # Verifica se colunas necessárias existem
        colunas_necessarias = ['Curso Alterado Matriz', 'Nome Alterado']
        if not all(col in df_matriz.columns for col in colunas_necessarias):
            st.error("As colunas 'Curso Alterado Matriz' e 'Nome Alterado' não foram encontradas na Matriz.")
            st.stop()

        cursos = df_matriz['Curso Alterado Matriz'].dropna().unique()

        for curso in cursos:
            df_filtrado = df_matriz[df_matriz['Curso Alterado Matriz'] == curso]
            nomes = df_filtrado['Nome Alterado'].dropna().unique()

            caminho_temporario = f"modelo_temp_{curso}.xlsx"
            with open(caminho_temporario, "wb") as f:
                f.write(modelo_bytes)

            wb = load_workbook(caminho_temporario)
            ws = wb.active

            # Preencher curso em A13
            ws["A13"] = curso

            # parametro que pega o curso que está escrito em F13 procura na lista de curso x carga horária e preenche com o equivalente
            carga_horaria = curso_e_carga_horaria.get(curso,
                                                      "")  # ------------------------------------------------------------------------------------
            ws["F13"] = carga_horaria

            # Preencher nomes de A23 até A53
            for i, nome in enumerate(nomes):
                if i < 31:
                    ws[f"A{23 + i}"] = nome

            # Filtrar agenda por curso
            agenda_curso = df_agenda[df_agenda['Curso'] == curso]

            if not agenda_curso.empty:
                # Preencher datas nas células D19-H19
                dias_unicos = agenda_curso['Dias da Semana'].dropna().unique()
                for i, dia in enumerate(dias_unicos[:5]):
                    ws.cell(row=19, column=4 + i).value = dia  # D=4, E=5, e por ai vai

                # Agora, agrupar os dados por Curso e Dia da Semana
                for i, dia in enumerate(dias_unicos[:5]):
                    agenda_dia = agenda_curso[agenda_curso['Dias da Semana'] == dia]

                    # Encontrar a menor Hora Menor e a maior Hora Maior para o mesmo Curso e Dia da Semana
                    hora_menor = agenda_dia['Hora Menor'].min()
                    hora_maior = agenda_dia['Hora Maior'].max()

                    # Preencher as células com as horas
                    ws.cell(row=21, column=4 + i).value = hora_maior
                    ws.cell(row=22, column=4 + i).value = hora_menor

                    # Preencher as informações dos inputs
                    ws.cell(row=9, column=1).value = nome_embarcacao
                    ws.cell(row=13, column=3).value = nome_instrutor
                    ws.cell(row=10, column=5).value = periodo_treinamento

            # Salvar arquivo  final
            nome_arquivo = f"Lista_Presenca_{curso}.xlsx"
            wb.save(nome_arquivo)
            # st.success(f"Lista de presença criada para o curso: {curso}")

            with open(nome_arquivo, "rb") as f:
                st.download_button(f"📑 Baixar Lista: {curso}", f, file_name=nome_arquivo)

            st.title("")

elif menu == "Grade":
    st.file_uploader("Insira a Grade Vazia")