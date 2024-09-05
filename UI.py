import streamlit as st
import hashlib
import json
import os
import pandas as pd
from PyPDF2 import PdfFileReader
from streamlit_option_menu import option_menu

#INICIO DATAFRAMES
if 'df_emps' not in st.session_state:
    st.session_state.df_emps = pd.DataFrame(pd.NA, index=[0], columns=['Coluna1', 'Coluna2', 'Coluna3', 'Coluna4','Coluna5'])


#INICIO DAS FUNÇOES
#########################################################################################

def cria_emprestimo(nome, valor, tempo, mes, planilha_selecionada):
    meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro','Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
    indice = 0
    for i in range(len(meses)):
        if meses[i]==mes:
            indice=i
            break
    planilha_selecionada = planilha_selecionada.append({'Coluna1':pd.NA,'Coluna2':pd.NA,'Coluna3':pd.NA,'Coluna4':pd.NA,'Coluna5':pd.NA},ignore_index=True)
    planilha_selecionada = planilha_selecionada.append({'Coluna1':nome,'Coluna2':'Mês','Coluna3':'Valor','Coluna4':'Situação','Coluna5':pd.NA},ignore_index=True)
    planilha_selecionada = planilha_selecionada.append({'Coluna1':pd.NA,'Coluna2':pd.NA,'Coluna3':pd.NA,'Coluna4':pd.NA,'Coluna5':pd.NA},ignore_index=True)
    for i in range(int(tempo)):
        planilha_selecionada = planilha_selecionada.append({'Coluna1':pd.NA,'Coluna2':pd.NA,'Coluna3':meses[indice+i],'Coluna4':int(valor)/int(tempo),'Coluna5':pd.NA},ignore_index=True)
    return planilha_selecionada




def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def load_users():
    if os.path.exists("users.json"):
        with open("users.json", "r") as file:
            return json.load(file)
    return {}

def save_users(users):
    with open("users.json", "w") as file:
        json.dump(users, file)

users_db = load_users()

def authenticate(username, password):
    if username in users_db and users_db[username]['password'] == hash_password(password):
        return True
    return False

def create_user(username, password, role):
    if username in users_db:
        return False
    users_db[username] = {"password": hash_password(password), "role": role}
    save_users(users_db)
    return True

def change_password(username, old_password, new_password):
    if authenticate(username, old_password):
        users_db[username]['password'] = hash_password(new_password)
        save_users(users_db)
        return True
    return False

def has_permission(username, role):
    return users_db.get(username, {}).get("role") == role

#FIM DAS FUNCOES
####################################################################################

st.title("Sistema de Planilhas AP Marinha")
st.sidebar.image('imgmar.jpeg')
st.sidebar.header("Login")
username = st.sidebar.text_input("Usuário")
password = st.sidebar.text_input("Senha", type="password")
login_button = st.sidebar.button("Login")

if login_button:
    if authenticate(username, password):
        st.sidebar.success(f"Bem-vindo {username}!")
        st.session_state['authenticated'] = True
        st.session_state['username'] = username
        st.session_state['role'] = users_db[username]['role']
    else:
        st.sidebar.error("Usuário ou senha incorretos.")
####################################################################################


# COMEÇA AQUI A PARTIR DE LOGADO


####################################################################################
if 'authenticated' in st.session_state and st.session_state['authenticated']:
    st.write(f"Usuário logado: {st.session_state['username']} ({st.session_state['role']})")


    #CASO EDITOR:
    if st.session_state['role'] == 'editor':

        with st.sidebar:
            selected = option_menu(
                    "Menu",
                    ["Planilhas", 'Empréstimo', 'Alterar Senha', 'Criar Usuário'],
                    icons=['info', 'calculator'],)

        if selected=='Planilhas':
            st.title('Planilhas')
            xls = pd.ExcelFile('big planilha.xlsx')
            planilha_selecionada = st.selectbox('Selecione a planilha desejada', ['BP_Pagamento','Condomínio Papem', 'Taxa_de_Condomínio', 'Despesas', 'ReceitasxDespesas', 'Previsão orçamentária', 'Taxa complementar', 'Empréstimo', 'Demonstrativo Junho'])
            df = pd.read_excel(xls, sheet_name=planilha_selecionada)
            df = st.data_editor(df, num_rows='dynamic')
            if planilha_selecionada == 'Empréstimo':
                with st.form("Formulário de empréstimo"):
                    st.write("Insira as informações do novo empréstimo")
                    nome_emprestimo = st.text_input('Nome')
                    valor_emprestimo = st.text_input('Valor (utilize o formato xx.xxx.xxx,xx)')
                    tempo_meses = st.text_input('Número de parcelas')
                    mes_inicio = st.text_input('Mês da primeira parcela')
                    submit_button = st.form_submit_button('Criar empréstimo')
                    if submit_button:
                        st.session_state.df_emps = cria_emprestimo(nome_emprestimo, valor_emprestimo, tempo_meses, mes_inicio, st.session_state.df_emps)
                    print(st.session_state.df_emps)
            elif planilha_selecionada == 'Previsão orçamentária':
                uploaded_file_0 = st.file_uploader("Upload da Planilha Serviços Diversos", type=["xls", "xlsx"])
                algo = pd.ExcelFile('11. Desconto Vnavi.ods')
                df_3 = pd.read_excel(algo)
                dsc = df_3['VALOR'].sum()
                outro = pd.ExcelFile('1. Planilha condomínios_compilada.xls')
                df_2 = pd.read_excel(outro, sheet_name='Previsão orçamentária')
                df_2.loc[49, 'Unnamed: 8'] = dsc
                df = st.data_editor(df_2, num_rows='dynamic')
                if uploaded_file_0 is not None:
                    df = pd.read_excel(uploaded_file_0)
                    st.write("Here is the content of the file:")
                    st.dataframe(df)

                uploaded_file_1 = st.file_uploader("Upload da Planilha Desconto VNAVI", type=["xls", "xlsx"])
                if uploaded_file_1 is not None:
                    df = pd.read_excel(uploaded_file_1)
                    st.write("Here is the content of the file:")
                    st.dataframe(df)
            
            elif planilha_selecionada == 'Demonstrativo Junho':
                uploaded_file = st.file_uploader("Upload da Planilha Taxa Extra", type=["xls", "xlsx"])
                if uploaded_file is not None:
                    taxa_extra = pd.read_excel(uploaded_file)
                    st.write("Here is the content of the file:")
                    st.dataframe(taxa_extra)

                uploaded_file_2 = st.file_uploader("Upload da Planilha Desocupados", type=["xls", "xlsx"])
                if uploaded_file_2 is not None:
                    desocupados = pd.read_excel(uploaded_file_2)
                    st.write("Here is the content of the file:")
                    st.dataframe(desocupados)

                uploaded_file_3 = st.file_uploader("Upload da Planilha Isolados", type=["xls", "xlsx"])
                if uploaded_file_3 is not None:
                    isolados = pd.read_excel(uploaded_file_3)
                    st.write("Here is the content of the file:")
                    st.dataframe(isolados)

                uploaded_file_4 = st.file_uploader("Upload da Planilha Restituições", type=["xls", "xlsx"])
                if uploaded_file_4 is not None:
                    restituicoes = pd.read_excel(uploaded_file_4)
                    st.write("Here is the content of the file:")
                    st.dataframe(restituicoes)
            st.session_state.df_emps = st.data_editor(st.session_state.df_emps, num_rows = 'dynamic')

        elif selected=='Alterar Senha':
                st.subheader("Alterar senha")
                old_password = st.text_input("Senha antiga", type="password")
                new_password = st.text_input("Nova senha", type="password")
                change_password_button = st.button("Alterar senha")

                if change_password_button:
                    if change_password(st.session_state['username'], old_password, new_password):
                        st.success("Senha alterada com sucesso!")
                    else:
                        st.error("Senha antiga incorreta.")
        elif selected == 'Criar Usuário':
            st.subheader("Criar novo usuário")
            new_username = st.text_input("Novo usuário")
            new_password = st.text_input("Primeira Senha", type="password")
            new_role = st.selectbox("Função", ["editor", "viewer"])
            create_user_button = st.button("Criar usuário")

            if create_user_button:
                if create_user(new_username, new_password, new_role):
                    st.success("Usuário criado com sucesso!")
                else:
                    st.error("Usuário já existe.")
    #CASO NÃO EDITOR:
    else:
        with st.sidebar:
            selected = option_menu(
                    "Menu",
                    ["Planilhas", 'Empréstimo', 'Alterar Senha', 'Criar Usuário'],
                    icons=['info', 'calculator'],)

        if selected=='Planilhas':
            st.title('Planilhas')
            xls = pd.ExcelFile('big planilha.xlsx')
            planilha_selecionada = st.selectbox('Selecione a planilha desejada', ['BP_Pagamento','Condomínio Papem', 'Taxa_de_Condomínio', 'Despesas', 'ReceitasxDespesas', 'Previsão orçamentária', 'Taxa complementar', 'Empréstimo'])
            df = pd.read_excel(xls, sheet_name=planilha_selecionada)
            df = st.dataframe(df)
                    
        elif selected=='Alterar Senha':
            st.subheader("Alterar senha")
            old_password = st.text_input("Senha antiga", type="password")
            new_password = st.text_input("Nova senha", type="password")
            change_password_button = st.button("Alterar senha")

            if change_password_button:
                if change_password(st.session_state['username'], old_password, new_password):
                    st.success("Senha alterada com sucesso!")
                else:
                    st.error("Senha antiga incorreta.")



else:
    st.write("Por favor, faça login para continuar.")
