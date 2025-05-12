import streamlit as st
from datetime import datetime

year = datetime.now().strftime("%Y")
last_year = int(year)-1

st.title(f'Conferência Ranking STN {last_year}')
st.write("Instruções")

with st.container():
    col1, col2 = st.columns(2)

    with col1:
        st.subheader('D2')
        st.write('Matriz usada para preenchimento da sheet D2 na planilha de conferência')
        matriz_d2 = st.file_uploader("Escolha um arquivo para a extração do D2", type="csv")
     
    with col2:
        st.subheader('D4')
        st.write('Matriz usada para preenchimento da sheet D4 na planilha de conferência')
        matriz_d4 = st.file_uploader("Escolha um arquivo para a extração do D4", type="csv")

with st.container():
    col1,col2 = st.columns(2)

    with col1:
        st.subheader('DCA')
        st.write('Matriz usada para preenchimento da sheet D2 na planilha de conferência')
        matriz_dca = st.file_uploader("Escolha um arquivo para a extração do DCA", type="csv")
        
    with col2:
        st.subheader('RREO')
        st.write('Matriz usada para preenchimento referente ao RREO na planilha de conferência')
        matriz_dca = st.file_uploader("Escolha um arquivo para a extração do RREO", type="csv")

with st.container():
    col1,col2 = st.columns(2)

    with col1:
        st.subheader('RGF')
        st.write('Matriz usada para preenchimento referente ao RGF na planilha de conferência')
        matriz_dca = st.file_uploader("Escolha um arquivo para a extração do RGF", type="csv")
        
    with col2:
        st.subheader('Planilha de Conferência')
        st.write('Matriz usada para preenchimento da sheet D2 na planilha de conferência')
        matriz_dca = st.file_uploader("Escolha o arquivo que será preenchido", type="csv")

with st.container():
    col1, col2, col3 = st.columns([1.5, 1, 1])  # col2 é maior (no centro)

    with col2:
        if st.button('INICIAR'):
            print('Processo Iniciado')