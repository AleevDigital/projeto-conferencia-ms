import streamlit as st
from datetime import datetime

year = datetime.now().strftime("%Y")
print(year)

st.title(f'Conferência Ranking STN {year}')
st.write("Seja bem vindo!")