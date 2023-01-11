#!/usr/bin/env python
# coding: utf-8

# ##  Bibliotecas:

# In[12]:


import pandas as pd 
import numpy as np 
import streamlit as st
from PIL import Image
from datetime import date
import base64
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
from xlsxwriter import Workbook




# ##  Barra Lateral: 

# In[16]:


#Barra Lateral
barra_lateral = st.sidebar.empty()
image = Image.open('Logo-Escuro.png')
st.sidebar.image(image)
st.sidebar.markdown('### Levantamentos Uso de Área FieldScan')
tipo_analise = st.sidebar.selectbox("📊 Tipo de Ánalise:", ['Solos','Drone'])

if tipo_analise == 'Solos':

    page_bg_img = """
    <style>
    [data-testid="stAppViewContainer"] > .main {
    background-image: url("https://s2.glbimg.com/-KiNeS2rsxOE3clNFotv5Y1LdQs=/780x440/e.glbimg.com/og/ed/f/original/2020/08/10/gettyimages-1025633076.jpg");
    background-size: cover;
    }

    [data-testid="stHeader"] {
    background: rgba(0,0,0,0);
    }

    </style>
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)


    # ##  Processamento da Planilha: 



    col1, col2 = st.columns(2)

    col1.title('Relatório Solos:')
    # Upload Arquivo csv 
    uploaded_files = col1.file_uploader("Upload Planilha de Solos 📥")

    col2.title('Relatório Drone:')
    # Upload Arquivo csv 
    uploaded_files_drone = col2.file_uploader("Upload Planilha de Drone 📥")

    #Solos
    tabela = pd.read_excel(uploaded_files)

    tabela_original = tabela

    #Excluir Dados Duplicados
    tabela.drop_duplicates(['Mapeamento','Fazenda','Talhão'], inplace = True)
    # Filtrar Colunas 
    tabela = tabela[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área (ha)','Data','Link']]
    n_mapas = tabela['Mapeamento'].count()

    tabela.head()

    tabela.drop_duplicates(['Talhão'], inplace = True)
    soma_area = tabela['Área (ha)'].sum()

    tabela_original.drop_duplicates(['Mapeamento','Fazenda','Talhão'], inplace = True)
    tabela_original = tabela_original[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área (ha)','Data','Link']]


    tabela_original.loc['Total'] = ' '
    tabela_original['Área (ha)']['Total'] = soma_area
    tabela_original['Mapeamento']['Total'] = n_mapas
    tabela_original['Cliente']['Total'] = 'Total'









