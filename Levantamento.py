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
tipo_analise = st.sidebar.selectbox("📊 Tipo de Levantamento:", ['Solos','Drone','Colheita','Falhas'])

if tipo_analise == 'Solos':

    page_bg_img = """
    <style>
    [data-testid="stAppViewContainer"] > .main {
    background-image: url("https://blog.sensix.ag/wp-content/uploads/2021/03/Steve_Boreham_Soil_Sampling.jpg");
    background-size: cover;
    }

    [data-testid="stHeader"] {
    background: rgba(0,0,0,0);
    }

    </style>
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)


    # ##  Processamento da Planilha: 

   
    st.title('Relatório Solos 🌱:')
    # Upload Arquivo csv 
    uploaded_files = st.file_uploader("Upload Planilha de Solos 📥")

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

   



   # DataFrame para Planilha Excel em xlsx

    def to_excel(tabela_original):
        output = BytesIO()
        writer = pd.ExcelWriter(output,engine='xlsxwriter')
        tabela_original.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df = to_excel(tabela_original)

    st.download_button(label=' ⬇️ Download Levantamento Solos', data=df,file_name= 'Planilha_Solos.xlsx')
    
if tipo_analise == 'Drone':

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

    st.title('Relatório Drone 🚁:')
    # Upload Arquivo csv 
    uploaded_files = st.file_uploader("Upload Planilha de Drone 📥")
         

    tabela_drone = pd.read_excel(uploaded_files)


    #Filtrar Dados de Drone
    filtro_drone = tabela_drone['Origem'] == 'Drone'
    tabela_drone = tabela_drone[filtro_drone]


    #Excluir Dados Duplicados
    tabela_drone.drop_duplicates(['Mapeamento','Fazenda','Talhão'], inplace = True)
    # Filtrar Colunas 
    tabela_drone = tabela_drone[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área (ha)','Data','Link']]
    n_mapas_drone = tabela_drone['Mapeamento'].count()
    soma_area_drone = tabela_drone['Área (ha)'].sum()
    tabela_drone.head()

    tabela_drone.loc['Total'] = ' '
    tabela_drone['Área (ha)']['Total'] = soma_area_drone
    tabela_drone['Mapeamento']['Total'] = n_mapas_drone
    tabela_drone['Cliente']['Total'] = 'Total'

    tabela_drone.head()



    # DataFrame para Planilha Excel em xlsx

    def to_excel(tabela_drone):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        tabela_drone.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df_drone = to_excel(tabela_drone)

    st.download_button(label=' ⬇️ Download Levantamento Drone', data=df_drone,file_name= 'Planilha_Drone.xlsx')
    
if tipo_analise == 'Colheita':
    
    page_bg_img = """
    <style>
    [data-testid="stAppViewContainer"] > .main {
    background-image: url("https://assets.revistacultivar.com.br/eaf49-Perdas-colheita-de-soja_04.jpg");
    background-size: cover;
    }

    [data-testid="stHeader"] {
    background: rgba(0,0,0,0);
    }

    </style>
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)
    
    
    # ##  Processamento da Planilha: 

   
    st.title('Relatório Colheita 🚜:')
    # Upload Arquivo csv 
    uploaded_files = st.file_uploader("Upload Planilha de Colheita 📥")

    tabela_colheita = pd.read_excel(uploaded_files)
    tabela_original_colheita = tabela_colheita
    #Excluir Dados Duplicados
    tabela_colheita.drop_duplicates(['Mapeamento','Fazenda','Talhão'], inplace = True)
    tabela_colheita.head()

    # Filtrar Colunas 
    tabela_colheita = tabela_colheita[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área (ha)','Data','Cultura','Link']]
    n_mapas_colheita = tabela_colheita['Mapeamento'].count()
    tabela_colheita.drop_duplicates(['Talhão'], inplace = True)
    soma_area_colheita = tabela_colheita['Área (ha)'].sum()

    tabela_colheita.head()

    tabela_colheita.drop_duplicates(['Talhão'], inplace = True)
    
    tabela_original_colheita = tabela_original_colheita[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área (ha)','Data','Cultura','Link']]
    tabela_original_colheita.loc['Total'] = ' '
    tabela_original_colheita['Área (ha)']['Total'] = soma_area_colheita
    tabela_original_colheita['Mapeamento']['Total'] = n_mapas_colheita
    tabela_original_colheita['Cliente']['Total'] = 'Total'

    # DataFrame para Planilha Excel em xlsx

    def to_excel(tabela_original_colheita):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        tabela_original_colheita.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df_colheita = to_excel(tabela_original_colheita)

    st.download_button(label=' ⬇️ Download Levantamento Colheita', data=df_colheita,file_name= 'Planilha_Colheita.xlsx')
    
if tipo_analise == 'Falhas':

    page_bg_img = """
    <style>
    [data-testid="stAppViewContainer"] > .main {
    background-image: url("https://blog.sensix.ag/wp-content/uploads/2021/03/Falhas-de-Plantio-Cana-de-acucar.jpg");
    background-size: cover;
    }
    [data-testid="stHeader"] {
    background: rgba(0,0,0,0);
    }
    </style>
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)


    # ##  Processamento da Planilha: 

    st.title('Relatório Falhas 🌿:')
    # Upload Arquivo csv 
    uploaded_files = st.file_uploader("Upload Planilha de Drone 📥")
         
    tabela_falhas = pd.read_excel(uploaded_files)
    #Excluir Dados Duplicados
    tabela_falhas.drop_duplicates(['Mapeamento','Fazenda','Talhão'], inplace = True)
    tabela_falhas.head()

    # Filtrar Colunas 
    tabela_falhas = tabela_falhas[['Cliente','E-mail','Fazenda','Talhão','Mapeamento','Área','Data','Link']]
    n_mapas_falhas = tabela_falhas['Mapeamento'].count()
    soma_area_falhas = tabela_falhas['Área'].sum()

    tabela_falhas.loc['Total'] = ' '
    tabela_falhas['Área']['Total'] = soma_area_falhas
    tabela_falhas['Mapeamento']['Total'] = n_mapas_falhas
    tabela_falhas['Cliente']['Total'] = 'Total'

    # DataFrame para Planilha Excel em xlsx

    def to_excel(tabela_falhas):
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        tabela_falhas.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format1 = workbook.add_format({'num_format': '0.00'}) 
        worksheet.set_column('A:A', None, format1)  
        writer.save()
        processed_data = output.getvalue()
        return processed_data

    df_falhas = to_excel(tabela_falhas)

    st.download_button(label=' ⬇️ Download Levantamento Falhas', data=df_falhas,file_name= 'Planilha_Falhas.xlsx')








