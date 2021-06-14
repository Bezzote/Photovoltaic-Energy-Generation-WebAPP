import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw
import base64
import operator
# import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
############### Hiding sreamlit menu and footer ############
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>

"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

############# Image banner ######################
page_bg_img = '''
<style>
body {
background-image: url("data:image/jpeg;base64,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");
  background-repeat: repeat;
  background-color: #cccccc;
}
</style>
'''

st.markdown(page_bg_img, unsafe_allow_html=True)
############## READ BOOK ######################

bk = xw.Book("Photovoltaic module_V10.xlsx")


input = bk.sheets['Input']
pv = "Photovoltaic Energy Generation"
st.markdown(
f'<body style="font-size:25px;border: 2px; background-color:skyblue; font-familly: Arial; padding: 10px; "><center>{pv}</center></body>'
, unsafe_allow_html=True)

#### Making Multiple columns ###########################
col1, col2 = st.beta_columns(2)

with col1:
        
        pv = "PV1"
        st.markdown(
        f'<div style="font-size:25px;border: 1px; background-color:skyblue; font-familly: Arial; padding: 12px; "><center>{pv}</center></div>'
        , unsafe_allow_html=True)

        
############### Inputs Form for PV1 ########################        
        with st.form(key='my_form'):
                st.text("Facility Name")
                #st.text("Enter a Location")
                location = st.selectbox("", options=["""Select Location""", "Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                #st.subheader("Envelope")
                Envelope_selection = st.selectbox("", options= ["Select Envelope", "North","South","East","West"])
                direction = st.selectbox("", options=["Select Direction", "North", "South", "East", "West"])
                Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
                #st.subheader("Azimuth Selection")
                Azimuth = st.selectbox("", options = ["Select Azimuth",0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                Slope = st.number_input("Enter a Slope", key='slope')
                
                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                Epv.dropna(subset=['Model'], inplace=True) 
                Epv = Epv[Epv['Model'] != 'Name']
               
                @st.cache()
                def load_data():
                        time.sleep(2) 
                        Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                        return Epv
                #st.subheader("""PV Specification Models""")
                model = st.selectbox("Select PV Model", Epv['Model'].values)
                #st.subheader("Scale")
                Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                inverter.dropna(subset=['Name'], inplace=True)
                inverter = inverter[inverter['Name'] != 'Units']
                
                @st.cache()
                def load_data():
                        time.sleep(2) 
                        inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                        return inverter
                        
                
                #st.subheader("""Inverter Models""")
                model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
                Rsurface = st.number_input("Enter Non-vertical Surface Solar Attenuation Rate", key='Rsurface')
                Total_equipment_cost = st.number_input("Enter Total Equipment Cost (KRW)", key='Total equipment cost')
                Equipment_cost = st.number_input("Enter Equipment Cost(Won)", key='Equipment_cost')
                Analysis_period = st.number_input("Enter Analysis period(Won)", key='Analysis_period')
                submit_button = st.form_submit_button(label='Submit')
####################    Other PVs Menu Forms    ##################
with col2:
        
        op = ['Select Other PV', 'PV2', 'PV3','PV4']
        option = st.selectbox("",op)      
        @st.cache()
        def load_data():
                time.sleep(2) 
                #inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                option = st.selectbox("",op)  
                return option
                        
        
        if option!=op[0]:    
                with st.form(key=option):
                        st.text("Facility Name")
                        #st.subheader("Enter a Location")
                        location = st.selectbox("", options=["Select Location", "Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                        #st.subheader("Envelope")
                        Envelope_selection = st.selectbox("", options= ["Select Envelope", "North","South","East","West"])
                        direction = st.selectbox("", options=["Select Direction", "North", "South", "East", "West"])
                        Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
                        #st.subheader("Azimuth Selection")
                        Azimuth = st.selectbox("", options = ["Select Azimuth",0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                        Slope = st.number_input("Enter a Slope", key='slope')
                        Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                        Epv.dropna(subset=['Model'], inplace=True) 
                        Epv = Epv[Epv['Model'] != 'Name']
                        
                        @st.cache()
                        def load_data():
                                time.sleep(2) 
                                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                                return Epv
                        #st.subheader("""PV Specification Models""")
                        model = st.selectbox("Select PV Model", Epv['Model'].values)
                        #st.subheader("Scale")
                        Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                        inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                        inverter.dropna(subset=['Name'], inplace=True)
                        inverter = inverter[inverter['Name'] != 'Units']
                        
                        @st.cache()
                        def load_data():
                                time.sleep(2) 
                                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                                return inverter
                        #st.subheader("""Inverter Models""")
                        model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
                        Rsurface = st.number_input("Enter Non-vertical Surface Solar Attenuation Rate", key='Rsurface')
                        Total_equipment_cost = st.number_input("Enter Total Equipment Cost (KRW)", key='Total equipment cost')
                        Equipment_cost = st.number_input("Enter Equipment Cost(Won)", key='Equipment_cost')
                        Analysis_period = st.number_input("Enter Analysis period(Won)", key='Analysis_period')
                        submit_button1 = st.form_submit_button(label='Compare PV1 and '+option)
########################  Writing into Other PVs   ######################
                        if submit_button1 and option=="PV2":
                                input = bk.sheets['Input']
                                input.range('D3:D10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                                input.range('D16:D18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                                input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]

                        if submit_button1 and option=="PV3":
                                input = bk.sheets['Input']
                                input.range('E3:E10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                                input.range('E16:E18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                                input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]
                        if submit_button1 and option=="PV4":
                                input = bk.sheets['Input']
                                input.range('F3:F10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                                input.range('F16:F18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                                input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]

########################    writting inputs into pv1   ######################### 
if submit_button:

        input = bk.sheets['Input']
        input.range('C3:C10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
        input.range('C16:C18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
        input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]
                
######################### Energy Generation OUTPUT     ############################
st.title("Output")        
st.subheader("Energy Generation (kWh)")
input.range("A27:M31").options(pd.DataFrame).value


#a, graph = st.beta_columns(2)
######################### Net Profit for 30 years Output ##########################
#with a:
st.subheader("Net Profit for 30 years")
input.range("A37:E40").options(pd.DataFrame).value
########################## Energy Generation Graphing  ############################                                
#with graph:
        
st.set_option('deprecation.showPyplotGlobalUse', False)

df = input.range("A27:M31").options(pd.DataFrame).value

pv1 = df[0:1][:]
pv2 = df[1:2][:]
pv3 = df[2:3][:]
pv4 = df[3:4][:]


df_revised = pd.concat([pv1, pv2,pv3,pv4])
df_revised.reset_index(inplace=True)
df_ = df_revised.T
df_.reset_index(inplace=True)

cols = np.array(df_[df_['index']=="Facility name"].values)
data =  np.array(df_[df_['index']!="Facility name"].values)
p = {'Months':data[0:,0], 'PV1':data[0:,1],'PV2':data[0:,2],'PV3':data[0:,3],'PV4':data[0:,4]}
        
pvs = pd.DataFrame(data=p)
pvs.set_index('Months', inplace=True)
pvs.plot.bar(rot=10, title="Energy Generation Graph")     
st.pyplot()
        

