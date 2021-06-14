import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw




st.write("""
 # WELCOME TO WEBSITE
 This is awesome!
 """)

#### Doing multiple columns ###########################
col1, col2 = st.beta_columns(2)
col1.success('PV1')
############# reading PV sheet ############################ 
with col1:
    
    Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
    Epv.dropna(subset=['Model'], inplace=True) 
    Epv = Epv[Epv['Model'] != 'Name'] #### remove name as title
    ######### Inputs Form #############


    st.subheader("Facility Name")
    location = st.selectbox("", options=["Select location", "Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
    st.subheader("Envelope")
    Envelope_selection = st.selectbox("", options= ["Select Envelope", "North","South","East","West"])
    direction = st.selectbox("", options=["Select Direction", "North", "South", "East", "West"])
    Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
    st.subheader("Azimuth Selection")
    Azimuth = st.selectbox("", options = ["Select Azimuth",0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
    Slope = st.number_input("Enter a Slope", key='slope')
##### Getting PV model#############
    st.subheader("""PV Specification Models""")
    model = st.selectbox("Select Model", Epv['Model'].values)
    # st.write(model)
    st.subheader("Scale")
    Amodule = st.number_input("Number of modules(EA)", key='Amodule')
    
########### reading Inverter sheet############
    inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
    inverter.dropna(subset=['Name'], inplace=True)
    inverter = inverter[inverter['Name'] != 'Units']
########## Getting Interter model###################
    st.subheader("""Inverter Models""")
    model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
    # st.text (Name)
    Rsurface = st.number_input("Non-vertical surface solar attenuation rate", key='Rsurface')
    Total_equipment_cost = st.number_input("Total equipment cost (KRW)", key='Total equipment cost')
   
    #####st.write(pv)
    ################ Getting inverter efficiency##########################################################
   
    #-----------------------------------------------------------------------------------------------------------------------

    #++++++++++++++++++++++++++[∑(Srad,month X Rcorr)]++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    with col1:
        st.subheader("Are these information you provide all correct? ")
        

        st.checkbox("YES")
        if st.button("SUBMIT PV1"):
            st.write()


#########################################################Output viewer #############################################


#### xlwings code##########################

bk = xw.Book("Photovoltaic module_V10.xlsx")

#st.write(bk.sheets[0].range('A1').value)

#bk.sheets[0].range('A1').value = 'Input2'


input = bk.sheets['Input']
input.range('C3:C13').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule],[model_units],[Rsurface],[Total_equipment_cost]]

input.range("A17:M21").options(pd.DataFrame).value
#with st.form(key='my_form'):
    # text_input = st.text_input(label='Enter Location')
	#location = st.selectbox("", options=["Select location", "Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
    # Envelope_selection = st.selectbox("", options= ["Select Envelope", "North","South","East","West"])
    
	#submit_button = st.form_submit_button(label='Submit')


form = st.form(key='my_form')
form.subheader("Facility Name")
form.selectbox("", options=["Select location", "Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"],key="location")
form.selectbox("", options= ["Select Envelope", "North","South","East","West"],key="envelope")
form.subheader("Envelope")
form.selectbox("", options= ["Select Envelope", "North","South","East","West"])
form.selectbox("", options=["Select Direction", "North", "South", "East", "West"])
form.number_input("Enter Area", min_value= 0, value= 0, step=0)
form.subheader("Azimuth Selection")
form.selectbox("", options = ["Select Azimuth",0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
form.number_input("Enter a Slope", key='slope')
form.subheader("""PV Specification Models""")
form.selectbox("Select Model", Epv['Model'].values,key="models")
form.subheader("Scale")
form.number_input("Number of modules(EA)", key='Amodule')
form.subheader("""Inverter Models""")
form.selectbox("Select Inverter Model", inverter['Name'].values,key="inverters")
form.number_input("Non-vertical surface solar attenuation rate", key='Rsurface')
form.number_input("Total equipment cost (KRW)", key='Total equipment cost')
submit_button = form.form_submit_button(label='Submit')
   
