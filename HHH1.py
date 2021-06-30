import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw

import matplotlib.pyplot as plt
import seaborn as sns


st.set_page_config(page_title=None, page_icon=None, layout='wide', initial_sidebar_state='auto')

sb=0
sb1=0
sb2=0
sb3=0

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
background-image: url("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS4JRrtBbUUcZ_A-LSRwZRlFerrHjFVvxE0U-47Kset1deiKz1OWZnhV7Y5jy0xEU86mFE&usqp=CAU");
background-size: cover;
}
</style>
'''

st.markdown(page_bg_img, unsafe_allow_html=True)
############## READ BOOK ######################
app = xw.App(visible=False,add_book=False)
bk = xw.Book("Photovoltaic module_V10.xlsx")
input = bk.sheets['Input']

@st.cache()
def load_data():
        time.sleep(2) 
        Epv = pd.read_excel("Photovoltaic module_V10.xlsx")
        return Epv

pv = "Photovoltaic Energy Generation"
st.markdown(
f'<body style="font-size:25px;border: 5px; background-color:skyblue; font-familly: Arial; padding: 10px; "><center>{pv}</center></body>'
, unsafe_allow_html=True)


#### Making Multiple columns ###########################
col1,col2,col3,col4= st.beta_columns(4)

with col1:
        pv = "PV1"
        st.markdown(
        f'<div style="font-size:16px;border: 2px; background-color:skyblue; font-familly: Arial; padding: 12px; "><center>{pv}</center></div>'
        , unsafe_allow_html=True)

     
############### Inputs Form for PV1 ########################        
        with st.form(key='my_form'):
                #st.text("Facility Name")
                #st.text("Enter a Location")
                location = st.selectbox("Select Location", options=["Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                #st.subheader("Envelope")
                Envelope_selection = st.selectbox("Select Envelope", options= ["North","South","East","West"])
                direction = st.selectbox("Select Direction", options=["North", "South", "East", "West"])
                Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
                #st.subheader("Azimuth Selection")
                Azimuth = st.selectbox("Select Azimuth", options = [0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                Slope = st.number_input("Enter a Slope", key='slope')
                
                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                Epv.dropna(subset=['Model'], inplace=True) 
                Epv = Epv[Epv['Model'] != 'Name']
                #st.subheader("""PV Specification Models""")
                model = st.selectbox("Select PV Model", Epv['Model'].values)
                #st.subheader("Scale")
                Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                inverter.dropna(subset=['Name'], inplace=True)
                inverter = inverter[inverter['Name'] != 'Units']
                #st.subheader("""Inverter Models""")
                model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
                Rsurface = st.number_input("Enter Non-vertical Surface Solar Attenuation Rate", key='Rsurface')
                Total_equipment_cost = st.number_input("Enter Total Equipment Cost (KRW)", key='Total equipment cost')
                Equipment_cost = st.number_input("Enter Equipment Cost(Won)", key='Equipment_cost')
                Analysis_period = st.number_input("Enter Analysis period(Won)", key='Analysis_period')
                submit_button = st.form_submit_button(label='Submit')
                

########################    writting inputs into pv1   ######################### 
                if submit_button:
                        input = bk.sheets['Input']
                        input.range('C3:C10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                        input.range('C16:C18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                        input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]
                        listOfGlobals = globals()
                        listOfGlobals['sb'] = 1
op = ['PV2', 'PV3','PV4']
#option = st.selectbox("",op)

options = st.multiselect('Select other PV', op)  

if len(options)==0:
        with col2:
                st.image("https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcS4JRrtBbUUcZ_A-LSRwZRlFerrHjFVvxE0U-47Kset1deiKz1OWZnhV7Y5jy0xEU86mFE&usqp=CAU", width=1000)

   
if options!=None:
        #st.sidebar(col1)
        if "PV2" in options:
                option = "PV2"
                
                ####################    Other PVs Menu Forms    ##################
                with col2:
                        st.markdown(
                f'<div style="font-size:16px;border: 2px; background-color:gray; font-familly: Arial; padding: 12px; "><center><b>{option}</b></center></div>'
                , unsafe_allow_html=True) 
                                        
                        with st.form(key=option):
                                #st.text("Facility Name")
                                #st.subheader("Enter a Location")
                                location = st.selectbox("Select Location", options=["Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                                #st.subheader("Envelope")
                                Envelope_selection = st.selectbox("Select Envelope", options= ["North","South","East","West"])
                                direction = st.selectbox("Select Direction",  options=["North", "South", "East", "West"])
                                Area = st.number_input("Enter Area", input.range('C6').value)
                                #st.subheader("Azimuth Selection")
                                Azimuth = st.selectbox("Select Azimuth", options = [0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                                Slope = st.number_input("Enter a Slope", key='slope')
                                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                                Epv.dropna(subset=['Model'], inplace=True) 
                                Epv = Epv[Epv['Model'] != 'Name']
                                
                                #st.subheader("""PV Specification Models""")
                                model = st.selectbox("Select PV Model", Epv['Model'].values)
                                #st.subheader("Scale")
                                Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                                inverter.dropna(subset=['Name'], inplace=True)
                                inverter = inverter[inverter['Name'] != 'Units']
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
                                        listOfGlobals = globals()
                                        listOfGlobals['sb1'] = 1
                               
                                
        if "PV3" in options:
                option = "PV3"
                ####################    Other PVs Menu Forms    ##################
                with col3: 
                        st.markdown(
                f'<div style="font-size:16px;border: 2px; background-color:gray; font-familly: Arial; padding: 12px; "><center><b>{option}</b></center></div>'
                , unsafe_allow_html=True)                 
                        with st.form(key=option):
                                #st.text("Facility Name")
                                #st.subheader("Enter a Location")
                                location = st.selectbox("Select Location", options=["Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                                #st.subheader("Envelope")
                                Envelope_selection = st.selectbox("Select Envelope", options= ["North","South","East","West"])
                                direction = st.selectbox("Select Direction",  options=["North", "South", "East", "West"])
                                Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
                                #st.subheader("Azimuth Selection")
                                Azimuth = st.selectbox("Select Azimuth", options = [0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                                Slope = st.number_input("Enter a Slope", key='slope')
                                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                                Epv.dropna(subset=['Model'], inplace=True) 
                                Epv = Epv[Epv['Model'] != 'Name']
                                
                                #st.subheader("""PV Specification Models""")
                                model = st.selectbox("Select PV Model", Epv['Model'].values)
                                #st.subheader("Scale")
                                Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                                inverter.dropna(subset=['Name'], inplace=True)
                                inverter = inverter[inverter['Name'] != 'Units']
                                #st.subheader("""Inverter Models""")
                                model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
                                Rsurface = st.number_input("Enter Non-vertical Surface Solar Attenuation Rate", key='Rsurface')
                                Total_equipment_cost = st.number_input("Enter Total Equipment Cost (KRW)", key='Total equipment cost')
                                Equipment_cost = st.number_input("Enter Equipment Cost(Won)", key='Equipment_cost')
                                Analysis_period = st.number_input("Enter Analysis period(Won)", key='Analysis_period')
                                submit_button2 = st.form_submit_button(label='Compare PV1 and '+option)
                                        
                ########################  Writing into Other PVs   ######################
                                
                                if submit_button2 and option=="PV3":
                                        input = bk.sheets['Input']
                                        input.range('E3:E10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                                        input.range('E16:E18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                                        input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]
                                        listOfGlobals = globals()
                                        listOfGlobals['sb2'] = 1
                                                                




        if "PV4" in options:
                option = "PV4"
                ####################    Other PVs Menu Forms    ##################
                with col4: 
                        st.markdown(
                f'<div style="font-size:16px;border: 2px; background-color:gray; font-familly: Arial; padding: 12px; "><center><b>{option}</b></center></div>'
                , unsafe_allow_html=True)               
                        with st.form(key=option):
                                #st.text("Facility Name")
                                #st.subheader("Enter a Location")
                                location = st.selectbox("Select Location", options=["Seoul","Chuncheon","Kangrueng","WonJu","DaeJeon","ChungJu","SuSan","DaeGu","PoHang","YoungJu","Busan","JinJu","JeonJu","KwangJu","MokPo","JeJu"])
                                #st.subheader("Envelope")
                                Envelope_selection = st.selectbox("Select Envelope", options= ["North","South","East","West"])
                                direction = st.selectbox("Select Direction",  options=["North", "South", "East", "West"])
                                Area = st.number_input("Enter Area", min_value= 0, value= 0, step=0)
                                #st.subheader("Azimuth Selection")
                                Azimuth = st.selectbox("Select Azimuth", options = [0,10,20,30,40,50,60,70,80,90,100,110,120,130,140,150,160,170,180,190,200,210,220,230,240,250,260,270,280,290,300,310,320,330,340,350,360])
                                Slope = st.number_input("Enter a Slope", key='slope')
                                Epv = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="PV")
                                Epv.dropna(subset=['Model'], inplace=True) 
                                Epv = Epv[Epv['Model'] != 'Name']
                                
                                #st.subheader("""PV Specification Models""")
                                model = st.selectbox("Select PV Model", Epv['Model'].values)
                                #st.subheader("Scale")
                                Amodule = st.number_input("Enter Number of Modules(EA)", key='Amodule')
                                inverter = pd.read_excel("Photovoltaic module_V10.xlsx", sheet_name="Inverter")
                                inverter.dropna(subset=['Name'], inplace=True)
                                inverter = inverter[inverter['Name'] != 'Units']
                                #st.subheader("""Inverter Models""")
                                model_units = st.selectbox("Select Inverter Model", inverter['Name'].values)
                                Rsurface = st.number_input("Enter Non-vertical Surface Solar Attenuation Rate", key='Rsurface')
                                Total_equipment_cost = st.number_input("Enter Total Equipment Cost (KRW)", key='Total equipment cost')
                                Equipment_cost = st.number_input("Enter Equipment Cost(Won)", key='Equipment_cost')
                                Analysis_period = st.number_input("Enter Analysis period(Won)", key='Analysis_period')
                                submit_button3 = st.form_submit_button(label='Compare PV1 and '+option)
                                        
                ########################  Writing into Other PVs   ######################

                                if submit_button3 and option=="PV4":
                                        input = bk.sheets['Input']
                                        input.range('F3:F10').value = [[location],[Envelope_selection],[direction],[Area],[Azimuth],[Slope],[model],[Amodule]]
                                        input.range('F16:F18').value = [[model_units],[Rsurface],[Total_equipment_cost]]
                                        input.range('L4:L6').value = [[Equipment_cost],[Analysis_period]]
                                        listOfGlobals = globals()
                                        listOfGlobals['sb3'] = 1




### GET ALL VALUES AFTER 
df = input.range("A27:M31").options(pd.DataFrame).value
#df.reset_index(inplace=True)

#global pv1,pv2,pv3,pv4
pv1 = df[0:1][:]
pv2 = df[1:2][:]
pv3 = df[2:3][:]
pv4 = df[3:4][:]                               

#costs
profit = input.range("A37:E40").options(pd.DataFrame).value
profit.reset_index(inplace=True)
cost1 = profit[['index','cost1']]
cost1.set_index('index', inplace=True)
cost2 = profit[['index','cost2']]
cost2.set_index('index', inplace=True)
cost3 = profit[['index','cost3']]
cost3.set_index('index', inplace=True)
cost4 = profit[['index','cost4']]
cost4.set_index('index', inplace=True)


######################### Energy Generation OUTPUT     ############################
st.title("Simulation Results")        
st.subheader("Energy Generation (kWh)")


Output = input.range("A27:M31").options(pd.DataFrame).value

if sb == 1:
        st.table(pv1.assign(hack='').set_index('hack'))
        costs = cost1
        pvss = pv1
elif "PV2" in options and "PV3" not in options and "PV4" not in options and sb1==1:
        st.table(pv1.append(pv2, ignore_index=True).assign(hack='').set_index('hack'))
        #merge costs
        costs = cost1.merge(cost2,left_index=True, right_index=True)
        pvss = pd.concat([pv1, pv2])

elif "PV2" in options and "PV3" in options  and "PV4" not in options and (sb2==1 or sb1==1):
        st.table(pv1.append([pv2,pv3], ignore_index=True).assign(hack='').set_index('hack'))
        costs1 = cost1.merge(cost2,left_index=True, right_index=True)
        costs = costs1.merge(cost3,left_index=True, right_index=True)

elif "PV2" in options and "PV3" in options and  "PV4" in options and (sb3==1 or sb2==1 or sb3==1):
        st.table(pv1.append([pv2,pv3,pv4], ignore_index=True).assign(hack='').set_index('hack'))
        #combine all costs for selectected PVS
        costs1 = cost1.merge(cost2,left_index=True, right_index=True)
        costs2 = costs1.merge(cost3,left_index=True, right_index=True)
        costs = costs2.merge(cost4,left_index=True, right_index=True)

elif "PV2" not in options and "PV3" in options and  "PV4" in options and (sb3==1 or sb2==1):
        st.table(pv1.append([pv3,pv4], ignore_index=True).assign(hack='').set_index('hack'))
        #combine all costs for selectected PVS
        costs1 = cost1.merge(cost3,left_index=True, right_index=True)
        costs = costs1.merge(cost4,left_index=True, right_index=True)
        pvss = pd.concat([pv1, pv3])
elif "PV2" not in options and "PV3" not in options and  "PV4" in options and sb3==1:
        st.table(pv1.append(pv4, ignore_index=True).assign(hack='').set_index('hack'))
        #combine all costs for selectected PVS
        costs = cost1.merge(cost4,left_index=True, right_index=True)

elif "PV2" not in options and "PV3" in options and  "PV4" not in options and sb2==1:
        st.table(pv1.append(pv3, ignore_index=True).assign(hack='').set_index('hack'))
        #combine all costs for selectected PVS
        costs = cost1.merge(cost3,left_index=True, right_index=True)

elif "PV2" in options and "PV3" not in options and  "PV4" in options and (sb3==1 or sb1==1):
        st.table(pv1.append([pv2,pv4], ignore_index=True).assign(hack='').set_index('hack'))
        #combine all costs for selectected PVS
        costs1 = cost1.merge(cost2,left_index=True, right_index=True)
        costs = costs1.merge(cost4,left_index=True, right_index=True)
else:
        st.table(input.range("A27:M27").options(pd.DataFrame).value)

prof,grph= st.beta_columns((2,3))
######################### Net Profit for 30  yearsOutput ##########################
with prof:
        st.subheader("Net Profit for 30 years")
        if sb == 1:
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

        elif "PV2" in options and "PV3" not in options and "PV4" not in options and sb1==1:
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))
                
        elif "PV2" in options and "PV3" in options  and "PV4" not in options and (sb2==1 or sb1==1):
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))
        
        elif "PV2" in options and "PV3" in options and  "PV4" in options and (sb3==1 or sb2==1 or sb3==1):
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

        elif "PV2" not in options and "PV3" in options and  "PV4" in options and (sb3==1 or sb2==1):
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

        elif "PV2" not in options and "PV3" not in options and  "PV4" in options and sb3==1:
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

        elif "PV2" not in options and "PV3" in options and  "PV4" not in options and sb2==1:
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

        elif "PV2" in options and "PV3" not in options and  "PV4" in options and (sb3==1 or sb1==1):
                #profit = input.range("A37:E40").options(pd.DataFrame).value
                profit=costs
                profit.reset_index(inplace=True)
                st.table(profit.assign(hack='').set_index('hack'))

########################## Energy Generation Graphing  ############################                                
with grph:
        if sb == 1:
                st.set_option('deprecation.showPyplotGlobalUse', False)

                df_revised = pvss
                df_revised.reset_index(inplace=True)
                df_ = df_revised.T
                df_.reset_index(inplace=True)
                cols = np.array(df_[df_['index']=="Facility name"].values)
                data =  np.array(df_[df_['index']!="Facility name"].values)
                p = {'Months':data[0:,0], 'PV1':data[0:,1]}

                pvs = pd.DataFrame(data=p)
                pvs.set_index('Months', inplace=True)
                pvs.plot.bar(rot=10, title="Energy Generation Graph",figsize=(15, 3))
                st.pyplot()

        elif "PV2" in options and "PV3" not in options and "PV4" not in options and sb1==1:
                st.set_option('deprecation.showPyplotGlobalUse', False)

                df_revised = pvss
                df_revised.reset_index(inplace=True)
                df_ = df_revised.T
                df_.reset_index(inplace=True)
                cols = np.array(df_[df_['index']=="Facility name"].values)
                data =  np.array(df_[df_['index']!="Facility name"].values)
                p = {'Months':data[0:,0], 'PV1':data[0:,1],'PV2':data[0:,2]}

                pvs = pd.DataFrame(data=p)
                pvs.set_index('Months', inplace=True)
                pvs.plot.bar(rot=10, title="Energy Generation Graph",figsize=(15, 3))
                st.pyplot()
                