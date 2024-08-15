######## Library ##########################################
import streamlit as st
import pandas as pd
from PIL import Image
import numpy as np
import os
import glob
from datetime import datetime, timedelta
import datetime
import calendar
import plotly.graph_objects as go
import requests
import sys
########################################################
Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=680)
#########################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
#################### MENU Select Month ###################################
#########################################################
def formatted_display2(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
#################### MENU Select Month ###################################
# Month = st.sidebar.selectbox( 'Month',['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'])
#####################
Yinput = st.sidebar.selectbox('Input-Year', ['2024'])
# Dictionary mapping month numbers to month names
month_name_map = {
    '01': 'JANUARY',
    '02': 'FEBRUARY',
    '03': 'MARCH',
    '04': 'APRIL',
    '05': 'MAY',
    '06': 'JUNE',
    '07': 'JULY',
    '08': 'AUGUST',
    '09': 'SEPTEMBER',
    '10': 'OCTOBER',
    '11': 'NOVEMBER',
    '12': 'DECEMBER'
}

Minput = st.sidebar.selectbox('Input-Month', [Yinput+'-01', Yinput+'-02', Yinput+'-03', Yinput+'-04', Yinput+'-05', Yinput+'-06', Yinput+'-07', Yinput+'-08', Yinput+'-09', Yinput+'-10', Yinput+'-11', Yinput+'-12'])

year, month_num = Minput.split('-')
Month = month_name_map[month_num]

###############################
def generate_weeks(year):
    start_date = datetime(year, 1, 1)
    start_date += timedelta(weeks=1)
    end_date = datetime(year, 12, 31)
    current_date = start_date
    weeks = []
    week_number = 1  # Start week number from 40

    while current_date <= end_date:
        week_start = current_date.strftime('%Y-%m-%d')
        week_end = (current_date + timedelta(days=6)).strftime('%Y-%m-%d')
        weeks.append((week_number, f"{week_start} - {week_end}"))
        current_date += timedelta(days=7)
        week_number += 1

    return weeks

def get_week_range_for_month(year, month):
    first_day = datetime.date(year, month, 1)
    last_day = datetime.date(year, month, calendar.monthrange(year, month)[1])

    start_week = first_day.isocalendar()[1]
    end_week = last_day.isocalendar()[1]

    if first_day.isocalendar()[0] < year:
        start_week = 1
    if last_day.isocalendar()[0] > year:
        end_week = 53

    return start_week, end_week
##################### New Month +1 #####################################
Minput = Minput
year, month = map(int, Minput.split('-'))
start_week, end_week = get_week_range_for_month(year, month)
# start_week, end_week
#########
Process = st.sidebar.selectbox('Process',['Die_casting','Finishing','Finishing_SUB','Sand_Blasting','Machine','QC','Other'] )
# Cost_Type = st.sidebar.selectbox('Cost Type',['Variable Cost','Fixed Cost'] )
#######################
st.sidebar.button('Please press here!. When data not update')
st.cache_data.clear()
########################## Read File #####################
DCDATA=pd.read_excel('Shot Wight Update 15-05-2024.xlsx')
# DCDATA
########################## Read File #####################
FNSBMC_DATA=pd.read_excel('Cycle Time FN 25-10-23.xlsx',header=1)
FNSBMC_DATA=FNSBMC_DATA[['Part no.','Shot Blasting /Pcs','TT-FN (Sec)','MC-CT(SEC)']]
FNSBMC_DATA.set_index('Part no.',inplace=True)
# FNSBMC_DATA
#########
########################## Production 2024 ############################
@st.cache_data 
def load_data_File_A(start_week, end_week):
    file = "Production-"+Yinput+".xlsx"
    all_sheets = pd.read_excel(file, header=7, sheet_name=None)  # Load all sheets
    # Convert sheet names to strings
    all_sheets = {str(sheet_name): df for sheet_name, df in all_sheets.items()}
    
    # Filter sheets based on start_week and end_week
    selected_sheets = {}
    for week_num in range(start_week, end_week + 1):
        sheet_name = str(week_num)
        if sheet_name in all_sheets:
            selected_sheets[sheet_name] = all_sheets[sheet_name]
    
    return selected_sheets

start_week = int(start_week)
end_week = int(end_week)
data2024 = load_data_File_A(start_week, end_week)
data2024 = pd.concat(data2024.values(), ignore_index=True)
data2024['Part no.']=data2024['Part no.'].astype(str)
values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
'1731A','KOSHIN','043061102-']
# Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
mask = ~data2024['Part no.'].isin(values_to_exclude)
data2024 = data2024[mask]
data2024=pd.merge(data2024,DCDATA,on='Part no.',how='left')
########################## Sales 2024 ############################
@st.cache_data 
def load_data_Sales(start_week, end_week):
    file = "Production-"+Yinput+".xlsx"
    all_sheets = pd.read_excel(file, header=[6,7],sheet_name=None)  # Load all sheets
    # Convert sheet names to strings
    all_sheets = {str(sheet_name): df for sheet_name, df in all_sheets.items()}
    
    # Filter sheets based on start_week and end_week
    selected_sheets = {}
    for week_num in range(start_week, end_week + 1):
        sheet_name = str(week_num)
        if sheet_name in all_sheets:
            selected_sheets[sheet_name] = all_sheets[sheet_name]
    
    return selected_sheets

start_week = int(start_week)
end_week = int(end_week)
Sales2024 = load_data_Sales(start_week, end_week)
Sales2024 = pd.concat(Sales2024.values(), ignore_index=True)
####################################################################
def flatten_column(col):
    if isinstance(col[0], str) and col[0].startswith('Unnamed:'):
        return col[1]
    else:
        return '_'.join([str(e) for e in col if e])

# Apply the function to rename the columns
Sales2024.columns = Sales2024.columns.map(flatten_column)
#####################################################################
Sales2024['Part no.']=Sales2024['Part no.'].astype(str)
values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
'1731A','KOSHIN','043061102-']
# Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
mask = ~Sales2024['Part no.'].isin(values_to_exclude)
Sales2024 = Sales2024[mask]
Sales2024=pd.merge(Sales2024,DCDATA,on='Part no.',how='left')
##############
Sales2024.columns = Sales2024.columns.astype(str)
# prod_mask = Sales2024.columns.str.startswith(Minput+1)
ACT_mask = Sales2024.columns.str.endswith('ACT')
# combine_mask = (prod_mask & month_mask)
Date_columns = Sales2024.loc[:, ACT_mask]
Sales2024=Sales2024[['Part no.']+ Date_columns.columns.tolist()]
###########
Sales2024=Sales2024.groupby('Part no.').sum()
Sales2024=Sales2024.fillna(0)
###########
# Sales2024
# ########################## Production-NG 2024 ############################
@st.cache_data 
def load_data_NG(start_week, end_week):
    file = "Production-NG-"+Yinput+".xlsx"
    all_sheets = pd.read_excel(file, header=6, sheet_name=None)  # Load all sheets
    # Convert sheet names to strings
    all_sheets = {str(sheet_name): df for sheet_name, df in all_sheets.items()}
    
    # Filter sheets based on start_week and end_week
    selected_sheets = {}
    for week_num in range(start_week, end_week + 1):
        sheet_name = str(week_num)
        if sheet_name in all_sheets:
            selected_sheets[sheet_name] = all_sheets[sheet_name]
    
    return selected_sheets

start_week = int(start_week)
end_week = int(end_week)
####################################
Prod2024 = data2024
####################################
NG2024 = load_data_NG(start_week, end_week)
NG2024 = pd.concat(NG2024.values(), ignore_index=True)
NG2024['Part no.']=NG2024['Part no.'].astype(str)
values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
'1731A','KOSHIN','043061102-']
# Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
mask = ~NG2024['Part no.'].isin(values_to_exclude)
NG2024 = NG2024[mask]
NG2024=pd.merge(NG2024,DCDATA,on='Part no.',how='left')
# #################################################################################################################################################################
# ##################################################################################################################################################################
if Process=='Die_casting':
    st.write('DC Production Report@',Month,Yinput)

    ############# DC NG #################################
    prod_mask = NG2024.columns.str.startswith(Minput)
    month_mask = NG2024.columns.str.endswith(':00.1')
    combine_mask = (prod_mask & month_mask)
    Date_columns = NG2024.loc[:, combine_mask]
    NG2024=NG2024[['Part no.']+ Date_columns.columns.tolist()]
    NG2024=NG2024.groupby('Part no.').sum()
    NG2024['SUM-FN-NG']=NG2024.sum(axis=1)
    # NG2024
    NG2024=NG2024['SUM-FN-NG']
    # Ending Stock #######################################################################################
    EndCheck = load_data_File_A(start_week, end_week+1)
    EndCheck = pd.concat(EndCheck.values(), ignore_index=True)
    EndCheck['Part no.']=EndCheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EndCheck['Part no.'].isin(values_to_exclude)
    EndCheck = EndCheck[mask]
    ######################
    EndCheck=EndCheck[['Part no.','Beginning Balance']]
    EndCheck=EndCheck.groupby('Part no.').agg({'Beginning Balance':'last'})
    EndCheck.rename(columns={'Beginning Balance':'Ending ST (Chk)'},inplace=True)
    # EndCheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.1')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    DCOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # DCOver.set_index('Part no.',inplace=True)
    DCOver=DCOver.groupby('Part no.').sum()
    DCOver=DCOver.apply(pd.to_numeric, errors='coerce')
    ################# Over DC SUM ##################
    col2SUM=DCOver.columns.str.endswith(':00.1')
    DCOver = DCOver.loc[:, col2SUM]
    DCOver['FN-Over(Pcs)']=DCOver.sum(axis=1)
    DCOver=DCOver['FN-Over(Pcs)']
    ############ DC Prod Over ##########################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    DCPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    DCPro_Over=DCPro_Over.groupby('Part no.').agg(agg_funcs)
    DCPro_Over.rename(columns={'Beginning Balance':'Beginning Stock'},inplace=True)
    DCPro_Over=DCPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM DC Pcs ##############
    SUMDCPcs=DCPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    DCPro_Over['DC-Prod-Over(Pcs)']=SUMDCPcs.sum(axis=1)
    DCPro_Over=DCPro_Over['DC-Prod-Over(Pcs)']
    # DC Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    DCProd=Prod2024[['Part no.','Weight (g)','Beginning Balance']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    DCProd=DCProd.groupby('Part no.').agg(agg_funcs)
    DCProd.rename(columns={'Beginning Balance':'Beginning Stock'},inplace=True)
    DCProd=DCProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM DC Pcs ##############
    SUMDCPcs=DCProd.drop(columns=['Weight (g)','Beginning Stock'])
    DCProd['DC-Prod-(Pcs)']=SUMDCPcs.sum(axis=1)
    ############### Merge Begining #############################################
    DCProd=pd.merge(DCProd,EndCheck,left_index=True,right_index=True,how='left')
    DCProd=pd.merge(DCProd,DCOver,left_index=True,right_index=True,how='left')
    DCProd=pd.merge(DCProd,DCPro_Over,left_index=True,right_index=True,how='left')
    DCProd=pd.merge(DCProd,NG2024,left_index=True,right_index=True,how='left')
    
    ##############################################################################
    DCProd['DC-ST-(Pcs)']=DCProd['DC-Prod-(Pcs)']+DCProd['Beginning Stock']
    DCProd['DC-Ending-(Pcs)']=(DCProd['Ending ST (Chk)']-DCProd['DC-Prod-Over(Pcs)'])+DCProd['FN-Over(Pcs)']
    #################### Clean DC Part #######################
    Part_to_exclude = [
    '220-00331',
    '220-00016-1',
    '220-00016-2',
    '220-00014',
    '220-00015',
    'MD372348',
    'HP-001',
    'HP-002'
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~DCProd.index.isin(Part_to_exclude)
    DCProd= DCProd[mask]
    DCProd=DCProd.fillna(0)
    ##########################################################
    static_columns=['Weight (g)','Beginning Stock','DC-Prod-(Pcs)','DC-ST-(Pcs)','DC-Prod-Over(Pcs)','Ending ST (Chk)','FN-Over(Pcs)','DC-Ending-(Pcs)','SUM-FN-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    DCThai=DCProd[all_columns]
    DCThai.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','DC-Prod-(Pcs)':'ยอดผลิต','DC-ST-(Pcs)':'ยอดผลิต+ยกมา','Ending ST (Chk)':'WK-Stock ตรวจนับ','FN-Over(Pcs)':'ยอดเบิกเกิน','DC-Ending-(Pcs)':'Stock คงเหลือ','SUM-FN-NG':'ยอดงานเสีย','DC-Prod-Over(Pcs)':'ยอดผลิตเกิน'},inplace=True)
    DCProd
    ###############
    DCThai['Stock คงเหลือ'] = DCThai['Stock คงเหลือ'].apply(lambda x: 0 if x < 0 else x)
    ###############
    DCThai
    Part_Items = len(DCProd.index)
    Part_Items
    ################ Export To Excel #####################################
    location=r'C:\Users\utaie\Production-2024\Production-App\\'
    DCThai.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
    ######################################################################
    
    # SUM Table ###########################################################################
    ######################
    SUM_Beg=DCProd['Beginning Stock'].sum()
    SUM_Prod=DCProd['DC-Prod-(Pcs)'].sum()
    SUM_Ending=DCProd['DC-Ending-(Pcs)'].sum()

    SUM_NG_Kgs=((DCProd['SUM-FN-NG']*DCProd['Weight (g)'])/1000).sum()

    data = {
    'Description': ['Total Beginning Stock','Total DC-Production','Total DC-Ending','Total NG-Weight',],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    # ########### Unit Cost ###############################################
    # st.write(f'Unit Cost DC @ {Minput}')
    # DCProd[['Weight (g)','Beginning Stock','DC-Prod-(Pcs)','DC-ST-(Pcs)','DC-Prod-Over(Pcs)','Ending ST (Chk)','FN-Over(Pcs)','DC-Ending-(Pcs)','SUM-FN-NG']]

    # st.write('---')
    ##################################################################################################################################################################
#   ##################################################################################################################################################################
if Process=='Finishing':
    st.write('FN Production Report@',Month,Yinput)
    ############# FN NG #################################
    prod_mask = NG2024.columns.str.startswith(Minput)
    month_mask = NG2024.columns.str.endswith(':00.3')
    combine_mask = (prod_mask & month_mask)
    Date_columns = NG2024.loc[:, combine_mask]
    NG2024=NG2024[['Part no.']+ Date_columns.columns.tolist()]
    NG2024=NG2024.groupby('Part no.').sum()
    NG2024['SUM-SB-NG']=NG2024.sum(axis=1)
    # NG2024
    NG2024=NG2024['SUM-SB-NG']
    # Ending Stock #######################################################################################
    EnFNheck = load_data_File_A(start_week, end_week+1)
    EnFNheck = pd.concat(EnFNheck.values(), ignore_index=True)
    EnFNheck['Part no.']=EnFNheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnFNheck['Part no.'].isin(values_to_exclude)
    EnFNheck = EnFNheck[mask]
    ######################
    EnFNheck=EnFNheck[['Part no.','Beginning Balance.1']]
    EnFNheck=EnFNheck.groupby('Part no.').agg({'Beginning Balance.1':'last'})
    EnFNheck.rename(columns={'Beginning Balance.1':'Ending ST (Chk)'},inplace=True)
    # EnFNheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.3')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    FNOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # FNOver.set_index('Part no.',inplace=True)
    FNOver=FNOver.groupby('Part no.').sum()
    FNOver=FNOver.apply(pd.to_numeric, errors='coerce')
    FNOver['SB-Over(Pcs)']=FNOver.sum(axis=1)
    # FNOver
    FNOver=FNOver['SB-Over(Pcs)']
    ###########################################################################################
    ############ FN Prod Over ##########################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.1')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    FNPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    FNPro_Over=FNPro_Over.groupby('Part no.').agg(agg_funcs)
    FNPro_Over.rename(columns={'Beginning Balance':'Beginning Stock'},inplace=True)
    FNPro_Over=FNPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM FN Pcs ##############
    SUMFNPcs=FNPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    FNPro_Over['FN-Prod-Over(Pcs)']=SUMFNPcs.sum(axis=1)
    FNPro_Over=FNPro_Over['FN-Prod-Over(Pcs)']
    # FN Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.1')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    FNProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.1']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.1':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    FNProd=FNProd.groupby('Part no.').agg(agg_funcs)
    FNProd.rename(columns={'Beginning Balance.1':'Beginning Stock'},inplace=True)
    FNProd=FNProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM FN Pcs ##############
    SUMFNPcs=FNProd.drop(columns=['Weight (g)','Beginning Stock'])
    FNProd['FN-Prod-(Pcs)']=SUMFNPcs.sum(axis=1)
    # FNProd
    ############### Merge Begining #############################################
    FNProd=pd.merge(FNProd,EnFNheck,left_index=True,right_index=True,how='left')
    FNProd=pd.merge(FNProd,FNPro_Over,left_index=True,right_index=True,how='left')
    FNProd=pd.merge(FNProd,FNOver,left_index=True,right_index=True,how='left')
    FNProd=pd.merge(FNProd,NG2024,left_index=True,right_index=True,how='left')
    ##############################################################################
    FNProd['FN-ST-(Pcs)']=FNProd['FN-Prod-(Pcs)']+FNProd['Beginning Stock']
    Koshin=[ '43061102',
    '32005102',
    '39025305',
    '39047501',
    '11526802',
    '11526902',
    '32005202',
    '320003304',
    '12208301',
    '12208101']
    if FNProd.index.isin(Koshin).any():
        FNProd['SB-Over(Pcs)']=FNProd['FN-Prod-Over(Pcs)']
        FNProd['FN-Ending-(Pcs)']=(FNProd['Ending ST (Chk)']-FNProd['FN-Prod-Over(Pcs)'])+FNProd['SB-Over(Pcs)']
    else:
        FNProd['FN-Ending-(Pcs)']=(FNProd['Ending ST (Chk)']-FNProd['FN-Prod-Over(Pcs)'])+FNProd['SB-Over(Pcs)']
    #################### Clean FN Part #######################
    Part_to_exclude = [
    '220-00331',
    '220-00016-1',
    '220-00016-2',
    '220-00014',
    '220-00015',
    'MD372348',
    'HP-001',
    'HP-002',
    'HP-003',
    'HP-004',
    'HP-005',
    'HP-006',
    'HP-007',
    'HP-008',
    'HP-009',
    'HP-010',
    'HP-011',
    'HP-012']

    # ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~FNProd.index.isin(Part_to_exclude)
    FNProd= FNProd[mask]
    FNProd=FNProd.fillna(0)
    ##########################################################
    static_columns=['Weight (g)','Beginning Stock','FN-Prod-(Pcs)','FN-ST-(Pcs)','FN-Prod-Over(Pcs)','Ending ST (Chk)','SB-Over(Pcs)','FN-Ending-(Pcs)','SUM-SB-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    FNThai=FNProd[all_columns]
    FNThai.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','FN-Prod-(Pcs)':'ยอดผลิต','FN-ST-(Pcs)':'ยอดผลิต+ยกมา','Ending ST (Chk)':'WK-Stock ตรวจนับ','SB-Over(Pcs)':'ยอดเบิกเกิน','FN-Ending-(Pcs)':'Stock คงเหลือ','SUM-SB-NG':'ยอดงานเสีย','FN-Prod-Over(Pcs)':'ยอดผลิตเกิน'},inplace=True)
    FNProd
    #############
    FNThai['Stock คงเหลือ'] = FNThai['Stock คงเหลือ'].apply(lambda x: 0 if x < 0 else x)
    #############
    FNThai
    ################ Export To Excel #####################################
    location=r'C:\Users\utaie\Production-2024\Production-App\\'
    FNThai.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
    ######################################################################
    # SUM Table ###########################################################################
    ######################
    SUM_Beg=FNProd['Beginning Stock'].sum()
    SUM_Prod=FNProd['FN-Prod-(Pcs)'].sum()
    SUM_Ending=FNProd['FN-Ending-(Pcs)'].sum()

    SUM_NG_Kgs=((FNProd['SUM-SB-NG']*FNProd['Weight (g)'])/1000).sum()

    data = {
    'Description': ['Total Beginning Stock','Total FN-Production','Total FN-Ending','Total NG-Weight',],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    # ##################################################################################################################################################################
    # ##################################################################################################################################################################
if Process=='Sand_Blasting':
    st.write('SB Production Report@',Month,Yinput)
    ############# SB NG #################################
    prod_mask = NG2024.columns.str.startswith(Minput)
    month_mask = NG2024.columns.str.endswith(':00.3')
    combine_mask = (prod_mask & month_mask)
    Date_columns = NG2024.loc[:, combine_mask]
    NG2024=NG2024[['Part no.']+ Date_columns.columns.tolist()]
    NG2024=NG2024.groupby('Part no.').sum()
    NG2024['SUM-SB-NG']=NG2024.sum(axis=1)
    # NG2024
    NG2024=NG2024['SUM-SB-NG']
    # Ending Stock #######################################################################################
    EnSBheck = load_data_File_A(start_week, end_week+1)
    EnSBheck = pd.concat(EnSBheck.values(), ignore_index=True)
    EnSBheck['Part no.']=EnSBheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnSBheck['Part no.'].isin(values_to_exclude)
    EnSBheck = EnSBheck[mask]
    ######################
    EnSBheck=EnSBheck[['Part no.','Beginning Balance.3']]
    EnSBheck=EnSBheck.groupby('Part no.').agg({'Beginning Balance.3':'last'})
    EnSBheck.rename(columns={'Beginning Balance.3':'Ending ST (Chk)'},inplace=True)
    # EnSBheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.5')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    SBOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # SBOver.set_index('Part no.',inplace=True)
    SBOver=SBOver.groupby('Part no.').sum()
    SBOver=SBOver.apply(pd.to_numeric, errors='coerce')
    SBOver['MC-Over(Pcs)']=SBOver.sum(axis=1)
    # SBOver
    SBOver=SBOver['MC-Over(Pcs)']
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.6')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    MCOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # MCOver.set_index('Part no.',inplace=True)
    MCOver=MCOver.groupby('Part no.').sum()
    MCOver=MCOver.apply(pd.to_numeric, errors='coerce')
    MCOver['QC-Over(Pcs)']=MCOver.sum(axis=1)
    # MCOver
    MCOver=MCOver['QC-Over(Pcs)']
    ############ SB Prod Over ##########################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.1')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    SBPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    SBPro_Over=SBPro_Over.groupby('Part no.').agg(agg_funcs)
    SBPro_Over.rename(columns={'Beginning Balance':'Beginning Stock'},inplace=True)
    SBPro_Over=SBPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM SB Pcs ##############
    SUMSBPcs=SBPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    SBPro_Over['SB-Prod-Over(Pcs)']=SUMSBPcs.sum(axis=1)
    SBPro_Over=SBPro_Over['SB-Prod-Over(Pcs)']
    # SB Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.3')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    SBProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.3']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.3':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    SBProd=SBProd.groupby('Part no.').agg(agg_funcs)
    SBProd.rename(columns={'Beginning Balance.3':'Beginning Stock'},inplace=True)
    SBProd=SBProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM SB Pcs ##############
    SUMSBPcs=SBProd.drop(columns=['Weight (g)','Beginning Stock'])
    SBProd['SB-Prod-(Pcs)']=SUMSBPcs.sum(axis=1)
    # SBProd
    ############### Merge Begining #############################################
    SBProd=pd.merge(SBProd,EnSBheck,left_index=True,right_index=True,how='left')
    SBProd=pd.merge(SBProd,SBOver,left_index=True,right_index=True,how='left')
    SBProd=pd.merge(SBProd,SBPro_Over,left_index=True,right_index=True,how='left')
    SBProd=pd.merge(SBProd,NG2024,left_index=True,right_index=True,how='left')
    SBProd=pd.merge(SBProd,MCOver,left_index=True,right_index=True,how='left')
    ##############################################################################
    SBProd['SB-ST-(Pcs)']=SBProd['SB-Prod-(Pcs)']+SBProd['Beginning Stock']
    OutSource= [
    '5611500702A',
    '5611505402A',
    '5611503102A',
    '5611506803A',
    '5611500802A',
    '5611507702A',
    '5611510201A',
    '5611502001A',
    '5611512200A',
    'Z0004946A',
    '5611514600A',
    '5612602102A',
    'Z0009524A',
    '5612604900A',
    '5612605000A',
    '5611510801A',
    'T26164BA',
    'T36744BA',
    'T35584CA',
    'T909088A',
    'Z0016091A']
    #######################################################
    MCPart = [
    '5612603000A',
    '5612603100A',
    'T96493CA',
    'Z0021771A',
    '5611514900A',
    'T46496AA',
    'T46497AA',
    '5611515600A',
    '5611516100A',
    '5611502001B',
    '5612604400A',
    '5612604500A'
]
    ########################################################
    # Function to apply conditional logic to each row
    def calculate_sb_ending(row):
        if row.name in MCPart:
            return (row['Ending ST (Chk)'] - row['SB-Prod-Over(Pcs)']) + row['MC-Over(Pcs)']
        elif row.name in OutSource:
            return (row['Ending ST (Chk)'] - row['SB-Prod-Over(Pcs)'])  
        else:
            return (row['Ending ST (Chk)'] - row['SB-Prod-Over(Pcs)']) + row['QC-Over(Pcs)']
    ################################################
    SBProd.loc[SBProd.index.isin(MCPart), 'QC-Over(Pcs)'] = 0
    SBProd.loc[SBProd.index.isin(OutSource), ['MC-Over(Pcs)', 'QC-Over(Pcs)']] = 0
    #################################################

    # Apply the function to each row of the DataFrame
    SBProd['SB-Ending-(Pcs)'] = SBProd.apply(calculate_sb_ending, axis=1)
    #################### Clean SB Part #######################
    Part_to_exclude = [
    '220-00331',
    '220-00016-1',
    '220-00016-2',
    '220-00014',
    '220-00015',
    '1050B375',
    'HP-001',
    'HP-002',
    'HP-003',
    'HP-004',
    'HP-005',
    'HP-006',
    'HP-007',
    'HP-008',
    'HP-009',
    'HP-010',
    'HP-011',
    'HP-012',
    '43061102',
    '32005102',
    '39025305',
    '39047501',
    '11526802',
    '11526902',
    '32005202',
    '320003304',
    '12208301',
    '12208101'
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~SBProd.index.isin(Part_to_exclude)
    SBProd= SBProd[mask]
    SBProd=SBProd.fillna(0)
    ##########################################################
    
    static_columns=['Weight (g)','Beginning Stock','SB-Prod-(Pcs)','SB-ST-(Pcs)','SB-Prod-Over(Pcs)','Ending ST (Chk)','MC-Over(Pcs)','QC-Over(Pcs)','SB-Ending-(Pcs)','SUM-SB-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    SBThai=SBProd[all_columns]
    SBThai.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','SB-Prod-(Pcs)':'ยอดผลิต','SB-ST-(Pcs)':'ยอดผลิต+ยกมา','SB-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','MC-Over(Pcs)':'MC เบิกเกิน','QC-Over(Pcs)':'QC เบิกเกิน','SB-Ending-(Pcs)':'Stock คงเหลือ','SUM-SB-NG':'ยอดงานเสีย'},inplace=True)
    SBProd
    ###############
    SBThai['Stock คงเหลือ'] = SBThai['Stock คงเหลือ'].apply(lambda x: 0 if x < 0 else x)
    ############
    SBThai
    ################ Export To Excel #####################################
    location=r'C:\Users\utaie\Production-2024\Production-App\\'
    SBThai.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
    ######################################################################
    # SUM Table ###########################################################################
    ######################
    SUM_Beg=SBProd['Beginning Stock'].sum()
    SUM_Prod=SBProd['SB-Prod-(Pcs)'].sum()
    SUM_Ending=SBProd['SB-Ending-(Pcs)'].sum()

    SUM_NG_Kgs=((SBProd['SUM-SB-NG']*SBProd['Weight (g)'])/1000).sum()

    data = {
    'Description': ['Total Beginning Stock','Total SB-Production','Total SB-Ending','Total NG-Weight',],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    # ##################################################################################################################################################################
    # ##################################################################################################################################################################
if Process=='Machine':
    st.write('MC Production Report@',Month,Yinput)
    ############# MC NG #################################
    prod_mask = NG2024.columns.str.startswith(Minput)
    month_mask = NG2024.columns.str.endswith(':00.5')
    combine_mask = (prod_mask & month_mask)
    Date_columns = NG2024.loc[:, combine_mask]
    NG2024=NG2024[['Part no.']+ Date_columns.columns.tolist()]
    NG2024=NG2024.groupby('Part no.').sum()
    NG2024= NG2024.apply(pd.to_numeric, errors='coerce')
    NG2024['SUM-MC-NG']=NG2024.sum(axis=1)
    # NG2024
    NG2024=NG2024['SUM-MC-NG']
    # Ending Stock #######################################################################################
    EnMCheck = load_data_File_A(start_week, end_week+1)
    EnMCheck = pd.concat(EnMCheck.values(), ignore_index=True)
    EnMCheck['Part no.']=EnMCheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnMCheck['Part no.'].isin(values_to_exclude)
    EnMCheck = EnMCheck[mask]
    ######################
    EnMCheck=EnMCheck[['Part no.','Beginning Balance.5']]
    EnMCheck=EnMCheck.groupby('Part no.').agg({'Beginning Balance.5':'last'})
    EnMCheck.rename(columns={'Beginning Balance.5':'Ending ST (Chk)'},inplace=True)
    # EnMCheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.6')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    MCOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # MCOver.set_index('Part no.',inplace=True)
    MCOver=MCOver.groupby('Part no.').sum()
    MCOver=MCOver.apply(pd.to_numeric, errors='coerce')
    MCOver['QC-Over(Pcs)']=MCOver.sum(axis=1)
    # MCOver
    MCOver=MCOver['QC-Over(Pcs)']
    ############ MC Prod Over ###############################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.5')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MCPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance.5']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.5':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MCPro_Over=MCPro_Over.groupby('Part no.').agg(agg_funcs)
    MCPro_Over.rename(columns={'Beginning Balance.5':'Beginning Stock'},inplace=True)
    MCPro_Over=MCPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC Pcs ##############
    SUMMCPcs=MCPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    MCPro_Over['MC-Prod-Over(Pcs)']=SUMMCPcs.sum(axis=1)
    MCPro_Over=MCPro_Over['MC-Prod-Over(Pcs)']
    # MC Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.5')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MCProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.5']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.5':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MCProd=MCProd.groupby('Part no.').agg(agg_funcs)
    MCProd.rename(columns={'Beginning Balance.5':'Beginning Stock'},inplace=True)
    MCProd=MCProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC Pcs ##############
    SUMMCPcs=MCProd.drop(columns=['Weight (g)','Beginning Stock'])
    MCProd['MC-Prod-(Pcs)']=SUMMCPcs.sum(axis=1)
    # MCProd
    ############### Merge Begining #############################################
    MCProd=pd.merge(MCProd,EnMCheck,left_index=True,right_index=True,how='left')
    MCProd=pd.merge(MCProd,MCOver,left_index=True,right_index=True,how='left')
    MCProd=pd.merge(MCProd,MCPro_Over,left_index=True,right_index=True,how='left')
    MCProd=pd.merge(MCProd,NG2024,left_index=True,right_index=True,how='left')
    
    ##############################################################################
    MCProd['Ending ST (Chk)'] = pd.to_numeric(MCProd['Ending ST (Chk)'], errors='coerce')
    MCProd['MC-Prod-Over(Pcs)'] = pd.to_numeric(MCProd['MC-Prod-Over(Pcs)'], errors='coerce')
    MCProd['QC-Over(Pcs)'] = pd.to_numeric(MCProd['QC-Over(Pcs)'], errors='coerce')
    #############################################################################
    MCProd['MC-TT-(Pcs)']=MCProd['MC-Prod-(Pcs)']+MCProd['Beginning Stock']
    MCProd['MC-Ending-(BL)'] = (MCProd['Ending ST (Chk)'] - MCProd['MC-Prod-Over(Pcs)']) + MCProd['QC-Over(Pcs)']
    #################### Clean MC Part #######################
    Part_to_include = [
    '5612603000A',
    '5612603100A',
    'T96493CA',
    'Z0021771A',
    '5611514900A',
    'T46496AA',
    'T46497AA',
    '5611515600A',
    '5611516100A',
    '5611502001B',
    '5612604400A',
    '5612604500A'
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = MCProd.index.isin(Part_to_include)
    MCProd= MCProd[mask]
    MCProd=MCProd.fillna(0)
    ##########################################################
    SUM_Beg=MCProd['Beginning Stock'].sum()
    SUM_Prod=MCProd['MC-Prod-(Pcs)'].sum()
    SUM_Ending=MCProd['MC-Ending-(BL)'].sum()
    #########################################################
    SUM_NG_Kgs=((MCProd['SUM-MC-NG']*MCProd['Weight (g)'])/1000).sum()
    static_columns=['Weight (g)','Beginning Stock','MC-Prod-(Pcs)','MC-TT-(Pcs)','MC-Prod-Over(Pcs)','Ending ST (Chk)','QC-Over(Pcs)','MC-Ending-(BL)','SUM-MC-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    MCProd=MCProd[all_columns]
    MCProd.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','MC-Prod-(Pcs)':'ยอดผลิต','MC-TT-(Pcs)':'ยอดผลิต+ยกมา','MC-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','QC-Over(Pcs)':'ยอดเบิกเกิน','MC-Ending-(BL)':'Stock คงเหลือ','SUM-MC-NG':'ยอดงานเสีย'},inplace=True)
    MCProd
    ######################################
    data = {
    'Description': ['Total Beginning Stock','Total MC-Production','Total MC-Ending','Total NG-Weight',],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    #### MC OP1 #################################################################################################################################################
    st.write('MC_OP1 Production Report@',Month,Yinput)
    # Ending Stock #######################################################################################
    EnMC_OP1heck = load_data_File_A(start_week, end_week+1)
    EnMC_OP1heck = pd.concat(EnMC_OP1heck.values(), ignore_index=True)
    EnMC_OP1heck['Part no.']=EnMC_OP1heck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'OP1TROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnMC_OP1heck['Part no.'].isin(values_to_exclude)
    EnMC_OP1heck = EnMC_OP1heck[mask]
    ######################
    EnMC_OP1heck=EnMC_OP1heck[['Part no.','Beginning Balance.4']]
    EnMC_OP1heck=EnMC_OP1heck.groupby('Part no.').agg({'Beginning Balance.4':'last'})
    EnMC_OP1heck.rename(columns={'Beginning Balance.4':'Ending ST (Chk)'},inplace=True)
    ####################
    # EnMC_OP1heck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.4')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    MC_OP1Over=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # MC_OP1Over.set_index('Part no.',inplace=True)
    MC_OP1Over=MC_OP1Over.groupby('Part no.').sum()
    MC_OP1Over=MC_OP1Over.apply(pd.to_numeric, errors='coerce')
    MC_OP1Over['QC-Over(Pcs)']=MC_OP1Over.sum(axis=1)
    # MC_OP1Over
    MC_OP1Over=MC_OP1Over['QC-Over(Pcs)']
    ############ MC_OP1 Prod Over ##################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.4')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_OP1Pro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance.4']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.4':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_OP1Pro_Over=MC_OP1Pro_Over.groupby('Part no.').agg(agg_funcs)
    MC_OP1Pro_Over.rename(columns={'Beginning Balance.4':'Beginning Stock'},inplace=True)
    MC_OP1Pro_Over=MC_OP1Pro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_OP1 Pcs ##############
    SUMMC_OP1Pcs=MC_OP1Pro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    MC_OP1Pro_Over['MC_OP1-Prod-Over(Pcs)']=SUMMC_OP1Pcs.sum(axis=1)
    MC_OP1Pro_Over=MC_OP1Pro_Over['MC_OP1-Prod-Over(Pcs)']
    # MC_OP1 Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.4')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_OP1Prod=Prod2024[['Part no.','Weight (g)','Beginning Balance.4']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.4':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_OP1Prod=MC_OP1Prod.groupby('Part no.').agg(agg_funcs)
    MC_OP1Prod.rename(columns={'Beginning Balance.4':'Beginning Stock'},inplace=True)
    MC_OP1Prod=MC_OP1Prod.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_OP1 Pcs ##############
    SUMMC_OP1Pcs=MC_OP1Prod.drop(columns=['Weight (g)','Beginning Stock'])
    MC_OP1Prod['MC_OP1-Prod-(Pcs)']=SUMMC_OP1Pcs.sum(axis=1)
    # MC_OP1Prod
    ############### Merge Begining #############################################
    MC_OP1Prod=pd.merge(MC_OP1Prod,EnMC_OP1heck,left_index=True,right_index=True,how='left')
    MC_OP1Prod=pd.merge(MC_OP1Prod,MC_OP1Over,left_index=True,right_index=True,how='left')
    MC_OP1Prod=pd.merge(MC_OP1Prod,MC_OP1Pro_Over,left_index=True,right_index=True,how='left')
    MC_OP1Prod=pd.merge(MC_OP1Prod,NG2024,left_index=True,right_index=True,how='left')
    ##############################################################################
    MC_OP1Prod['MC_OP1-ST-(Pcs)']=MC_OP1Prod['MC_OP1-Prod-(Pcs)']+MC_OP1Prod['Beginning Stock']
    MC_OP1Prod['MC_OP1-Ending-(Pcs)']=(MC_OP1Prod['Ending ST (Chk)']-MC_OP1Prod['MC_OP1-Prod-Over(Pcs)'])
    MC_OP1Prod['MC_OP1-Ending-(Pcs)'] = MC_OP1Prod['MC_OP1-Ending-(Pcs)'].apply(lambda x: 0 if x < 0 else x)
    #################### Clean MC_OP1 Part #######################
    Part_to_include = [
    '5612604900A',
    '5612605000A'
    ]
    ##########################################
    if Minput>='2024-05':
        mask = MC_OP1Prod.index.isin(Part_to_include)
        MC_OP1Prod= MC_OP1Prod[mask]
        MC_OP1Prod['MC_OP1-Prod-(Pcs)'] = MC_OP1Prod['MC_OP1-Prod-(Pcs)'].apply(lambda x: 0 if x > 0 else x)
    else:
        mask = MC_OP1Prod.index.isin(Part_to_include)
        MC_OP1Prod= MC_OP1Prod[mask]
        MC_OP1Prod=MC_OP1Prod.fillna(0)
    ####################################
    MC_OP1Prod.loc[MC_OP1Prod.index.isin(Part_to_include), 'QC-Over(Pcs)'] = 0
    ##########################################################
    SUM_Beg=MC_OP1Prod['Beginning Stock'].sum()
    SUM_Prod=MC_OP1Prod['MC_OP1-Prod-(Pcs)'].sum()
    SUM_Ending=MC_OP1Prod['MC_OP1-Ending-(Pcs)'].sum()
    SUM_NG_Kgs=((MC_OP1Prod['SUM-MC-NG']*MC_OP1Prod['Weight (g)'])/1000).sum()
    ############################################################
    static_columns=['Weight (g)','Beginning Stock','MC_OP1-Prod-(Pcs)','MC_OP1-ST-(Pcs)','MC_OP1-Prod-Over(Pcs)','Ending ST (Chk)','QC-Over(Pcs)','MC_OP1-Ending-(Pcs)','SUM-MC-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    MC_OP1Prod=MC_OP1Prod[all_columns]
    MC_OP1Prod.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','MC_OP1-Prod-(Pcs)':'ยอดผลิต','MC_OP1-ST-(Pcs)':'ยอดผลิต+ยกมา','MC_OP1-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','QC-Over(Pcs)':'ยอดเบิกเกิน','MC_OP1-Ending-(Pcs)':'Stock คงเหลือ',},inplace=True)
    MC_OP1Prod
    ###########################################################################
    data = {
    'Description': ['Total Beginning Stock','Total MC_OP1-Production','Total MC_OP1-Ending','Total NG-Weight'],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ",f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    ##################################################################################################################################################################
    # MC Electrolux
    st.write('MC_Electrolux Production Report@',Month,Yinput)
    # Ending Stock #######################################################################################
    EnMC_Elecheck = load_data_File_A(start_week, end_week+1)
    EnMC_Elecheck = pd.concat(EnMC_Elecheck.values(), ignore_index=True)
    EnMC_Elecheck['Part no.']=EnMC_Elecheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnMC_Elecheck['Part no.'].isin(values_to_exclude)
    EnMC_Elecheck = EnMC_Elecheck[mask]
    ######################
    EnMC_Elecheck=EnMC_Elecheck[['Part no.','Beginning Balance.1']]
    EnMC_Elecheck=EnMC_Elecheck.groupby('Part no.').agg({'Beginning Balance.1':'last'})
    EnMC_Elecheck.rename(columns={'Beginning Balance.1':'Ending ST (Chk)'},inplace=True)
    # EnMC_Elecheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.3')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    MC_ElecOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # MC_ElecOver.set_index('Part no.',inplace=True)
    MC_ElecOver=MC_ElecOver.groupby('Part no.').sum()
    MC_ElecOver=MC_ElecOver.apply(pd.to_numeric, errors='coerce')
    MC_ElecOver['QC-Over(Pcs)']=MC_ElecOver.sum(axis=1)
    # MC_ElecOver
    MC_ElecOver=MC_ElecOver['QC-Over(Pcs)']
    ############ MC_Elect Prod Over ######################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.1')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_ElectPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance.1']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.1':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_ElectPro_Over=MC_ElectPro_Over.groupby('Part no.').agg(agg_funcs)
    MC_ElectPro_Over.rename(columns={'Beginning Balance.1':'Beginning Stock'},inplace=True)
    MC_ElectPro_Over=MC_ElectPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_Elect Pcs ##############
    SUMMC_ElectPcs=MC_ElectPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    MC_ElectPro_Over['MC_Elect-Prod-Over(Pcs)']=SUMMC_ElectPcs.sum(axis=1)
    MC_ElectPro_Over=MC_ElectPro_Over['MC_Elect-Prod-Over(Pcs)']
    # MC_Elec Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.1')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_ElecProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.1']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.1':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_ElecProd=MC_ElecProd.groupby('Part no.').agg(agg_funcs)
    MC_ElecProd.rename(columns={'Beginning Balance.1':'Beginning Stock'},inplace=True)
    MC_ElecProd=MC_ElecProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_Elec Pcs ##############
    SUMMC_ElecPcs=MC_ElecProd.drop(columns=['Weight (g)','Beginning Stock'])
    MC_ElecProd['MC_Elec-Prod-(Pcs)']=SUMMC_ElecPcs.sum(axis=1)
    # MC_ElecProd
    ############### Merge Begining #############################################
    MC_ElecProd=pd.merge(MC_ElecProd,EnMC_Elecheck,left_index=True,right_index=True,how='left')
    MC_ElecProd=pd.merge(MC_ElecProd,MC_ElecOver,left_index=True,right_index=True,how='left')
    MC_ElecProd=pd.merge(MC_ElecProd,MC_ElectPro_Over,left_index=True,right_index=True,how='left')
    ##############################################################################
    MC_ElecProd['MC_Elec-ST-(Pcs)']=MC_ElecProd['MC_Elec-Prod-(Pcs)']+MC_ElecProd['Beginning Stock']
    MC_ElecProd['MC_Elec-Ending-(Pcs)']=(MC_ElecProd['Ending ST (Chk)']-MC_ElecProd['MC_Elect-Prod-Over(Pcs)'])+MC_ElecProd['QC-Over(Pcs)']
    #################### Clean MC_Elec Part #######################
    Part_to_include = [
    '220-00331',
    '220-00016-1',
    '220-00016-2'
   
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = MC_ElecProd.index.isin(Part_to_include)
    MC_ElecProd= MC_ElecProd[mask]
    MC_ElecProd=MC_ElecProd.fillna(0)
    ##########################################################
    SUM_Beg=MC_ElecProd['Beginning Stock'].sum()
    SUM_Prod=MC_ElecProd['MC_Elec-Prod-(Pcs)'].sum()
    SUM_Ending=MC_ElecProd['MC_Elec-Ending-(Pcs)'].sum()
    ##########################################################
    static_columns=['Weight (g)','Beginning Stock','MC_Elec-Prod-(Pcs)','MC_Elec-ST-(Pcs)','MC_Elect-Prod-Over(Pcs)','Ending ST (Chk)','QC-Over(Pcs)','MC_Elec-Ending-(Pcs)',]
    all_columns = Date_columns.columns.tolist()+static_columns
    MC_ElecProd=MC_ElecProd[all_columns]
    MC_ElecProd.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','MC_Elec-Prod-(Pcs)':'ยอดผลิต','MC_Elec-ST-(Pcs)':'ยอดผลิต+ยกมา','MC_Elect-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','QC-Over(Pcs)':'ยอดเบิกเกิน','MC_Elec-Ending-(Pcs)':'Stock คงเหลือ',},inplace=True)
    MC_ElecProd
    ################################################
    data = {
    'Description': ['Total Beginning Stock','Total MC_Elec-Production','Total MC_Elec-Ending'],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    # ##################################################################################################################################################################
    ##################################################################################################################################################################
    # MC Hometrolux
    st.write('MC_Homexpert Production Report@',Month,Yinput)
    # Ending Stock #######################################################################################
    EnMC_Homeheck = load_data_File_A(start_week, end_week+1)
    EnMC_Homeheck = pd.concat(EnMC_Homeheck.values(), ignore_index=True)
    EnMC_Homeheck['Part no.']=EnMC_Homeheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'HomeTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnMC_Homeheck['Part no.'].isin(values_to_exclude)
    EnMC_Homeheck = EnMC_Homeheck[mask]
    ######################
    EnMC_Homeheck=EnMC_Homeheck[['Part no.','Beginning Balance.2']]
    EnMC_Homeheck=EnMC_Homeheck.groupby('Part no.').agg({'Beginning Balance.2':'last'})
    EnMC_Homeheck.rename(columns={'Beginning Balance.2':'Ending ST (Chk)'},inplace=True)
    # EnMC_Homeheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    prod_mask = data2024.columns.str.startswith(Minput_next)
    month_mask = data2024.columns.str.endswith(':00.3')
    combine_mask = (prod_mask & month_mask)
    Date_columns2 = data2024.loc[:, combine_mask]
    ########################
    MC_HomeOver=data2024[['Part no.']+ Date_columns2.columns.tolist()]
    # MC_HomeOver.set_index('Part no.',inplace=True)
    MC_HomeOver=MC_HomeOver.groupby('Part no.').sum()
    MC_HomeOver=MC_HomeOver.apply(pd.to_numeric, errors='coerce')
    MC_HomeOver['QC-Over(Pcs)']=MC_HomeOver.sum(axis=1)
    # MC_HomeOver
    MC_HomeOver=MC_HomeOver['QC-Over(Pcs)']
    ############ MC_Home Prod Over #######################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.2')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_HomePro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance.2']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.2':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_HomePro_Over=MC_HomePro_Over.groupby('Part no.').agg(agg_funcs)
    MC_HomePro_Over.rename(columns={'Beginning Balance.2':'Beginning Stock'},inplace=True)
    MC_HomePro_Over=MC_HomePro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_Home Pcs ##############
    SUMMC_HomePcs=MC_HomePro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    MC_HomePro_Over['MC_Home-Prod-Over(Pcs)']=SUMMC_HomePcs.sum(axis=1)
    MC_HomePro_Over=MC_HomePro_Over['MC_Home-Prod-Over(Pcs)']
    # MC_Home Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.2')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    MC_HomeProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.2']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.2':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    MC_HomeProd=MC_HomeProd.groupby('Part no.').agg(agg_funcs)
    MC_HomeProd.rename(columns={'Beginning Balance.2':'Beginning Stock'},inplace=True)
    MC_HomeProd=MC_HomeProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM MC_Home Pcs ##############
    SUMMC_HomePcs=MC_HomeProd.drop(columns=['Weight (g)','Beginning Stock'])
    MC_HomeProd['MC_Home-Prod-(Pcs)']=SUMMC_HomePcs.sum(axis=1)
    # MC_HomeProd
    ############### Merge Begining #############################################
    MC_HomeProd=pd.merge(MC_HomeProd,EnMC_Homeheck,left_index=True,right_index=True,how='left')
    MC_HomeProd=pd.merge(MC_HomeProd,MC_HomeOver,left_index=True,right_index=True,how='left')
    MC_HomeProd=pd.merge(MC_HomeProd,MC_HomePro_Over,left_index=True,right_index=True,how='left')
    ##############################################################################
    MC_HomeProd['MC_Home-ST-(Pcs)']=MC_HomeProd['MC_Home-Prod-(Pcs)']+MC_HomeProd['Beginning Stock']
    MC_HomeProd['MC_Home-Ending-(Pcs)']=(MC_HomeProd['Ending ST (Chk)']-MC_HomeProd['MC_Home-Prod-Over(Pcs)'])+MC_HomeProd['QC-Over(Pcs)']
    
    #################### Clean MC_Home Part #######################
    Part_to_include = [
       'HP-001',
        'HP-002',
        'HP-003',
        'HP-004',
        'HP-005',
        'HP-006',
        'HP-007',
        'HP-008',
        'HP-009',
        'HP-011',
        'HP-012'
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = MC_HomeProd.index.isin(Part_to_include)
    MC_HomeProd= MC_HomeProd[mask]
    MC_HomeProd=MC_HomeProd.fillna(0)
    ##########################################################
    SUM_Beg=MC_HomeProd['Beginning Stock'].sum()
    SUM_Prod=MC_HomeProd['MC_Home-Prod-(Pcs)'].sum()
    SUM_Ending=MC_HomeProd['MC_Home-Ending-(Pcs)'].sum()
    ############################################################
    static_columns=['Weight (g)','Beginning Stock','MC_Home-Prod-(Pcs)','MC_Home-ST-(Pcs)','MC_Home-Prod-Over(Pcs)','Ending ST (Chk)','QC-Over(Pcs)','MC_Home-Ending-(Pcs)',]
    all_columns = Date_columns.columns.tolist()+static_columns
    MC_HomeProd=MC_HomeProd[all_columns]
    MC_HomeProd.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','MC_Home-Prod-(Pcs)':'ยอดผลิต','MC_Home-ST-(Pcs)':'ยอดผลิต+ยกมา','MC_Home-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','QC-Over(Pcs)':'ยอดเบิกเกิน','MC_Home-Ending-(Pcs)':'Stock คงเหลือ'},inplace=True)
    MC_HomeProd
    ######################
    data = {
    'Description': ['Total Beginning Stock','Total MC_Home-Production','Total MC_Home-Ending'],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    ##################################################################################################################################################################
    ALLMC=pd.concat([MC_HomeProd,MC_ElecProd,MCProd,MC_OP1Prod],axis=0)
    ALLMC=ALLMC.fillna(0)
    ALLMC = ALLMC.groupby(ALLMC.columns, axis=1).sum()
    ########################
    static_columns=['นน.(กรัม)','ยอดยกมา','ยอดผลิต','ยอดผลิต+ยกมา','ยอดผลิตเกิน','WK-Stock ตรวจนับ','ยอดเบิกเกิน','Stock คงเหลือ','ยอดงานเสีย']
    all_columns = Date_columns.columns.tolist()+static_columns
    st.write('ALL_MC Production Report@',Month,Yinput)
    ALLMC=ALLMC[all_columns ]
    ALLMC
#     # ####################
#     # ALLMC=pd.concat([MC_HomeThai,MC_ElecThai,MCThai,MC_OP1Thai],axis=0)
#     # ALLMC=ALLMC.fillna(0)
#     # ALLMC = ALLMC.groupby(ALLMC.columns, axis=1).sum()
#     # ###################################
#     # PARTOP1=['5612604900A','5612605000A']
#     # def calculate_mc_ending(row):
#     #     if row.name in PARTOP1:
#     #         return (row['WK-Stock ตรวจนับ'] - row['ยอดผลิตเกิน'])
#     #     else:
#     #         return (row['WK-Stock ตรวจนับ'] - row['ยอดผลิตเกิน']) + row['ยอดเบิกเกิน']
#     # ######################
#     # ALLMC.loc[ALLMC.index.isin(PARTOP1), 'ยอดเบิกเกิน'] = 0
#     # ###########################
#     # ALLMC['Stock คงเหลือ'] = ALLMC['Stock คงเหลือ'].apply(lambda x: 0 if x < 0 else x)
#     # ##########################################
#     # static_columns=['นน.(กรัม)','ยอดยกมา','ยอดผลิต','ยอดผลิต+ยกมา','ยอดผลิตเกิน','WK-Stock ตรวจนับ','ยอดเบิกเกิน','Stock คงเหลือ',]
#     # all_columns = Date_columns.columns.tolist()+static_columns
#     # #########################
#     # st.write('ALL-MC Production Report@',Month,Yinput)
#     # ALLMC=ALLMC[all_columns]
#     ALLMC
#     ################ Export To Excel #####################################
#     location=r'C:\Users\utaie\Production-2024\Production-App\\'
#     ALLMC.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
#     ######################################################################
#     ######################
    SUM_Beg=ALLMC['ยอดยกมา'].sum()
    SUM_Prod=ALLMC['ยอดผลิต'].sum()
    SUM_Ending=ALLMC['Stock คงเหลือ'].sum()
    SUM_NG_Kgs=((ALLMC['ยอดงานเสีย']*ALLMC['นน.(กรัม)'])/1000).sum()


    data = {
    'Description': ['Total Beginning Stock','Total MC-Production','Total MC-Ending','Total MC-NG'],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f}"],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']}

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')
    ################ Export To Excel #####################################
    location=r'C:\Users\utaie\Production-2024\Production-App\\'
    ALLMC.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
    ######################################################################
    # #################################################################################################################################################################
    # ##################################################################################################################################################################
if Process=='QC':
    st.write('QC Production Report@',Month,Yinput)
    ############# QC NG Final #################################
    prod_mask1 = NG2024.columns.str.startswith(Minput)
    month_mask1 = NG2024.columns.str.endswith(':00.6')
    month_mask2 = NG2024.columns.str.endswith(':00.7')
    combine_mask1 = (prod_mask1 & month_mask1)|(prod_mask1 & month_mask2)
    Date_columns1 = NG2024.loc[:, combine_mask1]
    NG2024_1=NG2024[['Part no.']+ Date_columns1.columns.tolist()]
    NG2024_1=NG2024_1.groupby('Part no.').sum()
    #################
    NG_1=NG2024_1.columns.str.endswith(':00.6')
    FIAL_NG=NG2024_1.loc[:, NG_1]
    FIAL_NG= FIAL_NG.apply(pd.to_numeric, errors='coerce')
    NG2024_1['SUM-Final-NG']=FIAL_NG.sum(axis=1)
    #################
    NG_2=NG2024_1.columns.str.endswith(':00.7')
    CSL_NG=NG2024_1.loc[:, NG_2]
    CSL_NG= CSL_NG.apply(pd.to_numeric, errors='coerce')
    NG2024_1['SUM-CSL1-NG']=CSL_NG.sum(axis=1)
    NG2024_1['SUM-QC-NG']=NG2024_1['SUM-Final-NG']+NG2024_1['SUM-CSL1-NG']
    ################
    # NG2024_1
    QC_NG2024=NG2024_1['SUM-QC-NG']
    # Ending Stock #######################################################################################
    EnQCheck = load_data_File_A(start_week, end_week+1)
    EnQCheck = pd.concat(EnQCheck.values(), ignore_index=True)
    EnQCheck['Part no.']=EnQCheck['Part no.'].astype(str)
    values_to_exclude = ['nan', 'TBKK', 'ELECTROLUX', 'HOME EXPERT','0','Part no.','KAYAMA',
    '1731A','KOSHIN','043061102-']
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~EnQCheck['Part no.'].isin(values_to_exclude)
    EnQCheck = EnQCheck[mask]
    #####################
    EnQCheck=EnQCheck[['Part no.','Beginning Balance.6']]
    EnQCheck=EnQCheck.groupby('Part no.').agg({'Beginning Balance.6':'last'})
    EnQCheck.rename(columns={'Beginning Balance.6':'Ending ST (Chk)'},inplace=True)
    # EnQCheck
    # Over Next Process ######################################################################################
    from datetime import datetime
    from dateutil.relativedelta import relativedelta
    ################
    Minput = Minput
    date_obj = datetime.strptime(Minput, '%Y-%m')
    next_month_obj = date_obj + relativedelta(months=1)
    Minput_next = next_month_obj.strftime('%Y-%m')
    ################
    Sales_mask = Sales2024.columns.str.startswith(Minput_next)
    combine_mask = Sales_mask
    Date_columns = Sales2024.loc[:, combine_mask]
    QCOver=Sales2024[Date_columns.columns.tolist()]
    QCOver['Sales-Over(Pcs)']=QCOver.sum(axis=1)
    # QCOver
    QCOver=QCOver['Sales-Over(Pcs)']
    ############ QC Prod Over ##############################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput_next)
    month_mask = Prod2024.columns.str.endswith(':00.6')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    QCPro_Over=Prod2024[['Part no.','Weight (g)','Beginning Balance.6']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.6':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    QCPro_Over=QCPro_Over.groupby('Part no.').agg(agg_funcs)
    QCPro_Over.rename(columns={'Beginning Balance.6':'Beginning Stock'},inplace=True)
    QCPro_Over=QCPro_Over.apply(pd.to_numeric, errors='coerce')
    ############ SUM QC Pcs ##############
    SUMQCPcs=QCPro_Over.drop(columns=['Weight (g)','Beginning Stock'])
    QCPro_Over['QC-Prod-Over(Pcs)']=SUMQCPcs.sum(axis=1)
    QCPro_Over=QCPro_Over['QC-Prod-Over(Pcs)']
    # QC Production ###########################################################################################
    Prod2024.columns=Prod2024.columns.astype(str)
    prod_mask = Prod2024.columns.str.startswith(Minput)
    month_mask = Prod2024.columns.str.endswith(':00.6')
    combine_mask = prod_mask & month_mask
    Date_columns = Prod2024.loc[:, combine_mask]
    QCProd=Prod2024[['Part no.','Weight (g)','Beginning Balance.6']+ Date_columns.columns.tolist()]
    agg_funcs = {'Weight (g)': np.mean,'Beginning Balance.6':'first'}
    for col in Date_columns.columns:
        agg_funcs[col] = np.sum
    QCProd=QCProd.groupby('Part no.').agg(agg_funcs)
    QCProd.rename(columns={'Beginning Balance.6':'Beginning Stock'},inplace=True)
    QCProd=QCProd.apply(pd.to_numeric, errors='coerce')
    ############ SUM QC Pcs ##############
    SUMQCPcs=QCProd.drop(columns=['Weight (g)','Beginning Stock'])
    QCProd['QC-Prod-(Pcs)']=SUMQCPcs.sum(axis=1)
    ############### Merge Begining #############################################
    QCProd=pd.merge(QCProd,EnQCheck,left_index=True,right_index=True,how='left')
    QCProd=pd.merge(QCProd,QCOver,left_index=True,right_index=True,how='left')
    QCProd=pd.merge(QCProd,QC_NG2024,left_index=True,right_index=True,how='left')
    QCProd=pd.merge(QCProd,QCPro_Over,left_index=True,right_index=True,how='left')
    ##############################################################################
    QCProd['QC-ST-(Pcs)']=QCProd['QC-Prod-(Pcs)']+QCProd['Beginning Stock']
    QCProd['QC-Ending-(Pcs)']=(QCProd['Ending ST (Chk)']-QCProd['QC-Prod-Over(Pcs)'])+QCProd['Sales-Over(Pcs)']
    ##########################################################
    Part_to_exclude = [
    '5611500702A',
    '5611505402A',
    '5611503102A',
    '5611506803A',
    '5611500802A',
    '5611507702A',
    '5611510201A',
    '5611502001A',
    '5611512200A',
    'Z0004946A',
    '5612602102A',
    'Z0009524A',
    '5612604900A',
    '5612605000A',
    '5611510801A',
    '5611514600A',
    'T26164BA',
    'T36744BA',
    'T35584CA',
    'T909088A',
    'Z0016091A'
    ]
    # Create a boolean mask to filter rows where 'Part_No' is not in the exclusion list
    mask = ~QCProd.index.isin(Part_to_exclude)
    QCProd= QCProd[mask]

    static_columns=['Weight (g)','Beginning Stock','QC-Prod-(Pcs)','QC-ST-(Pcs)','QC-Prod-Over(Pcs)','Ending ST (Chk)','Sales-Over(Pcs)','QC-Ending-(Pcs)','SUM-QC-NG']
    all_columns = Date_columns.columns.tolist()+static_columns
    QCThai=QCProd[all_columns]
    QCThai.rename(columns={'Weight (g)':'นน.(กรัม)','Beginning Stock':'ยอดยกมา','QC-Prod-(Pcs)':'ยอดผลิต','QC-ST-(Pcs)':'ยอดผลิต+ยกมา','QC-Prod-Over(Pcs)':'ยอดผลิตเกิน','Ending ST (Chk)':'WK-Stock ตรวจนับ','MC-Over(Pcs)':'ยอดเบิกเกิน','QC-Ending-(Pcs)':'Stock คงเหลือ','SUM-QC-NG':'ยอดงานเสีย'},inplace=True)
    QCProd
    ####################
    QCThai['Stock คงเหลือ'] = QCThai['Stock คงเหลือ'].apply(lambda x: 0 if x < 0 else x)
    ###################
    QCThai
    ################ Export To Excel #####################################
    location=r'C:\Users\utaie\Production-2024\Production-App\\'
    QCThai.to_excel(location+Minput+'-'+Process+'Rev-Export.xlsx')
    ######################################################################
    ######################
    SUM_Beg=QCProd['Beginning Stock'].sum()
    SUM_Prod=QCProd['QC-Prod-(Pcs)'].sum()
    SUM_Ending=QCProd['QC-Ending-(Pcs)'].sum()

    SUM_NG_Kgs=((QCProd['SUM-QC-NG']*QCProd['Weight (g)'])/1000).sum()

    data = {
    'Description': ['Total Beginning Stock','Total QC-Production','Total QC-Ending','Total NG-Weight',],
    'Quantity': [ f"{SUM_Beg:,.0f} ",f"{SUM_Prod:,.0f} ", f"{SUM_Ending:,.0f} ", f"{SUM_NG_Kgs:,.0f} "],
    'Unit': ['Pcs','Pcs','Pcs','Kgs']
    }

    # Create a DataFrame
    df = pd.DataFrame(data)
    df.set_index('Description', inplace=True)
    # Display the table
    st.write("Summary Table")
    st.table(df)
    st.write('---')