import streamlit as st
import pandas as pd
from PIL import Image
from datetime import datetime, timedelta
import datetime
import calendar
##################################
# Css Style #####################
with open('style.css') as modi:
    css = f'<style>{modi.read()} </style>'
    st.markdown(css, unsafe_allow_html=True)
# Banner #################################
banner_image = Image.open('Banner-Prod.jpg')
st.image(banner_image)
## Production Menu ########################################
#####################
Yinput = st.sidebar.selectbox('Input-Year', ['2024'])
######################
# Dictionary mapping month numbers to month names
Process = st.sidebar.selectbox('Process',['Die_casting','Finishing','Finishing_SUB','Sand_Blasting','Machine','QC'] )
#######################################
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
########################## Production 2024 ############################
@st.cache_data 
def load_data_File_A(start_week, end_week):
    file = "https://docs.google.com/spreadsheets/u/1/d/1pbzO4YI-TkW3AO6yssJgHO9F3FwWb9Rs.xlsx"
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
data2024