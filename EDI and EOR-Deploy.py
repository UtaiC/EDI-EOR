import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image
import plotly.graph_objs as go
import plotly.graph_objects as go
import locale

##################################
Logo=Image.open('SIM-LOGO-02.jpg')
st.image(Logo,width=720)
#################################
#########################################################
def formatted_display0(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.0f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
######################################################
def formatted_display(label, value, unit):
    formatted_value = "<span style='color:yellow'>{:,.2f}</span>".format(value)  # Format value with comma separator and apply green color
    display_text = f"{formatted_value} {unit}"  # Combine formatted value and unit
    st.write(label, display_text, unsafe_allow_html=True)
######################################################
StartWeek = st.sidebar.selectbox('Input-Week',['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20'] )
@st.cache_data 
def load_dataframes(sheet_names):
    url = "https://docs.google.com/spreadsheets/d/1pbzO4YI-TkW3AO6yssJgHO9F3FwWb9Rs/export?format=xlsx"
    dataframes = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(url, header=7, engine='openpyxl', sheet_name=sheet_name)
        dataframes[sheet_name] = df
    return dataframes

# Load dataframes for the specified range of weeks
StartWeek = int(StartWeek)
EndWeek = int(StartWeek)
all_sheet_names = [str(week) for week in range(StartWeek, EndWeek + 1)]
dataframes = load_dataframes(all_sheet_names)
#########################################
#########################
DataMerges = pd.concat(dataframes.values(), ignore_index=True)
# DataMerges.set_index('Part no.',inplace=True)
DataMerges=DataMerges.fillna(0)
exclued=['Part no.','KOSHIN','HOME EXPERT','ELECTROLUX','TBKK',0,'043061102-02',
'032005102-02',
'039025305-02',
'039047501-02',
'011526802-00',
'011526902-01',
'0320003304-02',
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
'220-00331',
'220-00016-1',
'220-00016-2',
'220-00014',
'220-00015',
'1050B375',
'MD372348'
]
StartWeek=str(StartWeek)
DataMerges=DataMerges[~DataMerges['Part no.'].isin(exclued)]
##################################
if int(StartWeek) <10:
    WK_FC='WK0'+StartWeek
else:
    WK_FC='WK'+StartWeek
##################################
stock=DataMerges
# stock
#######################
stock['St-BL']=stock['TOTAL']-stock[WK_FC]
stock.rename(columns={'ERO.6':'ERO-Pcs'},inplace=True)
stock.rename(columns={'ACT.7':'Sales-Pcs'},inplace=True)
stock['ERO-FC']=stock['ERO-Pcs']-stock[WK_FC]
stock['ERO-FC'] = stock['ERO-Pcs'].where(stock['ERO-Pcs'] > 0, stock['ERO-FC'])
stock['ERO-Pcs'] = stock['ERO-FC'].where(stock['ERO-FC'] > 0, stock['ERO-Pcs'])
stock['ERO-%']=(stock['ERO-FC']/stock[WK_FC].where(stock[WK_FC]!=0))*100
stock['ERO-%']=stock['ERO-%'].fillna(0)
# ############## SUMMARIZE #############
SUMALL=stock[['Part no.',WK_FC,'ERO-Pcs']]
SUMALL.rename(columns={WK_FC:'EDI-'+WK_FC},inplace=True)
SUMALL['GAP']=SUMALL['ERO-Pcs']-SUMALL['EDI-'+WK_FC]
SUMALL.set_index('Part no.',inplace=True)
columns_to_check = ['EDI-'+WK_FC, 'ERO-Pcs', 'GAP']
SUMALL = SUMALL[~(SUMALL[columns_to_check] == 0).all(axis=1)]
st.success(f'Weekly Data Summary on EDI vs ERO Volumes (Part no and Pcs)')
SUMALL.rename(columns={'ERO-Pcs':'WK-ERO'},inplace=True)
# Assuming SUMALL is your DataFrame
columns_to_format = ['EDI-'+WK_FC, 'WK-ERO', 'GAP']
# SUMALL[columns_to_format]=pd.to_numeric(SUMALL[columns_to_format])
SUMALL[columns_to_format] = SUMALL[columns_to_format].applymap(lambda val: f"{float(val):,.0f}")
st.table(SUMALL)
############# Set Table #################
SUMALL['GAP'] = SUMALL['GAP'].str.replace(',', '')
SUMALL['GAP']=pd.to_numeric(SUMALL['GAP'])
SUMALL['WK-ERO'] = SUMALL['WK-ERO'].str.replace(',', '')
SUMALL['WK-ERO']=pd.to_numeric(SUMALL['WK-ERO'])
SUMALL['EDI-'+WK_FC] = SUMALL['EDI-'+WK_FC].str.replace(',', '')
SUMALL['EDI-'+WK_FC]=pd.to_numeric(SUMALL['EDI-'+WK_FC])
sum_FC=SUMALL['EDI-'+WK_FC].sum()
sum_ERO=SUMALL['WK-ERO'].sum()
sum_SALES=stock['Sales-Pcs'].sum()
SUM_MEET_ERO=SUMALL['WK-ERO'].where(SUMALL['GAP']==0).sum()
SUM_NoFC=SUMALL['WK-ERO'].where(SUMALL['EDI-'+WK_FC]==0).sum()
SUM_NO_ERO=SUMALL['EDI-'+WK_FC].where(SUMALL['WK-ERO']==0).sum()
SUM_LESS_ERO=SUMALL['WK-ERO'].where(SUMALL['GAP']<0).sum()
SUM_OVER_ERO=SUMALL['WK-ERO'].where(SUMALL['GAP']>0).sum()
#####################
EROPCT=(sum_ERO/sum_FC)*100
SalesPCT=(sum_SALES/sum_FC)*100
MeetPCT=(SUM_MEET_ERO/sum_FC)*100
NoFC_PCT=(SUM_NoFC/sum_FC)*100
NOPCT=(SUM_NO_ERO/sum_FC)*100
LessPCT=(SUM_LESS_ERO/sum_FC)*100
OverPCT=(SUM_OVER_ERO/sum_FC)*100

data={"Item":['Forecasted EDI Volumes',
              'Ordered and Repeated Order (ERO)',
              "Sales (Stock Availability and Production)",
              'EDI/ERO-Orders Met',
              'Non-EDI Orders',
              'No-ERO (Un-ordered Forecast)',
              'Under-ERO (Forecast Exceeds Order)',
              'Over-EDI (Ordered Exceeds Forecast)'],
              "Volumes (Pcs)":[sum_FC,
                               sum_ERO,
                               sum_SALES,
                               SUM_MEET_ERO,
                               SUM_NoFC,
                               -SUM_NO_ERO,
                               -SUM_LESS_ERO,
                               SUM_OVER_ERO],
              "PCT-%":['100%',
                       EROPCT,
                       SalesPCT,
                       MeetPCT,
                       NoFC_PCT,
                       -NOPCT,
                       -LessPCT,
                       OverPCT]}

data["PCT-%"] = [f"{float(val):.2f}%" if val != '100%' else val for val in data["PCT-%"]]
data["Volumes (Pcs)"] = [f"{float(val):,.0f}" if val != '100' else val for val in data["Volumes (Pcs)"]]
df = pd.DataFrame(data).set_index('Item')
###################### Dispaly EDI/EOR #######################

st.success(f'Weekly Data Summary on EDI vs ERO Volumes (Pcs)')
st.table(df)
week=str(StartWeek)
path=r"C:\Users\utaie\Desktop\Production-2024\EDI-Data\\"
file_name=path+'Sum EDI-'+week+'.xlsx'
# file_name
df.to_excel(file_name)

data = {
    "Item": ['Forecast', 'Order/Repeat', 'Delivered', 'Order/FC-Met', 'NO-FC/Order', 'FC/No-Oder', 'Under-FC/Oder', 'Over-FC/Order'],
    "Volumes (Pcs)": [sum_FC, sum_ERO, sum_SALES, SUM_MEET_ERO, SUM_NoFC, -SUM_NO_ERO, -SUM_LESS_ERO, SUM_OVER_ERO],
    "PCT-%": [100, EROPCT, SalesPCT, MeetPCT, NoFC_PCT, NOPCT, -LessPCT, OverPCT],
    "Colors": ['#8CEC12','#ECD812','#12DBEC','#EC8C12','#EC8C12', '#EC6E35', '#EC4A12', '#EC3312']
}
data['PCT-%'] = [f'({float(pct):.2f}%)' for pct in data['PCT-%']]
locale.setlocale(locale.LC_ALL, '')
formatted_volumes = [f'{locale.format_string("%d", volume, grouping=True)}\n{pct}' for volume, pct in zip(data['Volumes (Pcs)'], data['PCT-%'])]
# Create traces
fig = go.Figure()
fig.add_trace(go.Bar(x=data['Item'], y=data['Volumes (Pcs)'], text=formatted_volumes, textposition='outside', name='Volumes (Pcs)', marker_color=data['Colors']))

# Update layout
fig.update_layout(title=f'Weekly Chart Summary on EDI vs ERO Volumes (Pcs/PCT-%) @-{WK_FC}',
                  xaxis_title='',
                  yaxis_title='Value')

st.plotly_chart(fig)
