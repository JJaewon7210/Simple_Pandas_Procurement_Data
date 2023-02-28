import pandas as pd
import openpyxl
import re
import datetime

input_file_path = 'data/input/특정품목 조달 내역_2022.csv'
output_file_path = 'data/output/특정품목 조달 내역_2022.xlsx'
data = pd.read_csv(input_file_path, encoding='cp949', dtype=str )

# choose columns for analyze
data = data[['계약(납품요구)일자','계약(납품요구)번호','수요기관구분','수요기관지역명','물품분류번호', \
    '품명', '세부물품분류번호', '물품식별번호','세부품명', '품목', '단가', '수량', '단위', '금액']]

# change the data type
for name in ['세부물품분류번호','물품식별번호','단가','수량','금액']:
    data[name] = data[name].apply(lambda x: int(x.replace(',', '')))

# group the data with the same '계약(납품요구)번호'
grouped_data = data.groupby('계약(납품요구)번호')

# preprocessed data
preprocess_data = pd.DataFrame()

# In each group, find the '냉난방기' in the column '세부품명'
# Loop through each group
for group_name, group_data in grouped_data:
    
    # Check there is kW information
    kW_filtered = group_data[group_data['품목'].str.contains('kW')]
    if len(kW_filtered) == 0: continue
    
    # Check it is valid information
    kW_filtered = kW_filtered[kW_filtered['수량'] > 0]
    if len(kW_filtered) == 0: continue
    
    # Get the total kW information 
    pattern_cool = r'냉방(\d+\.\d+|\d+)' 
    pattern_heat = r'난방(\d+\.\d+|\d+)'
    kW_cool = 0
    kW_heat = 0
    
    for unit, words in zip(kW_filtered['수량'].to_list(),kW_filtered['품목'].to_list()):
        match_cool = re.search(pattern_cool,words)
        match_heat = re.search(pattern_heat,words)
        
        if match_cool: kW_cool += float(match_cool.group(1))*unit
        if match_heat: kW_heat += float(match_heat.group(1))*unit

    # Calculate total kW device cost
    device_cost = kW_filtered['금액'].sum()
    
    # Calculate total construction cost
    total_cost = group_data['금액'].sum()

    # Get the location
    area = group_data['수요기관지역명'].to_list()[0].split(' ')[0]
    
    # Get the organization type
    organization = group_data['수요기관구분'].to_list()[0].split(' ')[0]
    
    # Get timestamp
    timestamp = datetime.datetime.strptime(group_data['계약(납품요구)일자'].to_list()[0],"%Y%m%d")
    
    # Define a new row of data as a dictionary
    new_data = {'수요기관지역명': area, '수요기관구분':organization, '장비금액': device_cost, '계약금액': total_cost, '냉방용량': kW_cool, '난방용량': kW_heat, '날짜':timestamp}
    new_data = pd.DataFrame([new_data])
    
    # Concatenate the new DataFrame to the original empty DataFrame
    preprocess_data = pd.concat([preprocess_data, new_data], ignore_index=True)

preprocess_data.to_excel(output_file_path)