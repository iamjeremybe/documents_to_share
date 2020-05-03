#!/usr/bin/python3
# # Greater MSP 2019 Data Visualization Challenge
# 
# ## Team:
# * David Treiber - david.treiber@gmail.com
# * Jeremy Brezovan - jeremy.brezovan@gmail.com
# * Megan Menth - meganlmenth@gmail.com

import numpy as np
import pandas as pd
import math

file_path = '/home/jeremy/Documents/Greater MSP Challenge'
file_name = file_path + '/Historical Data - Prior Dashboard - Color Guides/2015-2019 Dashboard Trends_all.xlsx'

dfs = pd.read_excel(file_name, sheet_name=None, header=None)

kpis = dfs['Key Indicators'].loc[[3,4]].dropna(axis='columns')


# ### Create a dictionary for Key Performance Indicators
# Capture the Dashboard Category and Indicator Title of each indicator. We'll use this later to set a value indicating whether or not an indicator is a KPI.
kpi_dict = {}
for col in kpis.columns:
    if col == 0:
        continue
    category = kpis[col].iloc[0]
    indicator = kpis[col].iloc[1]
    if category in kpi_dict:
        kpi_dict[category].append(indicator)
    else:
        kpi_dict[category] = [indicator]


# ## Functions

# ### The year field often has an appended text string, indicating the actual date ranges involved.
# Preserve the existing Year to use as a descriptor, and create a new Year variable that is just the 4-digit year.
def cleanup_year(string,**kwargs):
    desc_or_year = kwargs.get('desc_or_year','year')
    if desc_or_year == 'year':
        try:
            return int(float(string.split()[0]))
# split() will fail when the value is a lone year (ex: 2015 vs. "2015\n(using 13-14 data)")
# because the lone year is interpreted as an int. 
# Rather than convert that to a string, just return it.
        except:
# string may be empty/null! You can convert that null to a float.
# However, you CANNOT convert that null float to an int!
            str_as_float = float(string)
            if np.isnan(str_as_float):
                return string
            else:
                return int(str_as_float)
    else:
        return " ".join(str(string).split())


# ### Create a Key Indicator variable, and set it to 1 for key indicators.
def set_key_indc(category,indicator,**kwargs):
    my_kpi_dict = kwargs.get('kpi_dict',kpi_dict)
    if category in my_kpi_dict:
        if indicator in kpi_dict[category]:
            return 1
    return 0  

# ### Replicate the Alteryx logic for deciding Data Type:
# ```
# if abs([Measure Value]) <=1 then 'Percent'
# elseif Contains([Indicator Title], 'Cost') then 'Dollar'
# elseif Contains([Indicator Title], 'Wage') then 'Dollar'
# elseif Contains([Indicator Title], 'Income') then 'Dollar'
# elseif Contains([Indicator Title], 'Price') then 'Dollar'
# elseif [Dashboard Category] = 'Business Vitality' 
# 	AND !(Contains([Indicator Title], 'Patent')
# 	or Contains([Indicator Title], 'Establishment'))
# 	then 'Dollar'
# else 'Numeric'
# endif 
# ```
def calculate_data_type(series):
    if series['Value'] <=1:
        return 'Percent'
    elif any(x in series['Indicator'].lower() for x in ['cost','wage','income','price']):
        return 'Dollar'
    elif (series['Category'] == 'Business Vitality')          and not any(x in series['Indicator'].lower() for x in ['patent','establishment']):
        return 'Dollar'
    else:
        return "Numeric"


# ### Figure out the data type for each indicator.
# * Build a small dataframe of unique combos of Category + Indicator, and the max Value for each combo. (I checked--max() ignores NaNs.)
# * Convert that dataframe into a dictionary containing Indicator/Data Type pairs.
def build_data_type_df(df):
    this_category = df['Category'].unique()[0]
    df_cat = df.loc[df['Category'] == this_category]
# There is something funky about these two sheets--the .groupby() doesn't behave the same way.
# I found that if I drop category, it works better.
    if this_category in ['Environment','Livability']:
        max_df = pd.DataFrame()
        for this_indicator in df_cat['Indicator'].unique():
            indicator_data_df = df_cat.loc[df['Indicator'] == this_indicator,['Indicator','Value']]
            append_df = indicator_data_df.loc[:,['Indicator','Value']].groupby(['Indicator']).max().reset_index()
            max_df = max_df.append(append_df,ignore_index=True)
# Add back the dropped Category column, so calculate_data_type() works.
        max_df['Category'] = [this_category] * max_df.shape[0]
    else:
        max_df = df.loc[:,['Category','Indicator','Value']].groupby(['Category','Indicator']).max().reset_index()
    max_df['Value'] = max_df['Value'].apply(lambda x: abs(x))
    max_df['Data_Type'] = max_df.apply(lambda row: calculate_data_type(row),axis='columns')

# Prep the dataframe so it can be easily transformed into a dictionary
    max_df.drop(columns=['Category','Value'],inplace=True)

    data_type_df = max_df.set_index(['Indicator']).to_dict('index')
# I couldn't find a parameter for .to_dict() that gave me exactly what I wanted--
# namely, a dict where key is Indicator, value is Data Type.
# So build a new dict that has this format, and extract what we need from data_type_df.
    return_dict = {}
    for this_key in data_type_df:
        return_dict[this_key] = data_type_df[this_key]['Data_Type']
    return return_dict


# ### Clean up Values.
# Null values are okay. We'd prefer to work with numeric data for the rest--no strings, please.
def cleanup_values(x):
    if str(x) == 'no data':
        return np.nan
    elif math.isnan(x):
        return np.nan
    else:
        return float(x)


# ### Create a formatted value, based on data type
# 
# This calculated Formatted Value spared us from having to figure out how to display each individual indicator in Tableau.
# * Dollar:
#      * leading \$ sign
#      * commas between thousands
#      * padding to 2 decimal places if underlying data contains decimal
# 
# * Numeric:
#     * commas between thousands
#     * padding to 2 decimal places if underlying data contains decimal
# 
# * Percent:
#     * padding to 2 decimal places
#     * trailing % sign
# 
# * KEEP negative (-) signs 
def calculate_formatted_value(series):
# Capture nulls and empty strings straight away and return those values with no formatting.
    if pd.isnull(series['Value']) or series['Value'] == '':
        return series['Value']
    elif series['Data_Type'] == 'Percent':
        try:
            return "{:.2%}".format(series['Value'])
        except:
            print("This is supposed to be a Percent, but it's not:",series['Value'])
            return series['Value']

    elif series['Data_Type'] == 'Dollar':
        try:
            if str(abs(series['Value'])).isdigit():
                return "${:,}".format(series['Value'])
            else:
                return "${:,.2f}".format(series['Value'])
        except:
            print("This is supposed to be Dollar value, but it's not:",series['Value'])
            return series['Value']

    elif series['Data_Type'] == 'Numeric':
        try:
            if str(abs(series['Value'])).isdigit():
                return "{:,}".format(series['Value'])
            else:
                return "{:,.2f}".format(series['Value'])
        except:
            print("This is supposed to be Numeric, but it's not:",series['Value'])
            return series['Value']
    else:
# Thankfully didn't see any instances of an unknown Data Type, but we'll need a catch-all.
        print("UNKNOWN value type", series)
    return series['Value']


# ### Reshape a sheet of the workbook.
# This is where the magic happens. Most of the other functions above are called from this one.
def reshape_indicator_sheet(df):
    out_df_columns = ['Category','Indicator','Metro','Year_Desc','Year',
                      'Value','Formatted_Value','Rank','Rank_Order','Key_Indicator','Data_Type']
    out_df = pd.DataFrame(columns=out_df_columns)

# Capture the index of the RANK row.
# This will help us split the sheet into its values vs. rank halves.
    rank_index = np.where(df[0] == 'RANK')[0][0]
    indicators_sheet = df.iloc[0:rank_index].copy()
    rank_sheet = df.iloc[rank_index:].copy()

# Loop through each city, snag its values and rank info, drop them into the correct columns
    for metro in indicators_sheet[0][3:]:
        city_df = pd.DataFrame(columns=out_df_columns)
        city_category = indicators_sheet.loc[[0]].values[0][0].strip()

# This will fill 'Category' with nulls, but it gives the dataframe the correct length.
# Then, fill it with the actual category name.
        city_df['Category'] = indicators_sheet.loc[[0]].values[0][1:]
        city_df['Category'].fillna(value=city_category,inplace=True)

# Indicator has some values populated, so we can use 'pad' to copy them forward to null rows.
        city_df['Indicator'] = indicators_sheet.loc[[1]].values[0][1:]
        city_df['Indicator'].fillna(method='pad',inplace=True)
        city_df['Indicator'] = city_df['Indicator'].apply(lambda x: x.strip())
        
# If we find Indicator in the KPI dictionary under this Category, this is a Key Indicator
        city_df['Key_Indicator'] = city_df.apply(lambda row: set_key_indc(row['Category'],row['Indicator']),
                                                 axis='columns')

        city_df['Year_Desc'] = indicators_sheet.iloc[2].values[1:]
        city_df['Year_Desc'] = city_df['Year_Desc'].apply(lambda x: cleanup_year(x,desc_or_year='desc'))
        city_df['Year'] = city_df['Year_Desc'].apply(lambda x: cleanup_year(x))

        indc_metro_value = np.where(indicators_sheet[0] == metro)[0][0]
        city_df['Metro'] = [metro] * city_df.shape[0]
        city_values = indicators_sheet.iloc[indc_metro_value].values[1:]
        city_df['Value'] = ([cleanup_values(x) for x in city_values])

# Find the index for our current Metro in the rank half of the sheet, and get Rank-related values.
        rank_metro_value = np.where(rank_sheet[0] == metro)[0][0]
        city_df['Rank'] = rank_sheet.iloc[rank_metro_value].values[1:]
        city_df['Rank_Order'] = rank_sheet.iloc[[0]].values[0][1:]
        city_df['Rank_Order'].fillna(method='pad',inplace=True)

# Some sheets have null columns at the end, recognizable in our dataframe because Year is null.
# It's safe to drop these rows.
        city_df = city_df.loc[(city_df['Year'].astype('float') > 0)]

# We're done with this city. Append its stats to the big spreadsheet.
        out_df = out_df.append(city_df,ignore_index=True)

    data_type_dict = build_data_type_df(out_df)
    out_df.loc[:,'Data_Type'] = out_df['Indicator'].apply(lambda x: data_type_dict[x])
    out_df.loc[:,'Formatted_Value'] = out_df.apply(lambda row: calculate_formatted_value(row),axis='columns')
    return out_df


# ## Restructure the data for each sheet. Concatenate the restructured output from all of the sheets (minus the Key Indicator sheet).
# 
# ### Each row of the data should have the following columns:
# 1. Category
# 2. Indicator
# 3. Metro area
# 4. Year description (as taken from sheet)
# 5. Year (integer value only)
# 6. Value
# 7. Formatted Value (based on Data Type)
# 8. Rank
# 9. Rank Order
# 10. Key Indicator
# 11. Data Type
output_df = pd.DataFrame()
for sheet_key in dfs.keys():
    if sheet_key == 'Key Indicators':
        continue
    output_df = output_df.append(reshape_indicator_sheet(dfs[sheet_key]),ignore_index=True)


# ## Write a CSV file.
output_df.to_csv(file_path + '/greater_msp_data.csv',index=False)

