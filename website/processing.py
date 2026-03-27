import pandas as pd
import os
from openpyxl import Workbook
from pathlib import Path
from dotenv import load_dotenv
from supabase import create_client
from collections import defaultdict, Counter

env_path = Path(__file__).parent / '.env'
load_dotenv(dotenv_path=env_path)

url = os.environ.get('SUPABASE_URL')
key = os.environ.get('SUPABASE_KEY')

supabase = create_client(url, key)

if not url:
    print(f"DEBUG: Attempted to load .env from: {env_path}")
    print("DEBUG: URL is still missing")

def save_weekly_to_db(table_name, date_range, count):
    try:
        if len(date_range) == 1:
            if table_name == 'daily_percentage_docked':
                final_count = float(count[0])
            else: 
                final_count = int(float(count[0]))
                
            data = {
                "date": date_range,
                "count": final_count
            }
            supabase.table(table_name).upsert(data, on_conflict=["date"]).execute()
            print(f"Successfully saved {len(date_range)} date to {table_name}")

        else:
            final_count = []
            if table_name == 'daily_percentage_docked':

                for i in count:
                    final_count.append(float(i))

            else:
                for i in count:
                    final_count.append(int(float(i)))

            for k, j in enumerate(date_range):
                data = {
                    "date": j,
                    "count": final_count[k]
                }
                supabase.table(table_name).upsert(data, on_conflict=["date"]).execute()
            print(f"Successfully saved to {len(date_range)} data to {table_name}")

    except Exception as e:
        print(f"Database Error: {e}")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def main(input_file):
    raw_data = pd.read_excel(input_file, engine = 'openpyxl')

    df = clean_Data(raw_data)
    df = get_bay_alphabet(df)

    date, day, date_range, date_period, year_period = get_Dates(df)
    
    commando_name_list, st_name_list = get_NameList()

    commando_pivot_table, commando_table, commando_Row, commando_daily_total = data_Commandos(df, commando_name_list)
    print(f'Commandos Daily Total: {commando_daily_total}')

    st_pivot_table, st_table, st_Row, st_daily_total = data_ST(df, st_name_list)
    print(f'STs Daily Total: {st_daily_total}')

    sq_daily_total, sq_pivot_table, sq_pivot_table_rows = SQ_Flight_Count(df)

    tr_daily_total, tr_pivot_table, tr_pivot_table_rows = TR_Flight_Count(df)

    oal_daily_total, oal_pivot_table, oal_pivot_table_rows = OAL_Flight_Count(df)

    sqtr_daily_total, sqtroal_daily_total, grand_daily_total = sum_Flight_Count(sq_daily_total, tr_daily_total, oal_daily_total) 

    percentage_daily_total = get_Percentage(commando_daily_total, sqtr_daily_total)
    print(f'Percentage Daily Total (Commandos): {percentage_daily_total}')

    get_JSON(date_range, percentage_daily_total, commando_daily_total, st_daily_total)

    bay_coverage = bay_aggregate(commando_Row)
    #output_file = new_File(date_period)

    date = pd.DataFrame([date])
    day = pd.DataFrame([day])
    commando_pivot_table = pd.DataFrame(commando_pivot_table)
    commando_table = pd.DataFrame(commando_table)
    commando_Row = pd.DataFrame(commando_Row)
    commando_daily_total = pd.DataFrame([commando_daily_total])
    
    st_pivot_table = pd.DataFrame(st_pivot_table)
    st_table = pd.DataFrame(st_table)
    st_Row = pd.DataFrame(st_Row)
    st_daily_total = pd.DataFrame([st_daily_total])

    tr_pivot_table = pd.DataFrame(tr_pivot_table)
    oal_pivot_table = pd.DataFrame(oal_pivot_table)
    sq_daily_total = pd.DataFrame([sq_daily_total])
    tr_daily_total = pd.DataFrame([tr_daily_total])
    oal_daily_total = pd.DataFrame([oal_daily_total])
    sqtroal_daily_total = pd.DataFrame([sqtroal_daily_total])
    sqtr_daily_total = pd.DataFrame([sqtr_daily_total])
    grand_daily_total = pd.DataFrame([grand_daily_total])
    percentage_daily_total = pd.DataFrame([percentage_daily_total])

    bay_coverage = pd.DataFrame(bay_coverage)

    results = {
            'df': df,
            'date': date,
            'day': day,
            'commando_pivot_table': commando_pivot_table,
            'commando_table': commando_table,
            'commando_Row': commando_Row,
            'commando_daily_total': commando_daily_total,
            'st_pivot_table': st_pivot_table,
            'st_table': st_table,
            'st_Row': st_Row,
            'st_daily_total': st_daily_total,
            'sq_pivot_table': sq_pivot_table,
            'tr_pivot_table': tr_pivot_table,
            'oal_pivot_table': oal_pivot_table,
            'sq_pivot_table_rows': sq_pivot_table_rows,
            'tr_pivot_table_rows': tr_pivot_table_rows,
            'oal_pivot_table_rows': oal_pivot_table_rows,
            'sq_daily_total':sq_daily_total,
            'tr_daily_total': tr_daily_total,
            'oal_daily_total': oal_daily_total,
            'sqtroal_daily_total': sqtroal_daily_total,
            'sqtr_daily_total': sqtr_daily_total,
            'percentage_daily_total': percentage_daily_total,
            'year_of_report': year_period,
            'date_of_report': date_period,
            'bay_coverage': bay_coverage
        }
    
    return results

def clean_Data(raw_df):
    df = raw_df.dropna(how = 'all') 
    df['Flight No.'] = df['IATA_Airline_Code'].astype(str) + df['flightNo'].astype('Int64').astype(str)
    df['flightDate'] = pd.to_datetime(df['flightDate'])
    df['flightDate'] = df['flightDate'].dt.date
    df = df.sort_values(by = 'flightDate')

    return df

def get_bay_alphabet(df):
    df['bay_alphabet'] = df['bay'].str.extract(r'^([A-Za-z]+)')

    return df

def get_Dates(df):
    date = []
    day = []
    date_range = []
    unique_year = []
    
    for d in df['flightDate']:
        year = d.year
        if year not in unique_year:
            unique_year.append(year)
            print(f'year of data found: {year}')
    
    year_period = max(unique_year)

    for i in df.index:
        date_found = df.loc[i, 'flightDate']
        if date_found not in date:
            date.append(date_found)
    
    date.sort()

    for j in date:
        index = j.weekday()
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        day.append(days[index])
        

    if len(date) == 1:
        date_range.append(date[0].strftime('%Y-%m-%d'))
        date_period = date_range[0]
        print(f'Date of report is/are: {date_period}')

    else:
        for j in date:
            date_in_text = j.strftime('%Y-%m-%d')
            date_range.append(date_in_text)
            
        start_date = date[0]
        end_date = date[-1]

        if start_date.month == end_date.month:
            date_period = f"{start_date.day}-{end_date.day} {start_date.strftime('%b')}"
        else:
            date_period = f"{start_date.day} {start_date.strftime('%b')}-{end_date.day} {end_date.strftime('%b')}"
        print(f'Dates of report are from {date_period}')

    print(f'Dates send to supabase will be: {date_range}')

    return date, day, date_range, date_period, year_period


def get_NameList():
    commando_data = supabase.table('commando_namelist').select('name').execute()
    commando_name_list = [row['name'] for row in commando_data.data]

    st_data = supabase.table("st_namelist").select("name").execute()
    st_name_list = [row['name'] for row in st_data.data]


    return commando_name_list, st_name_list
def data_Commandos(df, commando_name_list):

    df1 = df.drop([
        'MAB_BY',
        'MAB',
        'Pax step docking by',
        'Pax step docking',
        'PLB Docking By', 'PLB Docking End'],
        axis = 1)
    
    df1 = df1.loc[
        (df1['flightDirection'] == 'ARRIVAL') & 
        (df1['body_type'] == 'WB')
        ]

    Date = []
    Staff_ID_Name = []
    Row = []
    
    Row.append(df1.columns.tolist()) #Append header

    for i in df1.index:
        date = df1.loc[i, 'flightDate']
        row_values = df1.loc[i,df1.columns != 'Date'].values

        for name in commando_name_list:
            if name in row_values:
                Date.append(date)
                Staff_ID_Name.append(name)
                Row.append(row_values.tolist())

    table ={                                
        'Date': list(Date),
        'Staff ID & Name': list(Staff_ID_Name)
    }

    table = pd.DataFrame(table)
    Row = pd.DataFrame(Row)

    pivot_table = pd.pivot_table(
        table.assign(val=1),
        index = 'Staff ID & Name',
        values = 'val',
        columns= 'Date',
        aggfunc = 'count',
        fill_value = '',
        margins = True,
        margins_name = 'Total'
    )
    pivot_table = pivot_table.sort_values(by='Total', ascending = False)
    daily_total = pivot_table.loc['Total'].tolist()

    total_row = pivot_table.loc[['Total']]
    pivot_table_no_total = pivot_table.drop('Total')

    pivot_table_sorted = pivot_table_no_total.sort_values(by = 'Total', ascending = False)

    pivot_table_final = pd.concat([pivot_table_sorted, total_row])

    #Push to Supabase
    date_cols = [col for col in pivot_table_final.columns if col != 'Total']

    supabase_payload = []

    for staff_info, row in pivot_table_final.drop('Total', errors = 'ignore').iterrows():
        for dt in date_cols:
            count_val = row[dt]

            if count_val != '' and pd.notna(count_val) and count_val != 0:
                record = {
                    'staff_id_name' : staff_info,
                    'count': int(count_val),
                    'date': str(dt)
                }
                supabase_payload.append(record)
    if supabase_payload:
        print(f'Uploading {len(supabase_payload)} records...')
        supabase.table('commando_pivot_data').insert(supabase_payload).execute()
        print(f'Sucessfully saved {len(supabase_payload)} data to commando_pivot_data')

    return pivot_table_final, table, Row, daily_total

def data_ST(df2, st_name_list):
    Date = []
    Staff_ID_Name = []
    Row = []
    
    Row.append(df2.columns.tolist()) #Append Header

    for i in df2.index:
        date = df2.loc[i, 'flightDate']
        row_values = df2.loc[i,df2.columns != 'Date'].values

        for name in st_name_list:
            if name in row_values:
                Date.append(date)
                Staff_ID_Name.append(name)
                Row.append(row_values.tolist())

    table ={                                
        'Date': list(Date),
        'Staff ID & Name': list(Staff_ID_Name)
    }

    table = pd.DataFrame(table)
    Row = pd.DataFrame(Row)

    pivot_table = pd.pivot_table(
        table.assign(val=1),
        index = 'Staff ID & Name',
        values = 'val',
        columns= 'Date',
        aggfunc = 'count',
        fill_value = '',
        margins = True,
        margins_name = 'Total'
    )

    pivot_table = pivot_table.sort_values(by='Total', ascending = False)
    daily_total = pivot_table.loc['Total'].tolist()

    total_row = pivot_table.loc[['Total']]
    pivot_table_no_total = pivot_table.drop('Total')

    pivot_table_sorted = pivot_table_no_total.sort_values(by = 'Total', ascending = False)

    pivot_table_final = pd.concat([pivot_table_sorted, total_row])

    date_cols = [col for col in pivot_table_final.columns if col != 'Total']

    supabase_payload = []

    for staff_info, row in pivot_table_final.drop('Total', errors = 'ignore').iterrows():
        for dt in date_cols:
            count_val = row[dt]

            if count_val != '' and pd.notna(count_val) and count_val != 0:
                record = {
                    'staff_id_name' : staff_info,
                    'count': int(count_val),
                    'date': str(dt)
                }
                supabase_payload.append(record)
    supabase.table('st_pivot_data').insert(supabase_payload).execute()
    print('Sucessfully saved data to st_pivot_data')
    return pivot_table_final, table, Row, daily_total

def SQ_Flight_Count(df):
    df_sq_flight_count = df[(df['flightDirection'] == 'ARRIVAL') & 
                    (df['IATA_Airline_Code'] == 'SQ') & 
                    (df['body_type'] == 'WB') & 
                    (df['flightnature'] == 'PAX')
                    ]

    pivot_table = pd.pivot_table(
        df_sq_flight_count.assign(val=1),
        index = 'Flight No.',
        columns = 'flightDate', 
        values = 'val',
        aggfunc = 'count',
        fill_value = '',
        margins = True,
        margins_name = 'Total'
        )
    num_rows = pivot_table.shape[0]
    daily_total = pivot_table.loc['Total'].tolist()
    print('Daily Total for SQ Flights are: ' + ', '.join(str(j) for j in daily_total))

    return daily_total, pivot_table, num_rows

def TR_Flight_Count(df):
    df_tr_flight_count = df[(df['flightDirection'] == 'ARRIVAL') & 
                    (df['IATA_Airline_Code'] == 'TR') & 
                    (df['body_type'] == 'WB') & 
                    (df['flightnature'] == 'PAX')
                    ]

    pivot_table = pd.pivot_table(
        df_tr_flight_count.assign(val=1),
        index = 'Flight No.',
        columns = 'flightDate', 
        values = 'val',
        aggfunc = 'count',
        fill_value = '',
        margins = True,
        margins_name = 'Total'
        )
    num_rows = pivot_table.shape[0]
    daily_total = pivot_table.loc['Total'].tolist()
    print('Daily Total for TR Flights are: ' + ', '.join(str(j) for j in daily_total))

    return daily_total, pivot_table, num_rows

def OAL_Flight_Count(df):
    df_oal_flight_count = df[(df['flightDirection'] == 'ARRIVAL') &
                          (df['IATA_Airline_Code'] != 'SQ') &
                          (df['IATA_Airline_Code'] != 'TR') &
                          (df['body_type'] == 'WB') &
                          (df['flightnature'] == 'PAX')
    ]
    pivot_table = pd.pivot_table(
        df_oal_flight_count.assign(val=1),
        index = 'Flight No.',
        values = 'val',
        columns = 'flightDate',
        aggfunc = 'count',
        fill_value = '',
        margins = True,
        margins_name = 'Total'
    )
    num_rows = pivot_table.shape[0]
    daily_total = pivot_table.loc['Total'].tolist()
    print('Daily Total for OAL Flights are: ' + ', '.join(str(j) for j in daily_total))

    return daily_total, pivot_table, num_rows

def sum_Flight_Count(sq, tr, oal):
    sqtr = []
    sqtroal = []

    for i in range(len(sq)):
        sqtr_total = int(sq[i])+int(tr[i])
        sqtr.append(sqtr_total)
    

    for j in range(len(sq)):
        sqtroal_total = int(sq[j]) + int(tr[j]) + int(oal[j])
        sqtroal.append(sqtroal_total)

    print('Total for SQ & TR Flights are: ' + ', '.join(str(j) for j in sqtr))
    print(f'Total for SQ, TR & OAL Flights are: ' + ', '.join(str(j) for j in sqtroal))

    grand_total = sqtroal[-1]

    return sqtr, sqtroal, grand_total

def get_Percentage(commando, sqtr):
    percentage_list = []
    for i in range(len(commando)):
        percentage = float(round(commando[i]/sqtr[i],4))
        percentage_list.append(percentage)

    return percentage_list

def get_JSON(date_range, percentage, commando, st):
    commando_pop = commando[:-1]
    st_pop = st[:-1]
    percentage_pop = percentage[:-1]

    save_weekly_to_db('daily_commando', date_range, commando_pop)
    save_weekly_to_db('daily_st', date_range, st_pop)
    save_weekly_to_db('daily_percentage_docked', date_range, percentage_pop)

def bay_aggregate(commando_Row):
    commando_Row.columns = commando_Row.iloc[0]
    commando_Row = commando_Row[1:].reset_index(drop=True)

    commando_Row.columns = [str(c).strip() for c in commando_Row.columns]

    headers = commando_Row.columns.tolist()
    date_idx = headers.index('flightDate')
    bay_idx  = headers.index('bay_alphabet')

    agg = defaultdict(Counter)

    for row in commando_Row.itertuples(index=False):
        date = row[date_idx]
        bay  = row[bay_idx]
        if bay:
            agg[date][bay] += 1

    records_to_insert = []
    for date, bay_counts in agg.items():

        record = {"date": str(date)}  
        for bay_letter in ['A', 'B', 'C', 'D', 'E', 'F']:
            record[bay_letter] = bay_counts.get(bay_letter, 0)
        records_to_insert.append(record)
    
    supabase.table("bay_alphabet_data").upsert(records_to_insert, on_conflict=["date"]).execute()
    print(f"Successfully saved {len(records_to_insert)} data to bay_alphabet_data")

    return records_to_insert
