import streamlit as st
import pandas as pd
import plotly.express as px
from sklearn.linear_model import LinearRegression
from prophet import Prophet
from supabase import create_client
from datetime import timedelta
import plotly
import time
import random
import matplotlib.pyplot as plt
print(plotly.__version__)

import streamlit as st

st.set_page_config(
    layout="wide", 
    page_title='PLB Report Dashboard - Commandos & STs'
)

st.markdown(
    """
    <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
            margin-top: -2rem; /* Optional: Pulls the content even higher */
        }
    </style>
    """,
    unsafe_allow_html=True
)

with st.container(border=True):
    st.title(f"PLB Report Dashboard - Commandos & STs")

st.write("### Page Selector")
page = st.selectbox("Select Page", ["📈Commandos Dashboard", "📈STs Dashboard", "👥Commandos Staff Tracking", "% Commando's Docking Percentage"], label_visibility="collapsed")

url = st.secrets["connections"]["supabase"]["url"]
key = st.secrets["connections"]["supabase"]["key"]

supabase = create_client(url, key)


def fetch_table(table_name, columns=None):
    try:
        all_data = []
        page_size = 1000
        start = 0
        
        while True:
            res = supabase.table(table_name).select("*").range(start, start + page_size - 1).execute()
            
            if not res.data:
                break
                
            all_data.extend(res.data)
            
            if len(res.data) < page_size:
                break

            start += page_size
        
        if not all_data:
            return pd.DataFrame()
        
        df = pd.DataFrame(all_data)
        
        if columns:
            df = df[[c for c in columns if c in df.columns]]
            
        return df
        
    except Exception as e:
        st.error(f"Error fetching {table_name}: {e}")
        return pd.DataFrame()

columns_needed_id = ['date', 'count']
columns_needed_staff_id_name = ['staff_id_name','date', 'count']
columns_needed_bay_alphbaet = ['date', 'A','B','C','D','E','F']

daily_commando_df = fetch_table('daily_commando', columns_needed_id)
daily_st_df = fetch_table('daily_st', columns_needed_id)
daily_percentage_docked_df = fetch_table('daily_percentage_docked', columns_needed_id)

commando_pivot_data_df = fetch_table('commando_pivot_data', columns_needed_staff_id_name)
st_pivot_data_df = fetch_table('st_pivot_data', columns_needed_staff_id_name)

bay_alphabet_df = fetch_table('bay_alphabet_data', columns_needed_bay_alphbaet)


###################
###Data Handling###
###################

dataframes = {
    "commando": daily_commando_df,
    "st": daily_st_df,
    "docked": daily_percentage_docked_df,
    "bay": bay_alphabet_df
}

for name, df in dataframes.items():

    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    
    df = df.groupby('date', as_index=False).sum(numeric_only=True)
    
    df = df.sort_values(by='date', ascending=False)
    
    if name == "commando": daily_commando_df = df
    elif name == "st": daily_st_df = df
    elif name == "docked": daily_percentage_docked_df = df
    elif name == "bay": bay_alphabet_df = df

daily_commando_df['date'] = pd.to_datetime(daily_commando_df['date'])
daily_st_df['date'] = pd.to_datetime(daily_st_df['date'])
daily_percentage_docked_df['date'] = pd.to_datetime(daily_percentage_docked_df['date'])
st_pivot_data_df['date'] = pd.to_datetime(st_pivot_data_df['date']) 
bay_alphabet_df['date'] = pd.to_datetime(bay_alphabet_df['date'])


commando_pivot_data_df['date'] = pd.to_datetime(commando_pivot_data_df['date']).dt.strftime('%Y-%m-%d')
commando_pivot_data_df['staff_id_name'] = commando_pivot_data_df['staff_id_name'].astype(str).str.strip()

commando_pivot_df = commando_pivot_data_df.pivot_table(
    index='date',
    columns='staff_id_name',
    values='count',
    aggfunc='sum',
    fill_value=0
)

###########
#Functions#
###########

#Date Filter
min_date = min(df['date'].min() for df in [daily_commando_df, daily_st_df, daily_percentage_docked_df, bay_alphabet_df] if not df.empty)
max_date = max(df['date'].max() for df in [daily_commando_df, daily_st_df, daily_percentage_docked_df, bay_alphabet_df] if not df.empty)

try:
    if 'selected_date_range' not in st.session_state:
        st.session_state.selected_date_range = [min_date, max_date]
    if st.button('Reset for full Date Range'):
        st.session_state.selected_date_range = [min_date, max_date]
    date_range = st.date_input(
        "Select Date Range",
        key='selected_date_range'
    )
except Exception:
    print('No data available')

def filter_df_by_date(df, date_col = 'date'):
    if df.empty:
        return df
    start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
    return df[(df[date_col] >= start) & (df[date_col] <= end)]

daily_commando_df = filter_df_by_date(daily_commando_df)
daily_st_df = filter_df_by_date(daily_st_df)
daily_percentage_docked_df = filter_df_by_date(daily_percentage_docked_df)


#Staff Filter
all_staff = pd.concat([commando_pivot_data_df['staff_id_name'], st_pivot_data_df['staff_id_name']]).dropna().unique()
#selected_staff = st.sidebar.multiselect('Select Staff', options=all_staff)


########################
###DISPLAY ON WEBSITE###
########################
st.markdown(
    """
    <style>
    /* Hide Streamlit menu and footer */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    /* Remove container padding */
    .css-18e3th9 {padding-top: 0rem;}
    </style>
    """,
    unsafe_allow_html=True
)

try:

    if daily_commando_df.empty or daily_st_df.empty or daily_percentage_docked_df.empty:
        st.warning("No data available.")
    else:
        weekday_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']

        daily_commando_df['weekday'] = daily_commando_df['date'].dt.day_name()
        daily_commando_df['weekday'] = pd.Categorical(daily_commando_df['weekday'],
                                                    categories=weekday_order, ordered=True)
        
        daily_st_df['weekday'] = daily_st_df['date'].dt.day_name()
        daily_st_df['weekday'] = pd.Categorical(daily_st_df['weekday'],
                                                categories=weekday_order, ordered=True)
        
        daily_percentage_docked_df['weekday'] = daily_percentage_docked_df['date'].dt.day_name()
        daily_percentage_docked_df['weekday'] = pd.Categorical(daily_percentage_docked_df['weekday'],
                                                                categories=weekday_order, ordered=True)
        
        bay_alphabet_df['weekday'] = bay_alphabet_df['date'].dt.day_name()
        bay_alphabet_df['weekday'] = pd.Categorical(bay_alphabet_df['weekday'],
                                                    categories=weekday_order, ordered=True)

        ##############################
        ###COMMANDOS DASHBOARD PAGE###
        ##############################
        if page == "📈Commandos Dashboard":
            #Commandos Daily Docking Data & Graph
            st.header('Commandos - Daily & Weekly Docking Graph')
            
            col1, col2 = st.columns([1,4])

            #Commandos Table
            daily_commando_df_to_show = daily_commando_df.drop(columns=['weekday'])
            with col1:
                st.dataframe(daily_commando_df_to_show,
                            use_container_width=False,
                            height = 550,
                            width = 500,
                            column_config={ 
                                'date': st.column_config.DateColumn(
                                    'Date',
                                    format='YYYY-MM-DD',
                                    width = 80
                                    ),
                                'count': st.column_config.NumberColumn(
                                    'Count',
                                    width = 30
                                    )
                                }
                            )

            #Commandos Daily Graph
            with col2:
                daily_commando_df['date_ordinal'] = daily_commando_df['date'].map(pd.Timestamp.toordinal)
                X = daily_commando_df['date_ordinal'].values.reshape(-1,1)
                y = daily_commando_df['count'].values
                reg = LinearRegression().fit(X, y)
                trend = reg.predict(X)

                fig_daily_commando = px.line(
                    daily_commando_df.sort_values('date'),
                    x = 'date',
                    y = 'count',
                    title = 'Commandos Daily Total Docking Trend',
                    markers = True
                )

                daily_commando_df['Category'] = daily_commando_df['count'].apply(categorize)

                fig_daily_commando.add_scatter(
                    x = daily_commando_df['date'],
                    y = trend,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_daily_commando.update_layout(
                    height=630,
                    xaxis_title = 'Date',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13, color="black"),
                    yaxis_title_font=dict(size=13, color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_daily_commando.data):
                    trace.name = f"<b>{trace.name}</b>"

                
                st.plotly_chart(fig_daily_commando, use_container_width=True)

            col3, col4 = st.columns([1,4])
            with col3:
            #Commandos Weekly Graph
                weekly_commando_df = (
                    daily_commando_df
                    .set_index('date')
                    .resample('W-MON', label='left', closed='left')['count']
                    .sum()
                    .reset_index()
                )
                weekly_commando_df['week_start'] = weekly_commando_df['date']
                weekly_commando_df['week_end'] = weekly_commando_df['week_start'] + pd.Timedelta(days=6)

                weekly_commando_df['week_range'] = weekly_commando_df['week_start'].dt.strftime('%d %b') + " – " + \
                                                weekly_commando_df['week_end'].dt.strftime('%d %b')
                
                weekly_percentage_docked_df_to_show = weekly_commando_df.sort_values('date', ascending=False)
                weekly_commando_df_to_show = weekly_percentage_docked_df_to_show.drop(columns=['date', 'week_start','week_end'])
                weekly_commando_df_to_show = weekly_percentage_docked_df_to_show[['week_range','count']]
        
                st.dataframe(weekly_commando_df_to_show,
                use_container_width=False,
                height = 550,
                width = 500,
                column_config={
                    'week_range': st.column_config.Column(
                        'Week Range',
                        width = 80
                    ),
                    'count': st.column_config.NumberColumn(
                        'Count',
                        width = 30
                        )
                    }
                )

            with col4:
                Xa = weekly_commando_df.index.values.reshape(-1, 1)
                ya = weekly_commando_df['count'].values
                rega = LinearRegression().fit(Xa, ya)
                y_pred = rega.predict(Xa)

                fig_weekly_commando = px.line(
                    weekly_commando_df,
                    x = 'week_range',
                    y = 'count',
                    markers = True,
                    title = 'Commandos Weekly Docking Trend'
                )
                fig_weekly_commando.add_scatter(
                    x = weekly_commando_df['week_range'],
                    y = y_pred,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_weekly_commando.update_layout(
                    height=630,
                    xaxis_title = 'Week Range',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_weekly_commando.data):
                    trace.name = f"<b>{trace.name}</b>"
                st.plotly_chart(fig_weekly_commando, use_container_width=True)
                
            col4, col5 = st.columns([4,2])

            with col4:
                daily_commando_df['week_start'] = daily_commando_df['date'] - pd.to_timedelta(daily_commando_df['date'].dt.weekday, unit='d')

                weekly_by_day_commando = daily_commando_df.pivot_table(
                    index='week_start',
                    columns='weekday',
                    values='count',
                    aggfunc='sum'
                )

                weekly_by_day_reset_commando = weekly_by_day_commando.reset_index()

                weekdays_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
                weekly_by_day = weekly_by_day_commando.reindex(columns=weekdays_order)
                weekly_by_day_reset_commando['week_range'] = weekly_by_day_reset_commando['week_start'].dt.strftime('%d %b') + ' – ' + \
                                    (weekly_by_day_reset_commando['week_start'] + pd.Timedelta(days=6)).dt.strftime('%d %b')

                fig_trend_by_day_of_week_commando = px.line(
                    weekly_by_day_reset_commando,
                    x='week_range',
                    y=weekdays_order,
                    labels={'value':'Count', 'week_range':'Week'},
                    markers=True,
                    title='Weekly Trend by Weekday '
                )
                fig_trend_by_day_of_week_commando.update_layout(
                    height = 630,
                    xaxis_title = 'Day of Week',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_trend_by_day_of_week_commando.data):
                    trace.name = f"<b>{trace.name}</b>"

                st.plotly_chart(fig_trend_by_day_of_week_commando, use_container_width=True)


            with col5:
                weekday_counts = daily_commando_df.groupby('weekday')['count'].sum().reset_index()

                fig_weekday_commando_df = px.bar(
                    daily_commando_df,
                    x = 'weekday',
                    y = 'count',
                    color = 'weekday',
                    title = 'Docking by Day of Week'
                )

                fig_weekday_commando_df.update_layout(
                    height = 630,
                    xaxis_title = 'Day of Week',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_weekday_commando_df.data):
                    trace.name = f"<b>{trace.name}</b>"
                
                st.plotly_chart(fig_weekday_commando_df, use_container_width=True)


            bay_columns = ['A', 'B', 'C', 'D', 'E', 'F']

            start, end = pd.to_datetime(date_range[0]).normalize(), pd.to_datetime(date_range[1]).normalize()
            bay_alphabet_df_filtered = bay_alphabet_df[(bay_alphabet_df['date'] >= start) & (bay_alphabet_df['date'] <= end)]

            bay_long = bay_alphabet_df_filtered.melt(
                id_vars='date',
                value_vars=bay_columns,
                var_name='Bay',
                value_name='Count'
            )

            bay_long['Count'] = pd.to_numeric(bay_long['Count'], errors='coerce').fillna(0)

            fig_bay = px.line(
                bay_long,
                x='date',
                y='Count',
                color='Bay',
                markers=True,
                title='Bay Coverage Docking Trend'
            )

            fig_bay.update_layout(
                    height = 630,
                    xaxis_title = 'Week',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
            for i, trace in enumerate(fig_bay.data):
                trace.name = f"<b>{trace.name}</b>"
                

            st.plotly_chart(fig_bay, use_container_width=True)


        ########################
        ###STs DASHBOARD PAGE###
        ########################
        if page == '📈STs Dashboard':
            st.header('STs/23 - Daily & Weekly Usage Graph')

            col7, col8 = st.columns([1,4])

            daily_st_df_to_show = daily_st_df.drop(columns=['weekday'])

            with col7:
                st.dataframe(daily_st_df_to_show,
                            use_container_width=False,
                            height = 550,
                            width = 500,
                            column_config={
                                'date': st.column_config.DateColumn(
                                    'Date',
                                    format='YYYY-MM-DD',
                                    width = 80
                                    ),
                                'count': st.column_config.NumberColumn(
                                    'Count',
                                    width = 30
                                )
                                }
                            )

            with col8:
                daily_st_df['date_ordinal'] = daily_st_df['date'].map(pd.Timestamp.toordinal)
                X1 = daily_st_df['date_ordinal'].values.reshape(-1,1)
                y1 = daily_st_df['count'].values
                reg1 = LinearRegression().fit(X1, y1)
                trend1 = reg1.predict(X1)

                fig_daily_st = px.line(
                    daily_st_df.sort_values('date'),
                    x = 'date',
                    y = 'count',
                    color_discrete_sequence=['orange'],
                    title = 'STs Daily Total Usage Trend',
                    markers = True
                )

                daily_st_df['Category'] = daily_st_df['count'].apply(categorize)

                fig_daily_st.add_scatter(
                    x = daily_st_df['date'],
                    y = trend1,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_daily_st.update_layout(
                    height=630,
                    xaxis_title = 'Date',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13, color='black'),
                    yaxis_title_font=dict(size=13, color='black'),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_daily_st.data):
                    trace.name = f"<b>{trace.name}</b>"

                st.plotly_chart(fig_daily_st, use_container_width=True)


            col9, col10 = st.columns([1,4])
            with col9:
                weekly_st_df =(
                    daily_st_df
                    .set_index('date')
                    .resample('W-MON', label='left', closed='left')['count']
                    .sum()
                    .reset_index()
                )
                weekly_st_df['week_start'] = weekly_st_df['date']
                weekly_st_df['week_end'] = weekly_st_df['week_start'] + pd.Timedelta(days=6)

                weekly_st_df['week_range'] = weekly_st_df['week_start'].dt.strftime('%d %b') + " – " + \
                                            weekly_st_df['week_end'].dt.strftime('%d %b')
                
                weekly_st_df_to_show = weekly_st_df.sort_values('date', ascending=False)
                weekly_st_df_to_show = weekly_st_df_to_show.drop(columns=['date', 'week_start','week_end'])
                weekly_st_df_to_show = weekly_st_df_to_show[['week_range','count']]

                st.dataframe(weekly_st_df_to_show,
                            use_container_width=False,
                            height = 550,
                            width = 500,
                            column_config={
                                'week_range': st.column_config.Column(
                                    'Week Range',
                                    width = 80
                                ),
                                'count': st.column_config.NumberColumn(
                                    'Count',
                                    width = 30
                                )
                            })
            with col10:
                X1a = weekly_st_df.index.values.reshape(-1,1)
                y1b = weekly_st_df['count'].values
                rega = LinearRegression().fit(X1a, y1b)
                y_pred = rega.predict(X1a)

                fig_weekly_st = px.line(
                    weekly_st_df,
                    x = 'week_range',
                    y = 'count',
                    markers = True,
                    color_discrete_sequence=['orange'],
                    title = 'STs Weekly Usage Trend'
                )
                fig_weekly_st.add_scatter(
                    x = weekly_st_df['week_range'],
                    y = y_pred,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_weekly_st.update_layout(
                    height=630,
                    xaxis_title = 'Week Range',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13, color='black'),
                    yaxis_title_font=dict(size=13, color='black'),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_weekly_st.data):
                    trace.name = f"<b>{trace.name}</b>"
                st.plotly_chart(fig_weekly_st, use_container_width=True)

            col11, col12 = st.columns([4,2])


            with col11:
                daily_st_df['week_start'] = daily_st_df['date'] - pd.to_timedelta(daily_st_df['date'].dt.weekday, unit='d')

                weekly_by_day_st = daily_st_df.pivot_table(
                    index='week_start',
                    columns='weekday',
                    values='count',
                    aggfunc='sum'
                )

                weekly_by_day_st_reset = weekly_by_day_st.reset_index()
                
                weekdays_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']

                weekly_by_day_st = weekly_by_day_st.reindex(columns=weekdays_order)
                weekly_by_day_st_reset['week_range'] = weekly_by_day_st_reset['week_start'].dt.strftime('%d %b') + ' – ' + \
                                    (weekly_by_day_st_reset['week_start'] + pd.Timedelta(days=6)).dt.strftime('%d %b')

                fig_trend_by_day_of_week_st = px.line(
                    weekly_by_day_st_reset,
                    x='week_range',
                    y=weekdays_order,
                    labels={'value':'Count', 'week_range':'Week'},
                    markers=True,
                    title='Weekly Trend by Weekday'
                )
                fig_trend_by_day_of_week_st.update_layout(
                    height = 630,
                    xaxis_title = 'Day of Week',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_trend_by_day_of_week_st.data):
                    trace.name = f"<b>{trace.name}</b>"

                st.plotly_chart(fig_trend_by_day_of_week_st, use_container_width=True)


            with col12:
                weekday_counts = daily_commando_df.groupby('weekday')['count'].sum().reset_index()

                fig_weekday_st_df = px.bar(
                    daily_st_df,
                    x = 'weekday',
                    y = 'count',
                    color = 'weekday',
                    title = 'Usage by Day of Week'
                )
                fig_weekday_st_df.update_layout(
                    height = 630,
                    xaxis_title= 'Day of Week',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_weekday_st_df.data):
                    trace.name = f"<b>{trace.name}</b>"
                
                st.plotly_chart(fig_weekday_st_df, use_container_width=True)


        #############################################
        ###COMMANDOS STAFF TRACKING DASHBOARD PAGE###
        #############################################

        elif page == '👥Commandos Staff Tracking':
            st.header('Commandos Staff Tracking - Number of Flights Docked')
            staff_list = commando_pivot_df.columns.tolist()

                # 3. Match the Streamlit filter to the same string format
            start_str = pd.to_datetime(date_range[0]).strftime('%Y-%m-%d')
            end_str = pd.to_datetime(date_range[1]).strftime('%Y-%m-%d')
        
            # Filter using the strings
            plot_df_filtered_cst = commando_pivot_df.loc[start_str : end_str]
            # Filter using the matching 'date' objects


            fig_commando_staff_tracking = px.line(
                plot_df_filtered_cst,
                x=plot_df_filtered_cst.index,
                y=list(plot_df_filtered_cst.columns),
                markers=True,
                labels={'value':'Count', 'date':'Date'},
                title='Staff Docking Trend'
            )
            fig_commando_staff_tracking.update_layout(
                height=700,
                xaxis_title='Date',
                yaxis_title='Count',
                xaxis_title_font=dict(size=13, color="black"),
                yaxis_title_font=dict(size=13, color="black"),
                legend=dict(
                    font=dict(size=10, color='black')
                )
            )
            st.plotly_chart(fig_commando_staff_tracking, use_container_width=True, key='staff_tracking_chart')

            latest_date = plot_df_filtered_cst.index.max()

            last_docked_info = {}

            for staff in plot_df_filtered_cst.columns:
                non_zero = plot_df_filtered_cst[staff][plot_df_filtered_cst[staff] > 0]
                
                if not non_zero.empty:
                    last_timestamp = non_zero.last_valid_index()
                    last_docked_info[staff] = {"Last Docked": last_timestamp}
                else:
                    last_docked_info[staff] = {"Last Docked": pd.NaT}

            last_docked_df = pd.DataFrame(last_docked_info).T

            last_docked_df["Days Since Last Docked"] = last_docked_df["Last Docked"].apply(
                lambda x: (pd.Timestamp(latest_date) - pd.Timestamp(x)).days if pd.notnull(x) else None
            )

            last_docked_df = last_docked_df.sort_values("Days Since Last Docked", ascending=False)

            last_docked_df["Last Docked"] = last_docked_df["Last Docked"].apply(
                lambda x: pd.Timestamp(x).date() if pd.notnull(x) else None
            )

            st.subheader("Last Docked per Staff")
            
            st.dataframe(
                last_docked_df,
                use_container_width=True,
                key='last_docked_table',
                column_config={
                    "Days Since Last Dock": st.column_config.NumberColumn(
                        "Days Since Last Dock",
                        help="Days since last docking"
                    )
                }
            )
        
        elif page == "% Commando's Docking Percentage":
            st.header('Commandos - Daily & Weeky Percentage Docking')
            daily_percentage_docked_df['count'] = daily_percentage_docked_df['count'].round(4)
            daily_percentage_docked_df['count'] = daily_percentage_docked_df['count']*100

            col13, col14 = st.columns([1,4])

            daily_percentage_docked_df_to_show = daily_percentage_docked_df.drop(columns=['weekday'])
            with col13:
                st.dataframe(daily_percentage_docked_df_to_show,
                            use_container_width=False,
                            height=550,
                            width=500,
                            column_config={
                                'date': st.column_config.DateColumn(
                                    'Date',
                                    format='YYYY-MM-DD',
                                    width=80
                                ),
                            'count': st.column_config.NumberColumn(
                                'Count',
                                width=30,
                                format="%.2f%%",
                                )
                            }
                        )
                
            with col14:
                daily_percentage_docked_df['date_ordinal'] = daily_percentage_docked_df['date'].map(pd.Timestamp.toordinal)
                X2 = daily_percentage_docked_df['date_ordinal'].values.reshape(-1,1)
                y2 = daily_percentage_docked_df['count'].values
                reg2 = LinearRegression().fit(X2, y2)
                trend2 = reg2.predict(X2)

                fig_daily_percentage_docked = px.line(
                    daily_percentage_docked_df.sort_values('date'),
                    x = 'date',
                    y = 'count',
                    title = 'Commandos Daily Percentage Docked',
                    color_discrete_sequence=['green'],
                    markers = True
                )
                fig_daily_percentage_docked.add_scatter(
                    x = daily_percentage_docked_df['date'],
                    y = trend2,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_daily_percentage_docked.update_layout(
                    height=630,
                    xaxis_title = 'Date',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13, color="black"),
                    yaxis_title_font=dict(size=13, color="black"),
                    yaxis_tickformat='.2f',
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )

                for i, trace in enumerate(fig_daily_percentage_docked.data):
                    trace.name = f"<b>{trace.name}</b>"
                st.plotly_chart(fig_daily_percentage_docked, use_container_width=True)

            col15, col16 = st.columns([1,4])
            with col15:
            #Commandos Weekly Graph
                weekly_percentage_docked_df = (
                    daily_percentage_docked_df
                    .set_index('date')
                    .resample('W-MON', label='left', closed='left')['count']
                    .mean()
                    .reset_index()
                )

                weekly_percentage_docked_df['count'] = weekly_percentage_docked_df['count'].round(2)
                weekly_percentage_docked_df['week_start'] = weekly_percentage_docked_df['date']
                weekly_percentage_docked_df['week_end'] = weekly_percentage_docked_df['week_start'] + pd.Timedelta(days=6)

                weekly_percentage_docked_df['week_range'] = weekly_percentage_docked_df['week_start'].dt.strftime('%d %b') + " – " + \
                                                weekly_percentage_docked_df['week_end'].dt.strftime('%d %b')
                
                weekly_percentage_docked_df_to_show = weekly_percentage_docked_df.sort_values('date', ascending=False)
                weekly_percentage_docked_df_to_show = weekly_percentage_docked_df_to_show.drop(columns=['date', 'week_start','week_end'])
                weekly_percentage_docked_df_to_show = weekly_percentage_docked_df_to_show[['week_range','count']]
                
        
                st.dataframe(weekly_percentage_docked_df_to_show,
                use_container_width=False,
                height = 550,
                width = 500,
                column_config={
                    'week_range': st.column_config.Column(
                        'Week Range',
                        width = 80
                    ),
                    'count': st.column_config.NumberColumn(
                        'Count',
                        width = 30,
                        format="%.2f%%"
                        )
                    }
                )

            with col16:
                Xa = weekly_percentage_docked_df.index.values.reshape(-1, 1)
                ya = weekly_percentage_docked_df['count'].values
                rega = LinearRegression().fit(Xa, ya)
                y_pred = rega.predict(Xa)

                fig_weekly_percentage_docked = px.line(
                    weekly_percentage_docked_df,
                    x = 'week_range',
                    y = 'count',
                    markers = True,
                    color_discrete_sequence = ['green'],
                    title = 'Commandos Weekly Percentage Docking Trend'
                )
                fig_weekly_percentage_docked.add_scatter(
                    x = weekly_percentage_docked_df['week_range'],
                    y = y_pred,
                    mode = 'lines',
                    name = 'Trendline',
                    line = dict(color = 'black', dash = 'dot')
                )
                fig_weekly_percentage_docked.update_layout(
                    height=630,
                    xaxis_title = 'Week Range',
                    yaxis_title = 'Count',
                    xaxis_title_font=dict(size=13,  color="black"),
                    yaxis_title_font=dict(size=13,  color="black"),
                    legend=dict(
                        font=dict(
                            size=10,
                            color='black'
                        )
                    )
                )
                for i, trace in enumerate(fig_weekly_percentage_docked.data):
                    trace.name = f"<b>{trace.name}</b>"
                st.plotly_chart(fig_weekly_percentage_docked, use_container_width=True)

except Exception as e:
    print(e)