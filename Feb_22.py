import openpyxl
import pandas as pd
import plotly.graph_objects as go
import matplotlib
import numpy as np
import streamlit as st
from matplotlib.backends.backend_agg import RendererAgg
matplotlib.use('agg')
from st_btn_select import st_btn_select
import streamlit_authenticator as stauth
_lock = RendererAgg.lock



# -----------------------------------------set page layout-------------------------------------------------------------
st.set_page_config(page_title='iSMM Dashboard',
                    page_icon = ':bar_chart:',
                    layout='wide',
                    initial_sidebar_state='collapsed')

#-----------------------------------------------User Authentication-----------------------------------------------
names = ['wenna', 'Mr.Loh', 'Dennis']
usernames = ['wenna0306@gmail.com', 'booninn.loh@surbanajurong.com', 'dennis.tanjw@surbanajurong.com']
passwords = ['password', 'password', 'password']

hashed_passwords = stauth.hasher(passwords).generate()

authenticator = stauth.authenticate(names,usernames,hashed_passwords,
    'some_cookie_name','some_signature_key',cookie_expiry_days=30)

name, authentication_status = authenticator.login('Login','main')

if authentication_status:
    st.write('Welcome *%s*' % (name))



#==========================================================pd read all files==============================================================

    def fetch_file_faults(filename):
        cols = ['Fault Number', 'Trade', 'Trade Category',
                'Type of Fault', 'Impact', 'Site', 'Building', 'Floor', 'Room', 'Cancel Status', 'Reported Date',
                'Fault Acknowledged Date', 'Responded on Site Date', 'RA Conducted Date',
                'Work Started Date', 'Work Completed Date', 'Attended By', 'Action(s) Taken',
                'Other Trades Required Date', 'Cost Cap Exceed Date',
                'Assistance Requested Date', 'Fault Reference',
                 'Incident Report', 'Remarks']
        parse_dates = ['Reported Date',
                        'Fault Acknowledged Date', 'Responded on Site Date', 'RA Conducted Date',
                        'Work Started Date', 'Work Completed Date',
                        'Other Trades Required Date', 'Cost Cap Exceed Date',
                        'Assistance Requested Date']
        dtype_cols = {'Site': 'str', 'Building': 'str', 'Floor': 'str', 'Room': 'str'}
    
        return pd.read_excel(filename, header=1, usecols=cols, parse_dates=parse_dates, dtype=dtype_cols)


    def fetch_file_schedules(filename):
        cols = ['Schedule ID', 'Trade', 'Trade Category', 'Strategic Partner',
            'Frequency', 'Description', 'Type', 'Scope', 'Site', 'Building',
            'Floor', 'Room', 'Start Date',
            'End Date', 'Work Started Date', 'Work Completed Date',
            'Attended By', 'Portal ID', 'Service Report Completed Date',
            'Total Number of Service Reports',
            'Total Number of Quantity in Service Reports', 'Asset Tag Number',
            'Alert Case']
        parse_dates = ['Start Date', 'End Date', 'Work Started Date', 'Work Completed Date',
                        'Service Report Completed Date']
        dtype_cols = {'Site': 'str', 'Building': 'str', 'Floor': 'str', 'Room': 'str'}
        return pd.read_excel(filename, header=0, usecols=cols, parse_dates=parse_dates, dtype=dtype_cols)



    unvalid_sites = ['Test Site', 'Test Site (Inventory)', 'SIN Capital', 'ALICE @ Mediapolis', 'Tool Box Site 1', 'Tool Box Site 2', 'DFM']

    # valid_sites = ['Connection One', 'Gardens By The Bay (Central)', 'Gardens By The Bay (East)', 'Gardens By The Bay (South)', 'NIE', 'RMG', 
    # 'SIA-ALH/H1/ISQ/SB1/SB2/SB3', 'SIAEC-H4', 'SIAEC-H5', 'SIAEC-H6', 'CEPO', 'EBC Ubi', 'NTU', 'AMKH', 'Razer']
    
    df = fetch_file_faults('Fault.xlsx')   #fault reported and was not cancelled within Feb 22
    df = df[~df.Site.isin(unvalid_sites)]   
    df.columns = df.columns.str.replace(' ', '_')
    
    dfs = fetch_file_schedules('schedules.xlsx')   #all scheduled schedules within Feb 22
    dfs = dfs[~dfs.Site.isin(unvalid_sites)]
    dfs.columns = dfs.columns.str.replace(' ', '_')

    dfa = pd.read_excel('asset.xlsx', header=1)     #all active assets within Feb 22
    dfa = dfa[~dfa.Site.isin(unvalid_sites)]
    dfa.columns = dfa.columns.str.replace(' ', '_')

    dft_CO = pd.read_excel('Transaction_CO.xlsx', header=1)
    dft_CO['Site'] = 'Connection One'

    dft_GBB_CM18 = pd.read_excel('Transaction_CM18.xlsx', header=1)
    dft_GBB_CM18['Site'] = 'GBB-CM18 (Inventory)'

    dft_GBB_EC11 = pd.read_excel('Transaction_EC11.xlsx', header=1)
    dft_GBB_EC11['Site'] = 'GBB-EC11 (Inventory)'

    dft_SIA_ALH = pd.read_excel('Transaction_ALH.xlsx', header=1)
    dft_SIA_ALH['Site'] = 'SIA-ALH (Inventory)'

    dft_SIA_PTB = pd.read_excel('Transaction_PTB.xlsx', header=1)
    dft_SIA_PTB['Site'] = 'SIA-PTB (Inventory)'

    dft_SIA_STC = pd.read_excel('Transaction_STC.xlsx', header=1)
    dft_SIA_STC['Site'] = 'SIA-STC (Inventory)'

    dft_SIAEC_H4 = pd.read_excel('Transaction_H4.xlsx', header=1)
    dft_SIAEC_H4['Site'] = 'SIAEC-H4 (Inventory)'

    dft_SIAEC_H5 = pd.read_excel('Transaction_H5.xlsx', header=1)
    dft_SIAEC_H5['Site'] = 'SIAEC-H5 (Inventory)'

    dft_SIAEC_H6 = pd.read_excel('Transaction_H6.xlsx', header=1)
    dft_SIAEC_H6['Site'] = 'SIAEC-H6 (Inventory)'

    dft = pd.concat([dft_CO, dft_GBB_CM18, dft_GBB_EC11, dft_SIA_ALH, dft_SIA_PTB, dft_SIA_STC, dft_SIAEC_H4, dft_SIAEC_H5, dft_SIAEC_H6])
    dft = dft[~dft.Site.isin(unvalid_sites)]
    dft.columns = dft.columns.str.replace(' ', '_')



#=============================================Summary==============================================================

    html_card_title="""
    <div class="card">
        <div class="card-body" style="border-radius: 10px 10px 0px 0px; padding-top: 5px; width: 1000px;
        height: 50px;">
        <h1 class="card-title" style=color:#90b134; font-family:Georgia; text-align: left; padding: 0px 0;>iSMM USAGE MONITORING</h1>
        </div>
    </div>
    """

    report_month = df.Reported_Date[0].month_name()
    
    st.markdown(html_card_title, unsafe_allow_html=True)
    st.markdown('##')
    st.markdown('##')
    st.markdown(f"""
        Welcome to this Analysis App for the month ***{report_month} 2022***, get more details from :point_right: [iSMM](https://ismm.sg/ce/login)\n
        - Check detail analysis for [RMG](https://share.streamlit.io/wenna0306/feb_22_rmg/main/Feb_22.py)
        - Check detail analysis for [NIE](https://share.streamlit.io/wenna0306/feb_22_nie/main/Feb_22.py)
        - Check detail analysis for [Gendens By The Bay (south)](https://share.streamlit.io/wenna0306/feb_22_gbb_s/main/Feb_22.py)
        - Check detail analysis for [SIA-ALH/H1/ISQ/SB1/SB2/SB3](https://share.streamlit.io/wenna0306/feb_22_s/main/Feb_22.py)
        - Check detail analysis for [Connection One](https://share.streamlit.io/wenna0306/feb_22_co/main/Feb_22.py)
        """)
    st.markdown('##')


    total_fault = df.shape[0]
    total_schedule = dfs.shape[0]
    total_asset = dfa.shape[0]
    total_transaction = dft.shape[0]
    fault_incident = int(df['Incident_Report'].sum())

    column01_fault, column02_schedule, column03_asset, column04_transaction, column05_incident = st.columns(5)
    with column01_fault, _lock:
        st.markdown(f"<h6 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:bottom; color: #5f9e8f;'>Total Faults</h6>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:top; color: #5f9e8f;'>{total_fault}</h2>", unsafe_allow_html=True)

    with column02_schedule, _lock:
        st.markdown(f"<h6 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:bottom; color: #5f9e8f;'>Total Schedules</h6>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:top; color: #5f9e8f;'>{total_schedule}</h2>", unsafe_allow_html=True)

    with column03_asset, _lock:
        st.markdown(f"<h6 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:bottom; color: #5f9e8f;'>Total Assets</h6>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:top; color: #5f9e8f;'>{total_asset}</h2>", unsafe_allow_html=True)

    with column04_transaction, _lock:
        st.markdown(f"<h6 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:bottom; color: #5f9e8f;'>Total Transaction</h6>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:top; color: #5f9e8f;'>{total_transaction}</h2>", unsafe_allow_html=True)

    with column05_incident, _lock:
        st.markdown(f"<h6 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:bottom; color: #5f9e8f;'>Incident Report</h6>", unsafe_allow_html=True)
        st.markdown(f"<h2 style='background-color:#0e1117; width:120px; height:20px; text-align: left; vertical-align:top; color: red;'>{fault_incident}</h2>", unsafe_allow_html=True)

#=================================================color & width & opacity========================================================================

        # all chart
        titlefontcolor = '#5f9e8f' #for x and y axis

        # pie chart for outstanding and tier1
        # colorpieoutstanding = ['#116a8c', '#4c9085', '#50a747', '#59656d', '#06c2ac', '#137e6d', '#929906', '#ff9408']
        # colorpierecoveredtier1 = ['#116a8c', '#4c9085', '#50a747', '#59656d', '#06c2ac', '#137e6d', '#929906', '#ff9408']

        # all barchart include stack bar chart and individual barchart and linechart
        plot_bgcolor = 'rgba(0,0,0,0)'
        gridwidth = 0.001
        gridcolor = 'black'

        textcolor = 'white'
        textsize = 15
        textfamily = 'sana serif'

        titlecolor = '#5f9e8f'  # for chart title
        titlesize = 20
        titlefamily = 'sana serif'

        # stack barchart
        # colorstackbarpass = '#116a8c'
        # colorstackbarfail = '#ffdb58'

        # individual barchart color, barline color& width, bar opacity
        markercolor = '#5f9e8f'
        markerlinecolor = '#5f9e8f'
        markerlinewidth = 1
        opacity01 = 1
        opacity02 = 1
        opacity03 = 1

        # x&y axis width and color
        linewidth_xy_axis = 1
        linecolor_xy_axis = '#59656d'

        # # linechart
        # linecolor = '#96ae8d'
        # linewidth = 2

#==========================================================iSMM Faults=============================================================
    st.markdown('##')
    st.markdown("""<hr style="height:5px;border:none;color:#333;background-color:#333;" /> """, unsafe_allow_html=True)
    st.markdown('##')
    
    ser_fault = df.groupby(['Site']).Fault_Number.count().sort_values(ascending=True)
    x_fault= ser_fault.index
    y_fault = ser_fault.values

    ser_schedule = dfs.groupby(['Site']).Schedule_ID.count().sort_values(ascending=True)
    x_schedule = ser_schedule.index
    y_schedule = ser_schedule.values

    ser_asset = dfa.groupby(['Site']).Asset_Tag_Number.count().sort_values(ascending=True)
    x_asset = ser_asset.index
    y_asset = ser_asset.values

    dft.insert(0, 'identifier', range(len(dft))) #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ser_transaction = dft.groupby(['Site']).identifier.count().sort_values(ascending=True)
    x_transaction = ser_transaction.index
    y_transaction = ser_transaction.values

    fig_fault, fig_schedule = st.columns(2)

    with fig_fault, _lock:
        fig_fault = go.Figure(data=[go.Bar(x=y_fault, y=x_fault, orientation='h', text=y_fault,
                                        textfont=dict(family=textfamily, size=textsize, color=textcolor),
                                        textposition='auto', textangle=0)])
        fig_fault.update_xaxes(title_text="Site", title_font_color=titlefontcolor, showgrid=True, gridwidth=gridwidth,
                            gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis, tickangle=-45)
        fig_fault.update_yaxes(title_text='Number of Fault', title_font_color=titlefontcolor, showgrid=False,
                            gridwidth=gridwidth, gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis)
        fig_fault.update_traces(marker_color=markercolor, marker_line_color=markerlinecolor,
                            marker_line_width=markerlinewidth, opacity=opacity03)
        fig_fault.update_layout(title='Number of Fault vs Site', titlefont=dict(family=titlefamily, size=titlesize, color=titlecolor), plot_bgcolor=plot_bgcolor)
        st.plotly_chart(fig_fault, use_container_width=True)
    
    with fig_schedule, _lock:
        fig_schedule = go.Figure(data=[go.Bar(x=y_schedule, y=x_schedule, orientation='h', text=y_schedule,
                                        textfont=dict(family=textfamily, size=textsize, color=textcolor),
                                        textposition='auto', textangle=0)])
        fig_schedule.update_xaxes(title_text="Site", title_font_color=titlefontcolor, showgrid=True, gridwidth=gridwidth,
                            gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis, tickangle=-45)
        fig_schedule.update_yaxes(title_text='Number of Schedule', title_font_color=titlefontcolor, showgrid=False,
                            gridwidth=gridwidth, gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis)
        fig_schedule.update_traces(marker_color=markercolor, marker_line_color=markerlinecolor,
                            marker_line_width=markerlinewidth, opacity=opacity03)
        fig_schedule.update_layout(title='Number of Schedule vs Site', titlefont=dict(family=titlefamily, size=titlesize, color=titlecolor), plot_bgcolor=plot_bgcolor)
        st.plotly_chart(fig_schedule, use_container_width=True)

    fig_asset, fig_transaction = st.columns(2)

    with fig_asset, _lock:
        fig_asset = go.Figure(data=[go.Bar(x=y_asset, y=x_asset, orientation='h', text=y_asset,
                                        textfont=dict(family=textfamily, size=textsize, color=textcolor),
                                        textposition='auto', textangle=0)])
        fig_asset.update_xaxes(title_text="Site", title_font_color=titlefontcolor, showgrid=True, gridwidth=gridwidth,
                            gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis, tickangle=-45)
        fig_asset.update_yaxes(title_text='Number of Asset', title_font_color=titlefontcolor, showgrid=False,
                            gridwidth=gridwidth, gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis)
        fig_asset.update_traces(marker_color=markercolor, marker_line_color=markerlinecolor,
                            marker_line_width=markerlinewidth, opacity=opacity03)
        fig_asset.update_layout(title='Number of Asset vs Site', titlefont=dict(family=titlefamily, size=titlesize, color=titlecolor), plot_bgcolor=plot_bgcolor)
        st.plotly_chart(fig_asset, use_container_width=True)

    with fig_transaction, _lock:
        fig_transaction = go.Figure(data=[go.Bar(x=y_transaction, y=x_transaction, orientation='h', text=y_transaction,
                                        textfont=dict(family=textfamily, size=textsize, color=textcolor),
                                        textposition='auto', textangle=0)])
        fig_transaction.update_xaxes(title_text="Site", title_font_color=titlefontcolor, showgrid=True, gridwidth=gridwidth,
                            gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis, tickangle=-45)
        fig_transaction.update_yaxes(title_text='Number of Transaction', title_font_color=titlefontcolor, showgrid=False,
                            gridwidth=gridwidth, gridcolor=gridcolor, showline=True, linewidth=linewidth_xy_axis,
                            linecolor=linecolor_xy_axis)
        fig_transaction.update_traces(marker_color=markercolor, marker_line_color=markerlinecolor,
                            marker_line_width=markerlinewidth, opacity=opacity03)
        fig_transaction.update_layout(title='Number of Transaction vs Site', titlefont=dict(family=titlefamily, size=titlesize, color=titlecolor), plot_bgcolor=plot_bgcolor)
        st.plotly_chart(fig_transaction, use_container_width=True)





elif authentication_status == False:
    st.error('Username/password is incorrect')
elif authentication_status == None:
    st.warning('Please enter your username and password')



hide_menu_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """
st.markdown(hide_menu_style, unsafe_allow_html=True)