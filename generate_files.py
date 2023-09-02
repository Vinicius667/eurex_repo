from utils import *
from variables import *
from datetime import datetime, timedelta
import requests
import re
from scipy.stats import norm
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pyhtml2pdf import  converter
import shutil
from plotly.colors import n_colors
from typing import Union, List, Dict

def get_date_idx(list_idxs)->list:
    while(True):
        date_idxs = [0, 1]
        date_idxs = list(map(int,date_idxs))
        date_idxs = [idx for idx in date_idxs if idx in list_idxs]
        if len(date_idxs) < 1:
            print("No valid indexes passed.")
        else:
            return date_idxs
        

def get_heute():
    heute = datetime.today()
    if heute.weekday() > 4:
        #sunday => wd = 6 -> td = +1
        #sartuday => wd = 5 > td +2
        heute += timedelta(days= 7 - heute.weekday())
    print(f'heute = {heute.strftime("%d/%m/%Y")}')
    return heute


def get_overview(auswahl):
    ## Download tradingDates and future_date_col
    global url, headers, tradingDates, future_date_col, overview_df
    if auswahl == 1:
    # STOXX
        url = 'https://www.eurex.com/api/v1/overallstatistics/69660'
        print("STOXX")
    else:
    # DAX
        url = 'https://www.eurex.com/api/v1/overallstatistics/70044'
        print ("DAX")
        
    headers = {'Accept': '*/*'}
    params_ov = {'filtertype': 'overview', 'contracttype': 'M'}
    print('Getting data from EUREX...')
    response = requests.get(url, params=params_ov, headers=headers,timeout = 8)

    if not response.ok:
        raise ValueError("Conection failed.")

    tradingDates = pd.to_datetime(pd.Series(response.json()['header']['tradingDates']),format="%d-%m-%Y %H:%M")

    print("Index     Date")
    for idx,date in enumerate(tradingDates.dt.strftime("%d-%m-%Y")):
        print(f" {idx:<6} {date:^3}")

    #date_idxs = get_date_idx(tradingDates.index)
    params_ov['busdate'] = tradingDates[0].strftime("%Y%m%d")
    response = requests.get(url, params=params_ov, headers=headers,timeout = 8)
    if not response.ok:
            raise ValueError("Conection failed.")

    params_details = {}
    if 'dataRows' in response.json(): 
        overview_df = pd.DataFrame(response.json()['dataRows'])
        overview_df = overview_df[overview_df.contractType == 'M']
        overview_df.sort_values('date',ignore_index=True,inplace=True)
        future_date_col = overview_df.loc[:,'date']
        overview_df.date = pd.to_datetime(overview_df.date,format='%Y%m%d').dt.strftime('%d-%m-%Y')

    else:
        raise ValueError

    params_details['busdate'] = f"{tradingDates[1].strftime('%Y%m%d')}"
    response = requests.get(url, params=params_details, headers=headers,timeout = 8)
    if not response.ok:
        raise ValueError("Conection failed.")



def generate_files(auswahl:int, heute: Union[None, str] = None )->None:
    """
    auswahl: Option defined by the user. See variables.auswahl_dict.
    heute: it should be a date as string using dd/mm/yyyy format. If None, it will be defined by get_heute().
    """
    # First stage
    # Create results folder in case they still were not created
    list_folder_results =  [current_results_path, old_results_path]
    for folder in list_folder_results:
        create_folder(folder)

    # During the generation of the PDF, the CSS file must be in the same directory
    # as the generated html file. So in case the CSS file is still not the in the
    # Results folder, we copy this file to it
    if not os.path.exists(result_css):
        shutil.copyfile(src_css, result_css)

    #auswahl = int(input("DAX = 0; STOXX = 1   :"))
    
    get_overview(auswahl)

    # **tbd: heute shall be a working day (Mo - Fr)**
    expiry = datetime.strptime(future_date_col[0],"%Y%m%d")
    expiry_1 = datetime.strptime(future_date_col[1],"%Y%m%d")

    if not heute:
        heute = get_heute()
    else:
        heute = datetime.strptime(heute,"%d/%m/%Y")
    tage_bis_verfall = (expiry - heute).days +1


    df_StgrtDAX = pd.read_html("https://www.boerse-stuttgart.de/en/",match="L&S DAX")[0]

    # Cntralize diagram
    DAX_LS_e = df_StgrtDAX.loc[df_StgrtDAX['Indices GER'] == "L&S DAX","Price"].values[0] #X4

    print(f"ZentralKurs value was changed: {ZentralKurs} => ",end="")
    if (ZentralKurs - DAX_LS_e) > 500:
        NewZentralKurs = ZentralKurs - 500
    elif (ZentralKurs - DAX_LS_e) < -500:
        NewZentralKurs =  ZentralKurs + 500

    print(f"{NewZentralKurs}")
    Minkurs = NewZentralKurs - (Spannweite/ 2)
    Maxkurs = Minkurs + Spannweite
    Schritte = Spannweite / Schritt

    DetailMin = round(DAX_LS_e - (Spannweite)/4)
    DetailMax = DetailMin + (Spannweite/2)


    if tage_bis_verfall >= 1:
        delta = 0.5
    else:
        delta = 1

    # Second stage
    # Create Basis column that will be used for multiple tables
    Basis = pd.Series((Minkurs + np.arange(int(Schritte) +1)* Schritt)[::-1])

    # Transform series into dataframe to make dataframe creation easier 
    Basis_df = Basis.reset_index().rename(columns={0:"Basis"})[["Basis"]]


    # Ueberhaenge_df, Summery_df are created from a copy of the same Dataframe
    # so they have the same index. It means they could be in one  Dataframe but
    #  I decided to keep them separated as they were in the VBA code.

    Ueberhaenge_df = Basis_df.copy()
    Summery_df = Basis_df.copy()

    params_details = {}
    params_details['filtertype'] = 'detail'
    params_details['contracttype'] = 'M'

    for productdate_idx in [0,1]:
        params_details['productdate'] = future_date_col[productdate_idx]
        for busdate_idx in [0,1]:
            params_details['busdate'] = f"{tradingDates[busdate_idx].strftime('%Y%m%d')}" 
            print(f"busdate_idx = {busdate_idx}   | productdate_idx = {productdate_idx}")
            
            if (busdate_idx == 1) and (productdate_idx == 1):
                # No need to request data for this case so we skip this iteration
                continue
            
            # Request data
            response = requests.get(url, params=params_details, headers=headers, timeout = 10)
            if not response.ok:
                raise ValueError("Conection failed.")
            contractsCall_aux_df =  pd.DataFrame(response.json()['dataRowsCall'])
            contractsPut_aux_df =  pd.DataFrame(response.json()['dataRowsPut'])

            # Create dataframe with the just requested data
            aux_df = Basis_df.copy()

            aux_df = aux_df.merge(
                contractsCall_aux_df[['strike','openInterest']],
                left_on = "Basis",
                right_on = "strike",
                how= "left"
            )

            aux_df = aux_df.merge(
                contractsPut_aux_df[['strike','openInterest']],
                on = "strike",
                how="left",
                suffixes=('_CF', '_PF')
            )

            if productdate_idx == 0:
                Ueberhaenge_df[busdate_idx] = aux_df.openInterest_PF - aux_df.openInterest_CF

                if busdate_idx == 0:
                    Summery_df["openInterest_PF"] = aux_df["openInterest_PF"]
                    Summery_df["openInterest_CF"] = aux_df["openInterest_CF"]
            
            elif productdate_idx == 1:
                if busdate_idx == 0:
                    Ueberhaenge_df["nextContract"] = aux_df.openInterest_PF - aux_df.openInterest_CF
                    Summery_df["nextContract"] = Ueberhaenge_df["nextContract"]


    Ueberhaenge_df.rename(columns = {0 : "Front"},inplace=True)
    Summery_df["heute"] = Ueberhaenge_df["Front"] * 0.2* delta

    Ueberhaenge_df.rename(columns = {1 : "Last"},inplace=True)
    Summery_df["last_day"] = Ueberhaenge_df["Last"] * 0.2* delta


    Ueberhaenge_df["Summe"] = Ueberhaenge_df[["Front", "nextContract"]].sum(axis=1)
    Ueberhaenge_df = Ueberhaenge_df[['Basis', 'Summe',"Last",'Front', "nextContract"]]


    Summery_df["Änderung"] = Summery_df.heute - Summery_df.last_day

    if (delta == 1):
        Summery_df['nextContract'] =  Summery_df['nextContract'] / 2
    
    SummeryDetail_df = Summery_df[(Summery_df.Basis >= DetailMin) & (Summery_df.Basis < (DetailMax + Schritt))]
    # Third stage
    euribor_3m_df = pd.read_html("https://www.euribor-rates.eu/en/current-euribor-rates/2/euribor-rate-3-months/")[0].rename(columns= {0: "Date", 1 : "InterestRate"})
    euribor_row = euribor_3m_df.loc[0]
    print(f"Euribor 3M in {euribor_row['Date']} is {euribor_row['InterestRate']}")

    response_txt = requests.get("https://qontigo.com/index/vdax2m/").text
    data = re.findall(r"window.chart_data_two_week = (.*);", response_txt)[0]
    data = eval(f"np.array({data},float)")[-10:][::-1]
    vda_2m_df = pd.DataFrame(np.asarray(data[:,0], dtype='datetime64[ms]'),columns=["Date"])
    vda_2m_df["Price"] = data[:,1]
    vda_2m_row = vda_2m_df.loc[0]
    print(f"VDA-NEW 2M in {vda_2m_row['Date']} is € {vda_2m_row['Price']}")
    vda_2m_df
    SchrittWeite = 10
    InterestRate = float(euribor_row["InterestRate"].replace("%",""))/100
    VDAX = vda_2m_row["Price"]/100
    VDAX_Laufzeit = 60

    Tage = tage_bis_verfall
    if Tage == 0:
        Tage = 0.5

    Tage_1 = (expiry_1 - heute).days +1
    Kurs_count = int(Spannweite/SchrittWeite)+1 
    HedgeBedarf_kurs=  pd.DataFrame(Maxkurs - np.arange(Kurs_count)* SchrittWeite,columns= ["Basis"])


    Hedge_dimensions = Kurs_count,int(Schritte+1)
    HedgeBedarf_values = np.zeros(Hedge_dimensions)
    HedgeBedarf1_values = np.zeros(Hedge_dimensions)


    for k in range(Hedge_dimensions[1]):
        Basis_value = Basis[k]
        Kontrakte = Ueberhaenge_df.loc[k,"Front"]
        Kontrakte_1 = Ueberhaenge_df.loc[k,"nextContract"]
        for i in range(Hedge_dimensions[0]):
            Kurs = Maxkurs - i * SchrittWeite
            # In python np.log = natural log
            h1 = np.log(Kurs / Basis_value)
            sigma = VDAX * ((Tage / VDAX_Laufzeit) ** 0.5)
            sigma_1 = VDAX * ((Tage_1 / VDAX_Laufzeit) ** 0.5)
            h2 = InterestRate + sigma * sigma / 2
            h2_1 = InterestRate + sigma_1 * sigma_1 / 2
            d1 = (h1 + (h2 * (Tage / 365))) / (sigma * ((Tage / 365) ** 0.5))

            d1_1 = (h1 + (h2_1 * (Tage_1 / 365))) / (sigma_1 * ((Tage_1 / 365) ** 0.5))

            Phi = norm.pdf(d1, 0, 1)               
            Phi_1 = norm.pdf(d1_1, 0, 1)

            Gamma = Phi / (Kurs * (sigma * (Tage / 365) ** 0.5))
            Gamma_1 = Phi_1 / (Kurs * (sigma_1 * (Tage_1 / 365) ** 0.5))
            HedgeBedarf_values[i,k] = Gamma * Kontrakte / 5
            HedgeBedarf1_values[i,k] = Gamma_1 * Kontrakte_1 / 5


    HedgeSum = HedgeBedarf_values.sum(axis=1)/2
    HedgeSum_1 = HedgeBedarf1_values.sum(axis=1)/2
    HedgeBedarf_df = pd.DataFrame(data =HedgeBedarf_values, columns=Basis)
    HedgeSum_df = pd.DataFrame(HedgeSum,columns=["HedgeSum"])
    HedgeBedarf_df = pd.concat([HedgeBedarf_kurs,HedgeSum_df,HedgeBedarf_df], axis=1)

    HedgeBedarf1_df = pd.DataFrame(data =HedgeBedarf1_values, columns=Basis)
    HedgeSum1_df = pd.DataFrame(HedgeSum_1,columns=["HedgeSum"])
    HedgeBedarf1_df = pd.concat([HedgeBedarf_kurs,HedgeSum1_df,HedgeBedarf1_df], axis=1)

    # Fourth Stage


    info_df = pd.DataFrame({
        'ZentralKurs' : [NewZentralKurs],
        'Spannweite': [Spannweite],
        'SchrittWeite': [SchrittWeite],
        'Schritt' : [Schritt],
        'auswahl': auswahl_dict[auswahl],
        'expiry': expiry,
        'expiry_1': expiry_1,
        'heute': heute,
        #"nextContract": expiry_1,
        'tage_bis_verfall': tage_bis_verfall,
        'DAX_LS_e': DAX_LS_e,
        'Minkurs': Minkurs,
        'Maxkurs': Maxkurs,
        'Schritte': Schritte,
        'DetailMin': DetailMin,
        'DetailMax': DetailMax,
        'delta' : delta,
        'Euribor_IR': euribor_row['InterestRate'],
        'VDA_price': vda_2m_row['Price'],
    }).T.reset_index().rename(columns={0:"Value", 'index': "Info"})
    info_df
    for excel_file in list_excel_files:
        with pd.ExcelWriter(excel_file,datetime_format="DD/MM/YYYY") as writer:
            info_df.to_excel(writer,sheet_name='infos',index=False)
            Ueberhaenge_df.to_excel(writer,sheet_name='Ueberhaenge',index=False)
            Summery_df.to_excel(writer,sheet_name='Summery',index=False)
            SummeryDetail_df.to_excel(writer,sheet_name='SummeryDetail',index=False)
            overview_df.to_excel(writer,sheet_name='Overview',index=False)
            HedgeBedarf_df.to_excel(writer,sheet_name='HedgeBedarf',index=False)
            HedgeBedarf1_df.to_excel(writer,sheet_name='HedgeBedarf+01',index=False)
            vda_2m_df.to_excel(writer,sheet_name='VDAX-New 2M Hist',index=False)
            euribor_3m_df.to_excel(writer,sheet_name='EURIBOR 3M',index=False)
            
    print("Excel files have been exported.")
    HedgeBedarf_df["Sum"] = HedgeBedarf_df.HedgeSum + HedgeBedarf1_df.HedgeSum
    # Fifth stage
    # Values to create to arrow
    idx_closest = (HedgeBedarf_df.Basis - DAX_LS_e).abs().idxmin()
    closest_Basis =  HedgeBedarf_df.loc[idx_closest,"Basis"]
    closest_Sum = HedgeBedarf_df.loc[idx_closest,"Sum"]

    x_axis_length = HedgeBedarf_df.Sum.max() - HedgeBedarf_df.Sum.min()
    y_axis_length = HedgeBedarf_df.Basis.max() - HedgeBedarf_df.Basis.min()



    min_Kontrakte = 5000
    max_Differenz = 1000
    prozentual = 0.2

    # (CF > min_Kontrakte) AND (PF > min_Kontrakte) AND (ABS(CF - PF)  < MIN(PF,CF)*)
    mask_highlights = (
        (Summery_df.openInterest_CF > min_Kontrakte) & 
        (Summery_df.openInterest_PF > min_Kontrakte) & 
        (
            (Summery_df.openInterest_CF - Summery_df.openInterest_PF).abs() < 
            (Summery_df[['openInterest_PF',"openInterest_CF"]].min(axis=1)*prozentual)))

    ################################## Hilights for Basis column ####################################

    Summery_df["basis_color"] = "lavender"
    Summery_df.loc[
        mask_highlights, "basis_color"] = "yellow"
    col_basis_color = Summery_df.basis_color.to_numpy()
    #################################################################################################

    ######################### Gradient of red and blue for Anderung column ##########################
    aux = Summery_df.Änderung.reset_index()
    aux['r'] = aux['g']  = aux['b'] = 255
    aux.loc[aux.Änderung<0,'g'] = aux.loc[aux.Änderung<0,'b'] = 255 - 150*aux.Änderung/(aux.Änderung.min()) 
    aux.loc[aux.Änderung>0,'g'] = aux.loc[aux.Änderung>0,'r'] = 255 - 150*aux.Änderung/(aux.Änderung.max())
    aux = aux.astype(float)
    rb_shades = np.array([f"rgb({aux.loc[i,'r']},{aux.loc[i,'g']},{aux.loc[i,'b']})" for i in range(aux.shape[0]) ])
    col_anderung_color = rb_shades
    #################################################################################################

    ############################# Colors for heute and last_day columns #############################
    col_heute_color = posneg_binary_color(Summery_df.heute,"rgb(0, 204, 204)","rgb(77, 166, 255)")
    col_last_day_color = posneg_binary_color(Summery_df.last_day,"rgb(0, 204, 204)","rgb(77, 166, 255)")
    #################################################################################################

    ###################### Gradient of green and blue for Put anc Call columns ######################
    col_put_color = colors = np.array(n_colors('rgb(214, 245, 214)', 'rgb(40, 164, 40)',
        20, colortype='rgb'))[scale_col_range(Summery_df.openInterest_PF,19)]
    col_call_color = colors = np.array(n_colors('rgb(204, 224, 255)', 'rgb(0, 90, 179)',
        20, colortype='rgb'))[scale_col_range(Summery_df.openInterest_CF,19)]
    #################################################################################################

    row_height = 25
    # 41 row + header
    height = row_height*42

    # width of the image
    widht = 1000

    # Create figure containg two elements
    fig = make_subplots(
        # 2 Columns which are table and chart
        rows=1, cols=2,
        # Width ocupied by table and chart
        column_widths=[13,7],

        # Define types of the figures and margins
        specs=[[
            {"type": "table"},

            {"type": "scatter",
            # Add margins to the top and to the bottom so Basis column match the y axis of the chart 
            "t" : 1.5*(row_height/height), 'b':0.75*(row_height/height)
            }, 
            ]],
            )

    # Create table figure
    fig.add_trace(
        # Choose element 1 => table
        row = 1, col = 1,
        trace = go.Table(

        # Define some paremeters for the header
        header=dict(
            # Names of the columns
            values=(bold(["Basis","Änderung",heute.strftime("%d/%m/%Y"),tradingDates[1].strftime("%d/%m/%Y"),"Put", "Call"])),
            
            # Header style
            fill_color='paleturquoise',
                align='center',
                font = {'size': 11},
                height = row_height
            ),

        cells=dict(

            align='center',
            height = row_height,
            font = {'size': 11,},

            # Values of the table
            values=[
                (Summery_df.Basis),
                (Summery_df.Änderung.round(2)),
                (Summery_df.heute.round(2)),
                (Summery_df.last_day.round(2)),
                (Summery_df.openInterest_PF), 
                (Summery_df.openInterest_CF)
            ],

            # Colors of the columns
            fill_color = [
                col_basis_color,
                col_anderung_color,
                col_heute_color,
                col_last_day_color,
                col_put_color, 
                col_call_color
            ],
        )
    ))

    ##################################### Add plots in the chart #####################################
    fig.add_trace(go.Scatter(
        x = HedgeBedarf_df.Sum,
        y = HedgeBedarf_df.Basis,
        mode = "lines",
        name = heute.strftime("%Y-%m") +" + " + expiry.strftime("%Y-%m"),
        marker_color= "blue"
        ),
    )

    fig.add_trace(go.Scatter(
        x = HedgeBedarf1_df.HedgeSum,
        y = HedgeBedarf_df.Basis,
        mode = "lines",
        name = expiry.strftime("%Y-%m"),
        marker_color= "purple"
        ),
    )

    fig.add_trace(go.Scatter(
        x = HedgeBedarf_df.HedgeSum,
        y = HedgeBedarf_df.Basis,
        mode = "lines",
        name = heute.strftime("%Y-%m"),
        marker_color= "orange"
    ))

    # Create arrow
    fig.add_trace(go.Scatter(
            x=[closest_Sum + x_axis_length/5,closest_Sum + x_axis_length/50], 
        y=[closest_Basis,closest_Basis],
        marker= dict(size=20,symbol= "arrow-bar-up", color="red", angleref="previous"),
        showlegend = False
        )
    )
    ##################################################################################################
    dx =  0.1*(HedgeBedarf_df.Sum.max() - HedgeBedarf_df.Sum.min())

    fig.update_layout(
        # Set limits in the x and y axis
        yaxis_range= [HedgeBedarf_df.Basis.min(), HedgeBedarf_df.Basis.max()],
        xaxis_range= [HedgeBedarf_df.Sum.min() - dx, HedgeBedarf_df.Sum.max() + dx],
        
        # Remove margins
        margin=dict(l=0,r=0,b=0,t=0),
        paper_bgcolor="white",
        
        # Define width and height of the image
        width=widht,height=height + row_height,
        template = "seaborn",

        # Legend parameters
        legend=dict(
            yanchor="top",
            y=0.98 - (row_height/height),
            xanchor="right",
            x= 0.98,
            font=dict(
                size = 10
            ),
            # legend in the vertical
            orientation = "v"
            )
    )

    #fig.show()

    fig.write_image(result_image,scale=1)

    with open(src_html) as file:
        template = file.read()

    # Replace specific characters in the template by values
    dict_raplace = {
        "$PUT_SUM$": int(Summery_df.openInterest_PF.sum()),
        "$CALL_SUM$": int(Summery_df.openInterest_CF.sum()),
        "$TBF$": tage_bis_verfall,
        "$DELTA$":str(delta).replace(".",","),
        "$DATE$": heute.strftime("%d/%m/%Y")
    }

    for key, value in dict_raplace.items():
        template = template.replace(key, str(value))

    # Export html file
    with open(result_html,'w') as file:
        file.write(template)

    # Export pdf files
    for pdf_file in list_pdf_files:
        converter.convert("file://" + os.path.join(os.getcwd(),result_html), pdf_file)
    print("PDF has been generated.")


