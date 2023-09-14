from utils import *
from variables import *
from datetime import datetime, timedelta
import requests
import re
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pyhtml2pdf import  converter
import shutil
from plotly.colors import n_colors
from typing import Union, List, Dict
from bs4 import BeautifulSoup
import traceback
from utils import read_pickle, save_as_pickle, next_business_day




def generate_files(heute_input: Union[None, str] = None)->None:
    # Create results folder in case they still were not created
    list_folder_results =  [current_results_path, old_results_path, temp_results_path]
    for folder in list_folder_results:
        create_folder(folder)

    # During the generation of the PDF, the CSS file must be in the same directory
    # as the generated html file. So in case the CSS file is still not the in the
    # Results folder, we copy this file to it
    #if not os.path.exists(result_css):
    shutil.copyfile(src_css, result_css) # removed if bc file was not updated in Detlef's cloned

    list_email_send = []

    for auswahl in [1,0]:
        result = parse_eurex(auswahl, heute = heute_input)

        if result:
            auswahl, heute, stock_price, tage_bis_verfall, nbd_dict, Summery_df, HedgeBedarf_df, HedgeBedarf1_df, delta, future_date_col, list_email_send_selection = result

        list_email_send += list_email_send_selection


        file_path = os.path.join(current_results_path, f"{dict_index_stock[auswahl]}.pdf")

        generate_pdfs(auswahl, Summery_df, HedgeBedarf_df, HedgeBedarf1_df, stock_price, heute, nbd_dict, tage_bis_verfall, delta, future_date_col, file_path)


    return list_email_send
    
        

def generate_pdfs(auswahl, Summery_df, HedgeBedarf_df, HedgeBedarf1_df, stock_price, heute, nbd_dict, tage_bis_verfall, delta, future_date_col, file_path):

    # Fifth stage
    # Values to create arrow
    idx_closest = (HedgeBedarf_df.Basis - stock_price).abs().idxmin()
    closest_Basis =  HedgeBedarf_df.loc[idx_closest,"Basis"]
    HedgeBedarf_df["Sum"] = HedgeBedarf_df.HedgeSum + HedgeBedarf1_df.HedgeSum
    closest_Sum = HedgeBedarf_df.loc[idx_closest,"Sum"]

    x_axis_length = HedgeBedarf_df.Sum.max() - HedgeBedarf_df.Sum.min()
    y_axis_length = HedgeBedarf_df.Basis.max() - HedgeBedarf_df.Basis.min()

    min_Kontrakte = 5000
    max_Differenz = 1000
    prozentual = 0.2

    condition_close_verfall = (auswahl == 0) and tage_bis_verfall < 5

    if condition_close_verfall:
        indexes = Summery_df.loc[(Summery_df.Basis >= stock_price)].tail(10).index.to_list() + Summery_df.loc[(Summery_df.Basis < stock_price)].head(10).index.to_list()
        Summery_df =  Summery_df.loc[indexes].reset_index(drop=True)

        indexes = HedgeBedarf_df.loc[(HedgeBedarf_df.Basis >= Summery_df.Basis.min()) & (HedgeBedarf_df.Basis <= Summery_df.Basis.max())].index
        
        HedgeBedarf_df = HedgeBedarf_df.loc[indexes].reset_index(drop=True)
        HedgeBedarf1_df = HedgeBedarf1_df.loc[indexes].reset_index(drop=True)

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
    col_put_color  = np.array(n_colors('rgb(214, 245, 214)', 'rgb(40, 164, 40)',
        20, colortype='rgb'))[scale_col_range(Summery_df.openInterest_PF,19)]

    col_call_color  = np.array(n_colors('rgb(204, 224, 255)', 'rgb(0, 90, 179)',
        20, colortype='rgb'))[scale_col_range(Summery_df.openInterest_CF,19)]
    #################################################################################################

    num_rows = Summery_df.shape[0]+1
    height = 1050 #row_height*42
    row_height = int(height/(num_rows))

    # width of the image
    widht = 1000
    row_height_percent =  (row_height/height)



    values_body = [
                (Summery_df.Basis),
                (Summery_df.Änderung.round(0)),
                (Summery_df.heute.round(0)),
                (Summery_df.last_day.round(0)),

            ]
    values_header = bold(["Basis","Änderung",nbd_dict['heute'].strftime("%d/%m/%y"),nbd_dict['last'].strftime("%d/%m/%y")])


    if not  condition_close_verfall:
        values_body += [(Summery_df.openInterest_PF), (Summery_df.openInterest_CF)]
        values_header += bold(["Put","Call"])


    font_size =  int(440/num_rows)

    fig = go.Figure(
        data=[
            go.Table(

                # Define some paremeters for the header
                header=dict(
                    # Names of the columns
                    values= values_header,
                    
                    # Header style
                    fill_color='paleturquoise',
                        align='center',
                        font = {'size': int(font_size*0.8)},
                        height = row_height,
                ),

                cells=dict(

                    align='center',
                    height = row_height,
                    font = {'size': font_size,},

                    # Values of the table
                    values= values_body,

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
            )
        ],
    )
    fig.update_layout(
        height=1050, width=550,
        margin=dict(l=0,r=0,b=0.0,t=0)
    )

    fig.write_image(os.path.join(current_results_path,'image_table.svg'),scale=1)


    data =  [
        go.Scatter(
            x = HedgeBedarf_df.Sum,
            y = HedgeBedarf_df.Basis,
            mode = "lines",
            name = datetime.strptime(future_date_col[0],"%Y%m%d").strftime("%Y-%m") + " + "+ datetime.strptime(future_date_col[1],"%Y%m%d").strftime("%Y-%m"),
            marker_color= "blue"
        )
    ]


    if not condition_close_verfall:
        data += [
            go.Scatter(
                x = HedgeBedarf1_df.HedgeSum,
                y = HedgeBedarf_df.Basis,
                mode = "lines",
                name = datetime.strptime(future_date_col[1],"%Y%m%d").strftime("%Y-%m"),
                marker_color= "purple"
            ),
            go.Scatter(
            x = HedgeBedarf_df.HedgeSum,
            y = HedgeBedarf_df.Basis,
            mode = "lines",
            name = datetime.strptime(future_date_col[0],"%Y%m%d").strftime("%Y-%m"),
            marker_color= "orange"
            )
        ]

    ##################################################################################################
    dx =  0.1*(HedgeBedarf_df.Sum.max() - HedgeBedarf_df.Sum.min())



    fig = go.Figure(data=data)

    # Create arrow
    fig.add_trace(go.Scatter(
            x = [closest_Sum + x_axis_length/5,closest_Sum + x_axis_length/50], 
            y = [closest_Basis,closest_Basis],
        marker= dict(size=20,symbol= "arrow-bar-up", angleref="previous"),
        showlegend = False
        )
    )


    fig.update_layout(
        margin=dict(l=0,r=0,b=0.1,t=row_height),
        # Set limits in the x and y axis
        yaxis_range= [HedgeBedarf_df.Basis.min(), HedgeBedarf_df.Basis.max()],
        xaxis_range= [HedgeBedarf_df.Sum.min() - dx, HedgeBedarf_df.Sum.max() + dx],
        yaxis = dict(
            tickmode = 'array',
            tickvals = HedgeBedarf_df.Basis[HedgeBedarf_df.index % 3 == 0],
            #ticktext = ['One', 'Three', 'Five', 'Seven', 'Nine', 'Eleven']
        ),
        
        # Remove margins
        paper_bgcolor="white",
        
        # Define width and height of the image
        width=380,height=1050,
        template = "seaborn",

        # Legend parameters
        legend=dict(
            yanchor="top",
            y=1,# - (row_height/height),
            xanchor="right",
            x= 1,
            font=dict(
                size = 8
            ),
            # legend in the vertical
            orientation = "v",
            bgcolor  = 'rgba(0,0,0,0)'
            ),
    )


    fig.write_image(os.path.join(current_results_path,'image_graph.svg'),scale=1)



    fig = go.Figure()

    if not condition_close_verfall:
        fig.add_trace(
            trace = go.Bar(name='Put', y=Summery_df.Basis, x=Summery_df.openInterest_PF,orientation='h', marker_color = 'rgb(40, 164, 40)'),
        )

        fig.add_trace(
            trace = go.Bar(name='Call', y=Summery_df.Basis, x=Summery_df.openInterest_CF,orientation='h', marker_color = 'rgb(0, 90, 179)'),
        )



    fig.update_layout(
        barmode='overlay', margin=dict(l=0,r=0,b=0.1,t=row_height),
        # Set limits in the x and y axis
        #yaxis_range= [HedgeBedarf_df.Basis.min(), HedgeBedarf_df.Basis.max()],
        #xaxis_range= [HedgeBedarf_df.Sum.min() - dx, HedgeBedarf_df.Sum.max() + dx],
        
        # Define width and height of the image
        width=80,height=1050,
        template = "seaborn",
        showlegend=False,
        bargap =0.5 ,
        plot_bgcolor='rgba(0, 0, 0, 0)',
        paper_bgcolor='rgba(0, 0, 0, 0)',
    )


    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)

    fig.write_image(os.path.join(current_results_path,'image_bar.svg'),scale=1)


    with open(src_html) as file:
        template = file.read()


    dict_title = {
        0 : "OpenInterest und HedgeBedarf",
        1 : "STOXX 50 OpenInterest und HedgeBedarf",
    }

    # Replace specific characters in the template by values
    dict_raplace = {
        "$PUT_SUM$": int(Summery_df.openInterest_PF.sum()),
        "$CALL_SUM$": int(Summery_df.openInterest_CF.sum()),
        "$TBF$": tage_bis_verfall,
        "$DELTA$":str(delta).replace(".",","),
        "$DATE$": heute.strftime("%d/%m/%Y"),
        "$TITLE$" : dict_title[auswahl]
    }

    for key, value in dict_raplace.items():
        template = template.replace(key, str(value))

    # Export html file
    with open(result_html,'w') as file:
        file.write(template)

    # Results files
    

    converter.convert("file://" + os.path.join(os.getcwd(),result_html), file_path)
    #shutil.copyfile(list_pdf_files[0], list_pdf_files[1])
    
    print("PDF has been generated.")


