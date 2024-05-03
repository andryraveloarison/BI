import streamlit as st  # pip install streamlit
import pandas as pd  # pip install pandas
import plotly.express as px  # pip install plotly-express
import base64  # Standard Python Module
from io import StringIO, BytesIO  # Standard Python Module
import plotly.graph_objects as go



def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    # Credit Plotly: https://discuss.streamlit.io/t/download-plotly-plot-as-html/4426/2
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)


st.set_page_config(page_title='BI Streamlit')
st.title('BI avec Streamlit 📈')
st.subheader('By Andry & Fehizoro')

uploaded_file = st.file_uploader('Choisir un fichier xlsx', type='xlsx')
if uploaded_file:
    st.markdown('---')
   
    sheet_name = 'DATA'

    df = pd.read_excel(uploaded_file,
                    sheet_name=sheet_name,
                    usecols='B:C',
                    header=3)

    df_sortants = pd.read_excel(uploaded_file,
                                    sheet_name= sheet_name,
                                    usecols='F:G',
                                    header=3)

    df_evolution_sortie = pd.read_excel(uploaded_file,
                                    sheet_name= sheet_name,
                                    usecols='K:L',
                                    header=3)

    df_sortants.dropna(inplace=True)

        
    # --- STREAMLIT SELECTION
    st.title('Regroupement des genres par filiere')
    filiere = df['Filiere'].unique().tolist()
    anneUni = df_evolution_sortie['AnneUni'].unique().tolist()
    nbSortie = df_evolution_sortie['NbSortie'].unique().tolist()


    filiere_selection = st.multiselect('Filiere:',
                                        filiere,
                                        default=filiere)

    # --- FILTER DATAFRAME BASED ON SELECTION
    mask = df['Filiere'].isin(filiere_selection)
    number_of_result = df[mask].shape[0]
    #st.markdown(f'*Available Results: {number_of_result}*')
    print(len(filiere_selection))

    # --- GROUP DATAFRAME AFTER SELECTION
    # Supposons que 'mask' est votre condition de sélection
    df_filtered = df[mask]

    # Regrouper par 'Rating' et compter le nombre de lignes pour chaque 'Rating'
    df_grouped = df_filtered.groupby(by=['Filiere', 'Sexe']).size().reset_index(name='Nombre')

    df_combined = df_grouped.pivot(index='Filiere', columns='Sexe', values='Nombre').reset_index()

    print(number_of_result)

    if(number_of_result > 0):
    # --- PLOT BAR CHART
        bar_chart = px.bar(df_combined,
                   x='Filiere',
                   y=['H', 'F'],
                   labels={'value': 'Nombre'},
                   color_discrete_sequence=['#F63366', '#00CC99'],  # Couleurs pour Hommes et Femmes
                   template='plotly_white')

        

        st.plotly_chart(bar_chart)

    else :
        st.markdown('Veuillez selectionner une filiere ')


    # --- PLOT PIE CHART
    pie_chart = px.pie(df_sortants,
                    title='Taux de réussite par filière',
                    values='TauxDeReussite',
                    names='Filiere.1')

    histogram = px.bar(data_frame=df_sortants, x='Filiere.1', y='TauxDeReussite', title='Taux de réussite des elèves')

    st.plotly_chart(pie_chart)
    st.plotly_chart(histogram)


    st.title('Evolution des nombres de sortons a l\'ESMIA depuis 2019')

    # Création du graphique
    fig = go.Figure(go.Scatter(
        x= anneUni,
        y= nbSortie
    ))

    # Mise à jour de la mise en page pour l'axe X
    fig.update_layout(
        xaxis=dict(
            tickmode='linear',
            tick0=0,
            dtick=0
        )
    )

    # Affichage du graphique dans Streamlit
    st.plotly_chart(fig)


      # -- DOWNLOAD SECTION
    st.subheader('Downloads:')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(pie_chart)