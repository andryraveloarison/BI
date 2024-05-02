import streamlit as st  # pip install streamlit
import pandas as pd  # pip install pandas
import plotly.express as px  # pip install plotly-express
import base64  # Standard Python Module
from io import StringIO, BytesIO  # Standard Python Module


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


st.set_page_config(page_title='Excel Plotter')
st.title('Excel Plotter 📈')
st.subheader('Feed me with your Excel file')

uploaded_file = st.file_uploader('Choose a XLSX file', type='xlsx')
if uploaded_file:
    st.markdown('---')
   


    sheet_name = 'DATA'

    df = pd.read_excel(uploaded_file,
                    sheet_name=sheet_name,
                    usecols='B:D',
                    header=3)

    df_participants = pd.read_excel(uploaded_file,
                                    sheet_name= sheet_name,
                                    usecols='F:G',
                                    header=3)
    df_participants.dropna(inplace=True)

        
    # --- STREAMLIT SELECTION
    department = df['Department'].unique().tolist()
    ages = df['Age'].unique().tolist()

    age_selection = st.slider('Age:',
                            min_value= min(ages),
                            max_value= max(ages),
                            value=(min(ages),max(ages)))

    department_selection = st.multiselect('Department:',
                                        department,
                                        default=department)

    # --- FILTER DATAFRAME BASED ON SELECTION
    mask = (df['Age'].between(*age_selection)) & (df['Department'].isin(department_selection))
    number_of_result = df[mask].shape[0]
    st.markdown(f'*Available Results: {number_of_result}*')


    # --- GROUP DATAFRAME AFTER SELECTION
    df_grouped = df[mask].groupby(by=['Rating']).count()[['Age']]
    df_grouped = df_grouped.rename(columns={'Age': 'Votes'})
    df_grouped = df_grouped.reset_index()

    # --- PLOT BAR CHART
    bar_chart = px.bar(df_grouped,
                    x='Rating',
                    y='Votes',
                    text='Votes',
                    color_discrete_sequence = ['#F63366']*len(df_grouped),
                    template= 'plotly_white')
    st.plotly_chart(bar_chart)

    # --- PLOT PIE CHART
    pie_chart = px.pie(df_participants,
                    title='Total No. of Participants',
                    values='Participants',
                    names='Departments')

    st.plotly_chart(pie_chart)

    # -- DOWNLOAD SECTION
    st.subheader('Downloads:')
    generate_excel_download_link(df_grouped)
    generate_html_download_link(pie_chart)