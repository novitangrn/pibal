import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import io

def convert_to_wind_direction(degrees):
    directions = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW', 'N']
    index = np.round((degrees % 360) / 22.5).astype(int)
    return np.array(directions)[index]

def calculate_wind_frequency(file_path):
    # Membaca file Excel dengan multiple sheets
    xls = pd.ExcelFile(file_path)

    # Membuat dictionary untuk menyimpan DataFrame setiap sheet
    data_frames = {}
    data_frames_dropna = {}

    # Loop melalui setiap sheet
    for sheet_name in xls.sheet_names:
        # Membaca data di range kolom G11:I165
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="G:H", skiprows=9, nrows=155)
        
        # Menghapus nilai NaN dan non-finite
        df_cleaned = df.dropna().replace([np.inf, -np.inf], np.nan)
        
        # Mengubah nilai ddd menjadi mata angin
        df_cleaned['wind_direction'] = convert_to_wind_direction(df_cleaned['ddd'])
        
        data_frames[sheet_name] = df
        data_frames_dropna[sheet_name] = df_cleaned

    # Menghitung frekuensi mata angin berdasarkan kecepatan angin
    frequency_tables = {}
    for sheet_name, df_cleaned in data_frames_dropna.items():
        # Menghitung range kecepatan angin secara dinamis
        min_speed = df_cleaned['ff'].min()
        max_speed = df_cleaned['ff'].max()
        speed_bins = np.linspace(min_speed, max_speed, num=6)
        speed_labels = [f'{speed_bins[i]:.1f}-{speed_bins[i+1]:.1f}' for i in range(len(speed_bins)-1)]
        
        frequency_table = df_cleaned.groupby(['wind_direction', pd.cut(df_cleaned['ff'], bins=speed_bins, labels=speed_labels)]).size().reset_index(name='frequency')
        frequency_tables[sheet_name] = frequency_table
    
    return frequency_tables

# Menggunakan Streamlit untuk menampilkan grafik polar
st.title("Grafik Polar Frekuensi Mata Angin")
file_path = st.file_uploader("Upload file Excel", type=["xlsx"])

if file_path is not None:
    tables = calculate_wind_frequency(file_path)
    bulan = st.selectbox("Pilih Bulan", list(tables.keys()))

    if bulan:
        table = tables[bulan]
        fig = px.bar_polar(table, r="frequency", theta="wind_direction",
                           color="ff",
                           color_discrete_sequence=px.colors.sequential.Plasma_r)
        st.plotly_chart(fig)

    # Add a download button to rename the downloaded file
    if st.button('Download Chart'):
        # Convert the figure to an image
        image = fig.to_image(format='png')
    
        # Prompt the user to enter a new file name
        new_file_name = st.text_input('Enter a new file name', 'chart.png')
    
        # Download the image with the new file name
        st.download_button(label='Download', data=image, file_name=new_file_name, mime='image/png')
