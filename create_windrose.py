import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

def convert_to_wind_direction(degrees):
    directions = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW', 'N']
    index = np.round((degrees % 360) / 22.5).astype(int)
    return np.array(directions)[index]

def sort_wind_directions(directions):
    """Mengurutkan arah mata angin dari utara ke selatan, timur ke barat."""
    order = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW']
    return sorted(directions, key=lambda direction: order.index(direction))

def calculate_wind_frequency(file_path):
    # Membaca file Excel dengan multiple sheets
    xls = pd.ExcelFile(file_path)

    # Membuat dictionary untuk menyimpan DataFrame setiap sheet
    data_frames = {}

    # Loop melalui setiap sheet
    for sheet_name in xls.sheet_names:
        # Membaca data di range kolom G11:I165
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="G:H", skiprows=9, nrows=155)
        
        # Menghapus nilai NaN dan non-finite
        df_cleaned = df.dropna().replace([np.inf, -np.inf], np.nan)
        
        # Mengubah nilai ddd menjadi mata angin
        df_cleaned['wind_direction'] = convert_to_wind_direction(df_cleaned['ddd'])
        
        data_frames[sheet_name] = df_cleaned

    # Menghitung frekuensi mata angin berdasarkan kecepatan angin
    frequency_tables = {}
    for sheet_name, df_cleaned in data_frames.items():
        # Menghitung range kecepatan angin secara dinamis
        min_speed = df_cleaned['ff'].min()
        max_speed = df_cleaned['ff'].max()
        speed_bins = np.linspace(min_speed, max_speed, num=6)
        speed_labels = [f'{speed_bins[i]:.1f}-{speed_bins[i+1]:.1f}' for i in range(len(speed_bins)-1)]

        # Menghitung frekuensi mata angin
        frequency_table = df_cleaned.groupby(['wind_direction', pd.cut(df_cleaned['ff'], bins=speed_bins, labels=speed_labels)]).size().reset_index(name='frequency')

        # Menambah bar kosong untuk mata angin yang tidak memiliki nilai
        for direction in directions:
            if direction not in frequency_table['wind_direction'].tolist():
                frequency_table = frequency_table.append({'wind_direction': direction, 'ff': speed_labels[0], 'frequency': 0}, ignore_index=True)

        # Mengurutkan bar berdasarkan arah mata angin
        frequency_table['wind_direction'] = sort_wind_directions(frequency_table['wind_direction'])

        frequency_tables[sheet_name] = frequency_table

    return frequency_tables

if __name__ == "__main__":
    # Membuat sidebar untuk memilih sheet
    sheet_name = st.sidebar.selectbox('Pilih sheet', xls.sheet_names)

    # Menghitung frekuensi arah angin
    tables = calculate_wind_frequency("850 mb 00UTC.xlsx")
    table = tables[sheet_name]

    # Menampilkan grafik bar polar
    fig = px.bar_polar(table, r="frequency", theta="wind_direction",
                       color="ff",
                       color_discrete_sequence=px.colors.sequential.Plasma_r)
    fig.update_layout(
        polar_angularaxis_direction='clockwise',
        polar_angularaxis_rotation=0
    )
    st.plotly_chart(fig)