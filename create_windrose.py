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

def calculate_wind_frequency(file_path, selected_sheet):
    try:
        # Membaca data dari sheet yang dipilih
        df = pd.read_excel(file_path, sheet_name=selected_sheet, usecols="G:H", skiprows=9, nrows=155)

        # Menghapus nilai NaN dan non-finite
        df_cleaned = df.dropna().replace([np.inf, -np.inf], np.nan)

        # Mengubah nilai ddd menjadi mata angin
        df_cleaned['wind_direction'] = convert_to_wind_direction(df_cleaned['ddd'])

        # Menghitung frekuensi mata angin
        frequency_table = df_cleaned['wind_direction'].value_counts().reset_index()
        frequency_table.columns = ['wind_direction', 'frequency']

        return frequency_table
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")

def main():
    st.title("Wind Frequency Dashboard")

    # Upload file Excel
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Membaca file Excel dengan multiple sheets
        xls = pd.ExcelFile(uploaded_file)

        # Mendapatkan nama-nama sheet
        sheet_names = xls.sheet_names

        # Menampilkan dropdown untuk memilih sheet
        selected_sheet = st.selectbox("Pilih Sheet", sheet_names)

        # Menghitung frekuensi angin
        frequency_table = calculate_wind_frequency(uploaded_file, selected_sheet)

        if frequency_table is not None:
            # Menampilkan grafik frekuensi mata angin
            st.title("Grafik Polar Frekuensi Mata Angin")
            fig = px.bar_polar(frequency_table, r="frequency", theta="wind_direction",
                               color="frequency",
                               color_continuous_scale="Plasma",
                               template="plotly_dark")
            st.plotly_chart(fig)

if __name__ == "__main__":
    main()
