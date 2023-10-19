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
    try:
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
            speed_bins = [0, 5, 10, 15, 20, 25, np.inf]
            speed_labels = ['0-5 knot', '6-10 knot', '11-15 knot', '16-20 knot', '20-25 knot', '>25 knot']
            frequency_table = df_cleaned.groupby(['wind_direction', pd.cut(df_cleaned['ff'], bins=speed_bins, labels=speed_labels)]).size().reset_index(name='frequency')
     
            # Menambah bar kosong untuk mata angin yang tidak memiliki nilai
            order = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW']
            for direction in order:
                if direction not in frequency_table['wind_direction'].tolist():
                    frequency_table = frequency_table.append({'wind_direction': direction, 'ff': speed_labels[0], 'frequency': 0}, ignore_index=True)

            # Mengurutkan bar berdasarkan arah mata angin
            frequency_table['wind_direction'] = sort_wind_directions(frequency_table['wind_direction'])

            frequency_tables[sheet_name] = frequency_table

        return frequency_tables
    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")


def create_pivot_table(data_frame, bulan):
    pivot_table = data_frame[bulan].pivot(index='ff', columns='wind_direction', values='frequency')
    pivot_table_filled = pivot_table.fillna(0).astype(int)
    return pivot_table_filled

def main():
    st.title("Wind Frequency Dashboard")

    # Upload file Excel
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Menghitung frekuensi angin
        frequency_tables = calculate_wind_frequency(uploaded_file)
        bulan = st.selectbox("Pilih Bulan", list(frequency_tables.keys()))

        if bulan:
            # Menampilkan grafik untuk setiap sheet dalam file Excel sesuai bulan yang dipilih
            table = frequency_tables[bulan]
            st.subheader(f"Windrose Bulan: {bulan}")
            fig = px.bar_polar(sorted_table, r="frequency", theta="wind_direction",
                               color="ff",
                               color_discrete_sequence=px.colors.sequential.Rainbow_r,
                               start_angle=0,
                               direction="clockwise"
                              )
            fig.update_layout(polar_angularaxis_rotation=90)
            fig.show()

            # Menampilkan Pivot Table
            pivot_table = create_pivot_table(table, bulan)
            st.subheader(f"Pivot Table Bulan: {bulan}")
            st.dataframe(pivot_table)

if __name__ == "__main__":
    main()
