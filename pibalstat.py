from windrose import WindroseAxes
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt


def create_windrose_from_excel(file_path):
    # Membaca file Excel dengan multiple sheets
    xls = pd.ExcelFile(file_path)

    # Membuat dictionary untuk menyimpan DataFrame setiap sheet
    data_frames = {}
    data_frames_dropna = {}

    # Loop melalui setiap sheet
    for sheet_name in xls.sheet_names:
        # Membaca data di range kolom G11:I165
        df = pd.read_excel(file_path, sheet_name=sheet_name, usecols="G:H", skiprows=9, nrows=155)

        # Menghapus nilai NaN dan menyimpan DataFrame baru
        df_dropna = df.dropna()

        data_frames[sheet_name] = df
        data_frames_dropna[sheet_name] = df_dropna

    # Loop melalui setiap sheet
    for sheet_name, df_dropna in data_frames_dropna.items():
        # Mendapatkan data angin dari DataFrame yang telah dihapus nilai NaN
        directions = df_dropna['ddd'].values
        speeds = df_dropna['ff'].values

        # Membuat plot windrose
        fig, ax = plt.subplots()
        ax = WindroseAxes.from_ax(fig=fig)
        ax.bar(directions, speeds, normed=True, opening=0.8, edgecolor='white')
        ax.set_legend()

        # Menampilkan plot
        plt.title(f"Windrose - {sheet_name}")

        # Menampilkan plot menggunakan Streamlit
        st.pyplot(fig)

# Memanggil fungsi untuk membuat windrose dari file Excel
file_path = st.file_uploader("Upload file Excel", type=["xlsx"])
if file_path is not None:
    create_windrose_from_excel(file_path)
