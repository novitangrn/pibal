import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st

def convert_to_wind_direction(degrees):
    directions = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW', 'N']
    index = np.round((degrees % 360) / 45).astype(int)
    return np.array(directions)[index]

def sort_wind_directions(directions):
    order = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW']
    return sorted(directions, key=lambda direction: order.index(direction))

def calculate_wind_frequency(df):
    try:
        df_cleaned = df.dropna().replace([np.inf, -np.inf], np.nan)
        df_cleaned.columns = [col.split('.')[0] for col in df_cleaned.columns]  # Remove suffix after dot in column names
        df_cleaned['wind_direction'] = convert_to_wind_direction(df_cleaned['ddd'])

        # Calculate wind frequency based on wind speed
        speed_bins = [0, 5, 10, 15, 20, 25, np.inf]
        speed_labels = ['0-5 knot', '6-10 knot', '11-15 knot', '16-20 knot', '20-25 knot', '>25 knot']
        frequency_table = df_cleaned.groupby(['wind_direction', pd.cut(df_cleaned['ff'], bins=speed_bins, labels=speed_labels)]).size().reset_index(name='frequency')
        order = ['N', 'NE', 'E', 'SE', 'S', 'SW', 'W', 'NW']
        empty_directions = [direction for direction in order if direction not in frequency_table['wind_direction'].tolist()]
        empty_rows = pd.DataFrame({'wind_direction': empty_directions, 'ff': speed_labels[0], 'frequency': 0})
        frequency_table = pd.concat([frequency_table, empty_rows], ignore_index=True)
        frequency_table['wind_direction'] = sort_wind_directions(frequency_table['wind_direction'])

        # Convert frequency to percentage
        total_count = frequency_table['frequency'].sum()
        frequency_table['percentage'] = frequency_table['frequency'] / total_count * 100

        return frequency_table
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")

def main():
    st.title("Wind Frequency Analysis")

    option = st.sidebar.selectbox("Select Analysis Type", ("Dynamic Windrose", "Upload Windrose"))

    if option == "Dynamic Windrose":
        main_dynamic_windrose()
    elif option == "Upload Windrose":
        main_upload_windrose()

def main_dynamic_windrose():
    st.subheader("Dynamic Windrose Generator")

    # Your code for dynamic windrose here

def main_upload_windrose():
    st.subheader("Upload Windrose")
    file_key = "upload_file_key"
    file_path = st.file_uploader("Upload Excel File", type=["xlsx"], key=file_key)

    if file_path is not None:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        sheet_name = st.selectbox("Select Sheet", sheet_names)

        col_range = st.text_input("Column Range (e.g., A:D)", value="A:D")

        if st.button("Calculate"):
            # Your code for upload windrose here

if __name__ == "__main__":
    main()
