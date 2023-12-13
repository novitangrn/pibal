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

def main_dynamic_windrose():
    st.title("Dynamic Windrose Generator")

    file_path = st.file_uploader("Upload Excel File", type=["xlsx"])
    if file_path is not None:
        df = pd.read_excel(file_path)
        frequency_table = calculate_wind_frequency(df)

        fig = px.bar_polar(frequency_table, r="percentage", theta="wind_direction",
                           color="ff",
                           color_discrete_sequence=px.colors.sequential.Rainbow_r,
                           start_angle=0,
                           direction="clockwise"
                          )
        fig.update_layout(polar_angularaxis_rotation=90)
        st.plotly_chart(fig)

def main_upload_windrose():
    st.title("Wind Frequency Analysis")

    file_path = st.file_uploader("Upload Excel File", type=["xlsx"])
    if file_path is not None:
        sheet_names = pd.ExcelFile(file_path).sheet_names
        sheet_name = st.selectbox("Select Sheet", sheet_names)

        col_range = st.text_input("Column Range (e.g., A:D)", value="A:D")

        if st.button("Calculate"):
            df = pd.read_excel(file_path, sheet_name=sheet_name, usecols=col_range)
            frequency_table = calculate_wind_frequency(df)

            fig = px.bar_polar(frequency_table, r="percentage", theta="wind_direction",
                               color="ff",
                               color_discrete_sequence=px.colors.sequential.Rainbow_r,
                               start_angle=0,
                               direction="clockwise"
                              )
            fig.update_layout(polar_angularaxis_rotation=90)
            st.plotly_chart(fig)

            pivot_table = create_pivot_table(frequency_table, sheet_name)
            pivot_table_file_name = f"pivot_table_{sheet_name}_{col_range}.xlsx"
            pivot_table.to_excel(pivot_table_file_name, index=True)
            st.success(f"Pivot table saved as {pivot_table_file_name}")

def create_pivot_table(frequency_table, sheet_name):
    pivot_table = pd.pivot_table(frequency_table, values='frequency', index='wind_direction', columns='ff', fill_value=0)
    pivot_table.columns.name = sheet_name
    return pivot_table

if __name__ == "__main__":
    main_dynamic_windrose()
    main_upload_windrose()
