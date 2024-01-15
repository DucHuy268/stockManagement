import streamlit as st
import pandas as pd
import os
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb

def main():
    st.title("Stock Management")
    st.markdown("It's for you, Anh Hong :sparkling_heart:")

    # Ask for the Excel file name
    excel_file_name = st.text_input("Excel File Name (with extension):", "stock_data.xlsx")

    # Check if the file exists
    if os.path.isfile(excel_file_name):
        # Ask for the number of columns
        num_columns = st.number_input("Number of Columns:", min_value=1, value=2)

        # Ask for the column names
        column_names = []
        for i in range(num_columns):
            col_name = st.text_input(f"Column Name {i + 1}:", f"Column_{i + 1}")
            column_names.append(col_name)

        # Load existing data from the Excel file
        stock_data = load_data(excel_file_name, column_names)

        # Form to add a new item
        st.subheader("Add a New Item")
        new_item = {}
        for col_name in column_names:
            value = st.text_input(f"{col_name}:")
            new_item[col_name] = value

        add_button = st.button("Add to Stock")

        if add_button:
            # Add the new item to the stock
            data = add_item(stock_data, new_item)
            st.success("New item added to the stock!")

            # Show the current stock
            st.subheader("Current Stock")
            st.write(data)

            # Update the Excel file
            save_data(data, excel_file_name)

            df_xlsx = to_excel(data)
            st.download_button(label='📥 Download Current Result',
                               data=df_xlsx,
                               file_name=excel_file_name)

    else:
        st.warning(f"The file {excel_file_name} does not exist. Please create a new file.")
        create_file_button = st.button("Create File")
        if create_file_button:
            # Create a new file with default column names if it doesn't exist
            create_empty_file(excel_file_name)


def load_data(excel_file_name, column_names):
    # Load data from the Excel file (create the file if it doesn't exist)
    try:
        stock_data = pd.read_excel(excel_file_name)
    except FileNotFoundError:
        stock_data = pd.DataFrame(columns=column_names)
        stock_data.to_excel(excel_file_name, index=False)
    return stock_data


def add_item(data, new_item):
    # Add a new item to the stock
    new_row = pd.DataFrame(new_item, index=[0])
    data = pd.concat([data, new_row], ignore_index=True)
    return data


def save_data(data, excel_file_name):
    # Save the data to the Excel file
    data.to_excel(excel_file_name, index=False)


def create_empty_file(excel_file_name, column_names=None):
    # Create an empty Excel file with specified or default columns
    if column_names is None:
        column_names = ["Column_1", "Column_2"]
    empty_data = pd.DataFrame(columns=column_names)
    empty_data.to_excel(excel_file_name, index=False)


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'})
    worksheet.set_column('A:A', None, format1)
    writer.close()
    processed_data = output.getvalue()
    return processed_data


if __name__ == "__main__":
    main()
