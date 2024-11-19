import pandas as pd
from tkinter import messagebox
from constants import EXCEL_FILE

# Create or load the Excel file
def create_or_load_file():
    try:
        df = pd.read_excel(EXCEL_FILE)
    except FileNotFoundError:
        columns = ['Province', 'Area Code', 'Date', 'Transformer Serial No', 'KV', 'KVA',
                   'Reason for Movement', 'Present Condition',
                   'Movement From Cost Code', 'Movement From SIN Location',
                   'Movement To Cost Code', 'Movement To SIN Location',
                   'TV/WB No', 'Remark']
        df = pd.DataFrame(columns=columns)
        df.to_excel(EXCEL_FILE, index=False)

# Function to add data to the Excel sheet
def add_data(data):
    try:
        df = pd.read_excel(EXCEL_FILE)
        new_data = pd.DataFrame([data], columns=df.columns)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        messagebox.showinfo("Success", "Data added successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Error while adding data: {str(e)}")

# Function to read data
def read_data():
    try:
        df = pd.read_excel(EXCEL_FILE)
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Error while reading data: {str(e)}")
        return pd.DataFrame()

# Function to search data
def search_data(search_query):
    df = read_data()
    return df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]

# Function to filter data by column and value
def filter_data(column, value):
    df = read_data()
    return df[df[column].astype(str).str.contains(value, case=False)]

# Function to sort data by column
def sort_data(column):
    df = read_data()
    return df.sort_values(by=column)
