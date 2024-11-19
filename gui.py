import sys
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd  # Import pandas to use DataFrame
from excel_handler import  add_data, search_data, read_data
from constants import EXCEL_FILE
from tkinter import END, LEFT, RIGHT
import os
from datetime import datetime

# Excel file to store data
excel_file = 'transformer_data.xlsx'
last_entry_index = None  # Track the currently displayed entry index for deletion/editing
entries = {}  # Declare entries as a global variable
add_btn = None  # Declare add_btn globally
clear_btn = None  # Declare clear_btn globally
update_btn = None  # Declare update_btn globally
delete_btn = None  # Declare delete_btn globally
cancel_btn = None  # Declare cancel_btn globally
  # Declare add_btn globally
searched_entry_index = None  # Store the index of the searched entry
is_update_mode = False  # Track if the form is in update mode
row_num = 18 
row_num1 = 19


root = tk.Tk()
root.title("CEB Transformer Data App(Additional Finance Manager-DD4)")
root.geometry("1920x1080" )  # Set a fixed size for the window


def get_resource_path(relative_path):
    """Get the absolute path to the resource in a bundled executable or in the development environment."""
    try:
        if hasattr(sys, '_MEIPASS'):
            # If running as a bundled executable, the resources will be inside _MEIPASS
            base_path = os.path.join(sys._MEIPASS, 'src')  # Ensure we point to 'src' folder inside _MEIPASS
        else:
            # If running in development mode (from source)
            base_path = os.path.dirname(__file__)  # The directory of the current script
        
        # Join the base path with the relative resource path
        resource_path = os.path.join(base_path, relative_path)
        print(f"Resource Path: {resource_path}")  # Debugging line to check the path
        
        if not os.path.exists(resource_path):
            print(f"Error: Resource not found at {resource_path}")
        
        return resource_path
    except Exception as e:
        print(f"Error accessing resource: {e}")
        return None
    
root.iconbitmap(get_resource_path("assets.ico"))    

# Create or load the Excel file
def collect_input_data():
    collected_data = {}
    for field, entry in entries.items():
        # Check if entry is a combobox or text entry
        if isinstance(entry, ttk.Combobox):
            # If combobox is empty, set as "N/A"
            collected_data[field] = entry.get() or "N/A"
        else:
            # If text entry is empty or still has the placeholder, set as "N/A"
            entry_value = entry.get()
            placeholder = fields[field]
            collected_data[field] = entry_value if entry_value and entry_value != placeholder else "N/A"
    return collected_data

def create_or_load_file():
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        columns = [
            'Timestamp',
            'පළාත/Province', 
            'ප්‍රදේශ කේතය/Area Code', 
            'චලනය වූ දිනය/Date of move', 
            'තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No', 
            'KV / KV', 
            'KVA / KVA', 
            'ඉවත් කිරීමට හේතුව/Reason for removal', 
            'වර්තමාන තත්වය/Present Condition',
            'සිට කේතය/Movement From Cost Code',
            'සිට ස්ථානය/Movement From Location', 
            'සිට සංකේත අංකය/Movement From Sin No',
            'දක්වා කේතය/Movement To Cost Code',
            'දක්වා ස්ථානය/Movement To Location', 
            'දක්වා සංකේත අංකය/Movement To Sin No',
            'මාර්ග බිල්පතේ, පැවරුම් වවුචරයේ,\nභාණ්ඩ ලදුපතේ අංකය/Number of waybill,\ntransfer voucher, goods receipt',
            'කාර්ය අංකය/Task no',
            'සටහන/Remark'
        ]
        df = pd.DataFrame(columns=columns)
        df.to_excel(excel_file, index=False)

# Function to add data to the Excel sheet
        # Update add_data function to include automatic number and timestamp
province_code_map = {
    "Corp": [970.00, 971.00, 972.00, 972.03, 972.10, 972.11, 972.12, 972.20, 972.30, 973.00, 974.00, 975.00, 975.11, 975.20, 975.30, 976.00, 976.11, 977.00, 978.00],
    "WPS 1": [510.00, 510.10, 510.11, 510.14, 510.20, 510.30, 510.40, 510.70, 510.80, 510.81, 510.82, 510.83, 510.84, 510.85, 511.00],
    "SP 1": [520.00, 520.11, 520.12, 520.13, 520.14, 520.20, 520.30, 520.40, 520.60, 520.61, 520.62, 520.63],
    "SP 2": [940.00, 940.10, 940.11, 940.14, 940.20, 940.30, 940.40, 940.60, 940.61, 940.62, 940.70]
}

def get_province(area_code):
    for province, codes in province_code_map.items():
        if area_code in codes:
            return province
    return "Unknown"

def update_area_codes(selected_province, area_combobox):
    area_combobox['values'] = province_code_map.get(selected_province, [])
    area_combobox.current(0) if area_combobox['values'] else area_combobox.set('')  # Set first value or clear if no values


def on_combobox_keypress(event, combobox, field):
    typed_text = combobox.get()

    # Check if there is any input
    if typed_text:
        # Find matching values based on the typed text
        matching_values = [item for item in fields[field] if typed_text.lower() in item.lower()]
        
        if matching_values:
            combobox['values'] = matching_values  # Update dropdown with matching values
            
            # Show the dropdown if there are matches
            combobox.event_generate('<Down>')  # Open dropdown
            combobox.current(0)  # Select the first match
        else:
            # Clear dropdown if no matches found
            combobox['values'] = []
    else:
        # Reset values when input is empty
        combobox['values'] = fields[field]  # Show all options when input is empty

def check_mandatory_fields():
    global entries  # Declare entries as global

    # Check if the required fields are filled
    province = entries['පළාත/Province'].get() != ""
    area_code = entries['ප්‍රදේශ කේතය/Area Code'].get() != ""
    move_date = entries['චලනය වූ දිනය/Date of move'].get() != "" and entries['චලනය වූ දිනය/Date of move'].get() != "Enter Date"
    transformer_serial_no = entries['තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No'].get() != "" and \
                            entries['තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No'].get() != "තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No"

    # Enable or disable the add button based on the field status
    if province and area_code and move_date and transformer_serial_no:
        add_btn.config(state='normal')  # Enable button
        print("check_mandatory_fields", 'normal') 
    else:
        add_btn.config(state='disabled')  # Disable button
        print("check_mandatory_fields", 'disabled') 

def on_add_or_update_click():
    global searched_entry_index
    data = collect_input_data()

    if add_btn['text'] == "Add":
        add_data(data)  # Add new entry
    elif add_btn['text'] == "Update":
        update_data(data)  # Update existing entry

def toggle_buttons():
    global add_btn, update_btn, delete_btn, cancel_btn, is_update_mode

    if is_update_mode:
        if add_btn:
            add_btn.grid_remove()
        if clear_btn:
            clear_btn.grid_remove()    
        if update_btn:
            update_btn.grid(row=row_num, column=0, sticky="w", pady=10, padx=10)
        if delete_btn:
            delete_btn.grid(row=row_num, column=1, sticky="w", pady=10, padx=10)
        if cancel_btn:
            cancel_btn.grid(row=row_num, column=2, sticky="w", pady=10, padx=10)
    else:
        if update_btn:
            update_btn.grid_remove()
        if delete_btn:
            delete_btn.grid_remove()
        if cancel_btn:
            cancel_btn.grid_remove()
        if add_btn:
            add_btn.grid(row=row_num + 2, column=0, padx=(10, 5), pady=10, sticky="e")
            add_btn.config(state=tk.NORMAL)
        if clear_btn:
            clear_btn.grid(row=row_num + 2, column=1, padx=(5, 10), pady=10, sticky="w")
            clear_btn.config(state=tk.NORMAL)    

def add_data(data):
    try:
        df = pd.read_excel(excel_file)
        new_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Add timestamp
        data['Timestamp'] = new_timestamp
        new_data = pd.DataFrame([data], columns=df.columns)
        df = pd.concat([df, new_data], ignore_index=True)
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Success", "Data added successfully!")
        preview_all_entries()  # Refresh the preview of all entries
        update_last_entry_preview()
        #reset_form()  # Clear the form after adding
        preview_all_entries()  # Refresh all entries
        check_mandatory_fields()
    except Exception as e:
        messagebox.showerror("Error", f"Error while adding data: {str(e)}")

def update_data(data):
    print("update_data",data) 
    global searched_entry_index, is_update_mode
    if is_update_mode:
        delete_searched_entry()  # Delete the searched entry first
        add_data(data)  # Then add the new data
        #add_btn.config(text="Add Data")  # Reset button text after update
        is_update_mode = False  # Reset update mode
        toggle_buttons()
        #reset_form()
    else:
        add_data(collect_input_data())  # Regular add function if not in update mode

def cancel_action():
    global is_update_mode
    if is_update_mode:
        # If in update mode, cancel the update and reset the form
        is_update_mode = False
        toggle_buttons()
        #reset_form()        

def reset_form():
    
    for entry in entries.values():
        entry.delete(0, tk.END)
    add_btn.config(text="Add")  # Reset the button to "Add"

def edit_data():
    global last_entry_index, add_btn  # Ensure add_btn is recognized in this scope
    try:
        if last_entry_index is None:
            messagebox.showwarning("No Entry", "Please select an entry to edit.")
            return

        df = pd.read_excel(excel_file)
        selected_data = df.iloc[last_entry_index]

        # Populate input fields with the selected data, excluding 'Timestamp'
        for field in fields:
            if field != 'Timestamp':  # Exclude the timestamp field
                entries[field].delete(0, tk.END)
                entries[field].insert(0, selected_data[field])

        add_btn.config(text="Update", command=update_data)

    except Exception as e:
        messagebox.showerror("Error", f"Error while trying to edit: {str(e)}")


# Function to preview last entry
# Function to preview last entry
# Function to preview last entry
last_entry_data = None  # Global variable to store last entry values

def update_last_entry_preview(index=None):
    global last_entry_index, last_entry_data  # Add last_entry_data to globals

    try:
        df = pd.read_excel(excel_file)
        if index is not None:
            last_entry_index = index  # Update the last entry index

        # Get the last entry based on the last_entry_index
        last_entry = df.tail(last_entry_index + 1).head(1) if last_entry_index is not None else df.tail(1)
        last_entry_text.delete(1.0, tk.END)  # Clear previous content

        # Store the last entry values in a dictionary
        last_entry_data = {col: last_entry.iloc[0][col] for col in df.columns}

        for idx, col in enumerate(df.columns, 1):  # Add numbering (starts from 1)
            field_name = f"{idx}. {col}: "  # Numbered field name
            field_value = f"{last_entry.iloc[0][col]}"

            # Insert numbered field name in normal text
            last_entry_text.insert(tk.END, field_name)
            last_entry_text.tag_add('field_name_style', 'end-2c', 'end-1c')
            last_entry_text.tag_configure('field_name_style', font=("Nirmala UI", 8, "normal"))

            # Insert field value in bold text
            last_entry_text.insert(tk.END, field_value + "\n")
            last_entry_text.tag_add('field_value_style', 'end-2c', 'end-1c')
            last_entry_text.tag_configure('field_value_style', font=("Nirmala UI", 12, "bold"))

    except Exception as e:
        messagebox.showerror("Error", f"Error while retrieving last entry: {str(e)}")

def back_last_entry():
    global last_entry_index
    try:
        df = pd.read_excel(excel_file)
        if last_entry_index is None:
            last_entry_index = len(df) - 1  # Start from the last entry if none selected
        elif last_entry_index > 0:
            last_entry_index -= 1  # Move to the previous entry
        else:
            messagebox.showinfo("Info", "This is the first entry.")  # No previous entry
        update_last_entry_preview()  # Update the preview with the new index
    except Exception as e:
        messagebox.showerror("Error", f"Error while navigating back: {str(e)}")

def forward_last_entry():
    global last_entry_index
    try:
        df = pd.read_excel(excel_file)
        if last_entry_index is None:
            last_entry_index = 0  # Start from the first entry if none selected
        elif last_entry_index < len(df) - 1:
            last_entry_index += 1  # Move to the next entry
        else:
            messagebox.showinfo("Info", "This is the last entry.")  # No next entry
        update_last_entry_preview()  # Update the preview with the new index
    except Exception as e:
        messagebox.showerror("Error", f"Error while navigating forward: {str(e)}")

def reset_last_entry():
    global last_entry_index
    last_entry_index = 0  # Reset to the first entry
    update_last_entry_preview()  # Update the preview

# Function to preview all entries in a scrollable area
# Modify the preview_all_entries function to color the rows and display in the desired format
def preview_all_entries():
    try:
        df = pd.read_excel(excel_file)
        all_entries_text.delete(1.0, tk.END)  # Clear previous content
        
        for i, row in df.iterrows():
            row_color = "#D0FFD0" if i % 2 == 0 else "#E0FFE0"  # Alternate row colors

            for col in df.columns:
                field_name = f"{col}: "
                field_value = f"{row[col]}"

                # Insert field name in normal text
                all_entries_text.insert(tk.END, field_name)
                all_entries_text.tag_add('field_name_style', 'end-2c', 'end-1c')
                all_entries_text.tag_configure('field_name_style', font=("Nirmala UI", 8, "normal"))

                # Insert field value in bold text
                all_entries_text.insert(tk.END, field_value + " | ")
                all_entries_text.tag_add('field_value_style', 'end-2c', 'end-1c')
                all_entries_text.tag_configure('field_value_style', font=("Nirmala UI", 12, "bold"))

            # Add a newline after each row
            all_entries_text.insert(tk.END, "\n")
            all_entries_text.tag_add(f'row_style_{i}', all_entries_text.index('end-2c'), all_entries_text.index('end-1c'))
            all_entries_text.tag_configure(f'row_style_{i}', background=row_color)

    except Exception as e:
        messagebox.showerror("Error", f"Error while retrieving all entries: {str(e)}")
# Function to search data

# Function to fill input fields
def fill_input_fields(row):
    for field in fields.keys():
        entries[field].delete(0, tk.END)  # Clear the entry field
        entries[field].insert(0, row[field])  # Fill with the relevant data

# Function to filter data by column and value
def filter_data(column, value):
    try:
        df = pd.read_excel(excel_file)
        filtered_results = df[df[column].astype(str).str.contains(value, case=False)]
        if not filtered_results.empty:
            messagebox.showinfo("Filtered Results", filtered_results.to_string(index=False))
        else:
            messagebox.showinfo("No Results", f"No data found for {column}: {value}.")
    except Exception as e:
        messagebox.showerror("Error", f"Error while filtering data: {str(e)}")

# Function to sort data by column
def sort_data(column):
    try:
        df = pd.read_excel(excel_file)
        sorted_df = df.sort_values(by=column)
        messagebox.showinfo("Sorted Results", sorted_df.to_string(index=False))
    except Exception as e:
        messagebox.showerror("Error", f"Error while sorting data: {str(e)}")

def export_data():
    try:
        df = pd.read_excel(excel_file)
        # Create 'export' directory at the same level as the script if it doesn't exist
        export_dir = os.path.join(os.path.dirname(__file__), '../export')
        os.makedirs(export_dir, exist_ok=True)
        
        # Generate timestamped filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_file = os.path.join(export_dir, f"transformer_data_{timestamp}.xlsx")
        
        df.to_excel(export_file, index=False)
        messagebox.showinfo("Export Success", f"Data exported successfully to {export_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error while exporting data: {str(e)}")

    try:
        df = pd.read_excel(excel_file)
        # Create 'export' directory if it doesn't exist
        export_dir = 'export'
        os.makedirs(export_dir, exist_ok=True)
        
        # Generate timestamped filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_file = os.path.join(export_dir, f"transformer_data_{timestamp}.xlsx")
        
        df.to_excel(export_file, index=False)
        messagebox.showinfo("Export Success", f"Data exported successfully to {export_file}")
    except Exception as e:
        messagebox.showerror("Error", f"Error while exporting data: {str(e)}")
# Tkinter GUI setup

# Function to open the export directory in File Explorer
def open_export_folder():
    export_dir = os.path.join(os.path.dirname(__file__), '../export')
    os.makedirs(export_dir, exist_ok=True)  # Ensure the directory exists
    os.startfile(export_dir)  # Open the folder in File Explorer
# Add buttons for editing and deleting entries
# Add buttons for editing and deleting entries

def export_and_quit():
    export_data()
    root.quit()

def search_data(search_query):
    global searched_entry_index, is_update_mode  # Use global to track the searched entry
    print(f"Searching for query: '{search_query}'")  # Debug print
    try:
        df = pd.read_excel(excel_file)
        print(f"Loaded DataFrame with {len(df)} entries")  # Debug print

        if not search_query.strip():
            messagebox.showwarning("Empty Search Query", "Please enter a search query.")
            return

        matching_rows = df[df.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)]
        print(f"Found {len(matching_rows)} matching rows")  # Debug print

        if not matching_rows.empty:
            searched_entry_index = matching_rows.index[0]
            fill_input_fields(matching_rows.iloc[0])
            messagebox.showinfo("Search Success", f"Entry found for '{search_query}'.")

            # Set update mode to True and toggle buttons
            is_update_mode = True
            print("Update mode activated", is_update_mode)  # Debug print
            toggle_buttons()
            print("Update mode activated", is_update_mode)  # Debug print
        else:
            searched_entry_index = None
            messagebox.showinfo("No Results", f"No entry found for '{search_query}'.")
            print("No matching entries found.")  # Debug print

    except Exception as e:
        print(f"Error while searching: {str(e)}")  # Debug print
        messagebox.showerror("Error", f"Error while searching: {str(e)}")


def delete_searched_entry():
    global searched_entry_index
    try:
        if searched_entry_index is None:
            messagebox.showwarning("No Entry Selected", "Please search for an entry first.")
            return

        df = pd.read_excel(excel_file)
        df.drop(searched_entry_index, inplace=True)
        df.to_excel(excel_file, index=False)
        messagebox.showinfo("Success", "Searched entry deleted successfully!")
        searched_entry_index = None
        preview_all_entries()  # Refresh the preview of all entries
        update_last_entry_preview()
    except Exception as e:
        messagebox.showerror("Error", f"Error while deleting searched entry: {str(e)}")

def delete_data():
    global last_entry_data  # Use the global variable for the last entry data
    try:
        if last_entry_data is not None:
            df = pd.read_excel(excel_file)  # Read the current data from the Excel file

            # Find the row in the DataFrame that matches the last entry
            found_row = df[df.apply(lambda row: all(str(row[col]).strip() == str(last_entry_data[col]).strip() for col in last_entry_data), axis=1)]

            if not found_row.empty:
                # Remove the selected row
                df = df[~df.apply(lambda row: all(str(row[col]).strip() == str(last_entry_data[col]).strip() for col in last_entry_data), axis=1)]

                # Write the updated DataFrame back to the Excel file
                df.to_excel(excel_file, index=False)

                messagebox.showinfo("Success", "Data deleted successfully!")  # Show success message
                update_last_entry_preview()
                preview_all_entries()  # Refresh the preview
            else:
                messagebox.showwarning("Entry Not Found", "The selected entry could not be found in the data.")
        else:
            messagebox.showwarning("Select Entry", "Please select an entry to delete.")
    except Exception as e:
        messagebox.showerror("Error", f"Error while trying to delete: {str(e)}")

def fill_input_fields(row):
    for field in fields.keys():
        entries[field].delete(0, tk.END)  # Clear the entry field
        entries[field].insert(0, row[field])  # Fill with the relevant data

def fill_input_fields(row):
    for field in fields.keys():
        entries[field].delete(0, tk.END)  # Clear the entry field
        entries[field].insert(0, row[field])  # Fill with the relevant data

# Ensure to clear input fields on start

def setup_gui():
    global entries, fields  # Make entries and fields global for accessibility
    global font_sinhala
    global add_btn, update_btn, delete_btn ,cancel_btn,clear_btn
    row_num = 1
    font_sinhala = ("Nirmala UI", 12)
    
    fields = {
        'පළාත/Province': ["Corp","WPS 1","SP 1","SP 2"],
       'ප්‍රදේශ කේතය/Area Code': [
                        '970.00', '971.00', '972.00', '972.03', '972.10', '972.11', '972.12', '972.20', '972.30',
                        '562.00', '973.00', '974.00', '975.00', '975.11', '975.20', '975.30', '976.00', '976.11',
                        '977.00', '978.00', '510.00', '510.10', '510.11', '510.14', '510.20', '510.30', '510.40',
                        '510.70', '510.80', '510.81', '510.82', '510.83', '510.84', '510.85', '511.00', '511.03',
                        '511.10', '511.20', '511.30', '511.40', '511.90', '514.00', '514.03', '514.10', '514.20',
                        '514.30', '514.40', '514.50', '514.80', '514.90', '515.00', '515.03', '515.10', '515.20',
                        '515.30', '515.50', '515.90', '519.00', '519.03', '519.10', '519.20', '519.30', '519.40',
                        '519.90', '520.00', '520.11', '520.12', '520.13', '520.14', '520.20', '520.30', '520.40',
                        '520.60', '520.61', '520.62', '520.63', '520.70', '521.00', '521.03', '521.10', '521.20',
                        '521.40', '521.50', '521.70', '522.00', '522.03', '522.10', '522.20', '522.40', '522.50',
                        '522.60', '523.00', '523.03', '523.10', '523.40', '523.60', '523.70', '523.80', '524.00',
                        '524.03', '524.10', '524.20', '524.30', '524.40', '525.00', '525.03', '525.10', '525.20',
                        '525.30', '525.40', '526.00', '526.03', '526.10', '526.20', '526.30', '526.40', '526.50',
                        '527.00', '527.03', '527.10', '527.20', '527.30', '527.40', '940.00', '940.10', '940.11',
                        '940.14', '940.20', '940.30', '940.40', '940.60', '940.61', '940.62', '940.70', '941.00',
                        '941.03', '941.10', '941.20', '941.30', '941.40', '941.50', '942.00', '942.03', '942.10',
                        '942.20', '942.30', '943.00', '943.03', '943.10', '943.20', '943.30', '943.40', '943.50',
                        '944.00', '944.10', '944.20', '971.10'
                    ]

,
        'චලනය වූ දිනය/Date of move': "Enter Date (DD/MM/YYYY)",
        'තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No': "තාරා පැවි අනුක්‍රමික අංකය/Transformer Serial No",
        'KV / KV': "KV / KV",
        'KVA / KVA': "KVA / KVA",
        'ඉවත් කිරීමට හේතුව/Reason for removal': "ඉවත් කිරීමට හේතුව/Reason for removal",
        'වර්තමාන තත්වය/Present Condition': "වර්තමාන තත්වය/Present Condition",
        'සිට කේතය/Movement From Cost Code': [
                        '970.00', '971.00', '972.00', '972.03', '972.10', '972.11', '972.12', '972.20', '972.30',
                        '562.00', '973.00', '974.00', '975.00', '975.11', '975.20', '975.30', '976.00', '976.11',
                        '977.00', '978.00', '510.00', '510.10', '510.11', '510.14', '510.20', '510.30', '510.40',
                        '510.70', '510.80', '510.81', '510.82', '510.83', '510.84', '510.85', '511.00', '511.03',
                        '511.10', '511.20', '511.30', '511.40', '511.90', '514.00', '514.03', '514.10', '514.20',
                        '514.30', '514.40', '514.50', '514.80', '514.90', '515.00', '515.03', '515.10', '515.20',
                        '515.30', '515.50', '515.90', '519.00', '519.03', '519.10', '519.20', '519.30', '519.40',
                        '519.90', '520.00', '520.11', '520.12', '520.13', '520.14', '520.20', '520.30', '520.40',
                        '520.60', '520.61', '520.62', '520.63', '520.70', '521.00', '521.03', '521.10', '521.20',
                        '521.40', '521.50', '521.70', '522.00', '522.03', '522.10', '522.20', '522.40', '522.50',
                        '522.60', '523.00', '523.03', '523.10', '523.40', '523.60', '523.70', '523.80', '524.00',
                        '524.03', '524.10', '524.20', '524.30', '524.40', '525.00', '525.03', '525.10', '525.20',
                        '525.30', '525.40', '526.00', '526.03', '526.10', '526.20', '526.30', '526.40', '526.50',
                        '527.00', '527.03', '527.10', '527.20', '527.30', '527.40', '940.00', '940.10', '940.11',
                        '940.14', '940.20', '940.30', '940.40', '940.60', '940.61', '940.62', '940.70', '941.00',
                        '941.03', '941.10', '941.20', '941.30', '941.40', '941.50', '942.00', '942.03', '942.10',
                        '942.20', '942.30', '943.00', '943.03', '943.10', '943.20', '943.30', '943.40', '943.50',
                        '944.00', '944.10', '944.20', '971.10'
                    ],
        'සිට ස්ථානය/Movement From Location': "සිට ස්ථානය/Movement From Location",
        'සිට සංකේත අංකය/Movement From Sin No': "සිට සංකේත අංකය/Movement From Sin No",
        'දක්වා කේතය/Movement To Cost Code': [
                        '970.00', '971.00', '972.00', '972.03', '972.10', '972.11', '972.12', '972.20', '972.30',
                        '562.00', '973.00', '974.00', '975.00', '975.11', '975.20', '975.30', '976.00', '976.11',
                        '977.00', '978.00', '510.00', '510.10', '510.11', '510.14', '510.20', '510.30', '510.40',
                        '510.70', '510.80', '510.81', '510.82', '510.83', '510.84', '510.85', '511.00', '511.03',
                        '511.10', '511.20', '511.30', '511.40', '511.90', '514.00', '514.03', '514.10', '514.20',
                        '514.30', '514.40', '514.50', '514.80', '514.90', '515.00', '515.03', '515.10', '515.20',
                        '515.30', '515.50', '515.90', '519.00', '519.03', '519.10', '519.20', '519.30', '519.40',
                        '519.90', '520.00', '520.11', '520.12', '520.13', '520.14', '520.20', '520.30', '520.40',
                        '520.60', '520.61', '520.62', '520.63', '520.70', '521.00', '521.03', '521.10', '521.20',
                        '521.40', '521.50', '521.70', '522.00', '522.03', '522.10', '522.20', '522.40', '522.50',
                        '522.60', '523.00', '523.03', '523.10', '523.40', '523.60', '523.70', '523.80', '524.00',
                        '524.03', '524.10', '524.20', '524.30', '524.40', '525.00', '525.03', '525.10', '525.20',
                        '525.30', '525.40', '526.00', '526.03', '526.10', '526.20', '526.30', '526.40', '526.50',
                        '527.00', '527.03', '527.10', '527.20', '527.30', '527.40', '940.00', '940.10', '940.11',
                        '940.14', '940.20', '940.30', '940.40', '940.60', '940.61', '940.62', '940.70', '941.00',
                        '941.03', '941.10', '941.20', '941.30', '941.40', '941.50', '942.00', '942.03', '942.10',
                        '942.20', '942.30', '943.00', '943.03', '943.10', '943.20', '943.30', '943.40', '943.50',
                        '944.00', '944.10', '944.20', '971.10'
                    ],
        'දක්වා ස්ථානය/Movement To Location': "දක්වා ස්ථානය/Movement To Location",
        'දක්වා සංකේත අංකය/Movement To Sin No': "දක්වා සංකේත අංකය/Movement To Sin No",
        'මාර්ග බිල්පතේ, පැවරුම් වවුචරයේ,\nභාණ්ඩ ලදුපතේ අංකය/Number of waybill,\ntransfer voucher, goods receipt': "මාර්ග බිල්පතේ, පැවරුම් වවුචරයේ,\nභාණ්ඩ ලදුපතේ අංකය/Number of waybill,\ntransfer voucher, goods receipt",
        'කාර්ය අංකය/Task no': "කාර්ය අංකය/Task no",
        'සටහන/Remark': "සටහන/Remark"
    }

    
    #root = tk.Tk()
    #root.title("Transformer Data Entry")
    #root.geometry("800x600")  # Set a fixed size for the window
        # Load the icon images
    add_icon = ImageTk.PhotoImage(Image.open(get_resource_path("add_icon.png")).resize((20, 20)))
    update_icon = ImageTk.PhotoImage(Image.open(get_resource_path("update_icon.png")).resize((20, 20)))
    delete_icon = ImageTk.PhotoImage(Image.open(get_resource_path("delete_icon.png")).resize((20, 20)))
    cancel_icon = ImageTk.PhotoImage(Image.open(get_resource_path("cancel_icon.png")).resize((20, 20)))
    terminate_icon = ImageTk.PhotoImage(Image.open(get_resource_path("terminate_icon.png")).resize((20, 20)))
    Export_icon = ImageTk.PhotoImage(Image.open(get_resource_path("Export_icon.png")).resize((20, 20)))
    Export_View_icon = ImageTk.PhotoImage(Image.open(get_resource_path("Export_View_icon.png")).resize((20, 20)))
    Search_icon = ImageTk.PhotoImage(Image.open(get_resource_path("Search_icon.png")).resize((20, 20)))

    
    # Load background image
    bg_image = tk.PhotoImage(file=get_resource_path("BG.png"))  # For background image
    transformers_image = ImageTk.PhotoImage(Image.open(get_resource_path("Transformers.jpg")))  # For logo  
    bg_label = tk.Label(root, image=bg_image)
    bg_label.place(relwidth=1, relheight=1)  # Set background to cover the entire window

    

        # Create the left frame and set transformers image as background
    left_frame = tk.Frame(root, padx=10, pady=10)  # No bg color to avoid overlay
    left_frame.grid(row=0, column=0, sticky="nsew")

    # Label for the left frame background image
    left_bg_label = tk.Label(left_frame, image=transformers_image)
    left_bg_label.place(relwidth=1, relheight=1)  # Cover the entire left frame with the image


    # Create a frame for right section (Searching, sorting, and previewing all data)
    right_frame = tk.Frame(root, padx=10, pady=10)  # Remove bg color for the frame
    right_frame.grid(row=0, column=1, sticky="nsew")

    # Create a label to hold the background image for the right frame
    right_bg_label = tk.Label(right_frame, image=bg_image)
    right_bg_label.place(relwidth=1, relheight=1)  # Set to cover the entire right frame

    # Adjust grid weights to allow resizing
    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)

     # Adding a static instruction label for date entry
    # Adding a static instruction label for date entry
    #date_instruction_label = tk.Label(root, text="Please enter a full date in DD/MM/YYYY format.", fg="red")
    #date_instruction_label.grid(row=row_num + 1, columnspan=2)  # Use grid instead of pack


    def add_placeholder(entry, placeholder):
        entry.insert(0, placeholder)  # Add placeholder text
        entry.config(fg='gray')  # Set placeholder text color to gray
        entry.bind("<FocusIn>", lambda e: clear_placeholder(entry, placeholder))  # Clear on focus
        entry.bind("<FocusOut>", lambda e: set_placeholder(entry, placeholder))  # Set on focus out

    def clear_placeholder(entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)  # Clear placeholder
            entry.config(fg='black')  # Change text color to black

    def set_placeholder(entry, placeholder):
        if entry.get() == "":
            entry.insert(0, placeholder)  # Reset placeholder if empty
            entry.config(fg='gray')  # Change text color back to gray

    def format_and_validate_date(event, entry):
        current_text = entry.get()
        cleaned_text = ''.join(filter(str.isdigit, current_text))
        
        # Only format if we have at least 2 digits for day or 4 for day and month
        if len(cleaned_text) <= 7: 
            return  # Don't reformat yet if there's only one digit

        formatted_text = ""

        # Construct the formatted text based on the length of cleaned_text
        if len(cleaned_text) >= 2:
            formatted_text += cleaned_text[:2]  # Day
        if len(cleaned_text) >= 4:
            formatted_text += "/" + cleaned_text[2:4]  # Month
        if len(cleaned_text) >= 6:
            formatted_text += "/" + cleaned_text[4:8]  # Year

        # Update the entry field with the formatted text
        entry.delete(0, tk.END)
        entry.insert(0, formatted_text)

        # Cancel the previous validation if it's already scheduled
        if hasattr(entry, 'validation_job'):
            entry.after_cancel(entry.validation_job)

        # Schedule a new validation after a delay
        entry.validation_job = entry.after(500, lambda: validate_date(cleaned_text, entry))

    def validate_date(cleaned_text, entry):
        if len(cleaned_text) < 8:  # Full date must be present
            return  # Do not validate yet; just return

        day = int(cleaned_text[:2])
        month = int(cleaned_text[2:4])
        year = int(cleaned_text[4:8])

        # Basic validation for month, day, and year
        if month < 1 or month > 12 or day < 1 or day > 31 or year < 2000:
            show_invalid_date_message(entry)
            return

        # Check for the correct number of days in month
        if month in [4, 6, 9, 11] and day > 30:
            show_invalid_date_message(entry)
            return
        if month == 2:
            if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):  # Leap year
                if day > 29:
                    show_invalid_date_message(entry)
                    return
            else:
                if day > 28:
                    show_invalid_date_message(entry)
                    return

        print("Valid date!")

    def show_invalid_date_message(entry):
        messagebox.showerror("Invalid Date", "Please enter a valid date in the format DD/MM/YYYY.")
        entry.delete(0, tk.END)

    def restrict_input(event, entry):
        if event.char not in '0123456789/':
            if event.keysym == "Tab":  # Allow Tab, trigger date validation on Tab
                format_and_validate_date(event, entry)
                return
            if event.keysym != "BackSpace":  # Allow Backspace
                return "break"  # Prevent the character from being inserted


    # Left Section
    title_label = tk.Label(left_frame, text="Add Transformer Data/ඇතුළත් කරන්න", font=(font_sinhala, 18, "bold"), bg="#c6d3f5", fg="#003366")  # Changed font style and color
    title_label.grid(row=0, column=0, columnspan=2, pady=10)

    entries = {}
    row_num = 1

    for index, field in enumerate(fields):
        numbered_label = f"{index + 1}. {field}"
        tk.Label(left_frame, text=numbered_label, bg="#003366",fg="#fff", font=font_sinhala).grid(row=row_num, column=0, sticky="w", pady=5)

        if isinstance(fields[field], list):
            combobox = ttk.Combobox(left_frame, values=fields[field], state='normal', width=40)
            combobox.grid(row=row_num, column=1, padx=5, pady=5)

            # Bind the keyrelease event to the combobox
            combobox.bind("<KeyRelease>", lambda event, c=combobox, f=field: (on_combobox_keypress(event, c, f), check_mandatory_fields()))
            
            entries[field] = combobox

            # Also bind the focus event to check when user interacts
            combobox.bind("<FocusOut>", lambda event, f=field: check_mandatory_fields())
        else:
            entry = tk.Entry(left_frame, width=40, bd=2, relief="groove")
            add_placeholder(entry, fields[field])
            entry.grid(row=row_num, column=1, padx=5, pady=5)
            entries[field] = entry

            # Bind key release and focus events
            entry.bind("<KeyRelease>", lambda event, f=field: check_mandatory_fields())
            entry.bind("<FocusOut>", lambda event, f=field: check_mandatory_fields())
            
            if field == 'චලනය වූ දිනය/Date of move':
                entry.bind("<KeyPress>", lambda event, e=entry: restrict_input(event, e)) # Bind to restrict non-numeric input
                entry.bind("<KeyRelease>", lambda event, e=entry: format_and_validate_date(event, e))

        row_num += 1

    def clear_form():
        for field, entry in entries.items():
            if isinstance(entry, ttk.Combobox):
                entry.set('')  # Clear comboboxes
            else:
                entry.delete(0, tk.END)  # Clear the text in the entry
                add_placeholder(entry, fields[field])  # Reset to placeholder

    # Add and Clear buttons in the same row with spacing between them
    add_btn = tk.Button(left_frame, text="Add Data", font=("Nirmala UI", 12, "bold"), bg="#4682b4", fg="white", 
                        image=add_icon, compound="left", padx=10, command=lambda: add_data(collect_input_data()))
    add_btn.grid(row=row_num + 2, column=0, padx=(10, 5), pady=10, sticky="e")  # Place "Add Data" button in column 0 with padding
    add_btn.config(state='disabled')  # Initially disabled

    clear_btn = tk.Button(left_frame, text="Clear Form", font=("Nirmala UI", 12, "bold"), bg="gray", fg="white",  image=cancel_icon, compound="left",
                        padx=10, command=clear_form)
    clear_btn.grid(row=row_num + 2, column=1, padx=(5, 10), pady=10, sticky="w")  # Place "Clear Form" button in column 1 with padding


    # Call to set button state based on initial values
    check_mandatory_fields()

    update_btn = tk.Button(left_frame, text="Update Data", font=("Nirmala UI", 12, "bold"), bg="#32CD32", fg="white", 
                          image=update_icon, compound="left",  padx=10, command=lambda: update_data(collect_input_data()))
    update_btn.grid_forget()  # Initially hide the Update button

    # Delete Searched Entry Button (new)
    delete_btn = tk.Button(left_frame, text="Delete Entry", font=(font_sinhala, 12, "bold"), bg="red", fg="white",
                        image=delete_icon, compound="left",  padx=10,command=delete_searched_entry)
    delete_btn.grid_forget()  # Initially hide the Delete button
    
    def cancel_and_clear():
        clear_form()      # Call the clear_form function
        cancel_action()   # Call the cancel_action function

    # Cancel Button (new)
    cancel_btn = tk.Button(left_frame, text="Cancel", font=("Nirmala UI", 12, "bold"), bg="gray", fg="white",
                        image=cancel_icon, compound="left",  padx=10,command=cancel_and_clear)  # Add a command to handle cancellation
    cancel_btn.grid_forget()  # Initially hide the Cancel button

    # Add Edit and Delete buttons
    button_frame = tk.Frame(left_frame, bg="lightgreen")
    button_frame.grid(row=row_num + 1, column=0, columnspan=2, pady=10)

  

    # Right Section
    right_title_label = tk.Label(right_frame, text="Additional Finance (DD4)", font=(font_sinhala, 24, "bold"), bg="white", fg="#003366")  # Title label
    right_title_label.grid(row=0, column=0, columnspan=2, pady=(10, 5))  # Adjusted padding

    # Add Logo as a local PNG image
    logo_path = get_resource_path("LOGO.png")  # Get the path to the logo image
    if logo_path:
        logo_image = Image.open(logo_path)  # Open the image with Pillow
        logo_image = logo_image.resize((400, 80), Image.LANCZOS)  # Resize the image
        logo_photo = ImageTk.PhotoImage(logo_image)  # Convert to PhotoImage for Tkinter

        logo_label = tk.Label(right_frame, image=logo_photo, bg="white")  # Create the label with the logo
        logo_label.image = logo_photo  # Keep a reference to the image to prevent garbage collection
        logo_label.grid(row=1, column=0, columnspan=2, pady=(0, 10))

    # Search Bar
    search_entry = tk.Entry(right_frame, width=30, bd=2, relief="groove", font=("Nirmala UI", 12))  # Increased font size for better visibility
    search_entry.grid(row=2, column=0, padx=(10, 5), pady=5)  # Adjusted padding
    search_btn = tk.Button(right_frame, text="Search", font=(font_sinhala, 12, "bold"), bg="#4682b4", fg="white", 
                        image=Search_icon, compound="left",  padx=10,command=lambda: search_data(search_entry.get()))
    search_btn.grid(row=2, column=1, padx=(5, 10), pady=5)  # Adjusted padding

   # Preview Last Entry Section
    last_entry_label = tk.Label(right_frame, text="Last Entry Preview / අවසන් ඇතුළත් කරුනු", font=("Nirmala UI", 14, "bold"), bg="#c6d3f5", fg="#003366")  # Label for last entry preview
    last_entry_label.grid(row=3, column=0, columnspan=2, pady=(10, 5))  # Adjusted row and padding

    global last_entry_text
    last_entry_text = tk.Text(right_frame, height=15, width=70, bd=2, relief="groove", bg="lightgray", font=("Nirmala UI", 10))  # Text area for last entry preview
    last_entry_text.grid(row=4, column=0, columnspan=2, padx=10, pady=5)
    
   # Navigation Buttons
    button_frame = tk.Frame(right_frame, bg="white")
    button_frame.grid(row=5, column=0, columnspan=2, pady=(10, 5))

    back_button = tk.Button(button_frame, text="← Back", command=forward_last_entry, bg="#007BFF", fg="white")
    back_button.grid(row=0, column=0, padx=5)

    forward_button = tk.Button(button_frame, text="Forward", command=back_last_entry, bg="#007BFF", fg="white")
    forward_button.grid(row=0, column=1, padx=5)

    last_button = tk.Button(button_frame, text="Last", command=reset_last_entry, bg="#007BFF", fg="white")
    last_button.grid(row=0, column=2, padx=5)

    delete_btn1 = tk.Button(button_frame, text="Delete", command=delete_data, bg="red", fg="white")
    delete_btn1.grid(row=0, column=3, padx=10)  # Place delete_btn in the next column

    # Adding an edit button to edit the last previewed entry
    edit_btn = tk.Button(button_frame, text="Edit", command=edit_data, bg="#007BFF", fg="white")
    edit_btn.grid(row=0, column=4, padx=10)  # Place edit_btn in the next column

    # Preview All Entries Section
    preview_all_label = tk.Label(right_frame, text="All Entries Preview / සියලුම ඇතුළත් කරුනු", font=("Nirmala UI", 14, "bold"), bg="#c6d3f5", fg="#003366")  # Label for all entries preview
    preview_all_label.grid(row=6, column=0, columnspan=2, pady=(10, 5))  # Adjusted row and padding

    global all_entries_text
    all_entries_text = tk.Text(right_frame, height=5, width=70, bd=2, relief="groove", bg="lightgray", font=("Nirmala UI", 10))  # Text area for all entries preview
    all_entries_text.grid(row=7, column=0, columnspan=2, padx=10, pady=5)

    # Export Button
    export_btn = tk.Button(right_frame, text="Export Data", font=("Nirmala UI", 12, "bold"), bg="#04b037", fg="white", 
                        image=Export_icon, compound="left",  padx=10,command=export_data)
    export_btn.grid(row=8, column=0, padx=(10, 5), pady=(10, 5))  # Adjusted padding

        # View Exports Button
    view_exports_btn = tk.Button(right_frame, text="View Exports", font=("Nirmala UI", 12, "bold"), bg="#08c4c1", fg="white",  image=Export_View_icon, compound="left",  padx=10,command=open_export_folder)
    view_exports_btn.grid(row=8, column=1, padx=(5, 10), pady=(10, 5))


    # Add guidance text for export
    #guidance_label = tk.Label(right_frame, text="Click to export the data to Excel", font=("Nirmala UI", 10), bg="lightblue")
    #guidance_label.grid(row=8, column=2, padx=(5, 10), pady=(10, 5))  # Adjusted padding

    # Terminate Application Button
    terminate_btn = tk.Button(right_frame, text="Exit", font=("Nirmala UI", 12, "bold"), bg="red", fg="white", image=terminate_icon, compound="left",  padx=10,
                            command=export_and_quit)  # Call root.quit() to close the application
    terminate_btn.grid(row=9, column=0, columnspan=1, pady=(15, 5), padx=15)  # Adjusted padding

    # Watermark
    watermark_label = tk.Label(right_frame, text="Copyright Darshana Wijebahu©2024", bg="white", font=("Nirmala UI", 10, "italic"))
    watermark_label.grid(row=9, column=1, columnspan=3, pady=(5, 10))  # Adjusted padding

    # Initialize functions for data handling
    create_or_load_file()  # Call function to create or load Excel file
    update_last_entry_preview()  # Update last entry preview on startup
    preview_all_entries()  # Preview all entries on startup

    root.mainloop()
if __name__ == "__main__":
    setup_gui()