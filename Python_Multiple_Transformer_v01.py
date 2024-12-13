# -*- coding: utf-8 -*-
"""
Created on Tue Dec 10 17:49:40 2024

@author: Franz Guzman
"""

import pandas as pd
import mhi.pscad
import time

# Path to access the Excel file
excel_path = "E:/OneDrive/1. ESTUDIOS - RODRIGO/1. UNICAMP - M. INGENIERÍA ELÉCTRICA/0. Cursos y Capacitaciones/0. Curso_RTDS_2024_UNICAMP/1. Aula 01/Check_List_RTDS_3.xlsx"  # Replace with the path to your file

# Read a specific sheet from the Excel file
one_sheet = pd.read_excel(excel_path, sheet_name='TRANSFORMADORES') 
print(one_sheet)

# Start and end points for the table data
init_row, init_col = (7, 16)  # Starting point of the data table
end_row, end_col = (one_sheet.shape[0], one_sheet.shape[1]-6)  # End point of the data table
len_row = end_row - init_row
len_col = end_col - init_col
print("Dimension of one_sheet:", one_sheet.shape)  # Dimensions of rows and columns in the sheet
print("Starting point: ", init_row, " - ", init_col, "; Ending point: ", end_row, " - ", end_col)

# Read only specific columns from the sheet (columns 13 to 34)
input_cols = [i for i in range(init_col, end_col)]  # Range from column 13 to 34
print(input_cols)

# Load the data file again using the selected columns
dtype = { # Ensure dtype in 'float' are 'int'
    "YD1": "Int64", 
    "YD2": "Int64", 
    "YD3": "Int64", 
    "Lead": "Int64", 
    "Tap": "Int64", 
    "Dtls": "Int64", 
    "Ideal": "Int64", 
    "Enab": "Int64", 
    "Sat": "Int64", 
    "Hys": "Int64"
}
df = pd.read_excel(excel_path, sheet_name='TRANSFORMADORES', header=init_row, usecols=input_cols, dtype=dtype)  # Load data to read specific columns
print("Column names: ", df.columns)
print("New dimension of df: ", df.shape)

# =============================================================================
# create a list of column name's from SRC, necessary for load data in PSCAD
col_name_XMFR = ["Name", "YD1", "YD2", "YD3", "Lead", "Tap", "Dtls", "Xl12", "Xl13", "Xl23", 
             "CuL12", "CuL13", "CuL23", "Tmva", "f", "V1", "V2", "V3", "Ideal", "Enab", 
             "Sat", "Hys", "Xknee"]

# =============================================================================

with mhi.pscad.application() as pscad:

    # Load Workspace_Python workspace. Note: charge an especific worskpace .pswx
    # worskpace_python = pscad.load("E:\\OneDrive\\1. ESTUDIOS - RODRIGO\\1. UNICAMP - M. INGENIERÍA ELÉCTRICA\\0. Cursos y Capacitaciones\\1. Curso_PSCAD 2024_UNICAMP\\9. MeuPSCAD\\2. Python\\Python_Projects\\Workspace_Python.pswx")
    
    # Load Teste_01 case estudy. Note: charge an especific case study project .pscx
    # pscad.load("E:\\OneDrive\\1. ESTUDIOS - RODRIGO\\1. UNICAMP - M. INGENIERÍA ELÉCTRICA\\0. Cursos y Capacitaciones\\1. Curso_PSCAD 2024_UNICAMP\\9. MeuPSCAD\\2. Python\\Python_Projects\\Transformers.pscx")

    # Go to 'Transformers' case study
    Transformers = pscad.project("Transformers") # Go to case study 'Transformers'
    canvas_Transformers = Transformers.canvas("Main") # Go to canvas case study 'Transformers'
    Transformers.navigate_to() # Navigate to 'Transformers'
  
    # Reference coordinates of the element
    elements_in_line = 0  # Counter for the number of elements (Transformers) in a row
    x0_position, y0_position, theta = 7, 7, 0 # referene x0, y0 position
    delta_x, delta_y = 8, 10  # Spacing between elements (Transformers)
    x_position, y_position = x0_position, y0_position  # Initial position to create a Source in RSCAD

    # Depure headers
    col_name_load_str = list(map(str.strip, map(str, df.columns))) # Convert to string and remove spaces from columns in excel
    col_name_XMFR_str = list(map(str.strip, map(str, col_name_XMFR))) # Convert to string and remove spaces from 
    
    # Verify the values header from excel is equal to input data in PSCAD 
    state_data_load = True
    
    # Comparind inputs headers from excel with output headers from PSCAD 
    mismatches = [
    (index, header_XMFR, header_load)
    for index, (header_XMFR, header_load) in enumerate(zip(col_name_XMFR_str, col_name_load_str))
    if header_XMFR != header_load
    ]
    # Print errors for mismatches for excel with PSCAD
    if mismatches:
        state_data_load = False
        for index, header_XMFR, header_load in mismatches:
            print(f"__________________Mismatch at index {index}___________________")
            print(f"The column name's '{header_load}' is different to '{header_XMFR}' in excel")
            print(f"Please change the column name's in Excel from: '{header_load}' to -> '{header_XMFR}'")
            print("_________________________________________________________")
    else:
        print("Header verification passed!")
    
    # Continue if all headers are OK!   
    if state_data_load == True:
        keys_to_round = ["Xl12", "Xl13", "Xl23", "CuL12", "CuL13", "CuL23",] # Parameters that need around decimals
        round_digits_number = 6 # '6' digist to around
        
        for data in range(df.shape[0]):
            data_row = df.iloc[data]  # Retrieve data for each row during the iteration
            print(f"Transformer {data}:") # Number of source that is data loading

            # Depure values from excel before to export to PSCAD
            parameters_dictionary = {
            key: str(round(value, round_digits_number)) if key in keys_to_round and isinstance(value, (float, int)) # if corresponds round digits that are in 'keys_to_round'
            else str(value).strip() # Convert to string and remove spaces
            for key, value in zip(col_name_XMFR_str, data_row) # take a objetc
            }
           
            try:
                master_transformer_XMFR_3p3w2 = canvas_Transformers.create_component("master:xfmr-3p3w2", x=x_position, y=y_position, orient=theta)
                master_transformer_XMFR_3p3w2_str = str(master_transformer_XMFR_3p3w2)
                XMFR_id = int(master_transformer_XMFR_3p3w2_str.split("#")[1])  # Identify the element's by unique ID
                master_transformer_XMFR_3p3w2_XMFR_id = Transformers.component(XMFR_id) # Select the component by id
            except Exception as e:
                print(f"Error creating a Source {data}: {e}")
                print("------------The process was interrupted------------")
                state_data_load = False
                break
                
            # Adjust spacing between elements
            x_position = x_position + delta_x 
            elements_in_line += 1 
            if elements_in_line == 8:  # Ensure a maximum of 8 elements per row
                y_position = y_position + delta_y # Ensure the height elements per row
                elements_in_line = 0 # Reset elements in line
                x_position = x0_position # Reset position 'x'
            
            for key, value in parameters_dictionary.items(): # take 1st and 2nd parameter from dictionary and pass to PSCAD
                try:
                    master_transformer_XMFR_3p3w2_XMFR_id.parameters(**{key: value}) # create a dinamic code por execute in python for set paramenters in PSCAD
                    print(f"{key:<10}: {value:<20} -> ok")
                except Exception as e:
                    print(f"{key:<10}: {value:<20} -> Error: {e}")
                    state_data_load = False
                    break       
                
            if not state_data_load: # If have errors Exit
                print("Data load process interrupted. Exiting loop.")
                break  
                
            
    print("...Simulation_Finished_Successfully...")
    