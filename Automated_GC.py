#########################################################################
# Automated GC Data Extractor
# 
# Inputs:
#   (1) Directory with data reports (1 day at a time)
#   (2) Directory for output csv
# 
# Output:
#   (1) Csv file with calculated VOC concentrations for all input files
# 
# Written by
#   Bode Hoover (bodehoov@iu.edu)
# 
# Last updated March 25, 2024
#########################################################################

import os
import xlrd
import pandas as pd
from tkinter import Tk, filedialog

def get_folder_path(title="Select Folder"):
    """Function to get the folder path interactively using a dialog box."""
    root = Tk()  
    root.withdraw()  
    folder_path = filedialog.askdirectory(title="Select folder with input data folders")  
    root.destroy()  
    return folder_path

def get_output_folder(title="Select Folder to Save CSV"):
    """Function to get the folder path interactively using a dialog box for saving CSV files."""
    root = Tk()  
    root.withdraw() 
    folder_path = filedialog.askdirectory(title="Select output data folder")  
    root.destroy()  
    return folder_path

def extract_data(xls_file):
    try:
        wb = xlrd.open_workbook(xls_file)
        sheet = wb.sheet_by_index(0)
        # Back
        end_row_back = None
        start_row_back = 20
        for row in range(start_row_back, sheet.nrows):
            # Find when end of Back area (row before blank cell)
            if sheet.cell_value(row, 5) == 'Sum':
                end_row_back = row - 1
                break
        # Front
        start_row_front = end_row_back + 4
        for row in range(start_row_front, sheet.nrows):
            if sheet.cell_value(row, 5) == 'Sum':
                end_row_front = row - 1
                break
        if end_row_back is None:
            raise ValueError("Back end row not found. Unable to determine where the data ends.")
        if end_row_front is None:
            raise ValueError("Front end row not found. Unable to determine where the data ends.")
        
        rts_back = sheet.col_values(1, start_rowx=start_row_back, end_rowx = end_row_back)   # Column B, specified rows
        areas_back = sheet.col_values(9, start_rowx=start_row_back, end_rowx = end_row_back)  # Column J, specified rows
        
        rts_front = sheet.col_values(1, start_rowx=start_row_front, end_rowx = end_row_front)   # Column B, specified rows
        areas_front = sheet.col_values(9, start_rowx=start_row_front, end_rowx = end_row_front)  # Column J, specified rows
        
        # Back detector
        df_back = pd.DataFrame({
            'RT': rts_back,
            'Area': areas_back
        })
        df_back['Detector'] = 'Back'
        
        # Front detector
        df_front = pd.DataFrame({
            'RT': rts_front,
            'Area': areas_front
        })
        df_front['Detector'] = 'Front'
        
        df = pd.concat([df_back, df_front], ignore_index=True)
        try:
            file_path = sheet.cell_value(3, 7) # Cell H4, report format dependent
            start_index = file_path.find('Signals')
            if start_index != -1:
                start_index += len('Signals') + 1
                end_index = start_index + len('2024-03-25 12-10-52')
                df['date_time'] = file_path[start_index:end_index]
            else:
                print("'Signals' not found in file path:", file_path)
                df['date_time'] = None
        except Exception as e:
            print(f"An error occurred while accessing cell H4: {e}")

        return df

    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def get_compounds():
    Compound = ['Ethane', 'Ethylene', 'Propane', 'Propylene', 'Acetylene', 'Trans-2-butene', '1-Butene', 'i-butane', 
            'Cyclopentane', 'Isopentane', 'n-pentane', '1-pentene', 'Trans-2-pentene', '2,2-dimethyl-butane', 
            '2,3-dimethyl-butane', '2-methylpentane', 'Isoprene', 'Hexane', 'Methyl-cyclopentane', 
            '2,4-dimethylpentane', '2-methylhexane', 'Benzene', 'Cyclohexane', '2,3-dimethyl-pentane', 
            '3-methylhexane', 'Heptane', '2,3,4-trimethylpentane', 'Toluene', '2-methylheptane', 'n-octane', 
            'Ethylbenzene', 'isooctane', 'mp-xylene', 'Styrene', 'o-xylene', 'Propyl-benzene', 'Nonane', 
            '4-Ethyltoluene', '1,3,5-TMB', '2-ethyltoluene', '1,2,4-TMB', 'Decane', '1,23,TMB', 
            '1,3-diethyl benzene', '1,4-diethyl benzene', 'undecane']
    Bode_RT = [9.552, 9.596, 9.693, 13.108, 14.95, 15.606, 20.511, 21.772, 21.895, 22.49, 23.467, 24.878, 25.784, 26.209, 
           27.511, 27.603, 28.059, 15.215, 15.737, 16.051, 16.197, 16.527, 17.78, 19.035, 20.476, 20.727, 21.169, 
           21.554, 22.78, 25.216, 25.588, 26.371, 26.611, 27.183, 27.952, 28.5535, 29.155, 29.728, 30.168, 30.738, 
           31.097, 31.844, 32.724, 32.981, 34.659, 37.87]
    Calibration_y_int = [0.464654348, -1.564774348, -0.105978261, 0.485695652, 0.000478261, -5.85432, -1.427391304, 
                     -1.303391304, -1.556898261, -0.827953043, 7.649956522, 2.269217391, -1.647565217, -3.003652174, 
                     8.501695652, -79.24636957, 0.010310872, 0.883981304, -1.031217391, 0.629946957, 2.724028696, 
                     -0.266615652, 3.297102174, 2.497007391, -3.538778261, 8.593491304, -2.240989225, -3.85139687, 
                     1.670566087, 5.617443478, 0.580217391, 1.127471304, -7.632040435, 1.274698696, 8.40189414, 
                     0.152173913, 0.892750087, 2.401063043, 0.808134348, 1.819353119, 0.150913043, 5.066173913, 
                     0.669321739, 0.678695652, 1.710173913, 0.447521739]
    Calibration_slope = [39.41618609, 57.04490739, 9.122931522, 23.9541287, 13.31210217, 52.671, 17.75965217, 13.32366261, 
                     59.27553261, 67.21774348, 9.644302957, 12.14508522, 14.23729043, 9.618815217, 39.246224348, 
                     93.97274891, 63.82877242, 154.71614087, 22.62563478, 14.87566435, 20.85892717, 32.74605217, 
                     4.433768478, 4.439050174, 38.463490435, 11.92473391, 36.503130311, 46.373871696, 7.416391304, 
                     40.509308696, 18.538845217, 9.654591304, 25.476620652, 5.106779348, 8.064055894, 12.22304348, 
                     8.525983304, 18.22815652, 45.413045217, 8.673984628, 115.732791304, 10.442743478, 11.50169217, 
                     10.104805435, 16.287812174, 10.44833043]
    Detector = ['Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 'Back', 
            'Back', 'Back', 'Back', 'Back', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 
            'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 
            'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front', 'Front']
    
    calibration_data = {
        'Compound': Compound,
        'Bode_RT': Bode_RT,
        'Calibration_y_int': Calibration_y_int,
        'Calibration_slope': Calibration_slope,
        'Detector': Detector
    }
    calibration_data = pd.DataFrame(calibration_data)
    return calibration_data

def sum_area(df, calibration_data):
    df_back = df[df['Detector'] == 'Back'].copy()  
    df_front = df[df['Detector'] == 'Front'].copy()  
    df_back['RT'] = pd.to_numeric(df_back['RT'], errors='coerce')
    df_front['RT'] = pd.to_numeric(df_front['RT'], errors='coerce')
    summed_areas = []
    compound_names = []
    date_times = []
    
    for index, row in calibration_data.iterrows():
        rt = row['Bode_RT']
        detector = row['Detector']
        low = rt * 0.98  # lower bound for RT range
        high = rt * 1.02  # upper bound for RT range
        
        if detector == "Back":
            filtered_data = df_back[(df_back['RT'] >= low) & (df_back['RT'] <= high)].copy()  # Filtered data within RT range
            if not filtered_data.empty: 
                filtered_data['Area'] = pd.to_numeric(filtered_data['Area'], errors='coerce')
                area = filtered_data['Area'].sum()  # Sum of areas within RT range
                date_time = filtered_data['date_time'].iloc[0] 
            else:
                area = None
                date_time = None
        elif detector == "Front":
            filtered_data = df_front[(df_front['RT'] >= low) & (df_front['RT'] <= high)].copy()  # Filtered data within RT range
            if not filtered_data.empty:  
                filtered_data['Area'] = pd.to_numeric(filtered_data['Area'], errors='coerce')
                area = filtered_data['Area'].sum()  # Sum of areas within RT range
                date_time = filtered_data['date_time'].iloc[0]  # Get the date_time value from the first row
            else:
                area = None
                date_time = None
        else:
            print("Invalid 'Detector' value for compound:", row['Compound'])
            area = None  
            date_time = None
        
        summed_areas.append(area)
        compound_names.append(row['Compound'])
        date_times.append(date_time)
    df_sum = pd.DataFrame({'Compound': compound_names, 'Area': summed_areas, 'date_time': date_times})

    return df_sum

def calculate_conc(df, calibration_data):
    df['Area'] = pd.to_numeric(df['Area'])
    df_pivot = pd.DataFrame()
    for compound in calibration_data['Compound']:
        if compound not in df['Compound'].unique():
            df_pivot[compound] = 0  # Add compound column with all 0 values if not present in the DataFrame
    
    for compound, area in zip(df['Compound'], df['Area']):
        if compound in calibration_data['Compound'].values:
            idx = calibration_data.index[calibration_data['Compound'] == compound][0]
            y_int = calibration_data.at[idx, 'Calibration_y_int']
            slope = calibration_data.at[idx, 'Calibration_slope']
            df.loc[df['Compound'] == compound, 'Concentration_ppb'] = (area - y_int) / slope
    
    # Combine the existing DataFrame with the one containing all compounds (including 0 values)
    df_combined = pd.concat([df, df_pivot], axis=1)
    
    df_combined.drop(columns=['Area'], inplace=True)
    
    # Shifting time backwards to account for 45 minute sampling period
    try:
        df_combined['date_time'] = pd.to_datetime(df_combined['date_time'], format='%Y-%m-%d %H-%M-%S')
        # Shift the time backwards by 20 minutes
        df_combined['date_time'] -= pd.Timedelta(minutes=20)
    except ValueError as e:
        print(f"Error converting to datetime: {e}")

    df_pivot = df_combined.pivot_table(index='date_time', columns='Compound', values='Concentration_ppb')
    df_pivot.reset_index(inplace=True)  # Reset index to make date_time a column again
    
    return df_pivot

def process_folder(folder_path):
    data_frames = []
    
    # Recursive function to traverse through all subfolders
    def traverse_folders(folder):
        nonlocal data_frames
        for root, dirs, files in os.walk(folder):
            for file in files:
                if file.endswith('.XLS'):
                    xls_file = os.path.join(root, file)
                    try:
                        df = extract_data(xls_file)
                        if df is not None:  # Check if extract_data() returns None
                            data_frames.append(df)
                    except Exception as e:
                        print(f"An error occurred while processing {xls_file}: {e}")

    traverse_folders(folder_path)

    if not data_frames:
        print("No valid data frames found.")
        return None
    else:
        print("Data frames found:", len(data_frames))
        return data_frames

def write_to_csv(df, file_index, output_folder):
    """Function to write DataFrame to CSV with filename based on date_time column."""
    try:
        # Get the first and last values of the date_time column
        start_date = df['date_time'].iloc[0].strftime('%Y-%m-%d %H-%M-%S')
        end_date = df['date_time'].iloc[-1].strftime('%Y-%m-%d %H-%M-%S')
        filename = f"{start_date} to {end_date} ({file_index}).csv" 
        filepath = os.path.join(output_folder, filename)
        df.to_csv(filepath, index=False) 
        print(f"DataFrame {file_index} saved to {filepath}")
            
    except Exception as e:
        print(f"An error occurred while writing DataFrame {file_index} to CSV: {e}")

def combine_csv_files_and_delete(folder_path):
    """Combine multiple CSV files into a single DataFrame and delete the original files."""
    combined_df = pd.DataFrame()  
    files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]  
    for file in files:
        csv_file = os.path.join(folder_path, file)
        try:
            df = pd.read_csv(csv_file)
            combined_df = pd.concat([combined_df, df], ignore_index=True, sort=False)  
            os.remove(csv_file) 
            print(f"File '{file}' processed and deleted.")
        except Exception as e:
            print(f"An error occurred while processing {file}: {e}")
    return combined_df

def reorder_columns(df):
    desired_order = ['date_time', 'Compound', 'Ethane', 'Ethylene', 'Propane', 'Propylene', 'Acetylene', 'Trans-2-butene', 
                '1-Butene', 'i-butane', 'Cyclopentane', 'Isopentane', 'n-pentane', '1-pentene', 'Trans-2-pentene', 
                '2,2-dimethyl-butane', '2,3-dimethyl-butane', '2-methylpentane', 'Isoprene', 'Hexane', 'Methyl-cyclopentane', 
                '2,4-dimethylpentane', '2-methylhexane', 'Benzene', 'Cyclohexane', '2,3-dimethyl-pentane', '3-methylhexane', 
                'Heptane', '2,3,4-trimethylpentane', 'Toluene', '2-methylheptane', 'n-octane', 'Ethylbenzene', 'isooctane', 
                'mp-xylene', 'Styrene', 'o-xylene', 'Propyl-benzene', 'Nonane', '4-Ethyltoluene', '1,3,5-TMB', '2-ethyltoluene', 
                '1,2,4-TMB', 'Decane', '1,23,TMB', '1,3-diethyl benzene', '1,4-diethyl benzene', 'undecane']
    if 'date_time' in df.columns:
        desired_order.remove('date_time')
        desired_order.insert(0, 'date_time')
    if 'Compound' in df.columns:
        desired_order.remove('Compound')
        desired_order.insert(1, 'Compound')
    
    present_columns = list(df.columns)
    columns_to_remove = [col for col in desired_order if col not in present_columns]
    for col in columns_to_remove:
        desired_order.remove(col)
    
    columns_to_add = [col for col in present_columns if col not in desired_order]
    for col in columns_to_add:
        desired_order.append(col)
    
    df_reordered = df[desired_order]
    
    return df_reordered

calibration_data = get_compounds()
directory_path = get_folder_path() 
if directory_path: 
    output_folder = get_output_folder()  
    if output_folder:  
        processed_data = process_folder(directory_path)
        if processed_data is not None:
            for i, data_frame in enumerate(processed_data):  
                df_sum = sum_area(data_frame, calibration_data)
                df_conc = calculate_conc(df_sum, calibration_data)
                write_to_csv(df_conc, i, output_folder)  
                
            combined_df = combine_csv_files_and_delete(directory_path)
            combined_df_reordered = reorder_columns(combined_df)  # Reorder the columns
            combined_filename = "combined_data.csv"
            combined_filepath = os.path.join(output_folder, combined_filename)
            combined_df_reordered.to_csv(combined_filepath, index=False)
            print(f"Combined CSV file saved to {combined_filepath}")
        else:
            print("No data files found in the selected folder.")
    else:
        print("No folder selected for saving CSV files.")
else:
    print("No folder selected.")