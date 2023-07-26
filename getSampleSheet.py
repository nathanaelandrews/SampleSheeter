import pandas as pd
import os
import datetime
import re
import glob
import random
import string
import shutil

print('''
 Engelab brings you:
 __                       _      
/ _\\ __ _ _ __ ___  _ __ | | ___ 
\\ \\ / _` | '_ ` _ \\| '_ \\| |/ _ \\
_\\ \\ (_| | | | | | | |_) | |  __/
\\__/\\__,_|_| |_| |_| .__/|_|\\___|
                   |_|           
 __ _               _            
/ _\\ |__   ___  ___| |_ ___ _ __ 
\\ \\| '_ \\ / _ \\/ _ \\ __/ _ \\ '__|
_\\ \\ | | |  __/  __/ ||  __/ |   
\\__/_| |_|\\___|\\___|\\__\\___|_|   

Version 0.1.0 - The perpetually penultimate version
Author: Nathanael Andrews
                                 
''')


def create_sample_sheet(folder_name):
    n = int(input("Enter the number of samples: "))

    # Ask about library type before the loop
    library_type = input("What is the library type? (RNA/ATAC/DNA)")

    for i in range(n):
        print(f"Plate {i+1}")
        df = pd.read_excel('./Barcodes/template.xlsx', header=None)

        # Set temporary column names
        df.columns = [f'Column {i+1}' for i in range(df.shape[1])]

        # Fill column 1 based on the library type
        if library_type == "RNA":
            df.iloc[:,0] = "Premade-Single Cell RNA Library - NOT 10X"
        elif library_type == "ATAC":
            df.iloc[:,0] = "Premade-ATAC-Seq Library"
        elif library_type == "DNA":
            df.iloc[:,0] = "TBD"
        else:
            print("Error: Library type must be either 'RNA', 'DNA', or 'ATAC'.")
            return
        
        # Question 2
        barcode_plate = input("Which barcode plate has been used? (pattern: [1-4][A-D])")
        if not re.match(r'[1-4][A-D]', barcode_plate):
            print("Error: Barcode Plate doesn't match the required pattern.")
            return
        df = df[df.iloc[:,12] == barcode_plate]
        
        # Question 3
        library_id = input("What is the library ID?")
        if len(library_id) != 9:
            print("WARNING, Library ID:s should have 9 characters, did you forget the suffix?")
        df.iloc[:,1] = library_id
        df.iloc[:,2] = library_id + '_' + df.iloc[:,2].astype(str)
        
        # Question 4
        insert_sizes = input("What are the insert sizes (default is 200-1000)?")
        if insert_sizes == "0" or insert_sizes == "":
            df.iloc[:,6] = "200-1000"
        else:
            df.iloc[:,6] = insert_sizes
        
        # Question 5
        total_data_amount = input("What is the total data amount? (Default 35)")
        if total_data_amount == "0" or total_data_amount == "":
            df.iloc[:,8] = '35'
        else:
            df.iloc[:,8] = total_data_amount
        
        # Question 6
        library_concentration = input("What is the library concentration (Default 20)?")
        if library_concentration == "0" or library_concentration == "":
            df.iloc[:,10] = '20'
        else:
            df.iloc[:,10] = library_concentration
        
        # Question 7
        library_volume = input("Library volume (Default 20)?")
        if library_volume == "0" or library_volume == "":
            df.iloc[:,11] = '20'
        else:
            df.iloc[:,11] = library_volume

        # Remove column 13
        df.drop(df.columns[12], axis=1, inplace=True)
        
        # Save the sample as an excel file
        today_date = datetime.date.today().strftime('%Y%m%d')
        output_folder = f"./sampleSheets/{folder_name}/"
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
        df.to_excel(output_folder + today_date + '_' + library_id + '.xlsx', index=False, header=False)

        print("Library added.\n")
    
    print("All libraries added successfully.")

def create_sample_sheet_bea(num_samples, folder_name):
    initial_df = pd.DataFrame()  # create an initial empty dataframe
    for i in range(1, num_samples + 1):
        print(f"Sample {i}")
        df = pd.read_excel('./Barcodes/template_BEA.xlsx')

        # Question 1: Which barcode plate has been used?
        while True:
            barcode_plate = input("Which barcode plate has been used? (pattern: [1-4][A-D])")
            if re.match(r'[1-4][A-D]', barcode_plate):
                break
            else:
                print("Error: invalid barcode plate. Please enter again.")
        df = df[df.iloc[:, 0] == barcode_plate]

        # Question 2: What is the library ID?
        while True:
            library_id = input("What is the library ID?")
            if len(library_id) == 9:
                break
            else:
                print("WARNING, Library ID:s should have 9 characters, did you forget the suffix?")
        df.iloc[:, 1] = library_id + '.' + df.iloc[:, 2]

        # Drop the first column
        df.drop(df.columns[0], axis=1, inplace=True)

        # Append current df to initial_df
        initial_df = initial_df.append(df, ignore_index=True)

    # Save the sample sheet
    initial_df.to_excel(f'./sampleSheets/{folder_name}/' + 'merged_BEAsample_sheet.xlsx', index=False)


def merge_files(path, output_name, merge_limit):
    file_list = sorted(glob.glob(path))
    length = len(file_list)

    for i in range(0, length, merge_limit):
        df_list = []
        for file in file_list[i:i + merge_limit]:
            df = pd.read_excel(file, header=None)
            df_list.append(df)

        merged_df = pd.concat(df_list)
        merged_df.to_excel(f'./sampleSheets/{output_name}_{i // merge_limit + 1}.xlsx', index=False, header=False)

def generate_random_string(length=5):
    letters = string.ascii_lowercase
    return ''.join(random.choice(letters) for i in range(length))

def main():
    # Create folder for today's date and a random string to avoid conflicts
    global folder_name
    folder_name = '.' + datetime.date.today().strftime("%Y%m%d") + '_' + generate_random_string()
    if not os.path.exists(f'./sampleSheets/{folder_name}/'):
        os.makedirs(f'./sampleSheets/{folder_name}/')

    # Ask where the samples are being sent
    while True:
        destination = input("Where are samples being sent? (Novogene/BEA)")
        if destination == "Novogene":
            create_sample_sheet(folder_name)

            # We always merge 
            while True:
                merge_option = 'Y'
                if merge_option in ['Y', 'N']:
                    break
                else:
                    print("Invalid option. Please enter again.")
            if merge_option == 'Y':
                merge_limit = input("How many libraries do you want to include per batch?")
                output_name = input("Enter the output file name:")
                merge_files(f'./sampleSheets/{folder_name}/*.xlsx', output_name, int(merge_limit))
                shutil.rmtree(os.path.join('./sampleSheets', folder_name))
            break

        elif destination == "BEA":
            num_samples = int(input("How many samples will be included?"))
            create_sample_sheet_bea(num_samples, folder_name)

            # BEA samples are always merged
            output_name = input("Enter the output file name:")
            merge_files(f'./sampleSheets/{folder_name}/*.xlsx', output_name, 4)
            shutil.rmtree(os.path.join('./sampleSheets', folder_name))
            break

        else:
            print("Error: invalid destination. Please enter again.")

if __name__ == "__main__":
    main()
