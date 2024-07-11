import os
import subprocess
import pandas as pd

def run_tool(tool_path, databasefile, outputfolder):
    if not os.path.isfile(tool_path):
        print(f"{tool_path} IS NOT FOUND. Exiting...")
        return False
    try:
        default_output_folder = os.path.dirname(os.path.abspath(__file__))
        Amcache = os.path.join(default_output_folder,"AmcacheParser.exe")
        if tool_path == Amcache:
            result = subprocess.run([tool_path, '-f', databasefile, '--csv', outputfolder, '--dt', 'yyyy-MM-dd HH:mm:sszzz', '-i'], check=True, capture_output=True, text=True)
            print(result.stdout)
        else:
            result = subprocess.run([tool_path, '--csv', outputfolder, '--dt', 'yyyy-MM-dd HH:mm:sszzz'], check=True, capture_output=True, text=True)
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error running tool: {e}")
        print(f"Output: {e.output}")
        return False

def combine_and_sort_csv_files(csv_files, input_folder, output_file):
    AppCompatCacheCSVFilesFields = {
        'AppCompatCache': 'LastModifiedTimeUTC',
    }
    AmcacheCSVFilesFields = {
        "Amcache_AssociatedFileEntries":"FileKeyLastWriteTimestamp",
        "Amcache_DeviceContainers": "KeyLastWriteTimestamp",
        "Amcache_DevicePnps": "KeyLastWriteTimestamp",
        "Amcache_DriveBinaries": "KeyLastWriteTimestamp",
        "Amcache_DriverPackages": "KeyLastWriteTimestamp",
        "Amcache_ShortCuts": "KeyLastWriteTimestamp",
        "Amcache_UnassociatedFileEntries": "FileKeyLastWriteTimestamp",
        "Amcache_ProgramEntries":"KeyLastWriteTimestamp"
    }

    df_list = []

    for file in csv_files:
        file_path = os.path.join(input_folder, file)
        df = pd.read_csv(file_path)
        sf = file
        sd = file_path

        for identifier, timestamp_column in AppCompatCacheCSVFilesFields .items():
            if identifier in file:
                df.insert(0, 'Source File', sf)
                df.insert(1, 'SourceDirectory', sd)
                df.insert(2, 'Timestamp UTC+0', pd.to_datetime(df[timestamp_column], errors='coerce'))
      
                df.to_csv(file_path, index=False)
                break  

        for identifier, timestamp_column in AmcacheCSVFilesFields.items():
            if identifier in file:
                df.insert(0, 'Source File', sf)
                df.insert(1, 'SourceDirectory', sd)
                df.insert(2, 'Timestamp UTC+0', pd.to_datetime(df[timestamp_column], errors='coerce'))
      
                df.to_csv(file_path, index=False)
                break  

        df_list.append(df)

    combined_csv = pd.concat(df_list, ignore_index=True)

    if 'Timestamp UTC+0' in combined_csv.columns:
        combined_csv = combined_csv.sort_values(by='Timestamp UTC+0', ascending=False)

        combined_csv.to_csv(output_file, index=False)
        print(f"Combined and sorted CSV saved to {output_file}")

def get_default_input_directories(tool_name):
    if tool_name == 'amcache':
        default_database_file = "C:\\Windows\\appcompat\\Programs\\Amcache.hve"
        default_output_folder = os.path.dirname(os.path.abspath(__file__))
    elif tool_name == 'appcomcache':
        default_database_file = ""
        default_output_folder = os.path.dirname(os.path.abspath(__file__))
    else:
        default_database_file = ""
        default_output_folder = ""
    return default_database_file, default_output_folder

if __name__ == "__main__":
    print("Welcome!")

    while True:
        tool_choice = input("Which tool do you want to use? (appcomcache/amcache/both/exit): ").strip().lower()

        if tool_choice == 'exit':
            print("Exiting...")
            break
        elif tool_choice not in ['appcomcache', 'amcache', 'both']:
            print("Invalid choice.")
            continue

        print(f"You have chosen to use the {tool_choice} tool.")

        edit_directories = input("Do you want to edit the default directories? (yes/no): ").strip().lower()

        if edit_directories not in ['yes', 'no']:
            print("Invalid choice.")
            continue

        if edit_directories == 'yes' and tool_choice in ['appcomcache']:
            outputfolder = input("Enter the output directory for CSV files: ").strip()
        elif edit_directories == 'yes' and tool_choice in ['amcache']:
            databasefile_amcache = input("Enter the path to the Amcache database file (example: C:\\Windows\\appcompat\\Programs\\Amcache.hve): ").strip()
            outputfolder = input("Enter the output directory for CSV files: ").strip()
        elif edit_directories == 'yes' and tool_choice in ['both']:
            databasefile_amcache = input("Enter the path to the Amcache database file (example: C:\\Windows\\appcompat\\Programs\\Amcache.hve): ").strip()
            outputfolder = input("Enter the output directory for CSV files: ").strip()
        else:
            databasefile_amcache, _ = get_default_input_directories('amcache')
            _, outputfolder = get_default_input_directories('appcomcache')

        csv_files = []
        if tool_choice in ['appcomcache']:
            appcompatcache_parser_path = os.path.abspath("AppCompatCacheParser.exe")
            if run_tool(appcompatcache_parser_path, None, outputfolder):
                csv_files.extend([file for file in os.listdir(outputfolder) if file.endswith('.csv')])

        elif tool_choice in ['amcache']:
            if not os.path.isfile(databasefile_amcache):
                print(f"ERROR: The database file '{databasefile_amcache}' does not exist. Exiting...")
            else:
                amcache_parser_path = os.path.abspath("AmcacheParser.exe")
                if run_tool(amcache_parser_path, databasefile_amcache, outputfolder):
                    csv_files.extend([file for file in os.listdir(outputfolder) if file.endswith('.csv')])
        elif tool_choice in ['both']:
            if not os.path.isfile(databasefile_amcache):
                print(f"ERROR: The database file '{databasefile_amcache}' does not exist. Exiting...")
            else: 
                appcompatcache_parser_path = os.path.abspath("AppCompatCacheParser.exe")
                checkappcompatcache = run_tool(appcompatcache_parser_path, None, outputfolder)
                amcache_parser_path = os.path.abspath("AmcacheParser.exe")
                checkamcache = run_tool(amcache_parser_path, databasefile_amcache, outputfolder)
                
                if checkappcompatcache or checkamcache:
                    csv_files.extend([file for file in os.listdir(outputfolder) if file.endswith('.csv')])

        if csv_files:
            print("CSV files generated:")
            for csv_file in csv_files:
                print(f"- {csv_file}")
            combine_and_sort_csv_files(csv_files, outputfolder, os.path.join(outputfolder, tool_choice + '_combined_output.csv'))
        else:
            print("No CSV files generated.")

        print(f"CSV files have been saved to {outputfolder}")

        if os.name == 'nt':
            os.startfile(outputfolder)
        else:
            print(f"Please open the following directory to view the CSV files: {outputfolder}")