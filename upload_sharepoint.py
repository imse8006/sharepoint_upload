from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
import datetime
import win32com.client as win32

# SharePoint Configuration
SHAREPOINT_SITE_URL = "https://sysco.sharepoint.com/sites/PGMDatabaseSyscoandBain"
CLIENT_ID = "2e4aa039-2d8a-4974-991a-063b4aa97378"
CLIENT_SECRET = "C~68Q~f~5cvPOJXDqmXChjfF-kvBAcjRN0RLHbje"

# Files and corresponding SharePoint sites
FILES_TO_UPLOAD = [
    ("GB_BS_Traffic_lights.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/02 Traffic Lights"),
    ("GB_BS_Traffic_lights_sprint_tool.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/SPRINT tool/Chris Daniels - SOW/PGM Data"),
    ("PGM data_Sprint_tool_input.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/SPRINT tool/Chris Daniels - SOW/PGM Data"),
    ("PGM_GB_list23_swaps.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/SPRINT tool/Chris Daniels - SOW/PGM Data"),
    ("GB_Contracted Swaps.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Contracted tool/Tool input")
]

"""FILES_TO_UPLOAD = [
    ("GB_BS_Traffic_lights.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Test/Test 1"),
    ("GB_BS_Traffic_lights_sprint_tool.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Test/Test 2"),
    ("PGM data_Sprint_tool_input.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Test/Test 3"),
    ("PGM_GB_list23_swaps.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Test/Test 4"),
    ("GB_Contracted Swaps.xlsx", "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/Test/Test 5")
]"""


def convert_xlsx_to_xlsb(source_path, target_path):
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(source_path))
        wb.SaveAs(os.path.abspath(target_path), FileFormat=50)  # 50 corresponds to the XLSB format
        wb.Close()
        excel.Quit()
        print(f"Conversion successful: '{source_path}' ➜ '{target_path}'")
    except Exception as e:
        print(f"Error during conversion: {str(e)}")

# Check if file exists on SharePoint
def exists_on_sharepoint(ctx, target_folder, file_name):
    try:
        folder = ctx.web.get_folder_by_server_relative_url(target_folder)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()

        for file in files:
            if file.properties["Name"] == file_name:
                print(f"File '{file_name}' already exists in '{target_folder}', skipping upload.")
                return True
        return False
    except Exception as e:
        print(f"Error checking file existence on SharePoint: {str(e)}")
        return False

# Upload to SharePoint
def upload_to_sharepoint(site_url, client_id, client_secret, local_file_path, target_folder):
    try:
        ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
        ctx.load(ctx.web)
        ctx.execute_query()
        print("Authentication successful")

        if not os.path.exists(local_file_path):
            raise FileNotFoundError(f"The file '{local_file_path}' does not exist.")

        file_name = os.path.basename(local_file_path)
        if exists_on_sharepoint(ctx, target_folder, file_name):
            print(f"⚠File '{file_name}' already exists, skipping upload.")
            return

        target_library = ctx.web.get_folder_by_server_relative_url(target_folder)
        with open(local_file_path, "rb") as file_content:
            target_library.files.add(file_name, file_content, True)
            ctx.execute_query()
            print(f"File '{file_name}' successfully uploaded to '{target_folder}'.")

    except Exception as e:
        print(f"Error during upload: {str(e)}")


today = datetime.datetime.now().strftime('%Y%m%d')

for file_name, target_folder in FILES_TO_UPLOAD:
    if file_name == "PGM data_Sprint_tool_input.xlsx":
        converted_file_name = f"{today}_{file_name.replace('.xlsx', '.xlsb')}"
        converted_file_path = os.path.join(os.path.dirname(file_name), converted_file_name)

        convert_xlsx_to_xlsb(file_name, converted_file_path)

    else:
        converted_file_name = f"{today}_{file_name}"
        converted_file_path = os.path.join(os.path.dirname(file_name), converted_file_name)

        os.rename(file_name, converted_file_path)

    upload_to_sharepoint(SHAREPOINT_SITE_URL, CLIENT_ID, CLIENT_SECRET, converted_file_path, target_folder)
