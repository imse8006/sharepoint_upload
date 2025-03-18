from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
import os
import datetime
import win32com.client as win32

SHAREPOINT_SITE_URL = "https://sysco.sharepoint.com/sites/PGMDatabaseSyscoandBain"
CLIENT_ID = "2e4aa039-2d8a-4974-991a-063b4aa97378"
CLIENT_SECRET = "C~68Q~f~5cvPOJXDqmXChjfF-kvBAcjRN0RLHbje"
TARGET_FOLDER = "/sites/PGMDatabaseSyscoandBain/Shared Documents/General/GB- Brakes/SPRINT tool/Chris Daniels - SOW/PGM Data"


LOCAL_FILE_PATH = "PGM data_Sprint_tool_input.xlsx"


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

def upload_to_sharepoint(site_url, client_id, client_secret, local_file_path, target_folder):
    try:
        ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))
        ctx.load(ctx.web)
        ctx.execute_query()
        print("Authentication successful")

        if not os.path.exists(local_file_path):
            raise FileNotFoundError(f"The file '{local_file_path}' does not exist.")

        file_name = os.path.basename(local_file_path)
        target_library = ctx.web.get_folder_by_server_relative_url(target_folder)

        with open(local_file_path, "rb") as file_content:
            target_library.files.add(file_name, file_content, True)
            ctx.execute_query()
            print(f"File '{file_name}' successfully uploaded to '{target_folder}'.")

    except Exception as e:
        print(f"Error during upload: {str(e)}")


if __name__ == "__main__":
    # generates the date in the YYYYMMDD format
    today = datetime.datetime.now().strftime('%Y%m%d')
    original_file_name = os.path.basename(LOCAL_FILE_PATH).replace(".xlsx", ".xlsb")
    converted_file_name = f"{today}_{original_file_name}"

    converted_file_path = os.path.join(os.path.dirname(LOCAL_FILE_PATH), converted_file_name)

    convert_xlsx_to_xlsb(LOCAL_FILE_PATH, converted_file_path)
    upload_to_sharepoint(SHAREPOINT_SITE_URL, CLIENT_ID, CLIENT_SECRET, converted_file_path, TARGET_FOLDER)
