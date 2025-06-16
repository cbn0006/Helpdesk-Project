import os
import io
import json
import requests
import pandas as pd
from dotenv import load_dotenv
from msal import PublicClientApplication

# Load environment variables from .env file
load_dotenv()

def get_msal_token(client_id, tenant_id):
    """Uses MSAL to get an access token via the device flow."""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = PublicClientApplication(client_id, authority=authority)
    
    # The scopes our script needs to function
    scopes = ["Files.ReadWrite", "User.Read"]

    # First, try to get a token from the cache
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result:
            print("✅ Successfully acquired token from cache.")
            return result['access_token']

    # If no cached token, start the device flow for the user to log in
    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        raise ValueError("Failed to create device flow. Check Azure Portal configuration.")

    print("--- User Login Required ---")
    print(flow["message"]) # Prints the user_code and verification_uri
    
    result = app.acquire_token_by_device_flow(flow) # This line will wait until you log in

    if "access_token" in result:
        print("✅ Successfully acquired new token.")
        return result["access_token"]
    else:
        print(f"❌ Error acquiring token: {result.get('error_description')}")
        return None

def main():
    """Main function to run the file update process."""
    try:
        # --- Configuration ---
        CLIENT_ID = os.getenv('CLIENT_ID')
        TENANT_ID = os.getenv('TENANT_ID')
        FILE_PATH = 'Testing.xlsx' # The path to the file in your OneDrive
        WORKSHEET_NAME = 'devices'
        CURRENT_CSV_PATH = 'Current.csv'

        if not all([CLIENT_ID, TENANT_ID]):
            print("❌ Error: CLIENT_ID or TENANT_ID missing from .env file.")
            return

        # --- 1. Authenticate and get token ---
        access_token = get_msal_token(CLIENT_ID, TENANT_ID)
        if not access_token:
            return # Stop if authentication failed

        headers = {'Authorization': f'Bearer {access_token}'}

        # --- 2. Download the Excel file from OneDrive ---
        # Note: The file path needs to have a colon at the end for the Graph API URL
        graph_url_download = f"https://graph.microsoft.com/v1.0/me/drive/root:/{FILE_PATH}:/content"
        
        print(f"Downloading '{FILE_PATH}' from OneDrive...")
        response = requests.get(graph_url_download, headers=headers)
        response.raise_for_status() # Raises an error if the download fails
        
        # Read the Excel file content from memory
        df_sheet = pd.read_excel(io.BytesIO(response.content), sheet_name=WORKSHEET_NAME)
        df_current = pd.read_csv(CURRENT_CSV_PATH, encoding='latin1')
        print("✅ Data loaded from online Excel and local file.")

        # --- 3. Process the data (your original logic) ---
        current_device_names = set(df_current['Device Name'])
        df_sheet['Match w/ Current?'] = df_sheet['Device Name'].isin(current_device_names)
        print("🔄 'Match w/ Current?' column updated.")

        # --- 4. Upload the modified file back to OneDrive ---
        # Save the updated dataframe to an in-memory Excel file
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_sheet.to_excel(writer, index=False, sheet_name=WORKSHEET_NAME)
        excel_data = output.getvalue()

        graph_url_upload = graph_url_download # The URL is the same for upload
        
        print(f"Uploading changes to '{FILE_PATH}'...")
        upload_response = requests.put(graph_url_upload, headers=headers, data=excel_data)
        upload_response.raise_for_status() # Raises an error if the upload fails

        print("✔️ Successfully uploaded changes to OneDrive!")

    except requests.exceptions.HTTPError as e:
        # Try to parse the error message from the response
        error_json = e.response.json()
        error_message = error_json.get('error', {}).get('message', e.response.text)
        print(f"❌ A Microsoft Graph API error occurred: {error_message}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()