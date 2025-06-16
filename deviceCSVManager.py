import pandas as pd

def update_device_match_status(sheet_filepath='ImportantList.csv', current_filepath='Current.csv'):

    try:
        # Read both CSV files, specifying the 'latin1' encoding
        df_sheet = pd.read_csv(sheet_filepath, encoding='latin1')
        df_current = pd.read_csv(current_filepath, encoding='latin1')
        print("✅ Files loaded successfully.")

        df_sheet['Device Name'] = df_sheet['Device Name'].str.split(' ').str[0]
        print("🧹 'Device Name' column cleaned.")

        # Ensure the target column exists, add it if it doesn't
        if 'Match w/ Current?' not in df_sheet.columns:
            df_sheet['Match w/ Current?'] = False

        # Create a set of device names from Current.csv for efficient lookup
        current_device_names = set(df_current['Device Name'])
        
        # Use the .isin() method to check for each device name's presence
        match_status = df_sheet['Device Name'].isin(current_device_names)
        
        # Update the 'Match w/ Current?' column with the results
        df_sheet['Match w/ Current?'] = match_status
        print("🔄 'Match w/ Current?' column updated.")

        # Save the modified DataFrame back to the original CSV file, overwriting it
        df_sheet.to_csv(sheet_filepath, index=False, encoding='utf-8')
        print(f"✔️ Successfully saved changes to '{sheet_filepath}'.")

    except FileNotFoundError as e:
        print(f"❌ Error: {e}. Make sure both CSV files are in the correct directory.")
    except KeyError as e:
        print(f"❌ Error: A required column {e} was not found in one of the files.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# --- Main execution ---
if __name__ == "__main__":
    update_device_match_status()