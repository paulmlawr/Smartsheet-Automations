import smartsheet

def main():
    # Ask user for API key and Sheet ID
    access_token = input("🔑 Enter your Smartsheet API token: ").strip()
    sheet_id = input("📄 Enter your Smartsheet Sheet ID: ").strip()

    # Validate Sheet ID
    try:
        sheet_id = int(sheet_id)
    except ValueError:
        print("❌ Sheet ID must be a number.")
        return

    # Initialize Smartsheet client
    smartsheet_client = smartsheet.Smartsheet(access_token)
    smartsheet_client.errors_as_exceptions(True)

    # Load Smartsheet sheet
    print("\n📥 Fetching sheet details...")
    sheet = smartsheet_client.Sheets.get_sheet(sheet_id)

    # Print out all columns
    print("\n📊 Smartsheet Columns Detected:")
    for col in sheet.columns:
        print(f"- {col.title:<30} (type: {col.type}, id: {col.id})")
    print("")


if __name__ == "__main__":
    main()
