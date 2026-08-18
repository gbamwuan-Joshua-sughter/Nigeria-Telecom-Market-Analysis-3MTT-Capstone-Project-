import pandas as pd
import requests
from pathlib import Path
from io import StringIO

# ==========================================
# NCC INDUSTRY STATISTICS DATA EXTRACTION
# ==========================================

# Official NCC Industry Statistics page
URL = "https://es.ncc.gov.ng/informes-de-datos-de-mercado/estad%C3%ADsticas-de-la-industria"

# Create folders
RAW_FOLDER = Path("raw_data")
RAW_FOLDER.mkdir(exist_ok=True)

print("=" * 60)
print("NCC TELECOM DATA EXTRACTION")
print("=" * 60)

try:

    # ------------------------------------------
    # 1. Request the NCC webpage
    # ------------------------------------------

    print("\nConnecting to NCC website...")

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/120.0 Safari/537.36"
        )
    }

    response = requests.get(
        URL,
        headers=headers,
        timeout=30
    )

    response.raise_for_status()

    print("Successfully connected to NCC.")
    print(f"HTTP Status: {response.status_code}")

    # ------------------------------------------
    # 2. Read all HTML tables
    # ------------------------------------------

    print("\nSearching for tables...")

    tables = pd.read_html(
        StringIO(response.text)
    )

    print(f"Found {len(tables)} tables.")

    # ------------------------------------------
    # 3. Save each table separately
    # ------------------------------------------

    for i, table in enumerate(tables, start=1):

        print(
            f"\nProcessing Table {i}..."
        )

        print(
            f"Rows: {table.shape[0]}"
        )

        print(
            f"Columns: {table.shape[1]}"
        )

        # Save Excel
        excel_file = (
            RAW_FOLDER /
            f"NCC_Table_{i}.xlsx"
        )

        table.to_excel(
            excel_file,
            index=False
        )

        # Save CSV
        csv_file = (
            RAW_FOLDER /
            f"NCC_Table_{i}.csv"
        )

        table.to_csv(
            csv_file,
            index=False
        )

        print(
            f"Saved: {excel_file}"
        )

        print(
            f"Saved: {csv_file}"
        )

    # ------------------------------------------
    # 4. Save all tables into one Excel file
    # ------------------------------------------

    combined_file = (
        RAW_FOLDER /
        "NCC_All_Raw_Tables.xlsx"
    )

    with pd.ExcelWriter(
        combined_file,
        engine="openpyxl"
    ) as writer:

        for i, table in enumerate(
            tables,
            start=1
        ):

            sheet_name = (
                f"Table_{i}"
            )

            table.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    print(
        "\nAll tables saved successfully."
    )

    print(
        f"Combined Excel file: "
        f"{combined_file}"
    )

    print("\nExtraction completed successfully!")

except requests.exceptions.RequestException as error:

    print(
        "\nERROR: Could not connect to NCC."
    )

    print(error)

except ValueError as error:

    print(
        "\nERROR: Could not read the NCC tables."
    )

    print(error)

except Exception as error:

    print(
        "\nUNEXPECTED ERROR:"
    )

    print(error)
