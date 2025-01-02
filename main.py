#!/usr/bin/env uv run
# /// script
# requires-python = ">=3.1"
# dependencies = [
#     "pyproj",
#     "pandas",
#     "openpyxl",
# ]
# ///
# Then just: uv run main.py input.xlsx output.xlsx

# Or traditional installs:
# pip install pyproj
# pip install pandas
# pip install openpyxl
# python main.py input.xlsx output.xlsx

# #!/usr/bin/env python3

import sys
import pandas as pd
from pyproj import Transformer

def irish_grid_to_latlon(easting, northing, transformer):
    """
    Convert traditional Irish Grid (EPSG:29900) to WGS84 lat/lon (EPSG:4326).
    Returns (latitude, longitude).
    """
    # pyproj returns (lon, lat) when converting to EPSG:4326, so swap
    lon, lat = transformer.transform(easting, northing)
    return lat, lon

def main():
    """
    Usage:
        python irish_grid_xlsx_converter.py <input.xlsx> <output.xlsx>
    
    Requirements:
      - 'Easting' and 'Northing' columns exist in the spreadsheet.
      - Rows with non-numeric Easting/Northing will keep their original text 
        and have blank ('') Latitude/Longitude.
    """
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input.xlsx> <output.xlsx>")
        sys.exit(1)

    input_xlsx = sys.argv[1]
    output_xlsx = sys.argv[2]

    # 1. Create the transformer (Irish Grid -> WGS84)
    #    Change "EPSG:29900" to "EPSG:29902" if needed.
    transformer = Transformer.from_crs("EPSG:29900", "EPSG:4326", always_xy=True)

    # 2. Read the Excel file
    df = pd.read_excel(input_xlsx)

    # Optional: Debug prints to confirm columns and types
    print("Columns in spreadsheet:", df.columns.tolist())
    print("Data types:\n", df.dtypes)
    print(df.head())

    # 3. Check for 'Easting' and 'Northing' columns
    if "Easting" not in df.columns or "Northing" not in df.columns:
        raise ValueError("Missing 'Easting' or 'Northing' columns in the Excel file.")

    # 4. Create or initialize Latitude and Longitude columns as blank
    #    This preserves rows that don't have valid numeric data.
    df["Latitude"] = ""
    df["Longitude"] = ""

    # 5. Iterate row-by-row to convert valid numeric Easting/Northing
    for idx, row in df.iterrows():
        # Convert the *original* Easting/Northing text to numeric; NaN if invalid
        easting_num = pd.to_numeric(row["Easting"], errors="coerce")
        northing_num = pd.to_numeric(row["Northing"], errors="coerce")

        # If both easting and northing are valid numbers, compute lat/lon
        if pd.notna(easting_num) and pd.notna(northing_num):
            lat, lon = irish_grid_to_latlon(easting_num, northing_num, transformer)
            df.at[idx, "Latitude"] = lat
            df.at[idx, "Longitude"] = lon
        else:
            # Leave Latitude and Longitude blank in non-numeric cases
            df.at[idx, "Latitude"] = ""
            df.at[idx, "Longitude"] = ""

    # 6. Write out the updated DataFrame
    df.to_excel(output_xlsx, index=False)
    print(f"Converted file saved as: {output_xlsx}")

if __name__ == "__main__":
    main()