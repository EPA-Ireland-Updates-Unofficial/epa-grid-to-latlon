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

#!/usr/bin/env python3

import sys
import pandas as pd
from pyproj import Transformer

def irish_grid_to_latlon(easting, northing, transformer):
    """
    Convert traditional Irish Grid (e.g. EPSG:29900) to WGS84 lat/lon (EPSG:4326).
    Returns (latitude, longitude).
    """
    # pyproj returns (lon, lat) by default when going to EPSG:4326
    lon, lat = transformer.transform(easting, northing)
    return lat, lon

def main():
    """
    Usage:
        python irish_grid_xlsx_converter.py input.xlsx output.xlsx
        
    Assumes the input file has columns named 'Easting' and 'Northing'.
    If 'Easting' or 'Northing' is invalid (text, missing, etc.), 
    the script will leave 'Latitude' and 'Longitude' blank ("").
    """
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <input.xlsx> <output.xlsx>")
        sys.exit(1)

    input_xlsx = sys.argv[1]
    output_xlsx = sys.argv[2]

    # 1. Create the transformer: Irish Grid -> WGS84
    #    - Change to "EPSG:29902" if needed for Ireland 1965 / Irish Grid
    transformer = Transformer.from_crs("EPSG:29900", "EPSG:4326", always_xy=True)

    # 2. Read the Excel file
    df = pd.read_excel(input_xlsx)

    # Optional: Debug prints
    print("Columns in spreadsheet:", df.columns.tolist())
    print("Data types:\n", df.dtypes)
    print(df.head())

    # 3. Ensure Easting and Northing are numeric.
    #    - invalid parsing => NaN
    if "Easting" not in df.columns or "Northing" not in df.columns:
        raise ValueError("Missing 'Easting' or 'Northing' columns in the Excel file.")

    df["Easting"] = pd.to_numeric(df["Easting"], errors="coerce")
    df["Northing"] = pd.to_numeric(df["Northing"], errors="coerce")

    # 4. Initialize new columns for Latitude and Longitude
    df["Latitude"] = None
    df["Longitude"] = None

    # 5. Row-by-row conversion
    for idx, row in df.iterrows():
        easting = row["Easting"]
        northing = row["Northing"]
        
        # If easting or northing is invalid (NaN), put blank strings
        if pd.isna(easting) or pd.isna(northing):
            df.at[idx, "Latitude"] = ""
            df.at[idx, "Longitude"] = ""
        else:
            # Convert valid numeric easting/northing
            lat, lon = irish_grid_to_latlon(easting, northing, transformer)
            df.at[idx, "Latitude"] = lat
            df.at[idx, "Longitude"] = lon

    # 6. Write to new Excel file
    df.to_excel(output_xlsx, index=False)
    print(f"Converted file saved as: {output_xlsx}")

if __name__ == "__main__":
    main()