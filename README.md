# EPA Irish Grid Coordinates to Latitude/Longitiude
This simple tool takes an XLSX file containing Water Abstraction data from the EPA and adds Latitude and Longitude columns, by converting from the traditional Northing/Easting Irish Grid coordinates.

Note that this code relies on there only being one tab in the spreadsheet. If the source has more than one tab, make a copy of it and delete the other tabs in the copy.

Using it is relatively straightforward:

* Install [Python 3](****)
* Install [uv for python](https://github.com/astral-sh/uv?tab=readme-ov-file#installation)
* Grab this code from GitHub (click the big green Code button and download the Zip)
* Open a terminal/command-prompt
* Then run:

```bash
uv run main.py name_of_input_file.xlsx name_of_output_file.xlsx
```

* The first time it runs will be slow as UV has to download all the packages. After that it should only take a few seconds

LICENSE Apache-2.0

Copyright Conor O'Neill 2025, conor@conoroneill.com
