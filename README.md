```
~> python3 sheets.py --help
usage: sheets.py [-h] {generate_token,create_sheet,read_rows,append_row} ...

Create, append to, and read from Google Sheets that are associated with a
Google APIs project, without exposing access to any other Google Sheets or
Google Drive data.

positional arguments:
  {generate_token,create_sheet,read_rows,append_row}
    generate_token      Generate a token for accessing Google Drive files via
                        a Google APIs project
    create_sheet        Create a sheet
    read_rows           Read rows from a sheet
    append_row          Append a row to a sheet

optional arguments:
  -h, --help            show this help message and exit
```

For more power, consider https://github.com/xflr6/gsheets, https://github.com/nithinmurali/pygsheets, or https://pypi.org/project/gspread/
