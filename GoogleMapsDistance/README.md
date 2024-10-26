# GoogleMaps distance automation script
This script calculates the walking distance from addresses defined in an Excel spreadsheet to a destination. The distances are written to the "Distance (km)" column of the Excel table after execution.


## Table of contents
1. [Implemented Features](#Implemented-Features)
2. [Installation and usage](#installation-and-usage)
3. [Disclaimer](#disclaimer)

## Implemented features
- Distance calculation: Calculates the distance between a list of addresses from an Excel file and a specified destination using the Google Maps API.
- Data update: Writes the calculated distances back to the Excel file in the existing "Distance (km)" column.
- Error handling: Handles potential errors in the API request and informs the user of problems.
- Dynamic row processing: Allows processing of a user-defined number of rows from the Excel file, skipping empty rows.

### Project status
- In Progress

### Future implementations
- More googlemaps modes

## Installation and usage
### Requirements
- **Python**: This script requires Python 3.x.
- **API-Key**: Create a Google Cloud account and activate the Google Maps Distance Matrix API. Upload your API key in an `.env` file:
```
GOOGLE_API_KEY=your_api_key
```

#### Excel file
The Excel file requires two columns with the headings "Adress" and "Distance (km)", one entry per line.
You can simply change the column headings and customize them in the script.

| Adress  | Distance (km)  |
| ---   |  --- |
| Street 1, 12345 City  | |
| Street 2, 12345 City  |  |

Clone the repository and navigate into the directory
```
git clone https://github.com/sp8cky/Automation-Scripts/GoogleMapsDistance && cd Automation-Scripts/GoogleMapsDistance

```
Install dependencies
```
pip install requirements.txt
```

### Call 
### Parameters
- destination (string): The destination address to which the distances are to be calculated (e.g. "Street 1, 12345 City").
- excel_path (string): The full path to the Excel file containing the addresses.
- max_rows (integer): The maximum number of rows to be processed from the Excel file.
  
### Call
**Call with required parameters**:
``` 
python main.py "<destination>" "<excel_path>" <max_rows>
``` 
**Example call**:
``` 
python main.py "Industrigata 33, 4632 Kristiansand" "C:\Users\john\Desktop\adresses.xlsx" 100
```

## Disclaimer
You use it at your own risk. I take no responsibility for any damages or problems that may arise from the use of this mail script.
