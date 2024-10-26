import googlemaps
import os
import pandas as pd
from dotenv import load_dotenv
import sys

# Load .env file with Google API key
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

# Initialize Google Maps client
gmaps = googlemaps.Client(key=api_key)

# Calculate distance between two addresses
def calculate_distance(startadress, destination):
    try:
        # Use Google Maps API to calculate distance
        response = gmaps.distance_matrix(startadress, destination, mode="walking")
        
        # Get distance and duration from response
        distance = response['rows'][0]['elements'][0]['distance']['text']
        duration = response['rows'][0]['elements'][0]['duration']['text']
        
        return distance, duration
    except Exception as e:
        return None, str(e)

def main(destination, excel_path, max_rows):
    number_read_lines = 0  # Count number of rows

    # Load excel file
    df = pd.read_excel(excel_path)

    # Iterate over all rows in the Excel file
    for index, row in df.iterrows():
        # Break condition: Maximum number of lines reached
        if number_read_lines >= max_rows:
            print("Maximum number of lines reached, end iteration.")
            break
        
        # Skip empty rows
        if row.isnull().all():
            print("Empty line found, skipping.")
            continue

        # Get start address from the correct column
        startadress = row['Adress']  # Ensure this matches your column name

        # Calculate distance between start and destination
        distance, duration = calculate_distance(startadress, destination)

        if distance:
            print(f"From '{startadress}' to '{destination}': {distance} in {duration}")
            distance_value = float(distance.split()[0])  # Extract distance value
            df.at[index, 'Distance (km)'] = distance_value  # Update distance in the dataframe
        else:
            print(f"Error in the calculation for '{startadress}': {duration}")
        
        # Increase number of read lines
        number_read_lines += 1

    # Write updated dataframe back to the Excel file
    df.to_excel(excel_path, index=False)  # Overwrite the existing file

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python main.py <destination> <excel_path> <max_rows>")
        sys.exit(1)

    destination = sys.argv[1]
    excel_path = sys.argv[2]
    max_rows = int(sys.argv[3])
    
    main(destination, excel_path, max_rows)
