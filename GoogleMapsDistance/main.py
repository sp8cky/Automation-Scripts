import googlemaps, os
import pandas as pd
from dotenv import load_dotenv

# Lodad .env file with Google API key
load_dotenv()
api_key = os.getenv("GOOGLE_API_KEY")

# Initialize Google Maps client
gmaps = googlemaps.Client(key=api_key)

# Goal destination
destination = "Gleimstraße 49, 10437 Berlin"

# Path to Excel file with addresses
excel_path = ""

# Load excel file
df = pd.read_excel(excel_path)


# calculate distance between two addresses
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

def main():
    number_read_lines = 0 # Count number of rows

    # Iterate over all rows in the Excel file
    for index, row in df.iterrows():
        # Break condition: Maximum number of lines reached
        if number_read_lines >= 10:
            print("Maximum number of lines reached, end iteration.")
            break
        
        # Skip empty rows
        if row.isnull().all():
            print("Empty line found, skipping.")
            continue

        # Get start address from row
        startadress = row.to_string(index=False).strip()

        # Calculate distance between start and destination
        distance, duration = calculate_distance(startadress, destination)

        if distance:
            print(f"From '{startadress}' to '{destination}': {distance} in {duration}")
        else:
            print(f"Error in the calculation for '{startadress}': {duration}")
        
        # Increase number of read lines
        number_read_lines += 1


if __name__ == "__main__":
    main()