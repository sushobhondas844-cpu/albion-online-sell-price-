import os
import json
import requests
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime

# --- CONFIGURATION ---
# Just type the exact English names of the items you want to track here.
ITEMS_TO_TRACK = [
    "Journeyman's Hide",   # Tier 2 Hide
    "Adept's Hide",        # Tier 4 Hide
    "Journeyman's Fiber",  # Tier 2 Fiber
    "Pine Logs",           # Tier 4 Wood
    "Iron Ore"             # Tier 4 Ore
]

# Which cities do you care about?
CITIES = ["Martlock", "Caerleon", "Lymhurst"]

# The name of your Google Sheet
SHEET_NAME = "Albion Market Tracker"

def get_item_mapping():
    """Fetches the official Albion item dictionary to map English names to internal IDs"""
    url = "https://raw.githubusercontent.com/broderickhyman/ao-bin-dumps/master/formatted/items.json"
    response = requests.get(url)
    data = response.json()
    
    mapping = {}
    for item in data:
        name = item.get('LocalizedNames', {}).get('EN-US')
        unique_name = item.get('UniqueName')
        if name and unique_name:
            mapping[name.lower()] = unique_name
            
    return mapping

def fetch_market_data(item_ids, server="west"):
    """Fetches 24-hour historical data including average price and daily sell volume"""
    # west = Americas, east = Asia
    base_url = f"https://{server}.albion-online-data.com/api/v2/stats/history"
    id_string = ",".join(item_ids)
    location_string = ",".join(CITIES)
    
    url = f"{base_url}/{id_string}?locations={location_string}&time-scale=24"
    
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    return []

def process_data():
    print("Loading Item Dictionary...")
    name_to_id_map = get_item_mapping()
    
    # Reverse map so we can put English names in the Excel sheet
    id_to_name_map = {v: k.title() for k, v in name_to_id_map.items()}
    
    target_ids = []
    for item in ITEMS_TO_TRACK:
        item_id = name_to_id_map.get(item.lower())
        if item_id:
            target_ids.append(item_id)
        else:
            print(f"WARNING: Could not find internal ID for '{item}'. Check spelling.")

    if not target_ids:
        print("No valid items found. Exiting.")
        return

    print("Fetching Market Data from API...")
    data_west = fetch_market_data(target_ids, server="west")
    data_east = fetch_market_data(target_ids, server="east")
    
    # Combine and process data
    all_data = []
    for entry, server_name in [(data_west, "Americas"), (data_east, "Asia")]:
        for item in entry:
            city = item.get("location")
            item_id = item.get("item_id")
            english_name = id_to_name_map.get(item_id, item_id)
            
            # The API returns an array of history. We want the most recent day.
            history = item.get("data", [])
            if history:
                latest = history[-1] 
                all_data.append({
                    "Server": server_name,
                    "City": city,
                    "Item Name": english_name,
                    "Avg Price": latest.get("avg_price"),
                    "Daily Sell Volume": latest.get("item_count"),
                    "Last Updated": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
                })
                
    return pd.DataFrame(all_data)

def update_google_sheet(df):
    """Pushes the sterile, formatted pandas dataframe to Google Sheets"""
    print("Authenticating with Google Sheets...")
    # Load credentials from GitHub Secrets
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    # Open the sheet and clear it
    sheet = client.open(SHEET_NAME).sheet1
    sheet.clear()
    
    # Write the new data
    sheet.update([df.columns.values.tolist()] + df.values.tolist())
    print("Google Sheet updated successfully!")

if __name__ == "__main__":
    df = process_data()
    if df is not None and not df.empty:
        # Sort it nicely so it's clean in the spreadsheet
        df = df.sort_values(by=["Server", "Item Name", "City"])
        update_google_sheet(df)
