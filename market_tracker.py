import os
import json
import requests
import gspread
from gspread_formatting import *
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import re

# --- CONFIGURATION ---
# Just use conversational names. The script understands typos like "tire 5 log".
ITEMS_TO_TRACK = [
    "tier 2 hide",
    "tier 2 fiber",
    "tier 2 ore",
    "tier 3 log",
    "tier 8 log",
    "tier 3 leather"
]

# Add or remove any cities here. The sheet will automatically adjust columns.
CITIES = ["Martlock", "Caerleon", "Lymhurst", "Bridgewatch", "Thetford", "Fort Sterling"]
SHEET_NAME = "Albion Market Tracker"

# City color coding (RGB values for Google Sheets)
CITY_COLORS = {
    "Martlock": color(0.6, 0.8, 1.0),        # Light Blue
    "Caerleon": color(1.0, 0.4, 0.4),        # Light Red
    "Lymhurst": color(0.6, 1.0, 0.6),        # Light Green
    "Bridgewatch": color(1.0, 0.8, 0.4),     # Light Orange
    "Thetford": color(0.8, 0.6, 1.0),        # Light Purple
    "Fort Sterling": color(0.9, 0.9, 0.9)    # Light Gray
}

def parse_item_name(name):
    """Converts 'tire 5 log' into 'T5_WOOD'"""
    match = re.search(r't(?:ie|i)re?\s*(\d)\s*(hide|log|wood|fiber|ore|rock|stone)', name.lower())
    if match:
        tier = match.group(1)
        res = match.group(2)
        if res == 'log': res = 'wood'
        if res == 'stone': res = 'rock'
        return f"T{tier}_{res.upper()}"
    return None

def fetch_market_data(item_ids, server="west"):
    base_url = f"https://{server}.albion-online-data.com/api/v2/stats/history"
    id_string = ",".join(item_ids)
    location_string = ",".join(CITIES)
    url = f"{base_url}/{id_string}?locations={location_string}&time-scale=24"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    return []

def process_data(server):
    target_ids = []
    id_to_name = {}
    
    for item in ITEMS_TO_TRACK:
        internal_id = parse_item_name(item)
        if internal_id:
            target_ids.append(internal_id)
            id_to_name[internal_id] = item.title()
            
    raw_data = fetch_market_data(target_ids, server=server)
    
    table_data = []
    for item in raw_data:
        city = item.get("location")
        item_id = item.get("item_id")
        english_name = id_to_name.get(item_id, item_id)
        
        history = item.get("data", [])
        if history:
            latest = history[-1]
            table_data.append({
                "Item": english_name,
                "City": city,
                "Price": latest.get("avg_price", 0),
                "Volume": latest.get("item_count", 0)
            })
            
    df = pd.DataFrame(table_data)
    if df.empty: return None, None
    
    # Pivot table to make rows=Items, cols=Cities
    pivot_price = df.pivot(index='Item', columns='City', values='Price').fillna(0)
    pivot_volume = df.pivot(index='Item', columns='City', values='Volume').fillna(0)
    
    # Analyze best routes (Simple analysis for a summary section)
    best_routes = []
    for item_name in pivot_price.index:
        prices = pivot_price.loc[item_name]
        vols = pivot_volume.loc[item_name]
        
        # Filter out cities with 0 price or 0 volume
        valid_cities = [c for c in CITIES if c in prices.index and prices[c] > 0 and vols[c] > 0]
        
        if len(valid_cities) >= 2:
            buy_city = min(valid_cities, key=lambda c: prices[c])
            sell_city = max(valid_cities, key=lambda c: prices[c])
            profit = prices[sell_city] - prices[buy_city]
            
            if profit > 0:
                best_routes.append([
                    item_name, buy_city, prices[buy_city], 
                    sell_city, prices[sell_city], profit, vols[sell_city]
                ])
                
    best_df = pd.DataFrame(best_routes, columns=["Item", "Buy At", "Buy Price", "Sell At", "Sell Price", "Gross Profit", "Sell Vol"])
    best_df = best_df.sort_values(by="Gross Profit", ascending=False)
    
    return pivot_price, pivot_volume, best_df

def update_server_tab(client, spreadsheet, server_name, server_code):
    try:
        worksheet = spreadsheet.worksheet(server_name)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="100", cols="20")
        
    worksheet.clear()
    worksheet.clear_basic_filter()
    
    data = process_data(server_code)
    if not data[0] is not None: return
    
    price_df, vol_df, best_df = data
    
    # --- BUILD THE GRID ---
    cities_in_data = list(price_df.columns)
    header_row = ["Item Name"]
    for city in cities_in_data:
        header_row.extend([f"{city} Price", f"{city} Vol"])
        
    final_data = [header_row]
    
    for item in price_df.index:
        row = [item]
        for city in cities_in_data:
            # Convert NumPy types to standard Python int/float
            price = int(price_df.at[item, city])
            vol = int(vol_df.at[item, city])
            row.extend([price, vol])
        final_data.append(row)
        
    # Append the Summary Section at the bottom
    final_data.append([])
    final_data.append([])
    final_data.append(["--- BEST ROUTES SUMMARY ---"])
    final_data.append(list(best_df.columns))
    
    # Convert the pandas best_df to standard Python types before appending
    # .astype(object) prevents int64 conversion issues
    best_df_clean = best_df.astype(object).where(pd.notnull(best_df), None)
    final_data.extend(best_df_clean.values.tolist())
        
    # Write to Google Sheets (Updated syntax to fix the deprecation warning)
    worksheet.update(values=final_data, range_name='A1')
    
    # --- APPLY STYLING ---
    # Bold the top header
    format_cell_range(worksheet, 'A1:Z1', cellFormat(
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    
    # Color code city columns dynamically
    col_idx = 2
    for city in cities_in_data:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        
        # Convert index to Excel letters (B, C, D...)
        col_letter1 = chr(64 + col_idx) 
        col_letter2 = chr(64 + col_idx + 1) 
        
        # Apply color to the two columns for that city (Price and Volume)
        fmt = cellFormat(backgroundColor=bg_color)
        format_cell_range(worksheet, f'{col_letter1}1:{col_letter2}{len(price_df)+1}', fmt)
        
        col_idx += 2

def main():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    spreadsheet = client.open(SHEET_NAME)
        
    print("Updating Americas Tab...")
    update_server_tab(client, spreadsheet, "Americas", "west")
    
    print("Updating Asia Tab...")
    update_server_tab(client, spreadsheet, "Asia", "east")
    
    print("Update Complete!")

if __name__ == "__main__":
    main()
