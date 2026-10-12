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
CITIES = ["Martlock", "Caerleon", "Lymhurst", "Bridgewatch", "Thetford", "Fort Sterling"]
SHEET_NAME = "Albion Market Tracker"

# Colors for specific Cities (Applied only to Header)
CITY_COLORS = {
    "Martlock": color(0.6, 0.8, 1.0),        # Light Blue
    "Caerleon": color(1.0, 0.4, 0.4),        # Light Red
    "Lymhurst": color(0.6, 1.0, 0.6),        # Light Green
    "Bridgewatch": color(1.0, 0.8, 0.4),     # Light Orange
    "Thetford": color(0.8, 0.6, 1.0),        # Light Purple
    "Fort Sterling": color(0.9, 0.9, 0.9)    # Light Gray
}

# Colors for Resources (Applied to the data rows)
RESOURCE_COLORS = {
    "wood": color(0.8, 0.6, 0.4),       # Brown
    "log": color(0.8, 0.6, 0.4),
    "plank": color(0.8, 0.6, 0.4),      
    "hide": color(0.95, 0.8, 0.6),      # Tan
    "leather": color(0.95, 0.8, 0.6),
    "fiber": color(0.6, 0.9, 0.6),      # Green
    "cloth": color(0.6, 0.9, 0.6),
    "ore": color(0.7, 0.7, 0.7),        # Dark Grey
    "metal": color(0.7, 0.7, 0.7),      
    "rock": color(0.9, 0.9, 0.9),       # Light Grey
    "stone": color(0.9, 0.9, 0.9),
    "block": color(0.9, 0.9, 0.9)
}

def get_item_mapping():
    """Fetches the official Albion item dictionary to map exact names."""
    url = "https://raw.githubusercontent.com/broderickhyman/ao-bin-dumps/master/formatted/items.json"
    mapping = {}
    try:
        response = requests.get(url)
        if response.status_code == 200:
            for item in response.json():
                name = item.get('LocalizedNames', {}).get('EN-US')
                uid = item.get('UniqueName')
                if name and uid:
                    mapping[name.lower()] = uid
    except:
        pass
    return mapping

def parse_item_name(name, mapping):
    """Converts user input like 'tire 5 metal bar' or 'adept cloth' to 'T5_METALBAR'"""
    name_lower = name.lower().strip()
    
    # 1. Try exact English match
    if name_lower in mapping:
        return mapping[name_lower]
        
    # 2. Try Regex for raw & refined resources
    match = re.search(r't(?:ie|i)re?\s*(\d)\s*(hide|log|wood|fiber|ore|rock|stone|leather|cloth|metal\s*bar|metalbar|plank|block|stone\s*block)', name_lower)
    if match:
        tier = match.group(1)
        res = match.group(2).replace(" ", "") # removes spaces from metal bar
        
        # Maps human words to exact internal suffixes
        res_map = {
            'log': 'WOOD', 'wood': 'WOOD',
            'hide': 'HIDE', 'fiber': 'FIBER',
            'ore': 'ORE', 'rock': 'ROCK', 'stone': 'ROCK',
            'leather': 'LEATHER', 'cloth': 'CLOTH',
            'metalbar': 'METALBAR', 'plank': 'PLANKS',
            'block': 'STONEBLOCK', 'stoneblock': 'STONEBLOCK'
        }
        mapped_res = res_map.get(res, res.upper())
        return f"T{tier}_{mapped_res}"
        
    # 3. Fallback: assume they typed the UniqueName (e.g., T4_HIDE)
    return name.upper()

def get_row_color(item_name):
    """Returns the color based on the resource name string."""
    name_lower = item_name.lower()
    for key, c in RESOURCE_COLORS.items():
        if key in name_lower:
            return c
    return color(1.0, 1.0, 1.0) # White fallback

def fetch_market_data(item_ids, server="west"):
    base_url = f"https://{server}.albion-online-data.com/api/v2/stats/history"
    id_string = ",".join(item_ids)
    location_string = ",".join(CITIES)
    url = f"{base_url}/{id_string}?locations={location_string}&time-scale=24"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()
    return []

def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        # Read the items typed manually into Column A (Skipping Row 1 and Row 2)
        items_to_track = worksheet.col_values(1)[2:]
        items_to_track = [i.strip() for i in items_to_track if i.strip() and not i.startswith("---")]
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="100", cols="20")
        items_to_track = []
        
    # If the sheet is completely blank, inject some starter values
    if not items_to_track:
        items_to_track = [
            "tier 2 hide", "tier 3 leather", "tier 4 cloth", 
            "tier 5 metal bar", "tier 6 plank", "tier 7 stone block"
        ]
        
    target_ids = []
    original_names = {}
    
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        if internal_id:
            target_ids.append(internal_id)
            # Save exactly what the user typed to write it back out
            original_names[internal_id] = item 
            
    if not target_ids:
        print(f"No valid items found to track for {server_name}.")
        return
        
    raw_data = fetch_market_data(target_ids, server=server_code)
    
    table_data = []
    for item in raw_data:
        city = item.get("location")
        item_id = item.get("item_id")
        
        history = item.get("data", [])
        if history:
            latest = history[-1]
            table_data.append({
                "Item": original_names.get(item_id, item_id),
                "City": city,
                "Price": latest.get("avg_price", 0),
                "Volume": latest.get("item_count", 0)
            })
            
    df = pd.DataFrame(table_data)
    if df.empty: return
    
    # Pivot tables
    pivot_price = df.pivot(index='Item', columns='City', values='Price').fillna(0)
    pivot_volume = df.pivot(index='Item', columns='City', values='Volume').fillna(0)
    
    # Analyze best routes
    best_routes = []
    for item_name in pivot_price.index:
        prices = pivot_price.loc[item_name]
        vols = pivot_volume.loc[item_name]
        valid_cities = [c for c in CITIES if c in prices.index and prices[c] > 0 and vols[c] > 0]
        
        if len(valid_cities) >= 2:
            buy_city = min(valid_cities, key=lambda c: prices[c])
            sell_city = max(valid_cities, key=lambda c: prices[c])
            profit = prices[sell_city] - prices[buy_city]
            if profit > 0:
                best_routes.append([item_name, buy_city, int(prices[buy_city]), sell_city, int(prices[sell_city]), int(profit), int(vols[sell_city])])
                
    best_df = pd.DataFrame(best_routes, columns=["Item", "Buy At", "Buy Price", "Sell At", "Sell Price", "Gross Profit", "Sell Vol"])
    best_df = best_df.sort_values(by="Gross Profit", ascending=False)
    
    # --- BUILD THE GRID ---
    cities_in_data = list(pivot_price.columns)
    
    # Row 1: Item Name, City1, [Blank], City2, [Blank]
    header_row_1 = ["Item Name"]
    # Row 2: [Blank], Price, Volume, Price, Volume
    header_row_2 = [""]
    
    for city in cities_in_data:
        header_row_1.extend([city, ""])
        header_row_2.extend(["Price", "Volume"])
        
    final_data = [header_row_1, header_row_2]
    
    # Append the item rows
    item_row_count = len(pivot_price.index)
    for item in items_to_track:  # Keep the order the user typed them in
        if item in pivot_price.index:
            row = [item]
            for city in cities_in_data:
                price = int(pivot_price.at[item, city])
                vol = int(pivot_volume.at[item, city])
                row.extend([price, vol])
            final_data.append(row)
        
    # Append Best Routes
    final_data.append([])
    final_data.append(["--- BEST ROUTES SUMMARY ---"])
    final_data.append(list(best_df.columns))
    best_df_clean = best_df.astype(object).where(pd.notnull(best_df), None)
    final_data.extend(best_df_clean.values.tolist())
        
    # Clear and write all data at once
    worksheet.clear()
    worksheet.clear_basic_filter()
    worksheet.update(values=final_data, range_name='A1')
    
    # --- APPLY STYLING ---
    requests_payload = []
    
    # 1. Merge the City Header Cells (Row 1)
    col_idx = 1 # 0-indexed in gspread API, but A is 0, B is 1
    for city in cities_in_data:
        requests_payload.append({
            "mergeCells": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0, "endRowIndex": 1,
                    "startColumnIndex": col_idx, "endColumnIndex": col_idx + 2
                },
                "mergeType": "MERGE_ALL"
            }
        })
        col_idx += 2
        
    if requests_payload:
        spreadsheet.batch_update({"requests": requests_payload})
        
    # 2. Format Header Rows (Bold, Centered)
    format_cell_range(worksheet, 'A1:Z2', cellFormat(
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    
    # 3. Apply Faction Colors to Row 1 Merged Cells
    col_idx = 2
    for city in cities_in_data:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        col_letter = chr(64 + col_idx) # 'B', 'D', 'F'
        format_cell_range(worksheet, f'{col_letter}1', cellFormat(backgroundColor=bg_color))
        col_idx += 2
        
    # 4. Color Code The Resource Rows
    current_row = 3 # Data starts on row 3
    end_col_letter = chr(64 + 1 + (len(cities_in_data) * 2))
    
    for item in items_to_track:
        if item in pivot_price.index:
            row_color = get_row_color(item)
            format_cell_range(worksheet, f'A{current_row}:{end_col_letter}{current_row}', cellFormat(backgroundColor=row_color))
            current_row += 1

def main():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    
    try:
        spreadsheet = client.open(SHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        spreadsheet = client.create(SHEET_NAME)
        
    print("Fetching global item dictionary...")
    global_mapping = get_item_mapping()
        
    print("Updating Americas Tab...")
    update_server_tab(client, spreadsheet, "Americas", "west", global_mapping)
    
    print("Updating Asia Tab...")
    update_server_tab(client, spreadsheet, "Asia", "east", global_mapping)
    
    print("Update Complete!")

if __name__ == "__main__":
    main()
