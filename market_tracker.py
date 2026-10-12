import os
import json
import requests
import gspread
from gspread_formatting import *
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import re

# ==========================================
# MODULE 1: CONFIGURATION & STYLING
# Defines the targeted cities and the color codes for the sheet
# ==========================================
CITIES = ["Martlock", "Caerleon", "Lymhurst", "Bridgewatch", "Thetford", "Fort Sterling"]
SHEET_NAME = "Albion Market Tracker"

CITY_COLORS = {
    "Martlock": color(0.6, 0.8, 1.0),
    "Caerleon": color(1.0, 0.4, 0.4),
    "Lymhurst": color(0.6, 1.0, 0.6),
    "Bridgewatch": color(1.0, 0.8, 0.4),
    "Thetford": color(0.8, 0.6, 1.0),
    "Fort Sterling": color(0.9, 0.9, 0.9)
}

RESOURCE_COLORS = {
    "wood": color(0.8, 0.6, 0.4), "log": color(0.8, 0.6, 0.4), "plank": color(0.8, 0.6, 0.4),      
    "hide": color(0.95, 0.8, 0.6), "leather": color(0.95, 0.8, 0.6),
    "fiber": color(0.6, 0.9, 0.6), "cloth": color(0.6, 0.9, 0.6),
    "ore": color(0.7, 0.7, 0.7), "metal": color(0.7, 0.7, 0.7),      
    "rock": color(0.9, 0.9, 0.9), "stone": color(0.9, 0.9, 0.9), "block": color(0.9, 0.9, 0.9)
}

# ==========================================
# MODULE 2: HELPER UTILITIES
# Functions to translate column numbers, map item names, and assign colors
# ==========================================
def col_to_letter(col):
    """Converts a column number to an Excel letter (1 -> A, 2 -> B)"""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result

def get_item_mapping():
    """Fetches the official Albion item dictionary to map English names to internal IDs"""
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
    """Parses user input (e.g., 'tier 4 hide') into the game's internal ID (e.g., 'T4_HIDE')"""
    name_lower = name.lower().strip()
    if name_lower in mapping:
        return mapping[name_lower]
        
    match = re.search(r't(?:ie|i)re?\s*(\d)\s*(hide|log|wood|fiber|ore|rock|stone|leather|cloth|metal\s*bar|metalbar|plank|block|stone\s*block)', name_lower)
    if match:
        tier = match.group(1)
        res = match.group(2).replace(" ", "")
        res_map = {
            'log': 'WOOD', 'wood': 'WOOD', 'hide': 'HIDE', 'fiber': 'FIBER',
            'ore': 'ORE', 'rock': 'ROCK', 'stone': 'ROCK', 'leather': 'LEATHER',
            'cloth': 'CLOTH', 'metalbar': 'METALBAR', 'plank': 'PLANKS',
            'block': 'STONEBLOCK', 'stoneblock': 'STONEBLOCK'
        }
        return f"T{tier}_{res_map.get(res, res.upper())}"
    return name.upper()

def get_row_color(item_name):
    """Determines the background color for an item row based on its resource type"""
    name_lower = item_name.lower()
    for key, c in RESOURCE_COLORS.items():
        if key in name_lower:
            return c
    return color(1.0, 1.0, 1.0)

# ==========================================
# MODULE 3: API COMMUNICATION
# Fetches the accurate 24-hour average price and daily volume from the Albion API
# ==========================================
def fetch_history_data(item_ids, server):
    url = f"https://{server}.albion-online-data.com/api/v2/stats/history/{','.join(item_ids)}?locations={','.join(CITIES)}&time-scale=24"
    res = requests.get(url)
    return res.json() if res.status_code == 200 else []

# ==========================================
# MODULE 4: GOOGLE SHEETS PROCESSING
# Handles reading user input, processing API data, and writing/formatting the sheet
# ==========================================
def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        raw_items = worksheet.col_values(1)[2:]
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="100", cols="20")
        raw_items = []
        
    items_to_track = []
    for i in raw_items:
        val = i.strip()
        if "BEST ROUTES" in val or "---" in val:
            break
        if val and val not in items_to_track:
            items_to_track.append(val)
            
    if not items_to_track:
        items_to_track = ["tier 2 hide", "tier 3 leather", "tier 4 cloth", "tier 5 metal bar"]
        
    target_ids = []
    original_names = {}
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        if internal_id:
            target_ids.append(internal_id)
            original_names[internal_id] = item 
            
    if not target_ids:
        return
        
    # Fetch exact accurate history data
    history_data = fetch_history_data(target_ids, server=server_code)
    
    market_data = {} 
    for h in history_data:
        uid = h.get("item_id")
        city = h.get("location")
        data = h.get("data", [])
        if data:
            latest = data[-1]
            market_data[(uid, city)] = {
                'price': latest.get("avg_price", 0),
                'volume': latest.get("item_count", 0)
            }
        
    best_routes = []
    for item in items_to_track:
        uid = parse_item_name(item, global_mapping)
        valid_cities = [c for c in CITIES if market_data.get((uid, c), {}).get("price", 0) > 0 and market_data.get((uid, c), {}).get("volume", 0) > 0]
        
        if len(valid_cities) >= 2:
            buy_city = min(valid_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            sell_city = max(valid_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            
            buy_price = market_data.get((uid, buy_city), {}).get("price", 0)
            sell_price = market_data.get((uid, sell_city), {}).get("price", 0)
            
            profit = int((sell_price * 0.895) - buy_price)
            vol = market_data.get((uid, sell_city), {}).get("volume", 0)
            
            if profit > 0:
                best_routes.append([item, buy_city, buy_price, sell_city, sell_price, profit, vol])
                
    best_df = pd.DataFrame(best_routes, columns=["Item", "Buy At", "Avg Buy Price", "Sell At", "Avg Sell Price", "Net Profit (After Tax)", "Daily Vol"])
    best_df = best_df.sort_values(by="Net Profit (After Tax)", ascending=False)
    
    # Structure Headers (2 Columns per City: Avg Price & Volume)
    header_row_1 = ["Item Name"]
    header_row_2 = [""]
    
    for city in CITIES:
        header_row_1.extend([city, ""])
        header_row_2.extend(["Avg Price", "Volume"])
        
    final_data = [header_row_1, header_row_2]
    
    for item in items_to_track: 
        uid = parse_item_name(item, global_mapping)
        row = [item]
        for city in CITIES:
            price = market_data.get((uid, city), {}).get("price", 0)
            vol = market_data.get((uid, city), {}).get("volume", 0)
            row.extend([price, vol])
        final_data.append(row)
        
    final_data.append([])
    final_data.append(["--- BEST ROUTES SUMMARY ---"])
    final_data.append(list(best_df.columns))
    best_df_clean = best_df.astype(object).where(pd.notnull(best_df), None)
    final_data.extend(best_df_clean.values.tolist())
        
    worksheet.clear()
    worksheet.clear_basic_filter()
    worksheet.update(values=final_data, range_name='A1')
    
    # Apply precise styling and merging logic to protect city names
    requests_payload = []
    
    # Completely unmerge all headers to prevent disappearing city names
    requests_payload.append({
        "unmergeCells": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0, "endRowIndex": 2,
                "startColumnIndex": 0, "endColumnIndex": 50
            }
        }
    })
    
    # Merge exactly 2 columns per city
    col_idx = 1
    for city in CITIES:
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
        
    end_col = col_to_letter(1 + (len(CITIES) * 2))
    format_cell_range(worksheet, f'A1:{end_col}2', cellFormat(
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    
    col_idx = 2
    for city in CITIES:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        letter1 = col_to_letter(col_idx)
        letter2 = col_to_letter(col_idx + 1)
        format_cell_range(worksheet, f'{letter1}1:{letter2}{len(items_to_track)+2}', cellFormat(backgroundColor=bg_color))
        col_idx += 2
        
    current_row = 3
    for item in items_to_track:
        row_color = get_row_color(item)
        format_cell_range(worksheet, f'A{current_row}:{end_col}{current_row}', cellFormat(backgroundColor=row_color))
        current_row += 1

# ==========================================
# MODULE 5: MAIN EXECUTION
# Initializes Google credentials and runs the update for both servers
# ==========================================
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
        
    global_mapping = get_item_mapping()
        
    print("Updating Americas Tab...")
    update_server_tab(client, spreadsheet, "Americas", "west", global_mapping)
    
    print("Updating Asia Tab...")
    update_server_tab(client, spreadsheet, "Asia", "east", global_mapping)
    
    print("Update Complete!")

if __name__ == "__main__":
    main()
