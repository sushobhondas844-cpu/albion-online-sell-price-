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
    "wood": color(0.8, 0.6, 0.4), "log": color(0.8, 0.6, 0.4), "plank": color(0.8, 0.6, 0.4),      
    "hide": color(0.95, 0.8, 0.6), "leather": color(0.95, 0.8, 0.6),
    "fiber": color(0.6, 0.9, 0.6), "cloth": color(0.6, 0.9, 0.6),
    "ore": color(0.7, 0.7, 0.7), "metal": color(0.7, 0.7, 0.7),      
    "rock": color(0.9, 0.9, 0.9), "stone": color(0.9, 0.9, 0.9), "block": color(0.9, 0.9, 0.9)
}

def get_item_mapping():
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
    name_lower = item_name.lower()
    for key, c in RESOURCE_COLORS.items():
        if key in name_lower:
            return c
    return color(1.0, 1.0, 1.0)

def fetch_market_data(item_ids, server="west"):
    base_url = f"https://{server}.albion-online-data.com/api/v2/stats"
    id_string = ",".join(item_ids)
    location_string = ",".join(CITIES)
    
    # 1. Fetch History (For Volume)
    hist_url = f"{base_url}/history/{id_string}?locations={location_string}&time-scale=24"
    hist_resp = requests.get(hist_url)
    hist_data = hist_resp.json() if hist_resp.status_code == 200 else []
    
    # 2. Fetch Live Prices (For Buy/Sell Orders)
    # quality=1 is default for raw resources
    price_url = f"{base_url}/prices/{id_string}?locations={location_string}&qualities=1"
    price_resp = requests.get(price_url)
    price_data = price_resp.json() if price_resp.status_code == 200 else []
    
    return hist_data, price_data

def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        items_to_track = worksheet.col_values(1)[2:]
        items_to_track = [i.strip() for i in items_to_track if i.strip() and not i.startswith("---")]
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="100", cols="30")
        items_to_track = []
        
    if not items_to_track:
        items_to_track = ["tier 2 hide", "tier 3 leather", "tier 4 cloth", "tier 5 metal bar", "tier 6 plank"]
        
    target_ids = []
    original_names = {}
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        if internal_id:
            target_ids.append(internal_id)
            original_names[internal_id] = item 
            
    if not target_ids:
        return
        
    hist_data, price_data = fetch_market_data(target_ids, server=server_code)
    
    # --- PROCESS DATA ---
    # Create a dictionary to hold combined data: data_dict[item_id][city] = {buy, sell, vol}
    data_dict = {}
    
    # Map Live Prices
    for p in price_data:
        item_id = p.get("item_id")
        city = p.get("city")
        if item_id not in data_dict: data_dict[item_id] = {}
        if city not in data_dict[item_id]: data_dict[item_id][city] = {"buy": 0, "sell": 0, "vol": 0}
        
        data_dict[item_id][city]["sell"] = p.get("sell_price_min", 0)
        data_dict[item_id][city]["buy"] = p.get("buy_price_max", 0)

    # Map Daily Volume
    for h in hist_data:
        item_id = h.get("item_id")
        city = h.get("location")
        if item_id not in data_dict: data_dict[item_id] = {}
        if city not in data_dict[item_id]: data_dict[item_id][city] = {"buy": 0, "sell": 0, "vol": 0}
        
        history_array = h.get("data", [])
        if history_array:
            data_dict[item_id][city]["vol"] = history_array[-1].get("item_count", 0)

    # --- BUILD BEST ROUTES ---
    best_routes = []
    for item_id, cities_data in data_dict.items():
        item_name = original_names.get(item_id, item_id)
        
        valid_buy_cities = [c for c, d in cities_data.items() if d["sell"] > 0] # We BUY from lowest Sell Orders
        valid_sell_cities = [c for c, d in cities_data.items() if d["buy"] > 0] # We SELL to highest Buy Orders
        
        if valid_buy_cities and valid_sell_cities:
            # Find the cheapest place to buy
            buy_city = min(valid_buy_cities, key=lambda c: cities_data[c]["sell"])
            buy_price = cities_data[buy_city]["sell"]
            
            # Find the best place to sell instantly
            sell_city = max(valid_sell_cities, key=lambda c: cities_data[c]["buy"])
            sell_price = cities_data[sell_city]["buy"]
            sell_vol = cities_data[sell_city]["vol"]
            
            profit = sell_price - buy_price
            if profit > 0:
                best_routes.append([item_name, buy_city, int(buy_price), sell_city, int(sell_price), int(profit), int(sell_vol)])
                
    best_df = pd.DataFrame(best_routes, columns=["Item", "Buy At (Sell Order)", "Buy Price", "Sell At (Buy Order)", "Sell Price", "Instant Profit", "Daily Vol"])
    best_df = best_df.sort_values(by="Instant Profit", ascending=False)
    
    # --- BUILD THE GRID ---
    # Now each city takes 3 columns
    header_row_1 = ["Item Name"]
    header_row_2 = [""]
    
    for city in CITIES:
        header_row_1.extend([city, "", ""]) # Merge 3 cols
        header_row_2.extend(["Buy Order", "Sell Order", "Daily Vol"])
        
    final_data = [header_row_1, header_row_2]
    
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        row = [item]
        for city in CITIES:
            stats = data_dict.get(internal_id, {}).get(city, {"buy": 0, "sell": 0, "vol": 0})
            row.extend([int(stats["buy"]), int(stats["sell"]), int(stats["vol"])])
        final_data.append(row)
        
    # Append Best Routes
    final_data.append([])
    final_data.append(["--- BEST INSTANT FLIP ROUTES ---"])
    final_data.append(list(best_df.columns))
    best_df_clean = best_df.astype(object).where(pd.notnull(best_df), None)
    final_data.extend(best_df_clean.values.tolist())
        
    worksheet.clear()
    worksheet.clear_basic_filter()
    worksheet.update(values=final_data, range_name='A1')
    
    # --- APPLY STYLING ---
    requests_payload = []
    
    # Merge 3 columns per City in Row 1
    col_idx = 1 
    for city in CITIES:
        requests_payload.append({
            "mergeCells": {
                "range": {
                    "sheetId": worksheet.id,
                    "startRowIndex": 0, "endRowIndex": 1,
                    "startColumnIndex": col_idx, "endColumnIndex": col_idx + 3
                },
                "mergeType": "MERGE_ALL"
            }
        })
        col_idx += 3
        
    if requests_payload:
        spreadsheet.batch_update({"requests": requests_payload})
        
    format_cell_range(worksheet, 'A1:Z2', cellFormat(
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    
    # Apply Faction Colors to Row 1 Merged Cells
    col_idx = 2
    for city in CITIES:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        # Handle columns beyond Z if needed (e.g., 6 cities * 3 cols = 18 cols. We are fine with single letters up to T)
        col_letter = chr(64 + col_idx) 
        format_cell_range(worksheet, f'{col_letter}1', cellFormat(backgroundColor=bg_color))
        col_idx += 3
        
    # Color Code The Resource Rows
    current_row = 3
    end_col_letter = chr(64 + 1 + (len(CITIES) * 3))
    
    for item in items_to_track:
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
