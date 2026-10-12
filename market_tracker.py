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

def col_to_letter(col):
    """Converts a column number to an Excel letter (1 -> A, 2 -> B)"""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result

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

def fetch_live_prices(item_ids, server):
    url = f"https://{server}.albion-online-data.com/api/v2/stats/prices/{','.join(item_ids)}?locations={','.join(CITIES)}"
    res = requests.get(url)
    return res.json() if res.status_code == 200 else []

def fetch_history_volume(item_ids, server):
    url = f"https://{server}.albion-online-data.com/api/v2/stats/history/{','.join(item_ids)}?locations={','.join(CITIES)}&time-scale=24"
    res = requests.get(url)
    return res.json() if res.status_code == 200 else []

def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        raw_items = worksheet.col_values(1)[2:]
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="100", cols="25")
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
        
    price_data = fetch_live_prices(target_ids, server=server_code)
    history_data = fetch_history_volume(target_ids, server=server_code)
    
    live_market = {} 
    for p in price_data:
        uid = p.get("item_id")
        city = p.get("city")
        if (uid, city) not in live_market:
            live_market[(uid, city)] = {'buy': p.get("buy_price_max", 0), 'sell': p.get("sell_price_min", 0)}
            
    volume_market = {} 
    for h in history_data:
        uid = h.get("item_id")
        city = h.get("location")
        data = h.get("data", [])
        vol = data[-1].get("item_count", 0) if data else 0
        volume_market[(uid, city)] = vol
        
    best_routes = []
    for item in items_to_track:
        uid = parse_item_name(item, global_mapping)
        valid_cities = [c for c in CITIES if live_market.get((uid, c), {}).get("sell", 0) > 0 and volume_market.get((uid, c), 0) > 0]
        
        if len(valid_cities) >= 2:
            buy_city = min(valid_cities, key=lambda c: live_market.get((uid, c), {}).get("sell", 0))
            sell_city = max(valid_cities, key=lambda c: live_market.get((uid, c), {}).get("sell", 0))
            
            buy_price = live_market.get((uid, buy_city), {}).get("sell", 0)
            sell_price = live_market.get((uid, sell_city), {}).get("sell", 0)
            
            profit = int((sell_price * 0.895) - buy_price)
            vol = volume_market.get((uid, sell_city), 0)
            
            if profit > 0:
                best_routes.append([item, buy_city, buy_price, sell_city, sell_price, profit, vol])
                
    best_df = pd.DataFrame(best_routes, columns=["Item", "Buy At", "Live Buy Cost", "Sell At", "Live Sell Value", "Net Profit (After Tax)", "Daily Vol"])
    best_df = best_df.sort_values(by="Net Profit (After Tax)", ascending=False)
    
    header_row_1 = ["Item Name"]
    header_row_2 = [""]
    
    for city in CITIES:
        header_row_1.extend([city, "", ""])
        header_row_2.extend(["Highest Buy Order", "Lowest Sell Order", "Volume"])
        
    final_data = [header_row_1, header_row_2]
    
    for item in items_to_track: 
        uid = parse_item_name(item, global_mapping)
        row = [item]
        for city in CITIES:
            buy_order = live_market.get((uid, city), {}).get("buy", 0)
            sell_order = live_market.get((uid, city), {}).get("sell", 0)
            vol = volume_market.get((uid, city), 0)
            row.extend([buy_order, sell_order, vol])
        final_data.append(row)
        
    final_data.append([])
    final_data.append(["--- BEST ROUTES SUMMARY ---"])
    final_data.append(list(best_df.columns))
    best_df_clean = best_df.astype(object).where(pd.notnull(best_df), None)
    final_data.extend(best_df_clean.values.tolist())
        
    worksheet.clear()
    worksheet.clear_basic_filter()
    worksheet.update(values=final_data, range_name='A1')
    
    # --- APPLY STYLING ---
    requests_payload = []
    
    # 0. FIX: Unmerge all existing cells in the header area first to clear the old 2-column layout
    requests_payload.append({
        "unmergeCells": {
            "range": {
                "sheetId": worksheet.id,
                "startRowIndex": 0, "endRowIndex": 2,
                "startColumnIndex": 0, "endColumnIndex": 50
            }
        }
    })
    
    # 1. Merge City Header Cells (Spanning 3 columns)
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
        
    end_col = col_to_letter(1 + (len(CITIES) * 3))
    format_cell_range(worksheet, f'A1:{end_col}2', cellFormat(
        textFormat=textFormat(bold=True),
        horizontalAlignment='CENTER'
    ))
    
    col_idx = 2
    for city in CITIES:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        letter1 = col_to_letter(col_idx)
        letter2 = col_to_letter(col_idx + 2)
        format_cell_range(worksheet, f'{letter1}1:{letter2}{len(items_to_track)+2}', cellFormat(backgroundColor=bg_color))
        col_idx += 3
        
    current_row = 3
    for item in items_to_track:
        row_color = get_row_color(item)
        format_cell_range(worksheet, f'A{current_row}:{end_col}{current_row}', cellFormat(backgroundColor=row_color))
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
        
    global_mapping = get_item_mapping()
        
    print("Updating Americas Tab...")
    update_server_tab(client, spreadsheet, "Americas", "west", global_mapping)
    
    print("Updating Asia Tab...")
    update_server_tab(client, spreadsheet, "Asia", "east", global_mapping)
    
    print("Update Complete!")

if __name__ == "__main__":
    main()
