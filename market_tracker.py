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
# ==========================================
CITIES = ["Martlock", "Caerleon", "Lymhurst", "Bridgewatch", "Thetford", "Fort Sterling"]
SHEET_NAME = "Albion Market Tracker"

# --- MASS ITEM GENERATOR ---
# Automatically generate all raw and refined resources based on this tier range
MIN_TIER = 1
MAX_TIER = 6

BASE_RESOURCES = ["hide", "log", "fiber", "ore", "rock"]
REFINED_RESOURCES = ["leather", "plank", "cloth", "metal bar", "stone block"]

# --- MANUAL ADDITIONS ---
# You can add any extra, random items here (weapons, mounts, tier 8 stuff)
EXTRA_ITEMS_TO_TRACK = [
    # "tier 8 hide",
    # "adept's mercenary jacket"
]

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
# ==========================================
def generate_tracking_list():
    """Generates the massive list of items based on the MIN_TIER and MAX_TIER variables"""
    generated = []
    for tier in range(MIN_TIER, MAX_TIER + 1):
        for res in BASE_RESOURCES + REFINED_RESOURCES:
            generated.append(f"tier {tier} {res}")
            
    # Combine generated list with any manual extras, removing duplicates
    combined = generated + EXTRA_ITEMS_TO_TRACK
    return list(dict.fromkeys(combined))

def col_to_letter(col):
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

def fetch_history_data(item_ids, server):
    url = f"https://{server}.albion-online-data.com/api/v2/stats/history/{','.join(item_ids)}?locations={','.join(CITIES)}&time-scale=24"
    res = requests.get(url)
    return res.json() if res.status_code == 200 else []

# ==========================================
# MODULE 3: GOOGLE SHEETS PROCESSING
# ==========================================
def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        existing_data = worksheet.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="300", cols="30")
        existing_data = []

    # 1. PARSE EXISTING ITEMS AND CUSTOM USER INPUTS
    items_to_track = generate_tracking_list()
    custom_inputs = {} 
    
    if len(existing_data) >= 3:
        headers1 = existing_data[0]
        headers2 = existing_data[1]
        
        city_cols = {}
        curr_city = None
        for idx, val in enumerate(headers1):
            if val and val not in ["Item Name", ""] and "---" not in val:
                curr_city = val
                city_cols[curr_city] = {'buy': None, 'sell': None}
                
        curr_city = None
        for idx, val in enumerate(headers2):
            if headers1[idx]: curr_city = headers1[idx]
            if curr_city and curr_city in city_cols:
                if val == "Custom Buy": city_cols[curr_city]['buy'] = idx
                if val == "Custom Sell": city_cols[curr_city]['sell'] = idx

        for row in existing_data[2:]:
            item_name = row[0].strip()
            if not item_name or "---" in item_name: break
            
            # If the user typed an item manually that wasn't generated by the config, keep it
            if item_name not in items_to_track:
                items_to_track.append(item_name)
                
            for city, cols in city_cols.items():
                b_idx = cols['buy']
                s_idx = cols['sell']
                b_val = row[b_idx].strip() if b_idx is not None and b_idx < len(row) else ""
                s_val = row[s_idx].strip() if s_idx is not None and s_idx < len(row) else ""
                
                b_num = int(re.sub(r'[^\d]', '', b_val)) if re.sub(r'[^\d]', '', b_val) else 0
                s_num = int(re.sub(r'[^\d]', '', s_val)) if re.sub(r'[^\d]', '', s_val) else 0
                
                custom_inputs[(item_name, city)] = {'buy': b_num, 'sell': s_num}
                
    # 2. FETCH API DATA
    target_ids = []
    original_names = {}
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        if internal_id:
            target_ids.append(internal_id)
            original_names[internal_id] = item 
            
    if not target_ids:
        return
        
    history_data = fetch_history_data(target_ids, server=server_code)
    market_data = {} 
    for h in history_data:
        uid = h.get("item_id")
        city = h.get("location")
        data = h.get("data", [])
        if data:
            latest = data[-1]
            market_data[(uid, city)] = {'price': latest.get("avg_price", 0), 'volume': latest.get("item_count", 0)}

    # 3. CALCULATE ROUTES
    safe_routes = []
    caerleon_routes = []
    custom_routes = []
    
    for item in items_to_track:
        uid = parse_item_name(item, global_mapping)
        valid_cities = [c for c in CITIES if market_data.get((uid, c), {}).get("price", 0) > 0 and market_data.get((uid, c), {}).get("volume", 0) > 0]
        safe_cities = [c for c in valid_cities if c != "Caerleon"]
        
        # --- API SAFE ROUTES ---
        if len(safe_cities) >= 2:
            buy_city = min(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            sell_city = max(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            
            buy_price = market_data.get((uid, buy_city), {}).get("price", 0)
            sell_price = market_data.get((uid, sell_city), {}).get("price", 0)
            profit = int((sell_price * 0.895) - buy_price)
            vol = market_data.get((uid, sell_city), {}).get("volume", 0)
            
            if profit > 0:
                safe_routes.append([item, buy_city, buy_price, sell_city, sell_price, profit, vol])
                
        # --- API CAERLEON ROUTES ---
        if "Caerleon" in valid_cities and len(safe_cities) >= 1:
            c_price = market_data.get((uid, "Caerleon"), {}).get("price", 0)
            c_vol = market_data.get((uid, "Caerleon"), {}).get("volume", 0)
            
            best_safe_buy = min(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            safe_buy_price = market_data.get((uid, best_safe_buy), {}).get("price", 0)
            profit_to_c = int((c_price * 0.895) - safe_buy_price)
            
            best_safe_sell = max(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            safe_sell_price = market_data.get((uid, best_safe_sell), {}).get("price", 0)
            safe_sell_vol = market_data.get((uid, best_safe_sell), {}).get("volume", 0)
            profit_from_c = int((safe_sell_price * 0.895) - c_price)
            
            if profit_to_c > 0 and profit_to_c >= profit_from_c:
                caerleon_routes.append([item, best_safe_buy, safe_buy_price, "Caerleon", c_price, profit_to_c, c_vol])
            elif profit_from_c > 0:
                caerleon_routes.append([item, "Caerleon", c_price, best_safe_sell, safe_sell_price, profit_from_c, safe_sell_vol])

        # --- HYBRID CUSTOM ROUTES (With Priority Logic) ---
        best_custom_profit = 0
        best_custom_route = None
        has_true_custom_pair = False
        
        for b_city in CITIES:
            for s_city in CITIES:
                if b_city == s_city: continue
                
                c_buy = custom_inputs.get((item, b_city), {}).get('buy', 0)
                c_sell = custom_inputs.get((item, s_city), {}).get('sell', 0)
                
                avg_buy = market_data.get((uid, b_city), {}).get('price', 0)
                avg_sell = market_data.get((uid, s_city), {}).get('price', 0)
                s_vol = market_data.get((uid, s_city), {}).get('volume', 0)
                
                # Check if this route uses 100% custom user inputs
                is_true_pair = (c_buy > 0 and c_sell > 0)
                
                actual_buy = c_buy if c_buy > 0 else avg_buy
                actual_sell = c_sell if c_sell > 0 else avg_sell
                
                if actual_buy == 0 or actual_sell == 0 or s_vol == 0: continue
                
                profit = int((actual_sell * 0.895) - actual_buy)
                
                # PRIORITY LOGIC:
                # If we found a true custom pair, ONLY allow other true custom pairs to beat it.
                # If we are currently evaluating a true pair, and the previous best was a hybrid/fallback, 
                # the true pair instantly overrides it regardless of profit size.
                should_update = False
                
                if is_true_pair and not has_true_custom_pair:
                    should_update = True # Instantly prioritize true custom pair
                    has_true_custom_pair = True
                elif is_true_pair and has_true_custom_pair and profit > best_custom_profit:
                    should_update = True # Beat existing true pair
                elif not has_true_custom_pair and profit > best_custom_profit:
                    should_update = True # Normal profit check for fallbacks
                
                if should_update:
                    best_custom_profit = profit
                    notes = []
                    if c_buy == 0: notes.append(f"Avg Price used for {b_city}")
                    if c_sell == 0: notes.append(f"Avg Price used for {s_city}")
                    
                    note_str = "PRIORITY: 100% Custom Pair" if is_true_pair else ", ".join(notes)
                    best_custom_route = [item, b_city, actual_buy, s_city, actual_sell, profit, s_vol, note_str]
                    
        if best_custom_route:
            custom_routes.append(best_custom_route)

    # 4. BUILD THE GRID
    route_columns = ["Item", "Buy At", "Buy Price", "Sell At", "Sell Price", "Net Profit (After Tax)", "Daily Vol"]
    safe_df = pd.DataFrame(safe_routes, columns=route_columns).sort_values(by="Net Profit (After Tax)", ascending=False)
    caerleon_df = pd.DataFrame(caerleon_routes, columns=route_columns).sort_values(by="Net Profit (After Tax)", ascending=False)
    
    custom_cols = route_columns + ["Notes (Fallback Status)"]
    custom_df = pd.DataFrame(custom_routes, columns=custom_cols).sort_values(by="Net Profit (After Tax)", ascending=False)
    
    header_row_1 = ["Item Name"]
    header_row_2 = [""]
    
    for city in CITIES:
        header_row_1.extend([city, "", "", ""])
        header_row_2.extend(["Avg Price", "Volume", "Custom Buy", "Custom Sell"])
        
    final_data = [header_row_1, header_row_2]
    
    for item in items_to_track: 
        uid = parse_item_name(item, global_mapping)
        row = [item]
        for city in CITIES:
            price = market_data.get((uid, city), {}).get("price", 0)
            vol = market_data.get((uid, city), {}).get("volume", 0)
            c_buy = custom_inputs.get((item, city), {}).get('buy', 0)
            c_sell = custom_inputs.get((item, city), {}).get('sell', 0)
            
            c_buy_str = c_buy if c_buy > 0 else ""
            c_sell_str = c_sell if c_sell > 0 else ""
            
            row.extend([price, vol, c_buy_str, c_sell_str])
        final_data.append(row)
        
    def append_table(title, df):
        final_data.append([])
        final_data.append([title])
        final_data.append(list(df.columns))
        clean_df = df.astype(object).where(pd.notnull(df), None)
        final_data.extend(clean_df.values.tolist())

    append_table("--- CUSTOM USER ROUTES (Hybrid) ---", custom_df)
    append_table("--- SAFE CITIES API ROUTES ---", safe_df)
    append_table("--- CAERLEON (RED ZONE) API ROUTES ---", caerleon_df)
        
    worksheet.clear()
    worksheet.clear_basic_filter()
    worksheet.update(values=final_data, range_name='A1')
    
    # 5. APPLY STYLING
    requests_payload = [{"unmergeCells": {"range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 50}}}]
    
    col_idx = 1
    for city in CITIES:
        requests_payload.append({
            "mergeCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 4},
                "mergeType": "MERGE_ALL"
            }
        })
        col_idx += 4
        
    if requests_payload:
        spreadsheet.batch_update({"requests": requests_payload})
        
    end_col = col_to_letter(1 + (len(CITIES) * 4))
    format_cell_range(worksheet, f'A1:{end_col}2', cellFormat(textFormat=textFormat(bold=True), horizontalAlignment='CENTER'))
    
    col_idx = 2
    for city in CITIES:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        letter1 = col_to_letter(col_idx)
        letter2 = col_to_letter(col_idx + 3)
        format_cell_range(worksheet, f'{letter1}1:{letter2}{len(items_to_track)+2}', cellFormat(backgroundColor=bg_color))
        
        cust1 = col_to_letter(col_idx + 2)
        cust2 = col_to_letter(col_idx + 3)
        format_cell_range(worksheet, f'{cust1}3:{cust2}{len(items_to_track)+2}', cellFormat(backgroundColor=color(0.97, 0.97, 0.97)))
        
        col_idx += 4
        
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
