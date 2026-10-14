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
# MODULE 1: CONFIGURATION AND STYLING
# Role: Defines target cities, resource tiers, and visual color themes.
# ==========================================
CITIES = ["Martlock", "Caerleon", "Lymhurst", "Bridgewatch", "Thetford", "Fort Sterling"]
SHEET_NAME = "Albion Market Tracker"

MIN_TIER = 1
MAX_TIER = 6

BASE_RESOURCES = ["hide", "log", "fiber", "ore", "rock"]
REFINED_RESOURCES = ["leather", "plank", "cloth", "metal bar", "stone block"]

EXTRA_ITEMS_TO_TRACK = []

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
# Role: Handles item ID translation, API requests, multi-cell filter matching, and viability math.
# ==========================================
def generate_tracking_list():
    generated = []
    for res in BASE_RESOURCES + REFINED_RESOURCES:
        for tier in range(MIN_TIER, MAX_TIER + 1):
            generated.append(f"tier {tier} {res}")
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

def calculate_sell_viability(profit, buy_price, vol):
    if buy_price <= 0: return "Unknown"
    margin_pct = (profit / buy_price) * 100
    if vol < 100: return "Low (High Undercut Risk)"
    elif vol >= 1000 and margin_pct >= 10: return "High (Fast Sell)"
    elif vol >= 500 and margin_pct >= 5: return "Good"
    else: return "Moderate"

def extract_clean_tokens(raw_cell_list):
    tokens = []
    for cell_val in raw_cell_list:
        if not cell_val:
            continue
        for part in str(cell_val).split(','):
            cleaned = part.strip().lower()
            if cleaned:
                tokens.append(cleaned)
    return tokens

def matches_multi_filter(item_name, item_tokens, tier_tokens):
    name_low = item_name.lower().strip()
    
    if tier_tokens:
        tier_matched = False
        for t_str in tier_tokens:
            t_digits = re.sub(r'[^\d]', '', t_str)
            if t_digits and re.search(rf'\b(?:tier\s*{t_digits}|t{t_digits})\b', name_low):
                tier_matched = True
                break
        if not tier_matched:
            return False
            
    if item_tokens:
        item_matched = False
        item_without_tier = re.sub(r'\b(?:tier\s*\d+|t\d+)\b', '', name_low).strip()
        for f_item_clean in item_tokens:
            if f_item_clean in item_without_tier or f_item_clean in name_low:
                item_matched = True
                break
        if not item_matched:
            return False
            
    return True

# ==========================================
# MODULE 3: NON-DESTRUCTIVE SHEET PROCESSING
# Role: Safely backs up custom data and multi-row filter columns, calculates routes, and updates cells.
# ==========================================
def update_server_tab(client, spreadsheet, server_name, server_code, global_mapping):
    try:
        worksheet = spreadsheet.worksheet(server_name)
        existing_data = worksheet.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=server_name, rows="350", cols="32")
        existing_data = []

    if worksheet.col_count < 30:
        worksheet.resize(rows=max(worksheet.row_count, 350), cols=32)

    items_to_track = generate_tracking_list()
    custom_inputs = {} 
    
    # Multi-cell column preservation lists (Columns AA to AD, starting at Row 3)
    custom_filter_items_col = []
    custom_filter_tiers_col = []
    api_filter_items_col = []
    api_filter_tiers_col = []
    
    if len(existing_data) >= 3:
        headers1 = existing_data[0]
        headers2 = existing_data[1]
        
        city_cols = {}
        curr_city = None
        for idx, val in enumerate(headers1):
            if idx >= 25: break
            if val and val not in ["Item Name", ""] and "CUSTOM" not in val and "SAFE" not in val and "CAERLEON" not in val and "BEST" not in val and "FILTER" not in val:
                curr_city = val.strip()
                city_cols[curr_city] = {'buy': None, 'sell': None}
                
        curr_city = None
        for idx, val in enumerate(headers2):
            if idx >= 25: break
            if idx < len(headers1) and headers1[idx].strip() in city_cols:
                curr_city = headers1[idx].strip()
            if curr_city and curr_city in city_cols:
                if "Custom Buy" in val: city_cols[curr_city]['buy'] = idx
                if "Custom Sell" in val: city_cols[curr_city]['sell'] = idx

        for row in existing_data[2:]:
            if not row: break
            item_name = row[0].strip()
            if not item_name or "CUSTOM USER" in item_name or "SAFE CITIES" in item_name or "CAERLEON" in item_name or "BEST LIQUIDATION" in item_name or item_name.startswith("#ERROR"):
                break
            if item_name not in items_to_track:
                items_to_track.append(item_name)
                
            # Read filter cells for this row across Columns AA to AD (indices 26 to 29)
            custom_filter_items_col.append(row[26].strip() if len(row) > 26 else "")
            custom_filter_tiers_col.append(row[27].strip() if len(row) > 27 else "")
            api_filter_items_col.append(row[28].strip() if len(row) > 28 else "")
            api_filter_tiers_col.append(row[29].strip() if len(row) > 29 else "")
                
            for city, cols in city_cols.items():
                b_idx = cols['buy']
                s_idx = cols['sell']
                b_val_str = row[b_idx].strip() if b_idx is not None and b_idx < len(row) else ""
                s_val_str = row[s_idx].strip() if s_idx is not None and s_idx < len(row) else ""
                
                b_num = int(re.sub(r'[^\d]', '', b_val_str)) if re.sub(r'[^\d]', '', b_val_str) else 0
                s_num = int(re.sub(r'[^\d]', '', s_val_str)) if re.sub(r'[^\d]', '', s_val_str) else 0
                
                custom_inputs[(item_name, city)] = {
                    'buy_num': b_num, 'sell_num': s_num,
                    'buy_str': b_val_str, 'sell_str': s_val_str
                }
                
    target_ids = []
    for item in items_to_track:
        internal_id = parse_item_name(item, global_mapping)
        if internal_id: target_ids.append(internal_id)
            
    if not target_ids: return
    history_data = fetch_history_data(target_ids, server=server_code)
    market_data = {} 
    for h in history_data:
        uid = h.get("item_id")
        city = h.get("location")
        data = h.get("data", [])
        if data:
            latest = data[-1]
            market_data[(uid, city)] = {'price': latest.get("avg_price", 0), 'volume': latest.get("item_count", 0)}

    api_safe_routes = []
    api_caerleon_routes = []
    custom_safe_routes = []
    custom_caerleon_routes = []
    
    custom_safe_liq = []
    custom_caerleon_liq = []
    api_safe_liq = []
    api_caerleon_liq = []
    
    c_item_tokens = extract_clean_tokens(custom_filter_items_col)
    c_tier_tokens = extract_clean_tokens(custom_filter_tiers_col)
    a_item_tokens = extract_clean_tokens(api_filter_items_col)
    a_tier_tokens = extract_clean_tokens(api_filter_tiers_col)
    
    custom_filter_active = bool(c_item_tokens or c_tier_tokens)
    api_filter_active = bool(a_item_tokens or a_tier_tokens)
    safe_city_names = [c for c in CITIES if c != "Caerleon"]
    
    for item in items_to_track:
        uid = parse_item_name(item, global_mapping)
        valid_cities = [c for c in CITIES if market_data.get((uid, c), {}).get("price", 0) > 0 and market_data.get((uid, c), {}).get("volume", 0) > 0]
        safe_cities = [c for c in valid_cities if c != "Caerleon"]
        
        # 1. SAFE CITIES LIQUIDATION (Multi-Item & Multi-Tier Filter Check)
        if custom_filter_active:
            if matches_multi_filter(item, c_item_tokens, c_tier_tokens):
                for c in safe_city_names:
                    c_buy = custom_inputs.get((item, c), {}).get('buy_num', 0)
                    api_avg = market_data.get((uid, c), {}).get('price', 0)
                    if c_buy > 0:
                        custom_safe_liq.append([item, c, c_buy, api_avg])
        else:
            best_safe_custom_price = 0
            best_safe_custom_city = ""
            best_safe_custom_api_fallback = 0
            has_any_safe_custom = False
            for c in safe_city_names:
                api_avg = market_data.get((uid, c), {}).get('price', 0)
                c_buy = custom_inputs.get((item, c), {}).get('buy_num', 0)
                if c_buy > 0:
                    has_any_safe_custom = True
                    if c_buy > best_safe_custom_price:
                        best_safe_custom_price = c_buy
                        best_safe_custom_city = c
                        best_safe_custom_api_fallback = api_avg
            if has_any_safe_custom:
                custom_safe_liq.append([item, best_safe_custom_city, best_safe_custom_price, best_safe_custom_api_fallback])

        if api_filter_active:
            if matches_multi_filter(item, a_item_tokens, a_tier_tokens):
                for c in safe_cities:
                    api_avg = market_data.get((uid, c), {}).get('price', 0)
                    c_buy = custom_inputs.get((item, c), {}).get('buy_num', 0)
                    c_buy_disp = c_buy if c_buy > 0 else ""
                    if api_avg > 0:
                        api_safe_liq.append([item, c, c_buy_disp, api_avg])
        else:
            best_safe_api_price = 0
            best_safe_api_city = ""
            has_any_safe_custom = any(custom_inputs.get((item, c), {}).get('buy_num', 0) > 0 for c in safe_city_names)
            for c in safe_cities:
                api_avg = market_data.get((uid, c), {}).get('price', 0)
                if api_avg > best_safe_api_price:
                    best_safe_api_price = api_avg
                    best_safe_api_city = c
            if not has_any_safe_custom and best_safe_api_price > 0:
                api_safe_liq.append([item, best_safe_api_city, "", best_safe_api_price])
            
        # 2. CAERLEON LIQUIDATION
        c_buy_caerleon = custom_inputs.get((item, "Caerleon"), {}).get('buy_num', 0)
        api_avg_caerleon = market_data.get((uid, "Caerleon"), {}).get('price', 0)
        if c_buy_caerleon > 0:
            custom_caerleon_liq.append([item, "Caerleon", c_buy_caerleon, api_avg_caerleon])
        elif "Caerleon" in valid_cities and api_avg_caerleon > 0:
            api_caerleon_liq.append([item, "Caerleon", "", api_avg_caerleon])

        # 3. ROUTES CALCULATION
        if len(safe_cities) >= 2:
            buy_city = min(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            sell_city = max(safe_cities, key=lambda c: market_data.get((uid, c), {}).get("price", 0))
            buy_price = market_data.get((uid, buy_city), {}).get("price", 0)
            sell_price = market_data.get((uid, sell_city), {}).get("price", 0)
            profit = int((sell_price * 0.895) - buy_price)
            vol = market_data.get((uid, sell_city), {}).get("volume", 0)
            
            if profit > 0:
                viability = calculate_sell_viability(profit, buy_price, vol)
                api_safe_routes.append([item, buy_city, buy_price, sell_city, sell_price, profit, vol, viability, "API Avg"])
                
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
                viability = calculate_sell_viability(profit_to_c, safe_buy_price, c_vol)
                api_caerleon_routes.append([item, best_safe_buy, safe_buy_price, "Caerleon", c_price, profit_to_c, c_vol, viability, "API Avg"])
            elif profit_from_c > 0:
                viability = calculate_sell_viability(profit_from_c, c_price, safe_sell_vol)
                api_caerleon_routes.append([item, "Caerleon", c_price, best_safe_sell, safe_sell_price, profit_from_c, safe_sell_vol, viability, "API Avg"])

        has_custom = any(custom_inputs.get((item, c), {}).get('buy_num', 0) > 0 or custom_inputs.get((item, c), {}).get('sell_num', 0) > 0 for c in CITIES)
                
        if has_custom:
            best_custom_profit = -999999 
            best_custom_route = None
            has_true_custom_pair = False
            
            for b_city in CITIES:
                for s_city in CITIES:
                    if b_city == s_city: continue
                    
                    market_sell_order = custom_inputs.get((item, b_city), {}).get('sell_num', 0)
                    market_buy_order = custom_inputs.get((item, s_city), {}).get('buy_num', 0)
                    
                    if market_sell_order == 0 and market_buy_order == 0: continue
                    
                    avg_buy = market_data.get((uid, b_city), {}).get('price', 0)
                    avg_sell = market_data.get((uid, s_city), {}).get('price', 0)
                    s_vol = market_data.get((uid, s_city), {}).get('volume', 0)
                    
                    is_true_pair = (market_sell_order > 0 and market_buy_order > 0)
                    actual_buy = market_sell_order if market_sell_order > 0 else avg_buy
                    actual_sell = market_buy_order if market_buy_order > 0 else avg_sell
                    
                    if actual_buy == 0 or actual_sell == 0 or s_vol == 0: continue
                    profit = int((actual_sell * 0.895) - actual_buy)
                    
                    if profit < 0 and not is_true_pair: continue
                    
                    should_update = False
                    if is_true_pair and not has_true_custom_pair:
                        should_update = True 
                        has_true_custom_pair = True
                    elif is_true_pair and has_true_custom_pair and profit > best_custom_profit:
                        should_update = True 
                    elif not has_true_custom_pair and profit > best_custom_profit:
                        should_update = True 
                    
                    if should_update:
                        best_custom_profit = profit
                        notes = []
                        if market_sell_order == 0: notes.append(f"Avg Price {b_city}")
                        else: notes.append(f"Custom Sell Order {b_city}")
                        if market_buy_order == 0: notes.append(f"Avg Price {s_city}")
                        else: notes.append(f"Custom Buy Order {s_city}")
                        
                        note_str = "PRIORITY: Custom Orders" if is_true_pair else " + ".join(notes)
                        viability = calculate_sell_viability(profit, actual_buy, s_vol)
                        best_custom_route = [item, b_city, actual_buy, s_city, actual_sell, profit, s_vol, viability, note_str]
                        
            if best_custom_route:
                if best_custom_route[1] == "Caerleon" or best_custom_route[3] == "Caerleon":
                    custom_caerleon_routes.append(best_custom_route)
                else:
                    custom_safe_routes.append(best_custom_route)

    # 4. BUILD THE GRID & DATAFRAMES
    route_columns = ["Item", "Buy At", "Buy Price", "Sell At", "Sell Price", "Net Profit", "Daily Vol", "Sell Viability", "Type / Notes"]
    liq_columns = ["Item", "City to Sell", "Custom Buy Order", "API Avg Price"]
    
    caerleon_all = custom_caerleon_routes + api_caerleon_routes
    caerleon_df = pd.DataFrame(caerleon_all, columns=route_columns)
    if not caerleon_df.empty:
        caerleon_df = caerleon_df.drop_duplicates(subset=["Item"], keep="first").sort_values(by="Net Profit", ascending=False).head(5)
        
    safe_df = pd.DataFrame(api_safe_routes, columns=route_columns).sort_values(by="Net Profit", ascending=False).head(5) if len(api_safe_routes) > 0 else pd.DataFrame()
    
    if len(custom_safe_routes) > 0:
        custom_df = pd.DataFrame(custom_safe_routes, columns=route_columns).sort_values(by="Net Profit", ascending=False)
    else:
        custom_df = safe_df.copy()
        
    safe_custom_liq_df = pd.DataFrame(custom_safe_liq, columns=liq_columns).sort_values(by="Custom Buy Order", ascending=False) if len(custom_safe_liq) > 0 else pd.DataFrame()
    caerleon_custom_liq_df = pd.DataFrame(custom_caerleon_liq, columns=liq_columns).sort_values(by="Custom Buy Order", ascending=False) if len(custom_caerleon_liq) > 0 else pd.DataFrame()
    
    if len(api_safe_liq) > 0:
        safe_api_liq_df = pd.DataFrame(api_safe_liq, columns=liq_columns).sort_values(by="API Avg Price", ascending=False)
        if not api_filter_active:
            safe_api_liq_df = safe_api_liq_df.head(10)
    else:
        safe_api_liq_df = pd.DataFrame()
        
    caerleon_api_liq_df = pd.DataFrame(api_caerleon_liq, columns=liq_columns).sort_values(by="API Avg Price", ascending=False).head(10) if len(api_caerleon_liq) > 0 else pd.DataFrame()
    
    header_row_1 = ["Item Name"]
    header_row_2 = [""]
    
    for city in CITIES:
        header_row_1.extend([city, "", "", ""])
        header_row_2.extend(["Avg Price", "Volume", "Custom Buy Order", "Custom Sell Order"])
        
    # Append Filter Panel Headers in Columns AA to AD (Indices 26 to 29)
    header_row_1.extend(["", "LIQUIDATION TRACKER FILTERS", "", "", ""])
    header_row_2.extend(["", "Custom Filter Item", "Custom Filter Tier", "API Filter Item", "API Filter Tier"])
        
    final_data = [header_row_1, header_row_2]
    
    for idx, item in enumerate(items_to_track): 
        uid = parse_item_name(item, global_mapping)
        row = [item]
        for city in CITIES:
            price = market_data.get((uid, city), {}).get("price", 0)
            vol = market_data.get((uid, city), {}).get("volume", 0)
            
            c_buy_str = custom_inputs.get((item, city), {}).get('buy_str', "")
            c_sell_str = custom_inputs.get((item, city), {}).get('sell_str', "")
            
            row.extend([price, vol, c_buy_str, c_sell_str])
            
        # Preserve User Filter Inputs across every row in Columns AA to AD
        c_f_item_val = custom_filter_items_col[idx] if idx < len(custom_filter_items_col) else ""
        c_f_tier_val = custom_filter_tiers_col[idx] if idx < len(custom_filter_tiers_col) else ""
        a_f_item_val = api_filter_items_col[idx] if idx < len(api_filter_items_col) else ""
        a_f_tier_val = api_filter_tiers_col[idx] if idx < len(api_filter_tiers_col) else ""
        
        row.extend(["", c_f_item_val, c_f_tier_val, a_f_item_val, a_f_tier_val])
        final_data.append(row)
        
    def append_table(title, df):
        if df is None or df.empty: return
        final_data.append([])
        final_data.append([title])
        final_data.append(list(df.columns))
        clean_df = df.astype(object).where(pd.notnull(df), None)
        final_data.extend(clean_df.values.tolist())

    custom_liq_title = "BEST LIQUIDATION (CUSTOM): SAFE CITIES [FILTERED ALL CITIES]" if custom_filter_active else "BEST LIQUIDATION (CUSTOM): SAFE CITIES"
    api_liq_title = "BEST LIQUIDATION (API): SAFE CITIES [FILTERED ALL CITIES]" if api_filter_active else "BEST LIQUIDATION (API): SAFE CITIES (Top 10)"

    append_table("CUSTOM USER ROUTES (Hybrid)", custom_df)
    append_table("SAFE CITIES API ROUTES (Top 5)", safe_df)
    append_table("CAERLEON (RED ZONE) API ROUTES (Top 5)", caerleon_df)
    append_table(custom_liq_title, safe_custom_liq_df)
    append_table("BEST LIQUIDATION (CUSTOM): CAERLEON", caerleon_custom_liq_df)
    append_table(api_liq_title, safe_api_liq_df)
    append_table("BEST LIQUIDATION (API): CAERLEON (Top 10)", caerleon_api_liq_df)
        
    # 5. NON-DESTRUCTIVE WRITE
    worksheet.update(values=final_data, range_name='A1', value_input_option='USER_ENTERED')
    new_row_count = len(final_data)
    worksheet.batch_clear([f'A{new_row_count + 1}:AD1000'])
    
    # 6. BATCH FORMATTING
    requests_payload = [{"unmergeCells": {"range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": 31}}}]
    
    col_idx = 1
    for city in CITIES:
        requests_payload.append({
            "mergeCells": {
                "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": col_idx, "endColumnIndex": col_idx + 4},
                "mergeType": "MERGE_ALL"
            }
        })
        col_idx += 4
        
    # Merge Filter Header Banner (AA1:AD1)
    requests_payload.append({
        "mergeCells": {
            "range": {"sheetId": worksheet.id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 26, "endColumnIndex": 30},
            "mergeType": "MERGE_ALL"
        }
    })
        
    if requests_payload:
        spreadsheet.batch_update({"requests": requests_payload})
        
    end_col = col_to_letter(1 + (len(CITIES) * 4))
    formatting_rules = []
    
    formatting_rules.append(('A1:AD2', cellFormat(textFormat=textFormat(bold=True), horizontalAlignment='CENTER')))
    formatting_rules.append(('AA1:AD2', cellFormat(backgroundColor=color(0.85, 0.92, 0.98), textFormat=textFormat(bold=True), horizontalAlignment='CENTER')))
    formatting_rules.append((f'AA3:AD{len(items_to_track)+2}', cellFormat(backgroundColor=color(1.0, 0.98, 0.8), textFormat=textFormat(bold=True), horizontalAlignment='CENTER')))
    
    col_idx = 2
    for city in CITIES:
        bg_color = CITY_COLORS.get(city, color(1.0, 1.0, 1.0))
        letter1 = col_to_letter(col_idx)
        letter2 = col_to_letter(col_idx + 3)
        formatting_rules.append((f'{letter1}1:{letter2}{len(items_to_track)+2}', cellFormat(backgroundColor=bg_color)))
        
        cust1 = col_to_letter(col_idx + 2)
        cust2 = col_to_letter(col_idx + 3)
        formatting_rules.append((f'{cust1}3:{cust2}{len(items_to_track)+2}', cellFormat(backgroundColor=color(0.97, 0.97, 0.97))))
        col_idx += 4
        
    for current_row, item in enumerate(items_to_track, start=3):
        row_color = get_row_color(item)
        formatting_rules.append((f'A{current_row}:{end_col}{current_row}', cellFormat(backgroundColor=row_color)))

    if formatting_rules:
        format_cell_ranges(worksheet, formatting_rules)

# ==========================================
# MODULE 4: MAIN EXECUTION
# Role: Authenticates Google credentials and runs updates for Americas and Asia.
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
