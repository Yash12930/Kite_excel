import datetime
import xlwings as xw
from kiteconnect.exceptions import KiteException
import time
import traceback

MAX_INPUT_ROWS = 200
MAX_PORTFOLIO_ROWS = 50 
MAX_HOLDINGS_ROWS = 100 
MAX_ORDERS_ROWS = 100

def dt_now_str_fn():
    return datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]

def get_default_lot_size(symbol):
    symbol_upper = str(symbol).upper()
    if "NIFTY" in symbol_upper and "BANK" not in symbol_upper and "MID" not in symbol_upper and "FIN" not in symbol_upper and "NEXT" not in symbol_upper: return 75
    elif "BANKNIFTY" in symbol_upper: return 15
    elif "FINNIFTY" in symbol_upper: return 40
    elif "MIDCPNIFTY" in symbol_upper or "NIFTY MIDCAP SELECT" in symbol_upper: return 75
    elif "SENSEX" in symbol_upper: return 10
    else: return 1

def fetch_holdings(kite):
    try:
        return kite.holdings()
    except KiteException as e:
        print(f"[{dt_now_str_fn()}] Kite API Error fetching holdings: {e}")
    except Exception as e:
        print(f"[{dt_now_str_fn()}] General Error fetching holdings: {e}")
    return []

def update_sheet_with_data(sheet, data_rows, start_cell_str, num_data_cols, max_total_rows_in_display_area, timestamp_cell_str=None):
    try:
        if data_rows and isinstance(data_rows, list) and len(data_rows) > 0:
            sheet.range(start_cell_str).value = data_rows
        if timestamp_cell_str:
            sheet.range(timestamp_cell_str).value = f"Last Updated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
    except Exception as e_update_sheet_helper:
        print(f"[{dt_now_str_fn()}] ERROR in update_sheet_with_data for {sheet.name}: {e_update_sheet_helper}")

def update_input_sheet(sheet, kite, holdings, quotes, current_positions):
    pos_qty_map = {}
    if current_positions:
        for pos in current_positions:
            exch = pos.get('exchange'); sym = pos.get('tradingsymbol'); qty = pos.get('quantity', 0)
            if exch and sym: pos_qty_map[f"{exch.upper()}:{sym.upper()}"] = qty

    try:
        symbols_data = sheet.range(f"A2:A{MAX_INPUT_ROWS + 1}").value
    except Exception as e_read_input_all:
        return
    price_updates = {}
    portfolio_qty_updates = {}
    timestamp_updates = {}
    now_str_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for i, symbol_cell in enumerate(symbols_data):
        if isinstance(symbol_cell, str) and ":" in symbol_cell:
            symbol_str = symbol_cell.strip().upper()
            row_num = i + 2
            qty = pos_qty_map.get(symbol_str, "")
            if qty == "" or qty == 0:
                holding_qty = 0
                for h in holdings:
                    exch = h.get('exchange', '').upper()
                    tsym = h.get('tradingsymbol', '').upper()
                    s = f"{exch}:{tsym}"
                    if s == symbol_str:
                        holding_qty = h.get('quantity', 0)
                        break
                qty = holding_qty
            portfolio_qty_updates[row_num] = qty
            price_row = get_price_fields_with_fallback(symbol_str, quotes, kite)
            price_updates[row_num] = price_row
            timestamp_updates[row_num] = now_str_timestamp
    try:
        for row_num, qty in portfolio_qty_updates.items():
            sheet.range(f"B{row_num}").value = qty
        
        for row_num, prices in price_updates.items():
            sheet.range(f"C{row_num}:K{row_num}").value = [prices]
            
        for row_num, timestamp in timestamp_updates.items():
            sheet.range(f"R{row_num}").value = timestamp
            
    except Exception as e_input_write_batch:
        print(f"ERROR writing updates to INPUT sheet: {e_input_write_batch}")

def update_portfolio_sheet(sheet_port, positions, quotes):
    portfolio_rows_data = []
    if positions:
        for pos in positions:
            if pos.get('quantity', 0) == 0:
                continue
            exch_p = pos.get('exchange', '').upper(); tsym_p = pos.get('tradingsymbol', '').upper()
            symbol_key_p = f"{exch_p}:{tsym_p}" if exch_p and tsym_p else None
            live_ltp_p = pos.get('last_price', '') 
            if symbol_key_p and symbol_key_p in quotes: live_ltp_p = quotes[symbol_key_p].get('last_price', live_ltp_p)
            pnl_p = pos.get('pnl', '') 
            day_mtm_pnl = pos.get('m2m', pos.get('unrealised_pnl', pos.get('pnl',''))) 
            avg_price_p = pos.get('average_price'); qty_p = pos.get('quantity')
            if live_ltp_p != "" and avg_price_p is not None and qty_p is not None:
                try: pnl_p = (float(live_ltp_p) - float(avg_price_p)) * int(qty_p)
                except (TypeError, ValueError): pass 
            portfolio_rows_data.append([
                pos.get('instrument_token', ''), tsym_p, exch_p, qty_p, 
                avg_price_p, live_ltp_p, pnl_p, 
                pos.get('realised_pnl', ''), 
                day_mtm_pnl 
            ])
    update_sheet_with_data(sheet_port, portfolio_rows_data, "A2", 9, MAX_PORTFOLIO_ROWS + 1, "K1")

def update_holdings_sheet(sheet_hold, holdings, quotes):
    holding_rows_data = []
    if holdings:
        for h_item in holdings:
            exch_h = h_item.get('exchange', '').upper(); tsym_h = h_item.get('tradingsymbol', '').upper()
            symbol_key_h = f"{exch_h}:{tsym_h}" if exch_h and tsym_h else None
            live_ltp_h = h_item.get('last_price', '')
            if symbol_key_h and symbol_key_h in quotes: live_ltp_h = quotes[symbol_key_h].get('last_price', live_ltp_h)
            pnl_h = h_item.get('pnl', '')
            avg_price_h = h_item.get('average_price'); qty_h = h_item.get('quantity')
            if live_ltp_h != "" and avg_price_h is not None and qty_h is not None and qty_h != 0:
                try: pnl_h = (float(live_ltp_h) - float(avg_price_h)) * int(qty_h)
                except (TypeError, ValueError): pass
            holding_rows_data.append([
                tsym_h, exch_h, h_item.get('isin', ''), qty_h, h_item.get('t1_quantity', ''), 
                avg_price_h, live_ltp_h, h_item.get('close_price', ''), pnl_h, 
                h_item.get('day_change', ''), h_item.get('day_change_percentage', '')
            ])
    update_sheet_with_data(sheet_hold, holding_rows_data, "A2", 11, MAX_HOLDINGS_ROWS + 1, "L1")

def update_orders_sheet(sheet_ords, kite):
    try:
        orders_api_data = kite.orders()
        orders_rows_for_sheet = []
        if orders_api_data:
            for o_item in orders_api_data:
                orders_rows_for_sheet.append([
                    o_item.get("order_id") and str(o_item.get("order_id")),  # A
                    o_item.get("variety"),                                    # B
                    o_item.get("status"),                                     # C
                    o_item.get("tradingsymbol"),                             # D
                    o_item.get("exchange"),                                   # E
                    o_item.get("order_type"),                                # F
                    o_item.get("product"),                                    # G
                    o_item.get("transaction_type"),                          # H
                    o_item.get("quantity"),                                   # I
                    o_item.get("price"),                                      # J
                    o_item.get("trigger_price"),                             # K
                    o_item.get("average_price"),                             # L
                    o_item.get("pending_quantity"),                          # M
                    o_item.get("filled_quantity"),                           # N
                    o_item.get("order_timestamp"),                           # O
                    o_item.get("parent_order_id", "")                        # P
                ])
        update_sheet_with_data(sheet_ords, orders_rows_for_sheet, "A2", 16, MAX_ORDERS_ROWS + 1, "AA1")
    except KiteException as ke_orders:
        print(f"[{dt_now_str_fn()}] Kite API Error fetching orders: {ke_orders}")
    except Exception as e_orders_update_gen:
        print(f"[{dt_now_str_fn()}] General Error updating orders sheet: {e_orders_update_gen}")

def process_order_modifications(sheet_ords, kite):
    try:
        orders_data_block = sheet_ords.range(f"A2:Z{min(MAX_ORDERS_ROWS, 200) + 1}").value
        
        for i, row_data in enumerate(orders_data_block):
            if not row_data or not row_data[0]:
                continue
                
            excel_row_num = i + 2
            order_id_raw = row_data[0]
            if isinstance(order_id_raw, float):
                order_id = str(int(order_id_raw))
            else:
                order_id = str(order_id_raw)
                
            variety = str(row_data[1] or 'regular').lower()
            status = str(row_data[2] or '').upper()
            tradingsymbol = row_data[3]
            exchange = row_data[4]
            order_type = row_data[5]
            product = row_data[6]
            transaction_type = row_data[7]
            quantity = row_data[8]
            price = row_data[9]
            trigger_price = row_data[10]
            
            modify_flag = str(row_data[16] or '').strip().upper()
            cancel_flag = str(row_data[17] or '').strip().upper()
            if cancel_flag in ["YES", "CANCEL", "C"]:
                try:
                    kite.cancel_order(variety=variety, order_id=order_id)
                    sheet_ords.range(f"R{excel_row_num}").value = "Cancelled"
                    print(f"Order {order_id} cancelled successfully")
                except Exception as e:
                    print(f"Order can't be cancelled with reason : {e}")
                    sheet_ords.range(f"R{excel_row_num}").value = f"Cancel Error: {str(e)[:20]}"
            elif modify_flag in ["YES", "REPLACE", "MODIFY"]:
                try:
                    new_price = row_data[18] if len(row_data) > 18 else None  
                    new_trigger = row_data[19] if len(row_data) > 19 else None   
                    new_qty = row_data[20] if len(row_data) > 20 else None  
                    new_order_type = row_data[21] if len(row_data) > 21 else None 
                    new_product = row_data[22] if len(row_data) > 22 else None  
                    new_validity = row_data[23] if len(row_data) > 23 else None 
                    order_params = {
                        "variety": variety,
                        "exchange": exchange,
                        "tradingsymbol": tradingsymbol,
                        "transaction_type": transaction_type,
                        "quantity": int(new_qty) if new_qty else int(quantity),
                        "order_type": str(new_order_type).upper() if new_order_type else str(order_type).upper(),
                        "product": str(new_product).upper() if new_product else str(product).upper(),
                        "validity": str(new_validity).upper() if new_validity else "DAY"
                    }
                    if order_params["order_type"] in ["LIMIT", "SL", "SL-M"]:
                        if new_price:
                            order_params["price"] = float(new_price)
                        elif price:
                            order_params["price"] = float(price)
                            
                    if order_params["order_type"] in ["SL", "SL-M"]:
                        if new_trigger:
                            order_params["trigger_price"] = float(new_trigger)
                        elif trigger_price:
                            order_params["trigger_price"] = float(trigger_price)
                    new_order_id = kite.modify_order(**order_params)
                    sheet_ords.range(f"Q{excel_row_num}").value = f"Replaced: {new_order_id}"
                    sheet_ords.range(f"S{excel_row_num}:Z{excel_row_num}").value = ""
                    print(f"Order {order_id} replaced with {new_order_id}")
                    
                except Exception as e:
                    error_msg = f"Replace Error: {str(e)[:30]}"
                    sheet_ords.range(f"Q{excel_row_num}").value = error_msg
                    print(f"Replace failed for {order_id}: {e}")
                    
    except Exception as e:
        print(f"Error in process_order_modifications: {e}")
        import traceback
        traceback.print_exc()

def update_settings_sheet(sheet_sett, kite):
    try:
        margins = kite.margins() 
        net_margin = margins.get("equity", {}).get("net", "") 
        used_margin = margins.get("equity", {}).get("utilised", {}).get("debits", "")
        available_cash = margins.get("equity", {}).get("available", {}).get("cash", "") 
        data_to_write_margins = [
            ["Available Margin", net_margin], ["Used Margin", used_margin], ["Available Cash", available_cash]
        ]
        sheet_sett.range("A4").value = data_to_write_margins 
        sheet_sett.range("C1").value = f"Last Updated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    except KiteException as ke_margins:
        print(f"[{dt_now_str_fn()}] Kite API Error fetching/writing margin info: {ke_margins}")
        if "timed out" in str(ke_margins).lower(): sheet_sett.range("B4").value = "Timeout" 
    except Exception as e_margin_gen:
        print(f"[{dt_now_str_fn()}] General Error fetching/writing margin info: {e_margin_gen}")
        sheet_sett.range("B4").value = "Error"

def autofill_input_sheet_with_portfolio_holdings(inp_sheet, positions, holdings, max_rows=200, clear_all=False):
    if clear_all:
        print("Clearing entire INPUT sheet and rebuilding from live data...")
        inp_sheet.range(f"A2:A{max_rows+1}").value = [[""]] * max_rows
        inp_sheet.range(f"B2:R{max_rows+1}").value = [[""]*17] * max_rows
    
    holdings_syms = []
    holdings_set = set()
    for h in holdings:
        exch = h.get('exchange', '').upper()
        tsym = h.get('tradingsymbol', '').upper()
        if exch and tsym:
            s = f"{exch}:{tsym}"
            if s not in holdings_set:
                holdings_syms.append(s)
                holdings_set.add(s)
    portfolio_syms = []
    portfolio_set = set()
    for pos in positions:
        exch = pos.get('exchange', '').upper()
        tsym = pos.get('tradingsymbol', '').upper()
        qty = pos.get('quantity', 0)
        if exch and tsym and qty != 0:
            s = f"{exch}:{tsym}"
            if s not in holdings_set and s not in portfolio_set:
                portfolio_syms.append(s)
                portfolio_set.add(s)
    
    manual_syms = []
    if not clear_all:
        existing = inp_sheet.range(f"A2:A{max_rows+1}").value
        for s in existing:
            if isinstance(s, str) and s.strip():
                s_up = s.strip().upper()
                if s_up not in holdings_set and s_up not in portfolio_set:
                    manual_syms.append(s_up)

    all_rows = holdings_syms + portfolio_syms + manual_syms
    all_rows = all_rows[:max_rows]
    inp_sheet.range(f"A2:A{max_rows+1}").value = [[s] for s in all_rows] + [[""]] * (max_rows - len(all_rows))

def get_price_fields_with_fallback(symbol_str, quotes, kite):
    q_data = quotes.get(symbol_str, {})
    if not q_data or not q_data.get("last_price"):
        try:
            rest_quote = kite.quote([symbol_str])
            q_data = rest_quote.get(symbol_str, q_data)
        except Exception:
            pass
    volume_val = q_data.get("volume")
    if not volume_val:
        volume_val = q_data.get("volume_traded", "")
    avg_price_val = q_data.get("average_price")
    if not avg_price_val:
        avg_price_val = q_data.get("average_trade_price", "")
    buy_price = ""
    sell_price = ""
    if "depth" in q_data:
        buy_depth = q_data["depth"].get("buy", [])
        sell_depth = q_data["depth"].get("sell", [])
        if buy_depth and isinstance(buy_depth, list):
            buy_price = buy_depth[0].get("price", "")
        if sell_depth and isinstance(sell_depth, list):
            sell_price = sell_depth[0].get("price", "")
    price_row = [
        q_data.get("ohlc", {}).get("open", ""),
        q_data.get("ohlc", {}).get("high", ""),
        q_data.get("ohlc", {}).get("low", ""),
        q_data.get("last_price", ""),
        volume_val,
        avg_price_val,
        buy_price,
        sell_price,
        q_data.get("ohlc", {}).get("close", "")
    ]
    return price_row