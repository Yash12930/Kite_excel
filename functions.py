import datetime
import os
from kiteconnect.exceptions import KiteException

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

def clear_row_except_column_a(sheet, row_num, max_col=25):
    sheet.range((row_num, 2), (row_num, max_col)).value = [[""] * (max_col - 1)]

def update_input_sheet(sheet, kite, holdings, quotes, current_positions):
    pos_qty_map, pos_pnl_map = {}, {}
    if current_positions:
        for pos in current_positions:
            key = f"{pos.get('exchange','').upper()}:{pos.get('tradingsymbol','').upper()}"
            pos_qty_map[key] = pos.get('quantity', 0)
            pos_pnl_map[key] = pos.get('pnl', None)

    hold_qty_map, hold_pnl_map = {}, {}
    if holdings:
        for h in holdings:
            key = f"{h.get('exchange','').upper()}:{h.get('tradingsymbol','').upper()}"
            hold_qty_map[key] = h.get('quantity', 0)
            hold_pnl_map[key] = h.get('pnl', None)

    symbols_data = sheet.range(f"A2:A{MAX_INPUT_ROWS + 1}").value
    now_str_timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    qty_list = []
    price_rows = []
    timestamp_list = []
    pnl_list = []

    for i, symbol_cell in enumerate(symbols_data):
        row_num = i + 2
        if not isinstance(symbol_cell, str) or not symbol_cell.strip():
            qty_list.append("")
            price_rows.append([""]*9)
            timestamp_list.append("")
            pnl_list.append("")
            continue
        symbol_str = symbol_cell.strip().upper()
        if symbol_str in ("HOLDINGS", "PORTFOLIO", "MANUAL") or ":" not in symbol_str:
            clear_row_except_column_a(sheet, row_num, max_col=25)
            qty_list.append("")
            price_rows.append([""]*9)
            timestamp_list.append("")
            pnl_list.append("")
            continue
        qty = pos_qty_map.get(symbol_str, "")
        if qty == "" or qty == 0:
            qty = hold_qty_map.get(symbol_str, "")
        qty_list.append(qty)
        price_row = get_price_fields_with_fallback(symbol_str, quotes, kite)
        price_rows.append(price_row)
        timestamp_list.append(now_str_timestamp)
        pnl_val = pos_pnl_map.get(symbol_str)
        if pnl_val in ("", None):
            pnl_val = hold_pnl_map.get(symbol_str)
        pnl_list.append(pnl_val if pnl_val not in (None, "") else "")

    # Now write in batch:
    sheet.range(f"B2:B{MAX_INPUT_ROWS+1}").value = [[q] for q in qty_list]
    sheet.range(f"C2:K{MAX_INPUT_ROWS+1}").value = price_rows
    sheet.range(f"R2:R{MAX_INPUT_ROWS+1}").value = [[t] for t in timestamp_list]
    sheet.range(f"Q2:Q{MAX_INPUT_ROWS+1}").value = [[p] for p in pnl_list]

def update_portfolio_sheet(sheet_port, positions, quotes):
    portfolio_rows_data = []
    if positions:
        for pos in positions:
            if pos.get('quantity', 0) == 0:
                continue
            exch_p = pos.get('exchange', '').upper()
            tsym_p = pos.get('tradingsymbol', '').upper()
            symbol_key_p = f"{exch_p}:{tsym_p}" if exch_p and tsym_p else None
            live_ltp_p = pos.get('last_price', '')
            if symbol_key_p and symbol_key_p in quotes:
                live_ltp_p = quotes[symbol_key_p].get('last_price', live_ltp_p)
            pnl_p = pos.get('pnl', '')
            portfolio_rows_data.append([
                pos.get('instrument_token', ''), tsym_p, exch_p, pos.get('quantity'),
                pos.get('average_price'), live_ltp_p, pnl_p,
                pos.get('realised_pnl', ''), pos.get('m2m', pos.get('unrealised_pnl', pnl_p))
            ])
    update_sheet_with_data(sheet_port, portfolio_rows_data, "A2", 9, MAX_PORTFOLIO_ROWS + 1, "K1")

def update_holdings_sheet(sheet_hold, holdings, quotes):
    holding_rows_data = []
    if holdings:
        for h_item in holdings:
            exch_h = h_item.get('exchange', '').upper()
            tsym_h = h_item.get('tradingsymbol', '').upper()
            symbol_key_h = f"{exch_h}:{tsym_h}" if exch_h and tsym_h else None
            live_ltp_h = h_item.get('last_price', '')
            if symbol_key_h and symbol_key_h in quotes:
                live_ltp_h = quotes[symbol_key_h].get('last_price', live_ltp_h)
            pnl_h = h_item.get('pnl', '')
            holding_rows_data.append([
                tsym_h, exch_h, h_item.get('isin', ''), h_item.get('quantity'), h_item.get('t1_quantity', ''),
                h_item.get('average_price'), live_ltp_h, h_item.get('close_price', ''), pnl_h,
                h_item.get('day_change', ''), h_item.get('day_change_percentage', '')
            ])
    update_sheet_with_data(sheet_hold, holding_rows_data, "A2", 11, MAX_HOLDINGS_ROWS + 1, "L1")

def update_orders_sheet(sheet_ords, kite, clear_all):
    try:
        if clear_all:
            sheet_ords.range("A2:P201").value = [[""]*16]*200

        orders_api_data = kite.orders()
        orders_rows_for_sheet = []
        if orders_api_data:
            for o_item in orders_api_data:
                orders_rows_for_sheet.append([
                    str(o_item.get("order_id")) if o_item.get("order_id") else "",
                    o_item.get("variety"),
                    o_item.get("status"),
                    o_item.get("tradingsymbol"),
                    o_item.get("exchange"),
                    o_item.get("order_type"),
                    o_item.get("product"),
                    o_item.get("transaction_type"),
                    o_item.get("quantity"),
                    o_item.get("price"),
                    o_item.get("trigger_price"),
                    o_item.get("average_price"),
                    o_item.get("pending_quantity"),
                    o_item.get("filled_quantity"),
                    o_item.get("order_timestamp"),
                    o_item.get("parent_order_id", "")
                ])
        update_sheet_with_data(sheet_ords, orders_rows_for_sheet, "A2", 16, MAX_ORDERS_ROWS + 1, "AA1")
    except Exception as e:
        print(f"Error updating orders sheet: {e}")

def process_order_modifications(sheet_ords, kite):
    try:
        orders_data_block = sheet_ords.range(f"A2:X{min(MAX_ORDERS_ROWS, 200) + 1}").value
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
            parent_order_id = row_data[15] if len(row_data) > 15 else None
            modify_flag = str(row_data[16] or '').strip().upper()
            cancel_flag = str(row_data[17] or '').strip().upper()
            if cancel_flag in ["YES", "CANCEL", "C"]:
                if status in ["OPEN", "TRIGGER PENDING", "AMO MODIFIED", "AMO REQ RECEIVED"]:
                    try:
                        if variety == 'co' and parent_order_id:
                            kite.cancel_order(variety=variety, order_id=order_id, parent_order_id=str(parent_order_id))
                        else:
                            kite.cancel_order(variety=variety, order_id=order_id)
                        sheet_ords.range(f"R{excel_row_num}").value = "Cancelled"
                        print(f"Order {order_id} cancelled successfully")
                    except Exception as e:
                        print(f"Order can't be cancelled with reason : {e}")
                        sheet_ords.range(f"R{excel_row_num}").value = f"Cancel Error: {str(e)[:20]}"
                else:
                    sheet_ords.range(f"R{excel_row_num}").value = "Not Cancellable"
            elif modify_flag in ["YES", "MODIFY"]:
                if status in ["OPEN", "TRIGGER PENDING", "AMO MODIFIED", "AMO REQ RECEIVED"]:
                    try:
                        new_price = row_data[18] if len(row_data) > 18 else None    
                        new_trigger = row_data[19] if len(row_data) > 19 else None    
                        new_qty = row_data[20] if len(row_data) > 20 else None       
                        new_order_type = row_data[21] if len(row_data) > 21 else None 
                        new_validity = row_data[22] if len(row_data) > 22 else None  
                        new_disclosed_qty = row_data[23] if len(row_data) > 23 else None 
                        modify_params = {
                            "variety": variety,
                            "order_id": order_id
                        }
                        if new_price is not None and str(new_price).strip() != "":
                            modify_params["price"] = float(new_price)
                        
                        if new_trigger is not None and str(new_trigger).strip() != "":
                            modify_params["trigger_price"] = float(new_trigger)
                        
                        if new_qty is not None and str(new_qty).strip() != "":
                            modify_params["quantity"] = int(new_qty)
                        
                        if new_order_type and str(new_order_type).strip() != "":
                            modify_params["order_type"] = str(new_order_type).upper()
                        
                        if new_validity and str(new_validity).strip() != "":
                            modify_params["validity"] = str(new_validity).upper()
                        
                        if new_disclosed_qty is not None and str(new_disclosed_qty).strip() != "":
                            modify_params["disclosed_quantity"] = int(new_disclosed_qty)
                        if variety == 'co' and parent_order_id:
                            modify_params["parent_order_id"] = str(parent_order_id)
                        
                        print(f"Modifying order with params: {modify_params}")
                        modified_order_id = kite.modify_order(**modify_params)
                        sheet_ords.range(f"Q{excel_row_num}").value = f"Modified: {modified_order_id}"
                        sheet_ords.range(f"S{excel_row_num}:X{excel_row_num}").value = ""
                        print(f"Order {order_id} modified successfully")
                        
                    except Exception as e:
                        error_msg = f"Modify Error: {str(e)[:30]}"
                        sheet_ords.range(f"Q{excel_row_num}").value = error_msg
                        print(f"Modify failed for {order_id}: {e}")
                else:
                    sheet_ords.range(f"Q{excel_row_num}").value = "Not Modifiable"
                    
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

def should_clear_today(config_file="last_clear_date.txt"):
    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    last_cleared = None
    if os.path.exists(config_file):
        with open(config_file, "r") as f:
            last_cleared = f.read().strip()
    if last_cleared != today_str:
        with open(config_file, "w") as f:
            f.write(today_str)
        return True
    return False

def autofill_input_sheet_with_portfolio_holdings(inp_sheet, positions, holdings, max_rows=200, clear_all=False):
    if clear_all:
        inp_sheet.range(f"A2:A{max_rows+1}").value = [[""]] * max_rows
        inp_sheet.range(f"B2:K{max_rows+1}").value = [[""]*10] * max_rows
        inp_sheet.range(f"Q2:R{max_rows+1}").value = [[""]*2] * max_rows
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
    if clear_all:
        all_rows = []
        if holdings_syms:
            all_rows.append("HOLDINGS")
            all_rows.extend(holdings_syms)
        if portfolio_syms:
            all_rows.append("PORTFOLIO")
            all_rows.extend(portfolio_syms)
        all_rows.append("MANUAL")
        inp_sheet.range(f"A2:A{max_rows+1}").value = [[s] for s in all_rows] + [[""]] * (max_rows - len(all_rows))
        for i, s in enumerate(all_rows):
            cell = inp_sheet.range(f"A{2 + i}")
            if s in ("HOLDINGS", "PORTFOLIO", "MANUAL"):
                cell.font.bold = True
            else:
                cell.font.bold = False
        return
    existing = inp_sheet.range(f"A2:A{max_rows+1}").value
    manual_start = None
    for i, s in enumerate(existing):
        if isinstance(s, str) and s.strip().upper() == "MANUAL":
            manual_start = i
            break
    all_rows = []
    if holdings_syms:
        all_rows.append("HOLDINGS")
        all_rows.extend(holdings_syms)
    if portfolio_syms:
        all_rows.append("PORTFOLIO")
        all_rows.extend(portfolio_syms)
    if manual_start is not None:
        inp_sheet.range(f"A2:A{2+manual_start}").value = [[s] for s in all_rows + ["MANUAL"]]
    else:
        inp_sheet.range(f"A2:A{len(all_rows)+2}").value = [[s] for s in all_rows]

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

def process_input_sheet_orders(sheet_inp, kite, order_id_map):
    try:
        data = sheet_inp.range("A2:Y201").value
        for i, row in enumerate(data):
            idx = i + 2
            row = (row + [None]*25)[:25]
            symbol = row[0]
            qty = row[11]
            direction = row[12]
            entry_signal = row[13]
            entry_status = row[15]
            variety = row[18]
            order_type = row[19]
            product = row[20]
            validity = row[21]
            price = row[22]
            trigger_price = row[23]
            ttl_minutes = row[24]
            if (symbol and qty and direction and entry_signal and variety and order_type and product and validity and
                str(entry_signal).strip().lower() == "yes" and (not entry_status or entry_status == "")):
                try:
                    symbol = symbol.strip().upper()
                    exchange, tradingsymbol = symbol.split(":")
                    order_params = dict(
                        variety=variety.strip().lower(),
                        exchange=exchange,
                        tradingsymbol=tradingsymbol,
                        transaction_type=direction.strip().upper(),
                        quantity=int(qty),
                        order_type=order_type.strip().upper(),
                        product=product.strip().upper(),
                        validity=validity.strip().upper()
                    )
                    if order_type.strip().upper() in ["LIMIT", "SL", "SL-M"] and price:
                        order_params["price"] = float(price)
                    if order_type.strip().upper() in ["SL", "SL-M"] and trigger_price:
                        order_params["trigger_price"] = float(trigger_price)
                    if validity.strip().upper() == "TTL" and ttl_minutes:
                        order_params["validity_ttl"] = int(ttl_minutes)
                    order_id = kite.place_order(**order_params)
                    order_id_map[idx] = order_id
                    sheet_inp.range(f"P{idx}").value = f"Placed: {order_id}"
                    sheet_inp.range(f"N{idx}").value = ""
                except Exception as e:
                    sheet_inp.range(f"P{idx}").value = f"ERROR: {str(e)}"
                    sheet_inp.range(f"N{idx}").value = "" 
            elif entry_status and isinstance(entry_status, str) and (
                "TRIGGER PENDING" in entry_status or "PENDING" in entry_status or "OPEN" in entry_status):
                order_id = order_id_map.get(idx)
                if order_id:
                    try:
                        order_history = kite.order_history(order_id=order_id)
                        if order_history:
                            last_update = order_history[-1]
                            status = last_update['status']
                            filled_qty = last_update['filled_quantity']
                            if status in ["TRIGGER PENDING", "OPEN"]:
                                entry_cell = f"TRIGGER PENDING ({filled_qty})"
                            elif status == "COMPLETE":
                                entry_cell = f"ORDERED ({filled_qty})"
                            elif status == "PARTIAL":
                                entry_cell = f"PARTIAL ({filled_qty})"
                            elif status in ["REJECTED", "CANCELLED"]:
                                entry_cell = f"ERROR: {status}"
                            else:
                                entry_cell = f"PENDING ({filled_qty})"
                            sheet_inp.range(f"P{idx}").value = entry_cell
                    except Exception as e:
                        sheet_inp.range(f"P{idx}").value = f"ERROR: {str(e)}"
    except Exception as e:
        print(f"Error in process_input_sheet_orders: {e}")

def set_input_sheet_defaults(inp_sheet, max_rows=200):
    for i in range(max_rows):
        row_num = i + 2
        symbol = inp_sheet.range(f"A{row_num}").value
        if not symbol or not isinstance(symbol, str) or not symbol.strip():
            continue
        symbol_up = symbol.strip().upper()
        if symbol_up in ("HOLDINGS", "PORTFOLIO", "MANUAL"):
            continue
        s_val = inp_sheet.range(f"S{row_num}").value
        if not s_val:
            inp_sheet.range(f"S{row_num}").value = "regular"
        u_val = inp_sheet.range(f"U{row_num}").value
        if not u_val:
            if "BFO" in symbol_up or "NFO" in symbol_up:
                inp_sheet.range(f"U{row_num}").value = "NRML"
            elif "BSE" in symbol_up or "NSE" in symbol_up:
                inp_sheet.range(f"U{row_num}").value = "CNC"
            else:
                inp_sheet.range(f"U{row_num}").value = ""
        v_val = inp_sheet.range(f"V{row_num}").value
        if not v_val:
            inp_sheet.range(f"V{row_num}").value = "DAY"
