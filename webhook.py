import threading
import queue
import datetime
import time
import traceback
import xlwings as xw
from flask import Flask, request, jsonify
from kiteconnect import KiteConnect, KiteTicker
from kiteconnect.exceptions import KiteException 

from functions import (
    update_input_sheet, update_portfolio_sheet, update_holdings_sheet,
    update_orders_sheet, process_order_modifications, update_settings_sheet,
    fetch_holdings, autofill_input_sheet_with_portfolio_holdings, process_input_sheet_orders,
    set_input_sheet_defaults
)

API_KEY = "v6khvkcvmrjba7fb" 
ACCESS_TOKEN_FILE = "excelapporsome/access_token.txt" 
EXCEL_FILE = "excelapporsome/options_live.xlsm" 
PREFETCH_EXCHANGES = ["NSE", "NFO", "BSE", "BFO", "MCX"]
GENERAL_UPDATE_INTERVAL_SECONDS = 2
app = Flask(__name__)
refresh_queue = queue.Queue() 
order_id_map = {}

kite, wb, inp, port, hold, ords, sett, kws = (None,) * 8
live_ticks, all_instruments_cache, symbol_to_token_map, token_to_symbol_map = {}, {}, {}, {}
subscribed_tokens, previous_symbol_in_row = set(), {}
live_ticks_lock = threading.Lock()

def dt_now_str(): 
    return datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]

def get_instrument_token_from_cache(symbol_str):
    global symbol_to_token_map, token_to_symbol_map, all_instruments_cache
    if not symbol_str or ":" not in symbol_str: return None
    if symbol_str in symbol_to_token_map: return symbol_to_token_map[symbol_str]
    try:
        exch, sym_only = symbol_str.split(":", 1)
        if exch in all_instruments_cache:
            for inst in all_instruments_cache[exch]:
                if inst["tradingsymbol"] == sym_only:
                    token = inst["instrument_token"]
                    symbol_to_token_map[symbol_str] = token 
                    token_to_symbol_map[token] = symbol_str   
                    return token
    except Exception as e:
        pass
    return None

@app.route("/refresh_symbol", methods=["POST"])
def refresh_symbol_route(): 
    data = request.json or {}; sym_vba = data.get("symbol", "").strip().upper(); row = data.get("row", 0)
    if row < 2: return jsonify(status="err", msg="Bad row from VBA"), 400
    refresh_queue.put((row, sym_vba)) 
    print(f"[{dt_now_str()}] Flask: Queued row {row} (VBA: '{sym_vba}')")
    return jsonify(status="ok"), 200

def on_ticks_background(ws, ticks):
    global live_ticks, token_to_symbol_map, live_ticks_lock 
    with live_ticks_lock: 
        for t in ticks:
            sym = token_to_symbol_map.get(t["instrument_token"]) 
            if sym: 
                live_ticks[sym] = t 

def on_connect_background(ws, response):
    print(f"[{dt_now_str()}] WS Connected (background thread).")

def on_close_background(ws, code, reason):
    print(f"[{dt_now_str()}] WS Closed (background thread): {code} - {reason}")

def on_error_background(ws, code, reason):
    print(f"[{dt_now_str()}] WS Error (background thread): {code} - {reason}")

def process_single_row_refresh_in_main_thread(row_req, sym_vba_sent_debug, current_positions_for_refresh):
    global live_ticks, token_to_symbol_map, inp, kite, subscribed_tokens, kws, previous_symbol_in_row, symbol_to_token_map, live_ticks_lock
    try:
        actual_current_sym_excel = ""
        try:
            val = inp.range(f"A{row_req}").value 
            if isinstance(val, str): actual_current_sym_excel = val.strip().upper()
        except Exception as e_xl_read_main_refresh:
            pass
        
        sym_to_process = actual_current_sym_excel 
        prev_sym_in_row = previous_symbol_in_row.get(row_req)
        
        should_clear_row = False
        if not sym_to_process: 
            should_clear_row = True
        elif ":" not in sym_to_process: 
            should_clear_row = True
        
        token_for_processing = None
        if not should_clear_row:
            token_for_processing = get_instrument_token_from_cache(sym_to_process)
            if not token_for_processing: 
                should_clear_row = True
        if should_clear_row:
            inp.range(f"B{row_req}").value = "" 
            inp.range(f"C{row_req}:K{row_req}").value = [[""]*9] 
            inp.range(f"R{row_req}").value = ""
            if prev_sym_in_row: 
                old_tok = symbol_to_token_map.pop(prev_sym_in_row, None) 
                if old_tok: 
                    token_to_symbol_map.pop(old_tok, None) 
                    if old_tok in subscribed_tokens: 
                        if kws and kws.is_connected(): kws.unsubscribe([old_tok]) 
                        subscribed_tokens.remove(old_tok) 
            previous_symbol_in_row[row_req] = None 
            return

        if prev_sym_in_row and prev_sym_in_row != sym_to_process:
            old_tok = symbol_to_token_map.pop(prev_sym_in_row, None)
            if old_tok: 
                token_to_symbol_map.pop(old_tok, None)
                if old_tok in subscribed_tokens:
                    if kws and kws.is_connected(): kws.unsubscribe([old_tok])
                    subscribed_tokens.remove(old_tok)
        
        if token_for_processing and token_for_processing not in subscribed_tokens:
            if kws and kws.is_connected():
                kws.subscribe([token_for_processing])
                kws.set_mode(kws.MODE_FULL, [token_for_processing])
            subscribed_tokens.add(token_for_processing)
        previous_symbol_in_row[row_req] = sym_to_process
        
        quote_data_for_this_row = None
        with live_ticks_lock: 
            quote_data_for_this_row = live_ticks.get(sym_to_process) 
        
        if not quote_data_for_this_row:
            try:
                api_resp = kite.quote([sym_to_process]) 
                specific_quote = api_resp.get(sym_to_process, {})
                quote_data_for_this_row = { 
                    "ohlc": specific_quote.get("ohlc",{}), "last_price": specific_quote.get("last_price",""),
                    "volume": specific_quote.get("volume",""), "average_price": specific_quote.get("average_price",""),
                    "depth": specific_quote.get("depth",{})
                }
            except Exception as e_q_main_refresh: 
                quote_data_for_this_row = {} 
        ohlc = quote_data_for_this_row.get("ohlc", {}); depth = quote_data_for_this_row.get("depth", {})
        prices = [
            ohlc.get("open",""), ohlc.get("high",""), ohlc.get("low",""),
            quote_data_for_this_row.get("last_price",""), 
            quote_data_for_this_row.get("volume",""), 
            quote_data_for_this_row.get("average_price",""), 
            depth.get("buy",[{}])[0].get("price","") if depth.get("buy") else "",
            depth.get("sell",[{}])[0].get("price","") if depth.get("sell") else "",
            ohlc.get("close","") 
        ]
        pos_map = {f"{p['exchange']}:{p['tradingsymbol']}":p["quantity"] for p in current_positions_for_refresh} 
        port_qty = pos_map.get(sym_to_process, "") 
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        inp.range(f"B{row_req}").value = port_qty 
        inp.range(f"C{row_req}:K{row_req}").value = [prices]
        inp.range(f"R{row_req}").value = timestamp

    except Exception as e_process_item_main_thread:
        pass

if __name__=="__main__":
    try:
        print("Initializing Kite Connect and Excel...")
        kite = KiteConnect(api_key=API_KEY)
        kite.set_access_token(open(ACCESS_TOKEN_FILE).read().strip())
        print(f"Pre-fetching instruments for: {PREFETCH_EXCHANGES}...")
        for exch in PREFETCH_EXCHANGES:
            try:
                fetched_inst_list_main = kite.instruments(exchange=exch) 
                all_instruments_cache[exch] = fetched_inst_list_main 
                for inst_detail_main in fetched_inst_list_main: 
                    full_sym_str_main = f"{exch}:{inst_detail_main['tradingsymbol']}"
                    token_val_main = inst_detail_main['instrument_token']
                    symbol_to_token_map[full_sym_str_main] = token_val_main
                    token_to_symbol_map[token_val_main] = full_sym_str_main
            except Exception as e_instr_fetch_main: pass
        wb = xw.Book(EXCEL_FILE) 
        sheet_names_list_main = ["INPUT", "Portfolio", "Holdings", "Orders", "Funds"]
        for sheet_name_str_main in sheet_names_list_main: 
            if sheet_name_str_main not in [s.name for s in wb.sheets]: 
                wb.sheets.add(sheet_name_str_main)
        inp, port, hold, ords, sett = ( wb.sheets[s_n_main] for s_n_main in sheet_names_list_main ) 
    except Exception as e_main_startup_init_block: 
        print(f"CRITICAL ERROR during Kite/Excel initialization: {e_main_startup_init_block}")
        import traceback; traceback.print_exc(); exit()
    try: 
        print("Starting Flask server...")
        threading.Thread(target=lambda: app.run(port=5000, debug=False, use_reloader=False), daemon=True).start()
    except Exception as e_main_startup_flask_block: 
        print(f"CRITICAL ERROR starting Flask: {e_main_startup_flask_block}")
        import traceback; traceback.print_exc(); exit()
    try:
        print("Initializing KiteTicker for background operation...")
        kws = KiteTicker(API_KEY, kite.access_token, reconnect=True, reconnect_max_tries=50, reconnect_max_delay=60)
        kws.on_ticks = on_ticks_background    
        kws.on_connect = on_connect_background 
        kws.on_close = on_close_background
        kws.on_error = on_error_background
        kws_thread = threading.Thread(target=lambda: kws.connect(threaded=True)) 
        kws_thread.daemon = True 
        kws_thread.start()
        time.sleep(5) 
        if kws.is_connected():
            initial_excel_symbols_main  = inp.range("A2:A201").value or []
            initial_tokens_to_subscribe_main_list = []
            for i_main, symbol_text_main_init in enumerate(initial_excel_symbols_main):
                excel_row_num_main_init = i_main + 2 
                current_excel_sym_main_init = None
                if isinstance(symbol_text_main_init, str) and ":" in symbol_text_main_init:
                    current_excel_sym_main_init = symbol_text_main_init.strip().upper()
                previous_symbol_in_row[excel_row_num_main_init] = current_excel_sym_main_init 
                if current_excel_sym_main_init:
                    token_main_init_val = get_instrument_token_from_cache(current_excel_sym_main_init)
                    if token_main_init_val and token_main_init_val not in subscribed_tokens: 
                        initial_tokens_to_subscribe_main_list.append(token_main_init_val)
            if initial_tokens_to_subscribe_main_list:
                kws.subscribe(initial_tokens_to_subscribe_main_list) 
                kws.set_mode(kws.MODE_FULL, initial_tokens_to_subscribe_main_list) 
                subscribed_tokens.update(initial_tokens_to_subscribe_main_list)
        else:
            print(f"[{dt_now_str()}] WARNING: KiteTicker did NOT connect after 5s.")
    except Exception as e_kws_startup_block:
        print(f"CRITICAL ERROR Initializing or Starting KiteTicker: {e_kws_startup_block}")
        import traceback; traceback.print_exc()
    print("Starting MAIN PROCESSING LOOP. Ctrl+C to exit.")
    last_general_sheets_update_timestamp = 0
    is_first_run = True

    try:
        while True:
            items_processed_in_main_loop_iter = 0
            while not refresh_queue.empty():
                if items_processed_in_main_loop_iter >= 1: break
                try:
                    row_req_from_q, sym_vba_sent_from_q = refresh_queue.get_nowait()
                    temp_current_positions_for_row_refresh = []
                    try: temp_current_positions_for_row_refresh = kite.positions().get("net", [])
                    except Exception as e_temp_pos_main_loop_refresh_iter: pass
                    process_single_row_refresh_in_main_thread(row_req_from_q, sym_vba_sent_from_q, temp_current_positions_for_row_refresh)
                    items_processed_in_main_loop_iter +=1
                except queue.Empty: break
                except Exception as e_main_q_item_processing_loop: pass
            try:
                current_main_loop_live_ticks_copy = {}
                with live_ticks_lock: 
                    current_main_loop_live_ticks_copy = live_ticks.copy()
                
                if current_main_loop_live_ticks_copy:
                    try: 
                        current_positions = kite.positions().get("net", [])
                        current_holdings = fetch_holdings(kite)
                    except Exception: 
                        current_positions, current_holdings = [], []
                    update_input_sheet(inp, kite, current_holdings, current_main_loop_live_ticks_copy, current_positions)
            except Exception as e_realtime_update: 
                pass

            if time.time() - last_general_sheets_update_timestamp > GENERAL_UPDATE_INTERVAL_SECONDS:
                current_main_loop_live_ticks_copy = {}
                with live_ticks_lock: current_main_loop_live_ticks_copy = live_ticks.copy()
                general_update_main_loop_positions, general_update_main_loop_holdings = [], []
                try:
                    general_update_main_loop_positions = kite.positions().get("net", [])
                    general_update_main_loop_holdings = fetch_holdings(kite)
                except Exception as e_gen_update_data_fetch_main_loop_iter: pass
                try:
                    autofill_input_sheet_with_portfolio_holdings(inp, general_update_main_loop_positions, general_update_main_loop_holdings, max_rows=200, clear_all=is_first_run)
                    set_input_sheet_defaults(inp, max_rows=200)
                    update_portfolio_sheet(port, general_update_main_loop_positions, current_main_loop_live_ticks_copy)
                    update_holdings_sheet(hold, general_update_main_loop_holdings, current_main_loop_live_ticks_copy)
                    update_orders_sheet(ords, kite, clear_all=is_first_run)
                    process_order_modifications(ords, kite)
                    is_first_run = False
                    try: update_settings_sheet(sett, kite)
                    except Exception as e_gen_update_margin_main_loop_iter: pass
                except Exception as e_general_sheet_update_main_loop_iter: pass
                try:
                    process_input_sheet_orders(inp, kite, order_id_map) 
                except Exception as e_input_orders:
                    pass
                last_general_sheets_update_timestamp = time.time()
            
            time.sleep(0.01)
    except KeyboardInterrupt: print("\nUser exit (Ctrl+C from main loop).")
    except Exception as e_main_loop_critical_outermost:
        print(f"CRITICAL UNHANDLED ERROR in main processing loop: {e_main_loop_critical_outermost}")
        import traceback; traceback.print_exc()
    finally: 
        print("Initiating final shutdown sequence...")
        if kws and kws.is_connected(): 
            try: kws.stop_retry(); kws.close(1000, "Program shutdown")
            except Exception as e_ws_final_shutdown_main_thread: pass
        if wb: 
            try: wb.close()
            except Exception as e_wb_final_shutdown_main_thread: pass
        print("Program exited.")
