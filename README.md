# Kite Excel Live Options Tracker

This project integrates Zerodha's Kite Connect API with Microsoft Excel to provide live Nifty and Bank Nifty options data. It uses `xlwings` to communicate between Python and Excel, `Flask` to handle webhook updates from Kite Ticker, and `kiteconnect` for API interactions.

## Features

*   Fetches live Nifty & Bank Nifty options data.
*   Displays data in an Excel spreadsheet (`options_live.xlsm`).
*   Handles Kite Connect API authentication.
*   Utilizes Kite Ticker for real-time data updates via webhooks.
*   Daily data clearing mechanism.

## Prerequisites

*   Python 3.7+
*   A Zerodha Kite Connect API Key and API Secret.
*   A Zerodha account.
*   Microsoft Excel installed on your system.
*   `pip` for installing Python packages.
*   Git (for cloning the repository).

## Installation

1.  **Clone the repository:**
    ```
    git clone https://github.com/Yash12930/Kite_excel.git
    cd Kite_excel
    ```

2.  **Create and activate a virtual environment (recommended):**
    ```
    python -m venv venv
    # On Windows
    venv\Scripts\activate
    # On macOS/Linux
    source venv/bin/activate
    ```

3.  **Install Python dependencies:**
    ```
    pip install -r requirements.txt
    ```

4.  **Enable Macros in Excel:**
    *   Open `options_live.xlsm`.
    *   Excel will likely show a security warning about macros. You need to enable macros for the project to function correctly.
    *   Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings` and select "Enable all macros" (not recommended for general use, but necessary for this workbook if you trust its source) or "Disable all macros with notification" (so you can enable them when you open the file).

## Configuration

1.  **API Credentials in `auth.py`:**
    *   Open the `auth.py` file.
    *   Locate the placeholders for `api_key` and `api_secret`.
    *   Replace these placeholders with your actual Zerodha Kite Connect API Key and API Secret. This script is primarily used for the initial generation of the `access_token`.
    ```
    # In auth.py (example - structure might vary slightly)
    api_key = "YOUR_API_KEY"
    api_secret = "YOUR_API_SECRET"
    ```

2.  **API Key in `webhook.py` (and potentially other scripts):**
    *   Open the `webhook.py` file (and any other scripts like `functions.py` if they independently initialize `KiteConnect`).
    *   Ensure that the `api_key` is also available to this script. It might be hardcoded (less secure for sharing), read from a configuration file, or passed from `auth.py` if your structure allows. Commonly, it's re-declared or read from a central config. For example:
    ```
    # In webhook.py, near the top or where KiteConnect is initialized
    api_key = "YOUR_API_KEY" # Make sure this is your actual API key
    # ...
    # kite = KiteConnect(api_key=api_key)
    # kite.set_access_token(access_token_from_file)
    ```
    *   **Important**: Be consistent. The `api_key` used in `webhook.py` to initialize `KiteConnect` must be the same one used in `auth.py` to generate the `access_token`.

3.  **Generate Access Token:**
    *   Run the `auth.py` script once:
    ```
    python auth.py
    ```
    *   This will guide you through the login process (usually opening a Zerodha login URL) and, upon successful authentication with the `request_token`, it will generate an `access_token.txt` file. This file contains the `access_token` that `webhook.py` will use.

4.  **Excel File (`options_live.xlsm`):**
    *   Ensure the `options_live.xlsm` file is in the same directory as the Python scripts.
    *   Enable macros in Excel as described previously.

## How It Works (Conceptual Flow)

1.  **Authentication (`auth.py`)**:
    *   User provides their `api_key` and `api_secret` within `auth.py`.
    *   User runs `auth.py`. The script uses the `api_key` to generate a login URL [1].
    *   User logs in, obtains a `request_token`.
    *   `auth.py` uses the `api_key`, `request_token`, and `api_secret` to generate an `access_token` and saves it to `access_token.txt` [1, 2].

2.  **Initialization (`webhook.py`)**:
    *   The `webhook.py` script reads the `api_key` (either hardcoded, from a config, or defined within the script).
    *   It reads the `access_token` from `access_token.txt`.
    *   It initializes `KiteConnect` using this `api_key` and then sets the `access_token` [2, 5].
    *   It initializes `KiteTicker` using the `api_key` and `access_token`.
    *   It starts a `Flask` web server to listen for POST requests (webhooks from Kite Ticker).
    *   It subscribes to instrument tokens for Nifty/Bank Nifty options.

3.  **Real-time Data (`KiteTicker` & `Flask`)**:
    *   Kite Ticker pushes live data to the `/webhook` endpoint defined in `webhook.py`.
    *   The Flask app receives this data.

4.  **Excel Integration (`xlwings` & `functions.py`)**:
    *   The `webhook.py` (or functions in `functions.py` called by it) uses `xlwings` to connect to the `options_live.xlsm` workbook.
    *   Live data is written to specific cells/ranges in the Excel sheet.

5.  **User Interface (`options_live.xlsm`)**:
    *   The user views and interacts with the live data in Excel.
    *   VBA macros within Excel might call Python functions (via `xlwings`) for actions like manual refresh, changing subscribed instruments, etc.

## How to Use

1.  **Start the Webhook Server:**
    *   Run the `webhook.py` script. This script starts a Flask web server to listen for real-time ticks from Kite Ticker and updates the Excel sheet.
    ```
    python webhook.py
    ```
    *   Keep this terminal window open while you are using the application. It will log incoming tick data and any errors.

2.  **Open the Excel File:**
    *   Open `options_live.xlsm`.
    *   If the Python scripts are running correctly and `xlwings` is set up, the Excel sheet should start receiving live data.

3.  **Interacting with the Sheet:**
    *   The `options_live.xlsm` sheet is where you will see the live options data.
    *   There might be buttons or cells within Excel to trigger specific actions (e.g., refresh data, subscribe to specific instruments) handled by VBA and `xlwings` calling Python functions in `functions.py`.

## File Descriptions

*   **`auth.py`**: Handles authentication with Kite Connect API to obtain an `access_token`. Requires your `api_key` and `api_secret`.
*   **`webhook.py`**: Main script to run. Initializes Kite Ticker, subscribes to instruments, and runs a Flask server to receive live tick data. Interacts with `xlwings` to update Excel.
*   **`functions.py`**: Contains various helper functions, potentially for tasks like fetching instrument lists, formatting data, and handling Excel updates via `xlwings`.
*   **`options_live.xlsm`**: The Excel macro-enabled workbook where live options data is displayed and potentially managed.
*   **`requirements.txt`**: Lists the necessary Python packages for the project.
*   **`access_token.txt`**: Stores the generated access token after successful authentication. This file is read by `webhook.py`.
*   **`last_clear_date.txt`**: Likely used to keep track of the date when certain data (e.g., daily option chain data) was last cleared or reset.

## How It Works (Conceptual Flow)

1.  **Authentication (`auth.py`)**: User runs `auth.py` once to enter API credentials, authenticate with Kite, and generate an `access_token.txt`.
2.  **Initialization (`webhook.py`)**:
    *   Reads the `access_token.txt`.
    *   Initializes `KiteConnect` and `KiteTicker`.
    *   Starts a `Flask` web server to listen for POST requests (webhooks from Kite Ticker).
    *   Subscribes to instrument tokens for Nifty/Bank Nifty options.
3.  **Real-time Data (`KiteTicker` & `Flask`)**:
    *   Kite Ticker pushes live data to the `/webhook` endpoint defined in `webhook.py`.
    *   The Flask app receives this data.
4.  **Excel Integration (`xlwings` & `functions.py`)**:
    *   The `webhook.py` (or functions in `functions.py` called by it) uses `xlwings` to connect to the `options_live.xlsm` workbook.
    *   Live data is written to specific cells/ranges in the Excel sheet.
5.  **User Interface (`options_live.xlsm`)**:
    *   The user views and interacts with the live data in Excel.
    *   VBA macros within Excel might call Python functions (via `xlwings`) for actions like manual refresh, changing subscribed instruments, etc.

## Troubleshooting

*   **Authentication Errors**:
    *   Ensure your `api_key` and `api_secret` in `auth.py` are correct.
    *   The `access_token` expires daily. You might need to re-run `auth.py` to generate a new token if you get authentication errors.
*   **`xlwings` Connection Issues**:
    *   Make sure `options_live.xlsm` is open or can be opened by `xlwings`.
    *   Ensure the `xlwings` add-in is correctly installed in Excel if you are using UDFs directly from Excel without running a Python script.
    *   Check that no other Python process is exclusively locking the Excel file.
*   **Webhook Not Receiving Data**:
    *   Verify that `webhook.py` is running without errors.
    *   Check your internet connection.
    *   Ensure your Kite API subscription is active and webhooks are enabled for your app.
*   **Data Not Updating in Excel**:
    *   Check the console output of `webhook.py` for errors.
    *   Ensure macros are enabled in `options_live.xlsm`.

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.
