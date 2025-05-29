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

1.  **API Credentials:**
    *   Open the `auth.py` file.
    *   Locate the placeholders for `api_key` and `api_secret`.
    *   Replace these placeholders with your actual Zerodha Kite Connect API Key and API Secret.
    ```
    # In auth.py (example - structure might vary slightly)
    api_key = "YOUR_API_KEY"
    api_secret = "YOUR_API_SECRET"
    ```

2.  **Generate Access Token:**
    *   Run the `auth.py` script once to generate an `access_token.txt` file. This script will likely prompt you to log in via a URL and provide a request token.
    ```
    python auth.py
    ```
    *   Follow the on-screen instructions. Upon successful authentication, an `access_token.txt` file containing your access token will be created in the project directory.

3.  **Excel File (`options_live.xlsm`):**
    *   Ensure the `options_live.xlsm` file is in the same directory as the Python scripts.
    *   The Excel file might have VBA macros that interact with the Python scripts via `xlwings`. Familiarize yourself with any buttons or UDFs (User Defined Functions) provided in the Excel sheet.

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
