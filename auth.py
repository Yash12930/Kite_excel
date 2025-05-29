from kiteconnect import KiteConnect

kite = KiteConnect(api_key="xxxx") # Replace with your actual API key
request_token = "xxxxxx" # Replace with your actual request token
data = kite.generate_session(request_token, api_secret="#####")  # Replace with your actual API secret

def save_access_token(token, filename="access_token.txt"):
    with open(filename, "w") as f:
        f.write(token)

access_token = data["access_token"]
save_access_token(access_token)

def load_access_token(filename="access_token.txt"):
    try:
        with open(filename, "r") as f:
            token = f.read().strip()
        return token
    except FileNotFoundError:
        return None

# Usage:
access_token = load_access_token()
if not access_token:
    raise Exception("Access token file not found. Please login and save token first.")

kite.set_access_token(access_token)
