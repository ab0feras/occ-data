import requests
import time
import json
import sys

# Standard headers to simulate a real browser
STANDARD_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image:apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'en-US,en;q=0.9,ar;q=0.8',
    'Connection': 'keep-alive',
    'Referer': 'https://www.google.com/', 
}

# Test URL: Batch Processing information page
TEST_URL = "https://www.theocc.com/Market-Data/Market-Data-Reports/Other-Market-Data-Info/Batch-Processing/Series-Search-Batch-Processing" 

def simple_occ_test():
    print("Starting test connection to The OCC from GitHub Actions...")

    try:
        print("Attempt 1...")
        # Use headers in the request
        response = requests.get(TEST_URL, headers=STANDARD_HEADERS, timeout=15)
        status = response.status_code

        # Check response status
        if status == 200:
            print(f"SUCCESS: Status Code 200 (OK).")
            print(f"Content length: {len(response.text)} bytes.")

            # Check content to ensure it's not a hidden error page
            if "Series Search - Batch Processing" in response.text:
                print("Content Check: The expected page title was found.")
                return "SUCCESS_CONTENT_MATCH"
            else:
                print("Content Check: WARNING! Page title not matched. Status 200 received but content might be unexpected.")
                return "SUCCESS_CONTENT_MISMATCH"

        elif status == 403:
            print(f"FAILURE: Status Code 403 (Forbidden). Request blocked.")
            return "FAILED_403"

        else:
            print(f"FAILURE: Received unexpected status code {status}.")
            return f"FAILED_{status}"

    except requests.exceptions.RequestException as e:
        print(f"FAILURE: Network or Timeout Error: {e}")
        return "FAILED_NETWORK_ERROR"

if __name__ == '__main__':
    result = simple_occ_test()
    print(f"\nFINAL RESULT: {result}")

    # Fail the GitHub Action if the test fails
    if "SUCCESS" not in result:
        sys.exit(1)
