import requests
import sys

PROXY_ADDRESS = "http://core-residential.evomi.com:1000:dhafersa5:VGpVsUraAdkBOKlBuy3L"

# عنوان OCC للاختبار
TEST_URL = "https://www.theocc.com/Market-Data/Market-Data-Reports/Other-Market-Data-Info/Batch-Processing/Series-Search-Batch-Processing" 

# رؤوس قياسية لمحاكاة متصفح حقيقي
STANDARD_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'en-US,en;q=0.9,ar;q=0.8',
    'Connection': 'keep-alive',
    'Referer': 'https://www.google.com/', 
}

def simple_occ_proxy_test():
    print("Starting test connection to The OCC using Oxylabs Proxy...")
    
    proxies = {
        "http": PROXY_ADDRESS,
        "https": PROXY_ADDRESS,
    }

    try:
        response = requests.get(TEST_URL, headers=STANDARD_HEADERS, proxies=proxies, timeout=15)
        status = response.status_code
        
        if status == 200:
            print(f"SUCCESS: Status Code 200 (OK). Proxy working.")
            # تحقق إضافي لضمان جلب محتوى صحيح
            if "Series Search - Batch Processing" in response.text:
                print("Content Check: Expected page title found.")
            return "SUCCESS_200"
        
        elif status == 403:
            print(f"FAILURE: Status Code 403 (Forbidden). Proxy blocked.")
            return "FAILED_403"
            
        else:
            print(f"FAILURE: Received unexpected status code {status}.")
            return f"FAILED_{status}"

    except requests.exceptions.ProxyError:
        print(f"FAILURE: Proxy connection error. Check proxy format and authentication.")
        return "FAILED_PROXY_ERROR"
    except Exception as e:
        print(f"FAILURE: Network or Timeout Error: {e}")
        return "FAILED_NETWORK_ERROR"

if __name__ == '__main__':
    result = simple_occ_proxy_test()
    print(f"\nFINAL RESULT: {result}")
    
    # فشل GitHub Action إذا لم يكن ناجحاً
    if "SUCCESS" not in result:
        sys.exit(1)
