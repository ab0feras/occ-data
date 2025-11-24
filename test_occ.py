import requests
import sys

# =======================================================
# بيانات البروكسي السكني (Evomi) - بناء يدوي للصيغة القياسية
# =======================================================
# نستخدم البيانات التي ظهرت في رسالة الخطأ الأخيرة
PROXY_USER = "dhafersa5"
PROXY_PASS = "VGpVsUraAdkBOKlBuy3L" 
PROXY_HOST = "core-residential.evomi.com"
PROXY_PORT = "1000"

# بناء سلسلة البروكسي بالصيغة القياسية لـ requests: http://user:pass@host:port
PROXY_ADDRESS = f"http://{PROXY_USER}:{PROXY_PASS}@{PROXY_HOST}:{PROXY_PORT}"

# عنوان OCC للاختبار
TEST_URL = "https://www.theocc.com/Market-Data/Market-Data-Reports/Other-Market-Data-Info/Batch-Processing/Series-Search-Batch-Processing" 

# رؤوس قياسية لمحاكاة متصفح حقيقي
STANDARD_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'Accept-Language': 'en-US,en;q=0.9,ar;q=0.8',
    'Connection': 'close',  # **التعديل الجديد: إغلاق الاتصال لمنع أخطاء IncompleteRead**
    'Referer': 'https://www.google.com/', 
}

def simple_occ_proxy_test():
    print("Starting test connection to The OCC using Evomi Residential Proxy...")
    
    # استخدام الصيغة القياسية
    proxies = {
        "http": PROXY_ADDRESS,
        "https": PROXY_ADDRESS,
    }

    try:
        # زيادة المهلة (Timeout) قليلاً لاحتساب عدم استقرار البروكسي
        response = requests.get(TEST_URL, headers=STANDARD_HEADERS, proxies=proxies, timeout=30)
        status = response.status_code
        
        if status == 200:
            print(f"SUCCESS: Status Code 200 (OK). Residential Proxy working!")
            return "SUCCESS_200"
        
        elif status == 403:
            print(f"FAILURE: Status Code 403 (Forbidden). Residential IP was also blocked or requires rotation.")
            return "FAILED_403"
            
        else:
            print(f"FAILURE: Received unexpected status code {status}.")
            return f"FAILED_{status}"

    except requests.exceptions.ProxyError:
        print(f"FAILURE: Proxy connection error. Evomi connection failed.")
        return "FAILED_PROXY_ERROR"
    except Exception as e:
        print(f"FAILURE: Network or Timeout Error: {e}")
        return "FAILED_NETWORK_ERROR"

if __name__ == '__main__':
    result = simple_occ_proxy_test()
    print(f"\nFINAL RESULT: {result}")
    
    if "SUCCESS" not in result:
        sys.exit(1)
