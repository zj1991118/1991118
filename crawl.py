import requests
import time

def main():
    url = "https://httpbin.org/get"
    resp = requests.get(url, timeout=10)
    print("✅ 爬虫成功！")
    print("响应：", resp.json())

if __name__ == "__main__":
    main()
