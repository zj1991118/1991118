# crawl.py
import pandas as pd
import requests
from bs4 import BeautifulSoup
from pathlib import Path

OUTPUT_FILE = Path("pe_result.xlsx")

def clear_old_output():
    if OUTPUT_FILE.exists():
        OUTPUT_FILE.unlink()
        print(f"✅ 已清除旧文件：{OUTPUT_FILE.name}")

def fetch_page_html():
    url = "https://worldperatio.com"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        print(f"📥 网页获取成功，长度：{len(resp.text)}")
        return resp.text
    except Exception as e:
        print(f"❌ 请求网页失败：{e}")
        return ""

def parse_pe_table(html):
    if not html:
        return None
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if not table:
        print("⚠️ 页面里没有找到表格！")
        return None
    # 使用内置 html.parser，不需要 lxml
    df = pd.read_html(str(table), flavor="html.parser")[0]
    print(f"📌 抓到表格，一共 {len(df)} 行数据")
    return df

def save_result(df):
    if df is None or len(df) == 0:
        print("⚠️ 没有抓到有效表格数据，不生成输出文件")
        return
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"✅ 已保存全部 PE 表格到 pe_result.xlsx")

if __name__ == "__main__":
    clear_old_output()
    html = fetch_page_html()
    df_data = parse_pe_table(html)
    save_result(df_data)
