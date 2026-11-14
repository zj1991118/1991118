# crawl.py
import pandas as pd
import requests
from pathlib import Path

# 输出文件，只用相对路径
OUTPUT_FILE = Path("pe_result.xlsx")

def clear_old_output():
    """清除上次遗留文件"""
    if OUTPUT_FILE.exists():
        OUTPUT_FILE.unlink()
        print(f"✅ 已清除旧文件：{OUTPUT_FILE.name}")

def fetch_raw_data():
    """在这里填入真实数据源链接"""
    url = "https://你的真实数据源地址"
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        raw = resp.json()
        print(f"📥 原始获取条数：{len(raw)}")
        return raw
    except Exception as e:
        print(f"❌ 请求数据源失败：{e}")
        return []

def get_latest_record(raw_list):
    """取最新一条记录"""
    if not raw_list:
        return None
    df_raw = pd.DataFrame(raw_list)
    date_col = "date"
    df_raw[date_col] = pd.to_datetime(df_raw[date_col])
    df_sorted = df_raw.sort_values(by=date_col, ascending=False)
    latest_row = df_sorted.head(1)
    print(f"📌 抓到最新一条，日期：{latest_row.iloc[0][date_col]}")
    return latest_row

def save_result(df):
    """保存结果，无数据不生成文件"""
    if df is None or len(df) == 0:
        print("⚠️ 没有抓到有效最新记录，不生成输出文件")
        return
    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"✅ 已保存最新单条记录到 pe_result.xlsx")

if __name__ == "__main__":
    clear_old_output()
    raw_data = fetch_raw_data()
    latest = get_latest_record(raw_data)
    save_result(latest)
