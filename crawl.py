# -*- coding: utf-8 -*-
"""
美股主要指数 P/E 数据抓取脚本
=============================================================
功能：
  1. 从 worldperatio.com 抓取主要指数的 P/E 值及网页数据日期；
  2. 以【追加】方式写入 Excel 文件，保留已有数据，不清空；
  3. 每次运行只新增一行。
=============================================================
"""

# ==================== 可修改配置 ====================
# ① 数据源网址
URL = "https://worldperatio.com/major-stock-index-pe-ratios/"

# 改成当前目录，适配GitHub Actions云环境
FILE_PATH = "美股PE.xlsx"

# ② 需要抓取的指数名称（必须与网页上显示的名称一致）
INDEXES = ["Nasdaq 100", "S&P 500", "Russell 2000", "Dow Jones", "FTSE China 50"]

# ③ 工作表名与表头所在行
SHEET_NAME = "Sheet1"
HEADER_ROW = 1
# ==================== 配置结束 ====================


import os
import re
import sys
from datetime import datetime

try:
    import requests
    from bs4 import BeautifulSoup
    from openpyxl import Workbook, load_workbook
except ImportError as e:
    print("缺少依赖库，请先安装： pip install requests beautifulsoup4 openpyxl")
    print("缺失模块：", e)
    sys.exit(1)


def get_pe_data():
    """抓取网页，返回 (网页数据日期(ISO字符串), {指数名: P/E浮点数})。"""
    headers = {
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/120.0 Safari/537.36"),
        "Accept-Language": "en-US,en;q=0.9",
    }
    resp = requests.get(URL, headers=headers, timeout=30)
    resp.raise_for_status()
    html = resp.text

    # 1) 网页数据日期：形如 "Last Update: 12 October 2026"
    m = re.search(r"Last Update:\s*([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})", html)
    if not m:
        raise ValueError("未能在网页中找到 'Last Update' 日期，请检查网址是否有效")
    date_str = m.group(1)
    date_iso = datetime.strptime(date_str, "%d %B %Y").strftime("%Y-%m-%d")

    # 2) 各指数 P/E：定位表格行，指数名位于第 3 个 td 的 <a>，P/E 位于第 4 个 td
    soup = BeautifulSoup(html, "html.parser")
    pe = {}
    for tr in soup.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 4:
            continue
        name_a = tds[2].find("a")
        if not name_a:
            continue
        name = name_a.get_text(strip=True)
        if name in INDEXES:
            pe_td = tds[3]
            custom = pe_td.get("sorttable_customkey")   # 网页自带的精确值
            value = float(custom) if custom else float(pe_td.get_text(strip=True))
            pe[name] = value

    missing = [n for n in INDEXES if n not in pe]
    if missing:
        raise ValueError("以下指数未抓到 P/E 值：" + ", ".join(missing))
    return date_iso, pe


def ensure_workbook(path):
    """若文件不存在，则新建一个带表头的空工作簿。"""
    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        headers = ["执行程序时间", "网页数据日期"] + INDEXES
        for idx, h in enumerate(headers, start=1):
            ws.cell(row=HEADER_ROW, column=idx, value=h)
        wb.save(path)
        print("未找到文件，已新建：", path)


def append_to_excel(date_iso, pe):
    """把 (日期, 各指数P/E) 追加为一行，保留已有数据。返回写入行号。"""
    wb = load_workbook(FILE_PATH)
    ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active

    # 找到每个指数名对应的表头列
    header_map = {}
    for cell in ws[HEADER_ROW]:
        if cell.value and str(cell.value).strip() in INDEXES:
            header_map[str(cell.value).strip()] = cell.column
    missing_col = [n for n in INDEXES if n not in header_map]
    if missing_col:
        raise ValueError("Excel 表头中缺少列：" + ", ".join(missing_col))

    # 找第一个空数据行（从表头下一行开始，A 列为空即视为空行）
    next_row = HEADER_ROW + 1
    while next_row <= ws.max_row:
        if ws.cell(row=next_row, column=1).value is None:
            break
        next_row += 1

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row=next_row, column=1, value=run_time)   # 执行程序时间
    ws.cell(row=next_row, column=2, value=date_iso)   # 网页数据日期
    for name in INDEXES:
        ws.cell(row=next_row, column=header_map[name], value=pe[name])

    wb.save(FILE_PATH)
    return next_row


def main():
    print("正在抓取网页数据 ...")
    date_iso, pe = get_pe_data()

    ensure_workbook(FILE_PATH)
    row = append_to_excel(date_iso, pe)

    print("==============================================")
    print("写入成功！")
    print("文件路径 ：", FILE_PATH)
    print("写入行号 ：第", row, "行")
    print("网页数据日期：", date_iso)
    print("执行程序时间：", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    for name in INDEXES:
        print("  %-14s : %s" % (name, pe[name]))
    print("==============================================")


if __name__ == "__main__":
    main()
