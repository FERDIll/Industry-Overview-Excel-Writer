#!/usr/bin/env python3
"""
RK-style dashboard updater
- Pulls latest and historical close prices from Yahoo Finance "chart" endpoint (public).
- Computes 1M/3M/6M/12M returns and Relative Strength vs SPY.
- Writes results into the existing Excel template (Data sheet), preserving formatting.

Usage:
  python3 update_dashboard.py --xlsx "/path/to/rk_dashboard_template.xlsx"

Notes:
- This uses requests (pip install requests openpyxl).
- Yahoo can rate limit; this script caches responses under ./cache by default.
"""

import argparse
import datetime as dt
import json
import os
import time
from typing import Dict, List, Optional, Tuple

import requests
from openpyxl import load_workbook

YF_CHART = "https://query1.finance.yahoo.com/v8/finance/chart/{ticker}"
DEFAULT_TICKERS = [
    ("INDEX","SPY","S&P 500 (SPY)"),
    ("INDEX","QQQ","Nasdaq 100 (QQQ)"),
    ("INDEX","IWM","Russell 2000 (IWM)"),
    ("INDEX","DIA","Dow 30 (DIA)"),
    ("SECTOR","XLK","Technology (XLK)"),
    ("SECTOR","XLF","Financials (XLF)"),
    ("SECTOR","XLE","Energy (XLE)"),
    ("SECTOR","XLV","Health Care (XLV)"),
    ("STYLE","IWF","Growth (IWF)"),
    ("STYLE","IWD","Value (IWD)"),
    ("STYLE","MTUM","Momentum (MTUM)"),
    ("STYLE","QUAL","Quality (QUAL)"),
    ("RISK","TLT","Long Treasuries (TLT)"),
    ("RISK","GLD","Gold (GLD)"),
    ("RISK","UUP","US Dollar (UUP)"),
    ("RISK","VIXY","VIX Short-Term (VIXY)"),
    ("COMMOD","USO","Oil (USO)"),
]

def ensure_dir(p: str) -> None:
    os.makedirs(p, exist_ok=True)

def cache_path(cache_dir: str, ticker: str, range_: str, interval: str) -> str:
    safe = ticker.replace("^","_")
    return os.path.join(cache_dir, f"{safe}_{range_}_{interval}.json")

def fetch_chart(ticker: str, range_: str, interval: str, cache_dir: str, max_age_sec: int = 3600) -> Dict:
    """
    Fetch Yahoo Finance chart JSON with basic caching.
    """
    ensure_dir(cache_dir)
    cp = cache_path(cache_dir, ticker, range_, interval)
    now = time.time()
    if os.path.exists(cp):
        age = now - os.path.getmtime(cp)
        if age <= max_age_sec:
            with open(cp, "r", encoding="utf-8") as f:
                return json.load(f)

    params = {"range": range_, "interval": interval, "includePrePost": "false", "events": "div,splits"}
    url = YF_CHART.format(ticker=ticker)
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
    }
    r = requests.get(url, params=params, headers=headers, timeout=20)
    r.raise_for_status()
    data = r.json()

    with open(cp, "w", encoding="utf-8") as f:
        json.dump(data, f)
    return data

def last_valid_close(chart_json: Dict) -> Optional[float]:
    try:
        closes = chart_json["chart"]["result"][0]["indicators"]["quote"][0]["close"]
        # walk backwards for last non-null
        for x in reversed(closes):
            if x is not None:
                return float(x)
    except Exception:
        return None
    return None

def close_n_days_ago(chart_json: Dict, n: int) -> Optional[float]:
    """
    Approximate close n trading days ago from the closes array.
    """
    try:
        closes = chart_json["chart"]["result"][0]["indicators"]["quote"][0]["close"]
        # collect non-null closes
        nn = [c for c in closes if c is not None]
        if len(nn) <= n:
            return None
        return float(nn[-(n+1)])
    except Exception:
        return None

def pct_return(last: Optional[float], past: Optional[float]) -> Optional[float]:
    if last is None or past is None or past == 0:
        return None
    return (last / past) - 1.0

def compute_returns(ticker: str, cache_dir: str) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]:
    """
    Returns: last, ret_1m, ret_3m, ret_6m, ret_12m
    Uses 2y daily data to compute approximate trading-day lookbacks.
    """
    # 2y daily gives enough points for 12m lookback
    j = fetch_chart(ticker, range_="2y", interval="1d", cache_dir=cache_dir, max_age_sec=3600)
    last = last_valid_close(j)
    # approx trading days: 1m=21, 3m=63, 6m=126, 12m=252
    r1 = pct_return(last, close_n_days_ago(j, 21))
    r3 = pct_return(last, close_n_days_ago(j, 63))
    r6 = pct_return(last, close_n_days_ago(j, 126))
    r12 = pct_return(last, close_n_days_ago(j, 252))
    return last, r1, r3, r6, r12

def write_to_excel(xlsx_path: str, rows: List[Dict]) -> None:
    wb = load_workbook(xlsx_path)
    if "Data" not in wb.sheetnames:
        raise RuntimeError("Workbook must contain a sheet named 'Data'")
    ws = wb["Data"]

    # Map existing tickers to row numbers (col A)
    ticker_to_row = {}
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, 1).value
        if isinstance(t, str) and t.strip():
            ticker_to_row[t.strip().upper()] = r

    updated_at = dt.datetime.now()

    for row in rows:
        t = row["Ticker"].upper()
        r = ticker_to_row.get(t)
        if r is None:
            # append new
            r = ws.max_row + 1
            ws.cell(r, 1).value = t

        ws.cell(r, 2).value = row["Name"]
        ws.cell(r, 3).value = row["Category"]
        ws.cell(r, 4).value = row["Last"]
        ws.cell(r, 5).value = row["Ret_1M"]
        ws.cell(r, 6).value = row["Ret_3M"]
        ws.cell(r, 7).value = row["Ret_6M"]
        ws.cell(r, 8).value = row["Ret_12M"]
        ws.cell(r, 9).value = row["RS_3M"]
        ws.cell(r,10).value = row["RS_6M"]
        ws.cell(r,11).value = row["RS_12M"]
        ws.cell(r,12).value = "Yahoo Finance (chart)"
        ws.cell(r,13).value = dt.datetime.now()
        ws.cell(r,14).value = updated_at

    wb.save(xlsx_path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="Path to the Excel template/workbook to update.")
    ap.add_argument("--cache", default="cache", help="Cache directory for Yahoo responses.")
    ap.add_argument("--no-cache", action="store_true", help="Disable caching (not recommended).")
    args = ap.parse_args()

    cache_dir = args.cache
    if args.no_cache:
        cache_dir = os.path.join(args.cache, "disabled_" + str(int(time.time())))

    # compute benchmark (SPY) returns for relative strength
    spy_last, spy_1, spy_3, spy_6, spy_12 = compute_returns("SPY", cache_dir=cache_dir)

    out_rows = []
    for cat, ticker, name in DEFAULT_TICKERS:
        try:
            last, r1, r3, r6, r12 = compute_returns(ticker, cache_dir=cache_dir)
        except Exception as e:
            last=r1=r3=r6=r12=None

        rs3 = (r3 - spy_3) if (r3 is not None and spy_3 is not None) else None
        rs6 = (r6 - spy_6) if (r6 is not None and spy_6 is not None) else None
        rs12 = (r12 - spy_12) if (r12 is not None and spy_12 is not None) else None

        out_rows.append({
            "Ticker": ticker,
            "Name": name,
            "Category": cat,
            "Last": last,
            "Ret_1M": r1,
            "Ret_3M": r3,
            "Ret_6M": r6,
            "Ret_12M": r12,
            "RS_3M": rs3,
            "RS_6M": rs6,
            "RS_12M": rs12,
        })

        time.sleep(0.2)  

    write_to_excel(args.xlsx, out_rows)

if __name__ == "__main__":
    main()
