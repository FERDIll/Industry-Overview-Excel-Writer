import requests, datetime, time, csv, os, sys

URL = "https://query1.finance.yahoo.com/v8/finance/chart/{}"
TICKERS = ["SPY","QQQ","XLK","XLF","IWF","IWD","TLT","GLD"]

def fetch(t):
    r = requests.get(
        URL.format(t),
        params={"range": "1y", "interval": "1d"},
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=20
    )
    r.raise_for_status()
    j = r.json()
    closes = j["chart"]["result"][0]["indicators"]["quote"][0]["close"]
    closes = [x for x in closes if x is not None]
    if len(closes) < 130:
        return None, None, None
    last = closes[-1]
    r3 = last / closes[-63] - 1
    r6 = last / closes[-126] - 1
    return last, r3, r6

def main(out_dir):
    spy_last, spy_r3, spy_r6 = fetch("SPY")
    if spy_last is None:
        raise RuntimeError("SPY data unavailable")

    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    out_csv = os.path.join(out_dir, "dashboard_data.csv")
    tmp_csv = out_csv + ".tmp"

    with open(tmp_csv, "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Ticker","Last","Ret3M","Ret6M","RS6M","UpdatedAt"])
        for t in TICKERS:
            last, r3, r6 = fetch(t)
            if last is None:
                w.writerow([t,"","","","",now])
            else:
                w.writerow([t,last,r3,r6,(r6 - spy_r6),now])
            time.sleep(0.2)

    os.replace(tmp_csv, out_csv)  # atomic swap
    print(f"Wrote {out_csv} at {now}")

if __name__ == "__main__":
    # write CSV to the script's folder
    out_dir = os.path.dirname(os.path.abspath(__file__))
    main(out_dir)
