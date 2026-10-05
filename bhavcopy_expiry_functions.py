import io
import os
import re
import time
import zipfile

import pandas as pd
import requests

# Anything smaller than this is an error page or a holiday placeholder, not a bhavcopy.
MIN_USABLE_BYTES = 500

# NSE rejects requests that don't look like a browser or lack the cookie its home page sets.
NSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": "https://www.nseindia.com/",
    "Connection": "keep-alive",
}

_MONTHS = {m: i for i, m in enumerate(
    ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"], 1
)}
_WEEKLY_MONTH = {**{str(i): i for i in range(1, 10)}, "O": 10, "N": 11, "D": 12}

# "NIFTY 29SEP2026 CE 22750"
_SYMBOL_SPACED = re.compile(r"^([A-Z]+) (\d{2})([A-Z]{3})(\d{4}) (CE|PE) (\d+)$")
# "NIFTY26SEP22900PE" (monthly), "SENSEX26O0170800PE" / "NIFTY2692922900PE" (weekly)
_SYMBOL_COMPACT = re.compile(r"^([A-Z]+?)(\d{2})(?:([A-Z]{3})|([1-9OND])(\d{2}))(\d+)(CE|PE)$")


def load_nfo_bhav_df(nfo_bhav_file, expiry_nfo):
    if not nfo_bhav_file:
        return None

    try:
        nfo_bhav_file.seek(0)
    except Exception:
        pass

    df_bhav_nfo = pd.read_csv(nfo_bhav_file)
    df_bhav_nfo.columns = df_bhav_nfo.columns.str.strip()

    if "CONTRACT_D" not in df_bhav_nfo.columns:
        raise ValueError("'CONTRACT_D' column missing in NFO Bhavcopy.")
    if "SETTLEMENT" not in df_bhav_nfo.columns:
        raise ValueError("'SETTLEMENT' column missing in NFO Bhavcopy.")

    contract = df_bhav_nfo["CONTRACT_D"].astype(str)
    df_bhav_nfo["Date"] = contract.str.extract(r"(\d{2}-[A-Z]{3}-\d{4})")
    df_bhav_nfo["Symbol"] = contract.str.extract(r"^(.*?)(\d{2}-[A-Z]{3}-\d{4})")[0]
    df_bhav_nfo["Strike_Type"] = contract.str.extract(r"(PE\d+|CE\d+)$")

    df_bhav_nfo["Date"] = pd.to_datetime(df_bhav_nfo["Date"], format="%d-%b-%Y", errors="coerce")
    df_bhav_nfo["Strike_Type"] = df_bhav_nfo["Strike_Type"].str.replace(
        r"^(PE|CE)(\d+)$", r"\2\1", regex=True
    )
    df_bhav_nfo["SETTLEMENT"] = pd.to_numeric(df_bhav_nfo["SETTLEMENT"], errors="coerce")
    if "UNDRLNG_ST" in df_bhav_nfo.columns:
        df_bhav_nfo["UNDRLNG_ST"] = pd.to_numeric(df_bhav_nfo["UNDRLNG_ST"], errors="coerce")

    df_bhav_nfo = df_bhav_nfo[
        (df_bhav_nfo["Date"] == pd.to_datetime(expiry_nfo))
        & (df_bhav_nfo["Symbol"] == "OPTIDXNIFTY")
    ].copy()

    if df_bhav_nfo.empty:
        return df_bhav_nfo[["Strike_Type", "SETTLEMENT"]]

    strike_parts = df_bhav_nfo["Strike_Type"].astype(str).str.extract(r"(\d+)(CE|PE)$")
    df_bhav_nfo["Strike"] = pd.to_numeric(strike_parts[0], errors="coerce")
    df_bhav_nfo["Option_Type"] = strike_parts[1]

    if "UNDRLNG_ST" in df_bhav_nfo.columns:
        underlying_series = df_bhav_nfo["UNDRLNG_ST"].fillna(df_bhav_nfo["SETTLEMENT"])
        expiry_day_like = (
            (df_bhav_nfo["SETTLEMENT"] - df_bhav_nfo["UNDRLNG_ST"]).abs().le(0.01).mean() >= 0.80
        )
    else:
        underlying_series = df_bhav_nfo["SETTLEMENT"]
        settlement_count = df_bhav_nfo["SETTLEMENT"].dropna().shape[0]
        unique_settlement_count = df_bhav_nfo["SETTLEMENT"].dropna().nunique()
        expiry_day_like = settlement_count > 0 and unique_settlement_count <= max(2, int(settlement_count * 0.05))

    if expiry_day_like:
        ce_mask = df_bhav_nfo["Option_Type"] == "CE"
        pe_mask = df_bhav_nfo["Option_Type"] == "PE"
        df_bhav_nfo.loc[ce_mask, "SETTLEMENT"] = (
            underlying_series.loc[ce_mask] - df_bhav_nfo.loc[ce_mask, "Strike"]
        ).clip(lower=0)
        df_bhav_nfo.loc[pe_mask, "SETTLEMENT"] = (
            df_bhav_nfo.loc[pe_mask, "Strike"] - underlying_series.loc[pe_mask]
        ).clip(lower=0)

    return df_bhav_nfo[["Strike_Type", "SETTLEMENT"]].drop_duplicates("Strike_Type")


def load_bfo_bhav_df(bfo_bhav_file, expiry_bfo):
    if not bfo_bhav_file:
        return None

    try:
        bfo_bhav_file.seek(0)
    except Exception:
        pass

    df_bhav_bfo = pd.read_csv(bfo_bhav_file)
    df_bhav_bfo.columns = df_bhav_bfo.columns.str.strip()

    required = ["Asset Code", "Expiry Date", "Series Code", "Close Price"]
    missing = [c for c in required if c not in df_bhav_bfo.columns]
    if missing:
        raise ValueError(f"Missing columns in BFO Bhavcopy: {missing}")

    df_bhav_bfo = df_bhav_bfo[df_bhav_bfo["Asset Code"] == "BSX"].copy()
    df_bhav_bfo["Expiry Date"] = pd.to_datetime(df_bhav_bfo["Expiry Date"], format="mixed", dayfirst=True)
    df_bhav_bfo = df_bhav_bfo[df_bhav_bfo["Expiry Date"] == pd.to_datetime(expiry_bfo)].copy()
    df_bhav_bfo["Symbols"] = df_bhav_bfo["Series Code"].astype(str).str[-7:]
    df_bhav_bfo["Close Price"] = pd.to_numeric(df_bhav_bfo["Close Price"], errors="coerce")

    if df_bhav_bfo.empty:
        return df_bhav_bfo[["Symbols", "Close Price"]]

    if "Strike Price" in df_bhav_bfo.columns:
        df_bhav_bfo["Strike"] = pd.to_numeric(df_bhav_bfo["Strike Price"], errors="coerce")
    else:
        df_bhav_bfo["Strike"] = pd.to_numeric(
            df_bhav_bfo["Symbols"].astype(str).str.extract(r"(\d+)(?:CE|PE)$")[0],
            errors="coerce",
        )

    if "Option Type (Call/Put)" in df_bhav_bfo.columns:
        df_bhav_bfo["Option_Type"] = df_bhav_bfo["Option Type (Call/Put)"].astype(str).str.strip().str.upper()
    else:
        df_bhav_bfo["Option_Type"] = df_bhav_bfo["Symbols"].astype(str).str.extract(r"(CE|PE)$")[0]

    underlying_col = "Underlying Asset Close Price"
    if underlying_col in df_bhav_bfo.columns:
        df_bhav_bfo[underlying_col] = pd.to_numeric(df_bhav_bfo[underlying_col], errors="coerce")
        underlying_series = df_bhav_bfo[underlying_col].fillna(df_bhav_bfo["Close Price"])
        expiry_day_like = (
            (df_bhav_bfo["Close Price"] - df_bhav_bfo[underlying_col]).abs().le(0.01).mean() >= 0.80
        )
    else:
        underlying_series = df_bhav_bfo["Close Price"]
        close_count = df_bhav_bfo["Close Price"].dropna().shape[0]
        unique_close_count = df_bhav_bfo["Close Price"].dropna().nunique()
        expiry_day_like = close_count > 0 and unique_close_count <= max(2, int(close_count * 0.05))

    if expiry_day_like:
        ce_mask = df_bhav_bfo["Option_Type"] == "CE"
        pe_mask = df_bhav_bfo["Option_Type"] == "PE"
        df_bhav_bfo.loc[ce_mask, "Close Price"] = (
            underlying_series.loc[ce_mask] - df_bhav_bfo.loc[ce_mask, "Strike"]
        ).clip(lower=0)
        df_bhav_bfo.loc[pe_mask, "Close Price"] = (
            df_bhav_bfo.loc[pe_mask, "Strike"] - underlying_series.loc[pe_mask]
        ).clip(lower=0)

    return df_bhav_bfo[["Symbols", "Close Price"]].drop_duplicates("Symbols")


def _parse_option_symbol(symbol):
    """Return (underlying, year, month, day|None, strike, CE/PE) or None if not an option symbol."""
    s = str(symbol).strip().upper()

    m = _SYMBOL_SPACED.match(s)
    if m:
        und, dd, mon, yyyy, opt, strike = m.groups()
        if mon not in _MONTHS:
            return None
        return und, int(yyyy), _MONTHS[mon], int(dd), int(strike), opt

    m = _SYMBOL_COMPACT.match(s)
    if m:
        und, yy, mon, weekly_m, weekly_d, strike, opt = m.groups()
        if mon:  # monthly contract: no day in the symbol
            if mon not in _MONTHS:
                return None
            return und, 2000 + int(yy), _MONTHS[mon], None, int(strike), opt
        return und, 2000 + int(yy), _WEEKLY_MONTH[weekly_m], int(weekly_d), int(strike), opt

    return None


def _price_map(bhav_df, key_col, price_col):
    if bhav_df is None or bhav_df.empty:
        return {}
    clean = bhav_df.dropna(subset=[price_col])
    return dict(zip(clean[key_col], clean[price_col]))


def _lookup_settlement(exchange, symbol, nfo_px, bfo_px, expiry_nfo, expiry_bfo):
    """Return (settlement_price, reason_if_missing)."""
    parsed = _parse_option_symbol(symbol)
    if parsed is None:
        return None, "OPEN - NOT AN INDEX OPTION (NOT SETTLED)"
    und, year, month, day, strike, opt = parsed

    if exchange == "NFO" and und == "NIFTY":
        prices, expiry, key = nfo_px, pd.to_datetime(expiry_nfo), f"{strike}{opt}"
    elif exchange == "BFO" and und == "SENSEX":
        prices, expiry, key = bfo_px, pd.to_datetime(expiry_bfo), f"{strike}{opt}"[-7:]
    else:
        return None, "OPEN - NO BHAVCOPY FOR THIS UNDERLYING (NOT SETTLED)"

    if (year, month) != (expiry.year, expiry.month) or (day is not None and day != expiry.day):
        return None, "OPEN - DIFFERENT EXPIRY (NOT SETTLED)"
    if key not in prices:
        return None, "OPEN - STRIKE NOT IN BHAVCOPY (NOT SETTLED)"
    return prices[key], None


def build_settled_positions(position_df, nfo_bhav_df, bfo_bhav_df, expiry_nfo, expiry_bfo):
    """
    Output 1 - compiled position file with settlement columns added.

    Every row is kept. Open rows (Net Qty != 0) whose contract expires on expiry_nfo / expiry_bfo
    are settled at the bhavcopy price:

        Settled P&L = Sell Value - Buy Value + Net Qty x Settlement Price

    Flat rows (Net Qty == 0) and open rows that can't be settled keep the reported P&L.
    """
    df = position_df.copy()
    nfo_px = _price_map(nfo_bhav_df, "Strike_Type", "SETTLEMENT")
    bfo_px = _price_map(bfo_bhav_df, "Symbols", "Close Price")

    net, buy, sell, pnl = (
        pd.to_numeric(df[c], errors="coerce").fillna(0)
        for c in ["Net Qty", "Buy Value", "Sell Value", "P&L"]
    )

    # parse each distinct contract once, not once per user row
    lookups = {
        (ex, sym): _lookup_settlement(ex, sym, nfo_px, bfo_px, expiry_nfo, expiry_bfo)
        for ex, sym in df[["Exchange", "Symbol"]].drop_duplicates().itertuples(index=False)
    }
    found = [lookups[k] for k in zip(df["Exchange"], df["Symbol"])]
    price = pd.Series([f[0] for f in found], index=df.index, dtype="float64")
    reason = pd.Series([f[1] for f in found], index=df.index)

    is_open = net != 0
    settled = is_open & price.notna()

    df["Position Status"] = is_open.map({True: "OPEN", False: "CLOSED"})
    df["Settlement Price"] = price.where(settled)
    df["Settled P&L"] = pnl.where(~settled, sell - buy + net * price).round(2)
    df["Settle Status"] = "FLAT - NO SETTLEMENT NEEDED"
    df.loc[is_open, "Settle Status"] = reason[is_open]
    df.loc[settled, "Settle Status"] = "SETTLED"
    df["Settle Formula"] = [
        f"{s:.2f} - {b:.2f} + ({n:g} x {p:.2f})" if ok else "P&L as reported"
        for s, b, n, p, ok in zip(sell, buy, net, price, settled)
    ]
    return df


def build_settled_summary(settled_df):
    """Output 2 - one row per user: realized, unrealized, MTM P&L and final settled P&L."""
    df = settled_df.copy()
    value_cols = ["Realized Profit", "Unrealized Profit", "P&L", "Settled P&L"]
    for c in value_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Open Rows"] = df["Position Status"].eq("OPEN")
    df["Settled Rows"] = df["Settle Status"].eq("SETTLED")
    df["Open Unsettled Rows"] = df["Open Rows"] & ~df["Settled Rows"]

    summary = (
        df.groupby(["UserID", "Server"], sort=True)[value_cols + ["Open Rows", "Settled Rows", "Open Unsettled Rows"]]
        .sum()
        .reset_index()
    )
    summary["Settlement Impact"] = summary["Settled P&L"] - summary["P&L"]
    summary[value_cols + ["Settlement Impact"]] = summary[value_cols + ["Settlement Impact"]].round(2)
    return summary


def save_settled_outputs(settled_df, summary_df, out_dir="."):
    """Write the two CSVs (Compiled_Position_Settled_<date>.csv, Compiled_Summary_<date>.csv)."""
    date = settled_df["Date"].dropna().iloc[0] if "Date" in settled_df.columns and settled_df["Date"].notna().any() else ""
    os.makedirs(out_dir, exist_ok=True)
    position_path = os.path.join(out_dir, f"Compiled_Position_Settled_{date}.csv")
    summary_path = os.path.join(out_dir, f"Compiled_Summary_{date}.csv")
    settled_df.to_csv(position_path, index=False)
    summary_df.to_csv(summary_path, index=False)
    return position_path, summary_path


# ---------------------------------------------------------------------------
# Fetching the bhavcopy from the exchanges
# ---------------------------------------------------------------------------
def nse_url(date_value):
    return (
        "https://www.nseindia.com/api/reports?"
        "archives=%5B%7B%22name%22%3A%22F%26O%20-%20Bhavcopy%20(fo.zip)%22%2C"
        "%22type%22%3A%22archives%22%2C%22category%22%3A%22derivatives%22%2C"
        f"%22section%22%3A%22equity%22%7D%5D&date={date_value.strftime('%d-%b-%Y')}"
        "&type=equity&mode=single"
    )


def bse_url(date_value):
    return ("https://www.bseindia.com/download/Bhavcopy/Derivative/"
            f"MS_{date_value.strftime('%Y%m%d')}-01.csv")


def _extract_op_csv(zip_bytes):
    """Pull the op*.csv settlement file out of NSE's fo.zip."""
    try:
        archive = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except zipfile.BadZipFile:
        raise ValueError("NSE returned a file that is not a readable ZIP.")
    with archive as z:
        names = [n for n in z.namelist()
                 if n.lower().endswith(".csv") and n.split("/")[-1].lower().startswith("op")]
        if not names:
            raise ValueError("NSE ZIP downloaded, but no op*.csv file was found inside it.")
        return z.read(names[0])


def fetch_nfo_bhavcopy(date_value):
    """Download the NSE F&O bhavcopy for a date; returns a seekable CSV buffer."""
    session = requests.Session()
    session.headers.update(NSE_HEADERS)
    session.get("https://www.nseindia.com/", timeout=10)  # sets the cookie the report API needs
    time.sleep(1)

    response = session.get(nse_url(date_value), timeout=30)
    if response.status_code != 200:
        raise ValueError(f"NSE returned HTTP {response.status_code}")
    if len(response.content) < MIN_USABLE_BYTES:
        raise ValueError("NSE returned no usable bhavcopy data. It may be a holiday or unavailable date.")
    return io.BytesIO(_extract_op_csv(response.content))


def fetch_bfo_bhavcopy(date_value):
    """Download the BSE derivatives bhavcopy for a date; returns a seekable CSV buffer."""
    response = requests.get(bse_url(date_value), headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    if response.status_code != 200:
        raise ValueError(f"BSE returned HTTP {response.status_code}")
    if len(response.content) < MIN_USABLE_BYTES:
        raise ValueError("BSE returned no usable bhavcopy data. It may be a holiday or unavailable date.")
    return io.BytesIO(response.content)


# ---------------------------------------------------------------------------
# One call: date + positions in, the two output files out
# ---------------------------------------------------------------------------
def run_settlement(date, position, out_dir=".", nfo_file=None, bfo_file=None):
    """
    Paste the expiry date, get both output files.

        run_settlement("29-09-2026", "Compiled_Position_29-09-2026.csv", out_dir="output")

    date      expiry / bhavcopy date (dd-mm-yyyy string, date or datetime). Contracts expiring on
              this date are settled; contracts of any other expiry are left as reported.
    position  compiled position DataFrame, or path to its CSV.
    nfo_file / bfo_file
              optional already-downloaded bhavcopy CSV (file object) to use instead of fetching,
              for when an exchange blocks the download.

    Returns (settled_df, summary_df, (position_path, summary_path)).
    """
    date_value = pd.to_datetime(date, dayfirst=True)
    position_df = pd.read_csv(position) if isinstance(position, (str, os.PathLike)) else position

    nfo_df = bfo_df = None
    for label, supplied, fetcher in (("NFO", nfo_file, fetch_nfo_bhavcopy), ("BFO", bfo_file, fetch_bfo_bhavcopy)):
        try:
            source = supplied if supplied is not None else fetcher(date_value)
            loaded = (load_nfo_bhav_df if label == "NFO" else load_bfo_bhav_df)(source, date_value)
        except Exception as error:
            print(f"[{label}] bhavcopy not used: {error}")
            continue
        print(f"[{label}] bhavcopy loaded: {len(loaded)} contracts expiring {date_value:%d-%b-%Y}")
        if label == "NFO":
            nfo_df = loaded
        else:
            bfo_df = loaded

    settled = build_settled_positions(position_df, nfo_df, bfo_df, date_value, date_value)
    summary = build_settled_summary(settled)
    paths = save_settled_outputs(settled, summary, out_dir)
    print("Saved:", *paths, sep="\n  ")
    print(settled["Settle Status"].value_counts().to_string())
    return settled, summary, paths


def main():
    path = input("Position file path: ").strip().strip('"').strip("'")
    if not os.path.isfile(path):
        raise SystemExit(f"File not found: {path}")

    # offer the date inside the position file as the default
    default = ""
    dates = pd.read_csv(path, usecols=["Date"])["Date"].dropna()
    if not dates.empty:
        default = str(dates.iloc[0])
    date = input(f"Bhavcopy date (dd-mm-yyyy) [{default}]: ").strip() or default
    if not date:
        raise SystemExit("A bhavcopy date is required.")

    default_out = os.path.join(os.path.dirname(os.path.abspath(path)), "output")
    out_dir = input(f"Output save folder [{default_out}]: ").strip().strip('"').strip("'") or default_out
    run_settlement(date, path, out_dir=out_dir)


if __name__ == "__main__":
    main()
