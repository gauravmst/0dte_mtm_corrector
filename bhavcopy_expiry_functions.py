import io
import os
import re
import time
import zipfile

import pandas as pd
import requests
from xlsxwriter.utility import xl_col_to_name

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
# index options that come in NSE's F&O bhavcopy
NFO_UNDERLYINGS = ("NIFTY", "BANKNIFTY")
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

    df_bhav_nfo["Underlying"] = df_bhav_nfo["Symbol"].str.replace("OPTIDX", "", n=1, regex=False)
    df_bhav_nfo = df_bhav_nfo[
        (df_bhav_nfo["Date"] == pd.to_datetime(expiry_nfo))
        & df_bhav_nfo["Underlying"].isin(NFO_UNDERLYINGS)
        & df_bhav_nfo["Symbol"].str.startswith("OPTIDX")
    ].copy()
    # NIFTY and BANKNIFTY share strikes, so the lookup key carries the underlying: "BANKNIFTY55000CE"
    df_bhav_nfo["Key"] = df_bhav_nfo["Underlying"] + df_bhav_nfo["Strike_Type"].astype(str)

    if df_bhav_nfo.empty:
        return df_bhav_nfo[["Key", "SETTLEMENT"]]

    strike_parts = df_bhav_nfo["Strike_Type"].astype(str).str.extract(r"(\d+)(CE|PE)$")
    df_bhav_nfo["Strike"] = pd.to_numeric(strike_parts[0], errors="coerce")
    df_bhav_nfo["Option_Type"] = strike_parts[1]

    # each underlying has its own spot, so decide "expiry day" and settle per underlying
    for underlying in df_bhav_nfo["Underlying"].unique():
        part = df_bhav_nfo[df_bhav_nfo["Underlying"] == underlying]
        if "UNDRLNG_ST" in part.columns:
            underlying_series = part["UNDRLNG_ST"].fillna(part["SETTLEMENT"])
            expiry_day_like = (part["SETTLEMENT"] - part["UNDRLNG_ST"]).abs().le(0.01).mean() >= 0.80
        else:
            underlying_series = part["SETTLEMENT"]
            settlement_count = part["SETTLEMENT"].dropna().shape[0]
            unique_settlement_count = part["SETTLEMENT"].dropna().nunique()
            expiry_day_like = settlement_count > 0 and unique_settlement_count <= max(2, int(settlement_count * 0.05))

        if expiry_day_like:
            ce = part.index[part["Option_Type"] == "CE"]
            pe = part.index[part["Option_Type"] == "PE"]
            df_bhav_nfo.loc[ce, "SETTLEMENT"] = (underlying_series.loc[ce] - part.loc[ce, "Strike"]).clip(lower=0)
            df_bhav_nfo.loc[pe, "SETTLEMENT"] = (part.loc[pe, "Strike"] - underlying_series.loc[pe]).clip(lower=0)

    return df_bhav_nfo[["Key", "SETTLEMENT"]].drop_duplicates("Key")


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

    if exchange == "NFO" and und in NFO_UNDERLYINGS:
        prices, expiry, key = nfo_px, pd.to_datetime(expiry_nfo), f"{und}{strike}{opt}"
    elif exchange == "BFO" and und == "SENSEX":
        prices, expiry, key = bfo_px, pd.to_datetime(expiry_bfo), f"{strike}{opt}"[-7:]
    else:
        return None, "OPEN - NO BHAVCOPY FOR THIS UNDERLYING (NOT SETTLED)"

    if (year, month) != (expiry.year, expiry.month) or (day is not None and day != expiry.day):
        return None, "OPEN - DIFFERENT EXPIRY (NOT SETTLED)"
    if key not in prices:
        return None, "OPEN - STRIKE NOT IN BHAVCOPY (NOT SETTLED)"
    return prices[key], None


# columns the realized / unrealized P&L is calculated from
REQUIRED_POSITION_COLS = ["Exchange", "Symbol", "Net Qty", "Buy Qty", "Sell Qty",
                          "Buy Avg Price", "Sell Avg Price", "P&L"]
SUMMARY_VALUE_COLS = ["Realized Profit", "Unrealized Profit", "P&L",
                      "Calculated_Realized_PNL", "Calculated_Unrealized_PNL", "Settled P&L"]


def build_settled_positions(position_df, nfo_bhav_df, bfo_bhav_df, expiry_nfo, expiry_bfo):
    """
    Output 1 - compiled position file with settlement columns added.

    Every row is kept. Open rows (Net Qty != 0) whose contract expires on expiry_nfo / expiry_bfo
    are settled at the bhavcopy price. P&L is built from average prices and quantities:

        Realized   = (Sell Avg - Buy Avg) x Sell Qty      if Net Qty >= 0
                     (Sell Avg - Buy Avg) x Buy Qty       if Net Qty <  0
        Unrealized = (Settlement - Buy Avg) x |Net Qty|   if Net Qty >  0
                     (Sell Avg - Settlement) x |Net Qty|  if Net Qty <  0
        Settled P&L = Realized + Unrealized

    Flat rows (Net Qty == 0) and open rows that can't be settled keep the reported P&L.
    """
    missing = [c for c in REQUIRED_POSITION_COLS if c not in position_df.columns]
    if missing:
        raise ValueError(f"Missing columns in position file: {missing}")

    df = position_df.copy()
    nfo_px = _price_map(nfo_bhav_df, "Key", "SETTLEMENT")
    bfo_px = _price_map(bfo_bhav_df, "Symbols", "Close Price")

    # the Excel formulas read these cells directly, so they must be real numbers
    for c in ["Net Qty", "Buy Qty", "Sell Qty", "Buy Avg Price", "Sell Avg Price", "P&L"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    net, buy_qty, sell_qty = df["Net Qty"], df["Buy Qty"], df["Sell Qty"]
    buy_avg, sell_avg, pnl = df["Buy Avg Price"], df["Sell Avg Price"], df["P&L"]

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

    realized = ((sell_avg - buy_avg) * sell_qty).where(net >= 0, (sell_avg - buy_avg) * buy_qty)
    unrealized = pd.Series(0.0, index=df.index)
    unrealized = unrealized.mask(settled & (net > 0), (price - buy_avg) * net.abs())
    unrealized = unrealized.mask(settled & (net < 0), (sell_avg - price) * net.abs())

    df["Position Status"] = is_open.map({True: "OPEN", False: "CLOSED"})
    df["Settlement Price"] = price.where(settled)
    df["Calculated_Realized_PNL"] = realized
    df["Calculated_Unrealized_PNL"] = unrealized
    df["Settled P&L"] = pnl.where(~settled, realized + unrealized).round(2)
    df["Settle Status"] = "FLAT - NO SETTLEMENT NEEDED"
    df.loc[is_open, "Settle Status"] = reason[is_open]
    df.loc[settled, "Settle Status"] = "SETTLED"
    return df


def build_settled_summary(settled_df):
    """Output 2 - one row per user: realized, unrealized, MTM P&L and final settled P&L."""
    df = settled_df.copy()
    for c in SUMMARY_VALUE_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["Open Rows"] = df["Position Status"].eq("OPEN")
    df["Settled Rows"] = df["Settle Status"].eq("SETTLED")
    df["Open Unsettled Rows"] = df["Open Rows"] & ~df["Settled Rows"]

    summary = (
        df.groupby(["UserID", "Server"], sort=True)[SUMMARY_VALUE_COLS + ["Open Rows", "Settled Rows", "Open Unsettled Rows"]]
        .sum()
        .reset_index()
    )
    summary["Settlement Impact"] = summary["Settled P&L"] - summary["P&L"]
    summary[SUMMARY_VALUE_COLS + ["Settlement Impact"]] = summary[SUMMARY_VALUE_COLS + ["Settlement Impact"]].round(2)
    return summary


def _write_position_formulas(ws, df):
    """Replace the calculated columns with live Excel formulas (F2 on a cell shows how it was built)."""
    col = {name: xl_col_to_name(i) for i, name in enumerate(df.columns)}
    idx = {name: i for i, name in enumerate(df.columns)}
    for n in range(len(df)):
        r = n + 2  # Excel row; row 1 is the header
        net, bq, sq = (f"{col[c]}{r}" for c in ("Net Qty", "Buy Qty", "Sell Qty"))
        ba, sa, sp, pnl = (f"{col[c]}{r}" for c in ("Buy Avg Price", "Sell Avg Price", "Settlement Price", "P&L"))
        cr, cu = f"{col['Calculated_Realized_PNL']}{r}", f"{col['Calculated_Unrealized_PNL']}{r}"
        formulas = {
            "Calculated_Realized_PNL": f"=IF({net}<0,({sa}-{ba})*{bq},({sa}-{ba})*{sq})",
            "Calculated_Unrealized_PNL": (
                f"=IF(ISNUMBER({sp}),IF({net}>0,({sp}-{ba})*ABS({net}),"
                f"IF({net}<0,({sa}-{sp})*ABS({net}),0)),0)"
            ),
            "Settled P&L": f"=IF(ISNUMBER({sp}),ROUND({cr}+{cu},2),{pnl})",
        }
        for name, formula in formulas.items():
            ws.write_formula(n + 1, idx[name], formula, None, float(df[name].iloc[n]))


def _write_summary_formulas(ws, summary_df, settled_df):
    """Per-user totals as SUMIFS / COUNTIFS over the Position sheet."""
    last = len(settled_df) + 1
    pos = {name: f"Position!${xl_col_to_name(i)}$2:${xl_col_to_name(i)}${last}"
           for i, name in enumerate(settled_df.columns)}
    col = {name: xl_col_to_name(i) for i, name in enumerate(summary_df.columns)}
    idx = {name: i for i, name in enumerate(summary_df.columns)}
    for n in range(len(summary_df)):
        r = n + 2
        keys = f"{pos['UserID']},{col['UserID']}{r},{pos['Server']},{col['Server']}{r}"
        formulas = {c: f"=ROUND(SUMIFS({pos[c]},{keys}),2)" for c in SUMMARY_VALUE_COLS}
        formulas["Open Rows"] = f'=COUNTIFS({keys},{pos["Position Status"]},"OPEN")'
        formulas["Settled Rows"] = f'=COUNTIFS({keys},{pos["Settle Status"]},"SETTLED")'
        formulas["Open Unsettled Rows"] = f"={col['Open Rows']}{r}-{col['Settled Rows']}{r}"
        formulas["Settlement Impact"] = f"=ROUND({col['Settled P&L']}{r}-{col['P&L']}{r},2)"
        for name, formula in formulas.items():
            ws.write_formula(n + 1, idx[name], formula, None, float(summary_df[name].iloc[n]))


def save_settled_outputs(settled_df, summary_df, out_dir="."):
    """Write Compiled_Settled_<date>.xlsx with a Position sheet and a Summary sheet, both formula-driven."""
    date = settled_df["Date"].dropna().iloc[0] if "Date" in settled_df.columns and settled_df["Date"].notna().any() else ""
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, f"Compiled_Settled_{date}.xlsx")
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        settled_df.to_excel(writer, sheet_name="Position", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        _write_position_formulas(writer.sheets["Position"], settled_df)
        _write_summary_formulas(writer.sheets["Summary"], summary_df, settled_df)
    return path


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
# One call: date + positions in, the Excel output out
# ---------------------------------------------------------------------------
def run_settlement(date, position, out_dir=".", nfo_file=None, bfo_file=None):
    """
    Paste the expiry date, get the Excel output (Position + Summary sheets).

        run_settlement("29-09-2026", "Compiled_Position_29-09-2026.csv", out_dir="output")

    date      expiry / bhavcopy date (dd-mm-yyyy string, date or datetime). Contracts expiring on
              this date are settled; contracts of any other expiry are left as reported.
    position  compiled position DataFrame, or path to its CSV.
    nfo_file / bfo_file
              optional already-downloaded bhavcopy CSV (file object) to use instead of fetching,
              for when an exchange blocks the download.

    Returns (settled_df, summary_df, xlsx_path).
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
    path = save_settled_outputs(settled, summary, out_dir)
    print("Saved:", path)
    print(settled["Settle Status"].value_counts().to_string())
    return settled, summary, path


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
