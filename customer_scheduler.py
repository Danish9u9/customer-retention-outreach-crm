#!/usr/bin/env python3
"""
Build Customer_Outreach_Schedule.xlsx from a Shopify orders export.

Usage:
    python customer_scheduler.py                          # reads orders_export.csv next to this file
    python customer_scheduler.py path/to/orders.csv       # custom CSV
    python customer_scheduler.py --single-order-days 0   # one-time buyers -> contact tomorrow
    python customer_scheduler.py --single-order-days 30  # (default) one-time buyers -> contact 30 days after their order

The Excel file is read by app.py (the dashboard), so run this first.
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import time as dtime
from datetime import timedelta
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CSV = BASE_DIR / "orders_export.csv"
DEFAULT_OUTPUT = BASE_DIR / "Customer_Outreach_Schedule.xlsx"

TIMEZONE = "Asia/Karachi"
COUNTRY_CODE = "92"  # used to turn 03xx numbers into WhatsApp-ready 923xx numbers
REATTEMPT_WORDS = ("re-attempt", "reattempt", "re attempt")
LAPSED_AFTER_DAYS = 14  # customers whose last order is older than this are counted as lapsed in the summary
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def fail(message: str) -> None:
    print(f"ERROR: {message}")
    sys.exit(1)


def find_col(df: pd.DataFrame, *names: str) -> str | None:
    """Find a column by name, ignoring case, spaces and underscores."""
    lookup = {re.sub(r"[\s_]+", " ", c).strip().lower(): c for c in df.columns}
    for name in names:
        hit = lookup.get(re.sub(r"[\s_]+", " ", name).strip().lower())
        if hit:
            return hit
    return None


def clean(series: pd.Series) -> pd.Series:
    """Strip text and turn NaN / 'nan' / 'None' into an empty string."""
    out = series.fillna("").astype(str).str.strip()
    return out.mask(out.str.lower().isin(["nan", "none", "nat", "null"]), "")


def normalize_phone(raw: str) -> str:
    """Return digits only, in international format (no +). 0300... -> 92300..."""
    digits = re.sub(r"\D", "", str(raw))
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith(COUNTRY_CODE) and len(digits) >= 12:
        return digits
    if digits.startswith("0") and len(digits) == 11:
        return COUNTRY_CODE + digits[1:]
    if len(digits) == 10 and digits.startswith("3"):
        return COUNTRY_CODE + digits
    return digits


def valid_whatsapp(number: str) -> bool:
    """Pakistani mobiles must be 923xxxxxxxxx (12 digits); other countries: 10-15 digits."""
    if number.startswith("0"):  # local format that could not be converted (e.g. landline)
        return False
    if number.startswith(COUNTRY_CODE):
        return len(number) == 12 and number[2] == "3"
    return 10 <= len(number) <= 15


def format_time(minutes: int) -> str:
    return dtime(minutes // 60, minutes % 60).strftime("%I:%M %p")


def valid_email(value: str) -> str:
    """Return the email lowercased if it looks valid, otherwise an empty string."""
    value = str(value).strip().lower()
    return value if EMAIL_RE.match(value) else ""


# --------------------------------------------------------------------------- #
# Load + prepare
# --------------------------------------------------------------------------- #
def load_orders(path: Path) -> pd.DataFrame:
    if not path.exists():
        fail(f"could not find '{path}'. Put your Shopify export next to this script.")
    print(f"Reading orders from '{path.name}'...")
    try:
        df = pd.read_csv(path, dtype=str, low_memory=False)
    except Exception as exc:  # noqa: BLE001 - show any read problem plainly
        fail(f"could not read the CSV: {exc}")
    df.columns = df.columns.str.strip()
    print(f"Loaded {len(df)} rows.")
    return df


def prepare_lines(df: pd.DataFrame) -> pd.DataFrame:
    """One row per line item, with order-level fields copied onto every line."""
    c_order = find_col(df, "Name") or df.columns[0]
    c_created = find_col(df, "Created at")
    if not c_created:
        fail("the CSV has no 'Created at' column.")

    c_cancel = find_col(df, "Cancelled at")
    c_ship = find_col(df, "Shipping Name")
    c_bill = find_col(df, "Billing Name")
    c_email = find_col(df, "Email")
    c_tags = find_col(df, "Tags")
    c_item = find_col(df, "Lineitem name")
    c_qty = find_col(df, "Lineitem quantity")
    c_total = find_col(df, "Total")

    phone_cols = [c for c in (find_col(df, "Shipping Phone"), find_col(df, "Phone"), find_col(df, "Billing Phone")) if c]
    if not phone_cols:
        fail("the CSV has no phone column (Shipping Phone / Phone).")

    # Shopify only fills the order name on the first row of some exports.
    order_id = clean(df[c_order].ffill())

    phone = pd.Series("", index=df.index)
    for col in reversed(phone_cols):  # first column in the list wins
        values = clean(df[col])
        phone = values.where(values != "", phone)

    empty = pd.Series("", index=df.index)
    lines = pd.DataFrame(
        {
            "order_id": order_id,
            "created": clean(df[c_created]),
            "cancelled": clean(df[c_cancel]) if c_cancel else empty,
            "name": clean(df[c_ship]) if c_ship else empty,
            "bill_name": clean(df[c_bill]) if c_bill else empty,
            "phone": phone,
            "email": clean(df[c_email]) if c_email else empty,
            "tags": clean(df[c_tags]) if c_tags else empty,
            "product": clean(df[c_item]) if c_item else empty,
            "qty": pd.to_numeric(df[c_qty], errors="coerce").fillna(1) if c_qty else 1,
            "total": pd.to_numeric(df[c_total], errors="coerce") if c_total else float("nan"),
        }
    )

    # Copy order-level fields from the first row of an order onto its other rows.
    order_cols = ["created", "cancelled", "name", "bill_name", "phone", "email", "tags"]
    filled = lines[order_cols].where(lines[order_cols] != "")
    filled = filled.groupby(lines["order_id"]).ffill()
    filled = filled.groupby(lines["order_id"]).bfill().fillna("")
    lines[order_cols] = filled

    lines["name"] = lines["name"].where(lines["name"] != "", lines["bill_name"])
    lines["email"] = lines["email"].map(valid_email)
    return lines


def filter_lines(lines: pd.DataFrame) -> pd.DataFrame:
    total_orders = lines["order_id"].nunique()

    lines = lines[(lines["order_id"] != "") & (lines["name"] != "") & (lines["phone"] != "")].copy()

    cancelled = lines["cancelled"] != ""
    reattempt = lines["tags"].str.lower().apply(lambda t: any(w in t for w in REATTEMPT_WORDS))
    kept = lines[~cancelled & ~reattempt].copy()

    kept["wa"] = kept["phone"].map(normalize_phone)
    kept["key"] = kept["wa"].str[-10:]
    good_phone = kept["wa"].map(valid_whatsapp)
    bad_phone_orders = kept.loc[~good_phone, "order_id"].nunique()
    kept = kept[good_phone].copy()

    kept["date"] = pd.to_datetime(kept["created"], errors="coerce", utc=True).dt.tz_convert(TIMEZONE)
    bad_date_orders = kept.loc[kept["date"].isna(), "order_id"].nunique()

    print(
        f"Orders in file: {total_orders} | skipped: {cancelled.groupby(lines['order_id']).any().sum()} cancelled, "
        f"{reattempt.groupby(lines['order_id']).any().sum()} re-attempt | used: {kept['order_id'].nunique()}"
    )
    if bad_phone_orders:
        print(f"WARNING: skipped {bad_phone_orders} order(s) with an invalid phone number (cannot be used for WhatsApp).")
    if bad_date_orders:
        print(f"WARNING: {bad_date_orders} order(s) have an unreadable 'Created at' date; their timing is guessed.")
    return kept


# --------------------------------------------------------------------------- #
# Customer summary
# --------------------------------------------------------------------------- #
def build_report(lines: pd.DataFrame, single_order_days: int) -> pd.DataFrame:
    now = pd.Timestamp.now(tz=TIMEZONE)
    today = now.date()
    tomorrow = today + timedelta(days=1)

    rows = []
    for _, group in lines.groupby("key", sort=False):
        group = group.sort_values("date", kind="stable", na_position="last")
        orders = group.drop_duplicates("order_id")  # first line of each order
        dated = orders.dropna(subset=["date"])
        order_count = len(orders)

        last = dated.iloc[-1] if not dated.empty else orders.iloc[-1]
        last_dt = last["date"] if not dated.empty else now
        last_order_date = last_dt.date()

        # Time of day customers usually order -> best time to message them.
        if not dated.empty:
            minutes = int((dated["date"].dt.hour * 60 + dated["date"].dt.minute).median())
        else:
            minutes = last_dt.hour * 60 + last_dt.minute
        reach_time = format_time(minutes)

        # Average gap between *orders* (not between rows of the same order).
        stamps = dated["date"].tolist()
        if len(stamps) > 1:
            gaps = [(b - a).total_seconds() / 86400 for a, b in zip(stamps, stamps[1:])]
            avg_days = round(sum(gaps) / len(gaps), 1)
            projected = last_order_date + timedelta(days=round(avg_days))
        else:
            avg_days = None
            projected = last_order_date + timedelta(days=single_order_days)

        next_contact = projected if projected > today else tomorrow

        # Favourite product = most units bought across all lines.
        items = group[group["product"] != ""]
        favourite = items.groupby("product")["qty"].sum().idxmax() if not items.empty else "N/A"

        ltv = orders["total"].sum(min_count=1)

        segment = "Priority" if order_count >= 3 else "Re-order" if order_count == 2 else "Regular"
        label = {"Priority": "Priority (3+ Orders)", "Re-order": "Re-order Due", "Regular": "Regular Customer"}[segment]

        # Most recent valid email on file for this customer.
        email = next((e for e in reversed(group["email"].tolist()) if e), "")

        rows.append(
            {
                "Customer Name": last["name"],
                "Phone": last["phone"],
                "WhatsApp Number": last["wa"],
                "Email": email,
                "Last Order ID": last["order_id"] or "N/A",
                "Last Order Date": last_order_date.isoformat(),
                "Segment": segment,
                "Priority": label,
                "Orders": order_count,
                "Next Contact Date": next_contact.isoformat(),
                "Next Contact Time": reach_time,
                "Next Contact Schedule": f"{next_contact:%A, %d %b %Y} @ {reach_time}",
                "Projected Reorder Date": projected.isoformat(),
                "Avg Days Between": avg_days if avg_days is not None else "",
                "Most Favorable Product": favourite,
                "Lifetime Value": round(float(ltv), 2) if pd.notna(ltv) else "",
            }
        )

    report = pd.DataFrame(rows)
    return report.sort_values(["Orders", "Next Contact Date"], ascending=[False, True]).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Excel output
# --------------------------------------------------------------------------- #
def write_excel(report: pd.DataFrame, path: Path) -> None:
    try:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Customers")
            ws = writer.sheets["Customers"]

            header_fill = PatternFill("solid", fgColor="0B6B5B")
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(vertical="center")
            ws.row_dimensions[1].height = 24

            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions

            for idx, col in enumerate(report.columns, start=1):
                longest = max([len(str(col))] + [len(str(v)) for v in report[col].head(500)])
                ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 3, 12), 48)

            # Keep phone numbers as text so Excel never strips leading zeros.
            for col_name in ("Phone", "WhatsApp Number"):
                col_idx = list(report.columns).index(col_name) + 1
                for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for cell in row:
                        cell.number_format = "@"
    except PermissionError:
        fail(f"could not write '{path.name}'. Close it in Excel and run the script again.")


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description="Build the customer outreach schedule from a Shopify export.")
    parser.add_argument("csv", nargs="?", default=str(DEFAULT_CSV), help="orders CSV (default: orders_export.csv)")
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT), help="Excel file to write")
    parser.add_argument(
        "--single-order-days",
        type=int,
        default=30,
        help="days after their only order to follow up with one-time buyers (default 30; 0 = contact tomorrow)",
    )
    args = parser.parse_args()

    lines = filter_lines(prepare_lines(load_orders(Path(args.csv))))
    if lines.empty:
        fail("no usable orders found (every order was cancelled, re-attempt, or missing a name/phone).")

    report = build_report(lines, args.single_order_days)
    write_excel(report, Path(args.out))

    segments = report["Segment"].value_counts()
    with_email = int((report["Email"] != "").sum())
    last_dates = pd.to_datetime(report["Last Order Date"], errors="coerce")
    cutoff = pd.Timestamp.now(tz=TIMEZONE).tz_localize(None).normalize() - pd.Timedelta(days=LAPSED_AFTER_DAYS)
    lapsed = int((last_dates < cutoff).sum())
    print(
        f"Done: {len(report)} customers saved to '{Path(args.out).name}' "
        f"({segments.get('Priority', 0)} priority, {segments.get('Re-order', 0)} re-order, "
        f"{segments.get('Regular', 0)} regular)."
    )
    print(f"Emails found for {with_email} of {len(report)} customers.")
    print(f"Lapsed (no order in {LAPSED_AFTER_DAYS}+ days): {lapsed} of {len(report)} customers.")
    print("Next: run  python app.py  and open http://127.0.0.1:5000")


if __name__ == "__main__":
    main()