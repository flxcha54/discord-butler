#!/usr/bin/env python3
import argparse
import csv
import datetime as dt
import os
import sys
import time
from typing import Dict, List, Optional, Tuple

import pandas as pd


POINTS_BY_CATEGORY: Dict[float, int] = {
    0.1: 0,
    0.3: 0,
    0.5: 1,
    1.0: 2,  # explicitly ignored per spec
    2.0: 2,
    5.0: 3,
    10.0: 10,
}


def _normalize_user_id(value: object) -> Optional[str]:
    """Normalize user identifiers to a canonical string for matching.

    - If value is int -> '123'
    - If value is float and very close to an integer -> '123'
    - If value is float non-integer -> string without scientific notation and without trailing .0
    - If value is string like '123.0' or '1.23E+03' -> convert similarly
    - Returns None if value is empty/None
    """
    if value is None:
        return None
    if isinstance(value, bool):
        # Avoid treating True/False as 1/0 IDs
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        # Non-integer float: represent compactly
        s = ("%f" % value).rstrip("0").rstrip(".")
        return s
    # Strings
    s = str(value).strip()
    if not s:
        return None
    # Try to interpret scientific or float-like strings
    try:
        f = float(s)
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        s2 = ("%f" % f).rstrip("0").rstrip(".")
        return s2
    except Exception:
        return s


def _parse_date_cell(value: object) -> Optional[dt.date]:
    """Try to parse a value from row 7 into a date.

    Accepts datetime.date/datetime, or strings in DD/MM/YYYY.
    Returns the date or None if not parseable.
    """
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    # Excel serial date number support (xlwings can sometimes return floats)
    if isinstance(value, (int, float)):
        try:
            serial = float(value)
            # Excel for Windows date origin is 1899-12-30 (accounting for 1900 leap year bug)
            base = dt.date(1899, 12, 30)
            delta = dt.timedelta(days=serial)
            candidate = base + delta
            # Heuristic: only accept reasonable calendar dates
            if dt.date(1900, 1, 1) <= candidate <= dt.date(2100, 12, 31):
                return candidate
        except Exception:
            pass
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
    return None


def _parse_category_cell(value: object) -> Optional[float]:
    """Parse category title from row 8 into float with 4 decimals accepted.
    Returns None if not parseable.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        s = s.replace(",", ".")  # tolerate commas
        try:
            return float(s)
        except ValueError:
            return None
    return None


def read_player_list(script_dir: str) -> List[str]:
    """Load allowed customer_ids from users_list.csv in the script directory.

    Accepts files with header; uses a column named 'customer_id' if present,
    otherwise picks the first column. Returns ids as strings.
    """
    path = os.path.join(script_dir, "users_list.csv")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Required file not found: {path}. Place users_list.csv next to bounty.py"
        )

    with open(path, "r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return []
    header = [h.strip().lower() for h in rows[0]]
    data_rows = rows[1:] if any(header) else rows

    # Prefer explicit 'customer_id' column, otherwise fallback to first column
    col_idx = 0
    if header and "customer_id" in header:
        col_idx = header.index("customer_id")
    else:
        for i, name in enumerate(header):
            if any(key in name for key in ("customer", "user", "id")):
                col_idx = i
                break
    user_ids: List[str] = []
    for row in data_rows:
        if not row:
            continue
        val = row[col_idx] if col_idx < len(row) else ""
        norm = _normalize_user_id(val)
        if norm:
            user_ids.append(norm)
    return user_ids


def _refresh_sheet_pivots_with_wait(book, sh, timeout_sec: int = 120) -> None:
    """Refresh pivot tables on the given sheet and wait until Excel finishes.

    Single attempt with timeout_sec waiting for calculation/async queries to finish.
    Raises RuntimeError if not completed. Requires COM availability (Windows with Excel).
    """
    app_api = getattr(book.app, "api", None)
    sh_api = getattr(sh, "api", None)
    if app_api is None or sh_api is None:
        raise RuntimeError(
            "Cannot ensure refresh: Excel COM API is unavailable. Run on a machine with Excel."
        )

    # Issue refresh for pivot tables on the sheet
    try:
        pivots = sh_api.PivotTables()
        count = int(pivots.Count)
        for i in range(1, count + 1):
            pt = pivots.Item(i)
            try:
                pt.RefreshTable()
            except Exception:
                try:
                    pt.PivotCache().Refresh()
                except Exception:
                    pass
    except Exception:
        # If sheet has no pivots, try workbook-wide refresh
        try:
            book.api.RefreshAll()
        except Exception:
            pass

    # Wait until Excel finishes calculations and async queries
    start = time.time()
    while time.time() - start < timeout_sec:
        try:
            # Try to complete async queries if supported
            try:
                app_api.CalculateUntilAsyncQueriesDone()
            except Exception:
                pass
            state = getattr(app_api, "CalculationState", None)
            # 0 == xlDone
            if state == 0:
                return
        except Exception:
            # Keep waiting if transient
            pass
        time.sleep(1)

    # If we reach here, attempt timed out
    raise RuntimeError("Pivot refresh did not complete within the allotted time.")


def _kill_excel_processes():
    """Force kill any remaining Excel processes to ensure clean restart."""
    import subprocess
    import platform
    
    system = platform.system().lower()
    try:
        if system == "windows":
            subprocess.run(["taskkill", "/f", "/im", "excel.exe"], 
                         capture_output=True, timeout=10)
        elif system == "darwin":  # macOS
            subprocess.run(["pkill", "-f", "Microsoft Excel"], 
                         capture_output=True, timeout=10)
        else:  # Linux
            subprocess.run(["pkill", "-f", "excel"], 
                         capture_output=True, timeout=10)
    except Exception:
        pass  # Ignore errors in process cleanup


def load_values_with_xlwings(excel_path: str, sheet: Optional[object]) -> List[List[object]]:
    """Load used range via xlwings after ensuring pivot refresh.

    Simulates manual cancel-and-retry: if refresh stalls/times out, fully kills
    Excel processes and reopens, retrying up to 5 times. Only reads after a successful refresh.
    """
    try:
        import xlwings as xw  # type: ignore
    except Exception as exc:  # pragma: no cover
        raise RuntimeError(
            "xlwings is required for this script. Please install it via 'pip install xlwings'."
        ) from exc

    last_error: Optional[Exception] = None
    for attempt in range(1, 6):
        app = None
        book = None
        try:
            print(f"Attempt {attempt}/5: Opening Excel and refreshing pivots...")
            app = xw.App(visible=False, add_book=False)
            book = app.books.open(excel_path, read_only=False, update_links=False)
            if sheet is None:
                sh = book.sheets[0]
            else:
                sh = book.sheets[sheet]

            # Ensure pivot tables on this sheet are refreshed before reading
            _refresh_sheet_pivots_with_wait(book, sh, timeout_sec=120)
            print(f"Attempt {attempt}/5: Refresh completed successfully!")

            used = sh.used_range
            values = used.value
            if not isinstance(values, list):
                values = [[values]]
            matrix: List[List[object]] = []
            for row in values:
                if isinstance(row, (list, tuple)):
                    matrix.append(list(row))
                else:
                    matrix.append([row])
            return matrix
        except Exception as exc:
            last_error = exc
            print(f"Attempt {attempt}/5 failed: {exc}")
        finally:
            # Always try to close cleanly first
            if book is not None:
                try:
                    book.close(save=False)
                except Exception:
                    pass
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass
            
            # Force kill any remaining Excel processes
            if attempt < 5:  # Don't kill on final attempt
                print(f"Killing Excel processes before retry...")
                _kill_excel_processes()
                time.sleep(3)  # Give time for processes to fully terminate

    if last_error is not None:
        raise RuntimeError(f"Failed to refresh and read after retries: {last_error}") from last_error
    raise RuntimeError("Failed to refresh and read after retries due to unknown error.")


def compute_rankings_for_date(
    values: List[List[object]],
    target_date: dt.date,
    allowed_user_ids: List[str],
) -> pd.DataFrame:
    """Compute ranking DataFrame for the given date from a values matrix.

    The matrix follows the specified structure:
    - Rows 1-6 ignored
    - Row 7: date headers (first column is label). Forward-fill date across category columns
    - Row 8: category headers as floats
    - Column 1: 'Row Labels' with user_ids, starting at row 9
    - Remaining cells: elimination counts (numbers)
    """
    if not values or len(values) < 8:
        raise ValueError("Input sheet appears too small to contain the required header rows.")

    # Normalize row lengths (ragged rows can appear)
    max_cols = max(len(r) for r in values)
    grid: List[List[object]] = [r + [None] * (max_cols - len(r)) for r in values]

    # Dynamically detect header layout: find the row containing 'Row Labels'
    def _norm_str(v: object) -> str:
        return str(v).strip().lower() if v is not None else ""

    category_row_idx = None
    row_labels_col_idx = None
    for r_idx, row in enumerate(grid):
        for c_idx, cell in enumerate(row):
            if _norm_str(cell) == "row labels":
                category_row_idx = r_idx
                row_labels_col_idx = c_idx
                break
        if category_row_idx is not None:
            break
    if category_row_idx is None:
        raise ValueError("Could not locate 'Row Labels' header in the sheet.")

    if category_row_idx - 1 < 0:
        raise ValueError("Date header row not found above the categories row.")
    date_row_idx = category_row_idx - 1

    # Forward-fill dates across columns on row 7
    filled_dates: List[Optional[dt.date]] = [None] * max_cols
    last_date: Optional[dt.date] = None
    start_col = (row_labels_col_idx or 0) + 1
    for col in range(start_col, max_cols):
        raw_date = grid[date_row_idx][col] if date_row_idx < len(grid) else None
        parsed = _parse_date_cell(raw_date)
        if parsed is not None:
            last_date = parsed
        filled_dates[col] = last_date

    # Parse categories from row 8 per column
    categories: List[Optional[float]] = [None] * max_cols
    for col in range(start_col, max_cols):
        categories[col] = _parse_category_cell(grid[category_row_idx][col])

    # Determine which columns belong to the target date
    target_cols: List[int] = [
        c for c in range(start_col, max_cols) if filled_dates[c] == target_date and categories[c] is not None
    ]
    if not target_cols:
        # No data for this date: return allowed players with zero points
        df = pd.DataFrame({"user_id": allowed_user_ids, "points": [0] * len(allowed_user_ids)})
        df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
        df.insert(0, "rank", range(1, len(df) + 1))
        return df.reset_index(drop=True)

    # Map column to its category
    col_to_category: Dict[int, float] = {c: float(categories[c]) for c in target_cols if categories[c] is not None}

    # Build user_id -> points
    points_by_user: Dict[str, float] = {uid: 0.0 for uid in allowed_user_ids}

    # Iterate player rows starting at row index 8 (row 9 in 1-based)
    data_start_row = category_row_idx + 1
    for row_idx in range(data_start_row, len(grid)):
        row = grid[row_idx]
        if not row:
            continue
        user_id_cell = row[row_labels_col_idx or 0] if len(row) > (row_labels_col_idx or 0) else None
        user_id = _normalize_user_id(user_id_cell)
        if not user_id:
            continue
        if user_id not in points_by_user:
            # Skip players not in the allowed list
            continue

        total_points = 0.0
        for col_idx, category in col_to_category.items():
            raw_val = row[col_idx] if col_idx < len(row) else None
            # Treat empty as 0
            try:
                count = float(raw_val) if raw_val not in (None, "") else 0.0
            except Exception:
                count = 0.0
            # Category points per spec
            per_elim_points = POINTS_BY_CATEGORY.get(round(category, 4), 0)
            total_points += count * per_elim_points
        points_by_user[user_id] = total_points

    df = pd.DataFrame(
        {"user_id": list(points_by_user.keys()), "points": list(points_by_user.values())}
    )
    df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
    df.insert(0, "rank", range(1, len(df) + 1))
    return df.reset_index(drop=True)


def compute_unfiltered_rankings_for_date(
    values: List[List[object]], target_date: dt.date
) -> pd.DataFrame:
    """Compute ranking DataFrame for the given date without filtering by player_list.csv."""
    if not values or len(values) < 8:
        raise ValueError("Input sheet appears too small to contain the required header rows.")

    max_cols = max(len(r) for r in values)
    grid: List[List[object]] = [r + [None] * (max_cols - len(r)) for r in values]

    # Detect 'Row Labels' position
    def _norm_str(v: object) -> str:
        return str(v).strip().lower() if v is not None else ""

    category_row_idx = None
    row_labels_col_idx = None
    for r_idx, row in enumerate(grid):
        for c_idx, cell in enumerate(row):
            if _norm_str(cell) == "row labels":
                category_row_idx = r_idx
                row_labels_col_idx = c_idx
                break
        if category_row_idx is not None:
            break
    if category_row_idx is None:
        raise ValueError("Could not locate 'Row Labels' header in the sheet.")
    if category_row_idx - 1 < 0:
        raise ValueError("Date header row not found above the categories row.")
    date_row_idx = category_row_idx - 1

    filled_dates: List[Optional[dt.date]] = [None] * max_cols
    last_date: Optional[dt.date] = None
    start_col = (row_labels_col_idx or 0) + 1
    for col in range(start_col, max_cols):
        raw_date = grid[date_row_idx][col] if date_row_idx < len(grid) else None
        parsed = _parse_date_cell(raw_date)
        if parsed is not None:
            last_date = parsed
        filled_dates[col] = last_date

    categories: List[Optional[float]] = [None] * max_cols
    for col in range(start_col, max_cols):
        categories[col] = _parse_category_cell(grid[category_row_idx][col])

    target_cols: List[int] = [
        c for c in range(start_col, max_cols) if filled_dates[c] == target_date and categories[c] is not None
    ]
    if not target_cols:
        # No columns for this date: return all users with 0 points
        user_ids: List[str] = []
        data_start_row = category_row_idx + 1
        for row_idx in range(data_start_row, len(grid)):
            row = grid[row_idx]
            if not row:
                continue
            uid = _normalize_user_id(row[row_labels_col_idx or 0] if len(row) > (row_labels_col_idx or 0) else None)
            if uid:
                user_ids.append(uid)
        user_ids = sorted(set(user_ids))
        df0 = pd.DataFrame({"user_id": user_ids, "points": [0] * len(user_ids)})
        df0.insert(0, "rank", range(1, len(df0) + 1))
        return df0

    col_to_category: Dict[int, float] = {c: float(categories[c]) for c in target_cols if categories[c] is not None}

    points_by_user: Dict[str, float] = {}
    data_start_row = category_row_idx + 1
    for row_idx in range(data_start_row, len(grid)):
        row = grid[row_idx]
        if not row:
            continue
        user_id = _normalize_user_id(row[row_labels_col_idx or 0] if len(row) > (row_labels_col_idx or 0) else None)
        if not user_id:
            continue
        total_points = 0.0
        for col_idx, category in col_to_category.items():
            raw_val = row[col_idx] if col_idx < len(row) else None
            try:
                count = float(raw_val) if raw_val not in (None, "") else 0.0
            except Exception:
                count = 0.0
            per_elim_points = POINTS_BY_CATEGORY.get(round(category, 4), 0)
            total_points += count * per_elim_points
        points_by_user[user_id] = total_points

    df = pd.DataFrame(
        {"user_id": list(points_by_user.keys()), "points": list(points_by_user.values())}
    )
    df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
    df.insert(0, "rank", range(1, len(df) + 1))
    return df.reset_index(drop=True)


def write_csv(df: pd.DataFrame, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    df.to_csv(output_path, index=False)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compute and rank player points per day and write a CSV. "
            "Uses xlwings to read the workbook per the specified sheet structure."
        )
    )
    parser.add_argument(
        "--excel",
        required=False,
        default=None,
        help=(
            "Path to the Excel workbook. If omitted, the script will look for "
            "'pkos.xlsx' (or the common typo 'pkos.xslx') next to bounty.py."
        ),
    )
    parser.add_argument(
        "--sheet",
        default=None,
        help="Worksheet name or index (0-based). Default: first sheet.",
    )
    parser.add_argument(
        "--date",
        required=True,
        help="Target date in DD/MM/YYYY format.",
    )
    parser.add_argument(
        "--out",
        default=None,
        help="Output CSV path. Default: rankings_<DD-MM-YYYY>.csv in current directory.",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> pd.DataFrame:
    args = parse_args(argv)
    # Resolve Excel path: use --excel if provided, otherwise default to pkos.xlsx (or pkos.xslx) next to script
    if args.excel:
        excel_path = os.path.abspath(args.excel)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidate1 = os.path.join(script_dir, "pkos.xlsx")
        candidate2 = os.path.join(script_dir, "pkos.xslx")  # tolerate common typo
        if os.path.exists(candidate1):
            excel_path = candidate1
        elif os.path.exists(candidate2):
            excel_path = candidate2
        else:
            raise FileNotFoundError(
                "Excel file not provided and neither 'pkos.xlsx' nor 'pkos.xslx' found next to bounty.py"
            )
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Parse date
    try:
        target_date = dt.datetime.strptime(args.date, "%d/%m/%Y").date()
    except ValueError as exc:
        raise SystemExit("--date must be in DD/MM/YYYY format") from exc

    # Sheet selection
    sheet_arg: Optional[str]
    if args.sheet is None:
        # Default to 'bounties' sheet if present, otherwise first sheet
        sheet_arg = "bounties"
    else:
        # allow numeric index
        s = str(args.sheet).strip()
        if s.isdigit():
            sheet_arg = str(int(s))  # xlwings accepts index via sheets[int]
        else:
            sheet_arg = s

    # Load player list
    script_dir = os.path.dirname(os.path.abspath(__file__))
    allowed_user_ids = read_player_list(script_dir)
    if not allowed_user_ids:
        raise SystemExit("users_list.csv is empty or missing valid customer ids")

    # Read matrix via xlwings
    # Determine sheet reference (name or index)
    sheet_ref: Optional[object]
    if sheet_arg is None:
        sheet_ref = None
    elif sheet_arg.isdigit():
        sheet_ref = int(sheet_arg)
    else:
        sheet_ref = sheet_arg

    values = load_values_with_xlwings(excel_path, sheet=sheet_ref)

    # Compute
    df = compute_rankings_for_date(values, target_date, allowed_user_ids)
    # Also compute general (unfiltered) rankings
    general_df = compute_unfiltered_rankings_for_date(values, target_date)

    # Output
    out_path = args.out
    if out_path is None:
        out_name = f"rankings_{target_date.strftime('%d-%m-%Y')}.csv"
        out_path = os.path.join(os.getcwd(), out_name)
    write_csv(df, out_path)
    # Write general rankings alongside, using a parallel filename
    general_out_path = out_path.replace("rankings_", "rankings_general_")
    write_csv(general_df, general_out_path)

    # Display summary
    print("Filtered (users_list) rankings:\n" + df.to_string(index=False))
    print(f"\nWrote CSV: {out_path}")
    print("\nGeneral rankings (unfiltered):\n" + general_df.to_string(index=False))
    print(f"\nWrote CSV: {general_out_path}")

    return df


if __name__ == "__main__":
    main()
