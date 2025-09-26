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
    """Deprecated for new layout; retained for compatibility in case of legacy sheets."""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
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
    """Parse category header into a float.

    Accepts numbers and strings containing a leading numeric token, e.g.:
    '0,1', '0.1', '0.1x', '0.1 pts', '  0,50  '. Returns None if not parseable.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        s = value.strip()
        if not s:
            return None
        # Replace comma decimal, then extract the first numeric token
        s = s.replace(",", ".")
        import re
        m = re.search(r"[-+]?\d+(?:\.\d+)?", s)
        if not m:
            return None
        try:
            return float(m.group(0))
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


def read_koseries_events_for_date(script_dir: str, target_date: dt.date) -> List[str]:
    """Load event names from koseries_info.csv for the given date.
    
    Returns a list of event names that occurred on the target date.
    """
    path = os.path.join(script_dir, "koseries_info.csv")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Required file not found: {path}. Place koseries_info.csv next to bounty.py"
        )
    
    events_for_date = []
    with open(path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            if 'date' in row and 'name' in row:
                try:
                    # Parse date in DD/MM/YYYY format
                    event_date = dt.datetime.strptime(row['date'], "%d/%m/%Y").date()
                    if event_date == target_date:
                        events_for_date.append(row['name'])
                except ValueError:
                    # Skip rows with invalid date format
                    continue
    
    return events_for_date


def load_values_from_excel_sheet(excel_path: str, sheet_name: str) -> List[List[object]]:
    """Load all cell values from a given Excel sheet using pandas (no pivot refresh).

    Returns a matrix of Python objects with None for empty cells. Raises a clear
    error if the sheet is missing.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None, engine=None)
    except ValueError as exc:
        # pandas raises ValueError if sheet is not found
        raise FileNotFoundError(
            f"Required sheet '{sheet_name}' not found in workbook: {excel_path}"
        ) from exc
    except ImportError as exc:
        raise RuntimeError(
            "Reading .xlsx requires an engine like 'openpyxl'. Install via 'pip install openpyxl'."
        ) from exc

    # Replace NaN with None and convert to list of lists
    matrix: List[List[object]] = df.where(pd.notna(df), None).values.tolist()
    # Ensure at least 1 row structure
    if not isinstance(matrix, list):
        matrix = [[matrix]]
    return matrix


def compute_rankings_for_date(
    values: List[List[object]],
    target_date: dt.date,
    allowed_user_ids: List[str],
    event_names: List[str],
) -> pd.DataFrame:
    """Compute ranking DataFrame for the given date from a values matrix.

    The matrix follows the new structure:
    - Row 8: Event names as strings
    - Row 9: Row Labels + categories as floats
    - Data starts from row 10
    """
    if not values or len(values) < 10:
        raise ValueError("Input sheet appears too small to contain headers and data.")

    # Normalize row lengths (ragged rows can appear)
    max_cols = max(len(r) for r in values)
    grid: List[List[object]] = [r + [None] * (max_cols - len(r)) for r in values]

    # Find event names row (row 8, index 7) and categories row (row 9, index 8)
    event_names_row = grid[7] if len(grid) > 7 else []
    categories_row = grid[8] if len(grid) > 8 else []
    
    # Find columns that match the event names for this date
    target_cols: List[int] = []
    col_to_category: Dict[int, float] = {}
    col_to_event_name: Dict[int, str] = {}
    
    for col_idx in range(len(event_names_row)):
        if col_idx < len(categories_row):
            event_name = str(event_names_row[col_idx]).strip() if event_names_row[col_idx] else ""
            category = _parse_category_cell(categories_row[col_idx])
            
            # Check if this event name is in our target events for this date
            if event_name in event_names and category is not None:
                target_cols.append(col_idx)
                col_to_category[col_idx] = float(category)
                col_to_event_name[col_idx] = event_name

    if not target_cols:
        # No matching events found, return empty dataframe with 0 points
        df = pd.DataFrame({"user_id": allowed_user_ids, "points": [0] * len(allowed_user_ids)})
        df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
        df.insert(0, "rank", range(1, len(df) + 1))
        return df.reset_index(drop=True)

    # Build user_id -> points and detailed bounty counts
    points_by_user: Dict[str, float] = {uid: 0.0 for uid in allowed_user_ids}
    bounty_counts_by_user: Dict[str, Dict[float, int]] = {uid: {} for uid in allowed_user_ids}
    total_bounties_by_user: Dict[str, int] = {uid: 0 for uid in allowed_user_ids}

    # Iterate player rows starting from row 10 (index 9)
    data_start_row = 9
    for row_idx in range(data_start_row, len(grid)):
        row = grid[row_idx]
        if not row:
            continue
        
        # Get user_id from first column (Row Labels column)
        user_id_cell = row[0] if len(row) > 0 else None
        user_id = _normalize_user_id(user_id_cell)
        if not user_id:
            continue
        if user_id not in points_by_user:
            # Skip players not in the allowed list
            continue

        total_points = 0.0
        total_bounties = 0
        
        for col_idx in target_cols:
            raw_val = row[col_idx] if col_idx < len(row) else None
            # Treat empty/invalid/NaN as 0
            try:
                count = float(raw_val) if raw_val not in (None, "", "nan", "NaN") else 0.0
                if count != count:  # NaN check
                    count = 0.0
            except Exception:
                count = 0.0
            
            category = col_to_category[col_idx]
            total_bounties += int(count)
            
            # Store bounty count by category
            if category not in bounty_counts_by_user[user_id]:
                bounty_counts_by_user[user_id][category] = 0
            bounty_counts_by_user[user_id][category] += int(count)
            
            # Category points per spec
            per_elim_points = POINTS_BY_CATEGORY.get(round(category, 4), 0)
            total_points += count * per_elim_points
        
        points_by_user[user_id] = total_points
        total_bounties_by_user[user_id] = total_bounties

    # Create detailed DataFrame with bounty breakdown
    data_rows = []
    for user_id in allowed_user_ids:
        if points_by_user[user_id] > 0:  # Only include users with points
            row_data = {
                "user_id": user_id,
                "points": points_by_user[user_id],
                "primes gagnées": total_bounties_by_user[user_id],
                "primes gagnées 1€": bounty_counts_by_user[user_id].get(0.1, 0),
                "primes gagnées 3€": bounty_counts_by_user[user_id].get(0.3, 0),
                "primes gagnées 5€": bounty_counts_by_user[user_id].get(0.5, 0),
                "primes gagnées 10€": bounty_counts_by_user[user_id].get(1.0, 0),
                "primes gagnées 20€": bounty_counts_by_user[user_id].get(2.0, 0),
                "primes gagnées 50€": bounty_counts_by_user[user_id].get(5.0, 0),
                "primes gagnées 100€": bounty_counts_by_user[user_id].get(10.0, 0),
            }
            data_rows.append(row_data)

    if not data_rows:
        # No users with points, return empty dataframe
        return pd.DataFrame(columns=["rank", "user_id", "points", "primes gagnées", 
                                   "primes gagnées 1€", "primes gagnées 3€", "primes gagnées 5€",
                                   "primes gagnées 10€", "primes gagnées 20€", "primes gagnées 50€", "primes gagnées 100€"])

    df = pd.DataFrame(data_rows)
    df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
    
    # Assign competition ranks: ties share rank, next rank skips by tie size (1,1,3,...)
    ranks: List[int] = []
    last_points: Optional[float] = None
    current_rank = 0
    position = 0
    for pts in df["points"].tolist():
        position += 1
        if last_points is None or pts != last_points:
            current_rank = position
            last_points = pts
        ranks.append(current_rank)
    df.insert(0, "rank", ranks)
    return df.reset_index(drop=True)


def compute_unfiltered_rankings_for_date(
    values: List[List[object]], target_date: dt.date, event_names: List[str]
) -> pd.DataFrame:
    """Compute ranking DataFrame for the given date without filtering by player_list.csv."""
    if not values or len(values) < 10:
        raise ValueError("Input sheet appears too small to contain headers and data.")

    max_cols = max(len(r) for r in values)
    grid: List[List[object]] = [r + [None] * (max_cols - len(r)) for r in values]

    # Find event names row (row 8, index 7) and categories row (row 9, index 8)
    event_names_row = grid[7] if len(grid) > 7 else []
    categories_row = grid[8] if len(grid) > 8 else []
    
    # Find columns that match the event names for this date
    target_cols: List[int] = []
    col_to_category: Dict[int, float] = {}
    
    for col_idx in range(len(event_names_row)):
        if col_idx < len(categories_row):
            event_name = str(event_names_row[col_idx]).strip() if event_names_row[col_idx] else ""
            category = _parse_category_cell(categories_row[col_idx])
            
            # Check if this event name is in our target events for this date
            if event_name in event_names and category is not None:
                target_cols.append(col_idx)
                col_to_category[col_idx] = float(category)

    if not target_cols:
        # No matching events found, return empty dataframe
        return pd.DataFrame(columns=["rank", "user_id", "points"])

    points_by_user: Dict[str, float] = {}
    
    # Iterate player rows starting from row 10 (index 9)
    data_start_row = 9
    for row_idx in range(data_start_row, len(grid)):
        row = grid[row_idx]
        if not row:
            continue
        
        # Get user_id from first column (Row Labels column)
        user_id_cell = row[0] if len(row) > 0 else None
        user_id = _normalize_user_id(user_id_cell)
        if not user_id:
            continue
        
        total_points = 0.0
        for col_idx in target_cols:
            raw_val = row[col_idx] if col_idx < len(row) else None
            try:
                count = float(raw_val) if raw_val not in (None, "", "nan", "NaN") else 0.0
                if count != count:
                    count = 0.0
            except Exception:
                count = 0.0
            
            category = col_to_category[col_idx]
            per_elim_points = POINTS_BY_CATEGORY.get(round(category, 4), 0)
            total_points += count * per_elim_points
        
        points_by_user[user_id] = total_points

    df = pd.DataFrame(
        {"user_id": list(points_by_user.keys()), "points": list(points_by_user.values())}
    )
    df.sort_values(["points", "user_id"], ascending=[False, True], inplace=True)
    
    # Competition ranking for ties
    ranks: List[int] = []
    last_points: Optional[float] = None
    current_rank = 0
    position = 0
    for pts in df["points"].tolist():
        position += 1
        if last_points is None or pts != last_points:
            current_rank = position
            last_points = pts
        ranks.append(current_rank)
    df.insert(0, "rank", ranks)
    return df.reset_index(drop=True)


def write_csv(df: pd.DataFrame, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    df.to_csv(output_path, index=False)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compute and rank player points per day and write a CSV. "
            "Reads a per-date sheet (DD_MM_YYYY) from pko.xlsx with categories as columns."
        )
    )
    parser.add_argument(
        "--excel",
        required=False,
        default=None,
        help=(
            "Path to the Excel workbook. If omitted, the script will look for "
            "'pko.xlsx' (or the common typo 'pko.xslx') next to bounty.py."
        ),
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
    # Resolve Excel path: use --excel if provided, otherwise default to pko.xlsx (or pko.xslx) next to script
    if args.excel:
        excel_path = os.path.abspath(args.excel)
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        candidate1 = os.path.join(script_dir, "pko.xlsx")
        candidate2 = os.path.join(script_dir, "pko.xslx")  # tolerate common typo
        if os.path.exists(candidate1):
            excel_path = candidate1
        elif os.path.exists(candidate2):
            excel_path = candidate2
        else:
            raise FileNotFoundError(
                "Excel file not provided and neither 'pko.xlsx' nor 'pko.xslx' found next to bounty.py"
            )
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    # Parse date
    try:
        target_date = dt.datetime.strptime(args.date, "%d/%m/%Y").date()
    except ValueError as exc:
        raise SystemExit("--date must be in DD/MM/YYYY format") from exc

    # Load player list and event names for the date
    script_dir = os.path.dirname(os.path.abspath(__file__))
    allowed_user_ids = read_player_list(script_dir)
    if not allowed_user_ids:
        raise SystemExit("users_list.csv is empty or missing valid customer ids")
    
    # Get event names for this date from koseries_info.csv
    event_names = read_koseries_events_for_date(script_dir, target_date)
    if not event_names:
        print(f"No events found for date {target_date.strftime('%d/%m/%Y')}")
        return pd.DataFrame()

    # Read matrix from the main sheet (no longer per-date sheets)
    values = load_values_from_excel_sheet(excel_path, "Sheet1")  # or whatever the main sheet is called

    # Compute filtered rankings (users_list.csv)
    df = compute_rankings_for_date(values, target_date, allowed_user_ids, event_names)
    # Also compute general (unfiltered) rankings
    general_df = compute_unfiltered_rankings_for_date(values, target_date, event_names)

    # Output with new naming convention
    out_path = args.out
    if out_path is None:
        out_name = f"classement_koseries_{target_date.strftime('%d-%m-%Y')}.csv"
        out_path = os.path.join(os.getcwd(), out_name)
    write_csv(df, out_path)
    # Write general rankings alongside, using a parallel filename
    general_out_path = out_path.replace("classement_koseries_", "rankings_general_")
    write_csv(general_df, general_out_path)

    # Display summary
    print("Filtered (users_list) rankings:\n" + df.to_string(index=False))
    print(f"\nWrote CSV: {out_path}")
    print("\nGeneral rankings (unfiltered):\n" + general_df.to_string(index=False))
    print(f"\nWrote CSV: {general_out_path}")

    return df


if __name__ == "__main__":
    main()
