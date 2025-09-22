#!/usr/bin/env python3
import argparse
import asyncio
import csv
import datetime as dt
import os
from typing import Dict, List, Optional, Tuple

import pandas as pd


# --------- Hardcoded configuration ---------
# Set your guild and channel IDs here
GUILD_ID: int = 123456789012345678  # TODO: replace with your guild ID
CHANNEL_ID: int = 123456789012345678  # TODO: replace with your channel ID

# File name for the banner image to display on top of the leaderboard (placed next to this script)
BANNER_FILENAME: str = "banner.png"  # TODO: replace with your banner filename

# Discord bot token (replace with your actual bot token)
DISCORD_BOT_TOKEN: str = "YOUR_BOT_TOKEN_HERE"


def _script_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _read_users_mapping(users_csv_path: str) -> Tuple[List[str], Dict[str, str]]:
    """Read users_list.csv and return:
    - list of customer_ids (non-empty, not 'staff') as strings
    - mapping from customer_id -> discord_id (string)
    """
    customer_ids: List[str] = []
    customer_to_discord: Dict[str, str] = {}

    with open(users_csv_path, "r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        # Expect columns: discord_id, customer_id, invite_url (extra columns tolerated)
        for row in reader:
            customer_id_raw = (row.get("customer_id") or "").strip()
            discord_id_raw = (row.get("discord_id") or "").strip()
            if not customer_id_raw or customer_id_raw.lower() == "staff":
                continue
            # Normalize: keep as-is string; bounty.py has its own normalization
            customer_ids.append(customer_id_raw)
            if discord_id_raw:
                customer_to_discord[customer_id_raw] = discord_id_raw

    # Deduplicate preserving order
    seen: set = set()
    deduped: List[str] = []
    for cid in customer_ids:
        if cid in seen:
            continue
        seen.add(cid)
        deduped.append(cid)
    return deduped, customer_to_discord


# player_list.csv generation is no longer needed because bounty.py now reads users_list.csv directly.


def _run_bounty_and_get_df(target_date: dt.date) -> pd.DataFrame:
    """Invoke bounty.main programmatically to compute filtered rankings for target_date.
    Expects bounty.py in the same directory and pkos.xlsx present alongside.
    """
    import importlib.util

    bounty_path = os.path.join(_script_dir(), "bounty.py")
    spec = importlib.util.spec_from_file_location("bounty", bounty_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Unable to import bounty.py module")
    bounty = importlib.util.module_from_spec(spec)  # type: ignore
    spec.loader.exec_module(bounty)  # type: ignore

    date_str = target_date.strftime("%d/%m/%Y")
    # Call bounty.main; it returns the filtered dataframe
    df: pd.DataFrame = bounty.main(["--date", date_str])  # type: ignore
    return df


def _format_leaderboard(df: pd.DataFrame, date_str: str, id_map: Dict[str, str]) -> Tuple[str, List[Tuple[int, str, float]]]:
    """Build French leaderboard text and return rows used.
    Rows are tuples (rank, discord_id_str, points).
    Shows only top 10 players.
    """
    # Prepare rows with discord mention (only players with points > 0, limit to top 10)
    rows: List[Tuple[int, str, float]] = []
    for _, r in df.iterrows():
        user_id = str(r.get("user_id"))
        rank = int(r.get("rank"))
        # robust numeric handling for points
        try:
            points_val = float(r.get("points"))
        except Exception:
            points_val = float('nan')
        if points_val != points_val:  # NaN
            continue
        points = points_val
        
        # Skip players with 0 or negative points
        if points <= 0:
            continue
            
        # Stop at 10 players
        if len(rows) >= 10:
            break
            
        discord_id = id_map.get(user_id) or id_map.get(str(int(float(user_id)))) if user_id.replace(".", "", 1).isdigit() else id_map.get(user_id)
        # Fallback to user_id if missing mapping
        display = f"<@{discord_id}>" if discord_id else f"ID:{user_id}"
        rows.append((rank, display, points))

    # Build French format
    lines = [f"```{date_str}```", ""]
    
    for rank, user, pts in rows:
        if rank == 1:
            lines.append(f"1er - {user} avec {pts:.0f} point(s).")
        elif rank == 2:
            lines.append(f"2ème - {user} avec {pts:.0f} point(s).")
        else:
            lines.append(f"{rank}ème - {user} avec {pts:.0f} point(s).")
    
    # Add congratulations messages
    if len(rows) > 0:
        lines.append("")  # Empty line before congratulations
        
        # First congratulations: all players ranked 1st to 5th (including ties at 5th)
        first_5_users = []
        for rank, user, _ in rows:
            if rank <= 5:
                first_5_users.append(user)
        
        if first_5_users:
            first_5_users_str = ", ".join(first_5_users)
            lines.append(f"Félicitations {first_5_users_str}, vous gagnez tous un ticket pour le Main Event des KO Series 50 000€ garanti de dimanche 5 octobre.")
        
        # Second congratulations: all players ranked 6th to 10th (including ties)
        sixth_to_tenth_users = []
        for rank, user, _ in rows:
            if 6 <= rank <= 10:
                sixth_to_tenth_users.append(user)
        
        if sixth_to_tenth_users:
            sixth_to_tenth_users_str = ", ".join(sixth_to_tenth_users)
            lines.append(f"Bien joué {sixth_to_tenth_users_str}, vous gagnez tous un ticket 20€ à utiliser dans les tournois des KO Series !")
    
    text = "\n".join(lines)
    return text, rows


async def _post_to_discord(token: str, guild_id: int, channel_id: int, content: str, banner_path: Optional[str]) -> None:
    import aiohttp

    url = f'https://discord.com/api/v10/channels/{channel_id}/messages'
    headers = {'Authorization': f'Bot {token}'}
    
    async with aiohttp.ClientSession() as session:
        try:
            if banner_path and os.path.exists(banner_path):
                # Send banner image first (above)
                form_data = aiohttp.FormData()
                
                # Read file content and add to form data
                with open(banner_path, 'rb') as f:
                    file_content = f.read()
                form_data.add_field('file', file_content, filename=os.path.basename(banner_path))
                
                # Send image first
                async with session.post(url, headers=headers, data=form_data) as response:
                    if response.status == 200:
                        print("Banner image posted successfully!")
                    else:
                        print(f"Error posting banner: {response.status} - {await response.text()}")
                
                # Then send the leaderboard text
                data = {'content': content}
                headers_json = {'Authorization': f'Bot {token}', 'Content-Type': 'application/json'}
                async with session.post(url, headers=headers_json, json=data) as response:
                    if response.status == 200:
                        print("Leaderboard text posted successfully!")
                    else:
                        print(f"Error posting leaderboard: {response.status} - {await response.text()}")
            else:
                # Send text only - use JSON
                headers_json = {'Authorization': f'Bot {token}', 'Content-Type': 'application/json'}
                data = {'content': content}
                async with session.post(url, headers=headers_json, json=data) as response:
                    if response.status == 200:
                        print("Message posted successfully!")
                    else:
                        print(f"Error posting message: {response.status} - {await response.text()}")
        except Exception as e:
            print(f"Error posting to Discord: {e}")


def main() -> None:
    parser = argparse.ArgumentParser(description="KO Leaderboard poster using bounty.py and users_list.csv")
    parser.add_argument("--date", required=True, help="Target date in DD/MM/YYYY")
    args = parser.parse_args()

    # Parse date
    try:
        target_date = dt.datetime.strptime(args.date, "%d/%m/%Y").date()
    except ValueError:
        raise SystemExit("--date must be in DD/MM/YYYY format")

    # Resolve paths
    base_dir = _script_dir()
    users_csv = os.path.join(base_dir, "users_list.csv")
    pkos_xlsx = os.path.join(base_dir, "pko.xlsx")
    if not os.path.exists(users_csv):
        raise SystemExit(f"Missing users_list.csv at {users_csv}")
    if not os.path.exists(pkos_xlsx):
        raise SystemExit(f"Missing pko.xlsx at {pkos_xlsx}")

    # Read users to build mapping customer_id -> discord_id for display
    _user_ids, id_map = _read_users_mapping(users_csv)
    if not id_map:
        # Proceed but warn: no mapping means we will show raw IDs
        pass

    # Run bounty
    df = _run_bounty_and_get_df(target_date)

    # Build message
    date_str = target_date.strftime("%d/%m/%Y")
    message, _rows = _format_leaderboard(df, date_str, id_map)

    # Discord
    token = DISCORD_BOT_TOKEN.strip()
    if not token or token == "YOUR_BOT_TOKEN_HERE":
        raise SystemExit(
            "Please replace 'YOUR_BOT_TOKEN_HERE' with your actual Discord bot token in ko_lb.py"
        )

    banner_path = os.path.join(base_dir, BANNER_FILENAME) if BANNER_FILENAME else None

    # Post
    asyncio.run(_post_to_discord(token, GUILD_ID, CHANNEL_ID, message, banner_path))


if __name__ == "__main__":
    main()

