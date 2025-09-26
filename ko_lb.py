#!/usr/bin/env python3
import argparse
import asyncio
import csv
import datetime as dt
import os
import json
import io
import zipfile
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

# Optional: external link button under the embed
DISCORD_BUTTON_LABEL: str = "Voir tous les classements"
DISCORD_BUTTON_URL: str = "https://example.com/ko-series"  # replace with your link


def _normalize_user_id(value: str) -> str:
    """Normalize identifiers like '123.0' -> '123', '00123' -> '123'.
    If not numeric-like, return stripped value.
    """
    s = (value or "").strip()
    if not s:
        return s
    try:
        f = float(s)
        if abs(f - round(f)) < 1e-9:
            return str(int(round(f)))
        # Non-integer but numeric; keep canonical float string without trailing .0
        return ("%f" % f).rstrip("0").rstrip(".")
    except Exception:
        return s


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
            # Track raw and normalized ids
            customer_ids.append(customer_id_raw)
            if discord_id_raw:
                customer_to_discord[customer_id_raw] = discord_id_raw
                norm = _normalize_user_id(customer_id_raw)
                if norm and norm not in customer_to_discord:
                    customer_to_discord[norm] = discord_id_raw

    # Deduplicate preserving order
    seen: set = set()
    deduped: List[str] = []
    for cid in customer_ids:
        if cid in seen:
            continue
        seen.add(cid)
        deduped.append(cid)
    # Also provide a normalized-only list to help with matching later
    normalized_deduped: List[str] = []
    seen_norm: set = set()
    for cid in deduped:
        norm = _normalize_user_id(cid)
        if norm and norm not in seen_norm:
            seen_norm.add(norm)
            normalized_deduped.append(norm)
    return normalized_deduped, customer_to_discord


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


async def _fetch_display_names(token: str, guild_id: int, discord_ids: List[str]) -> Dict[str, str]:
    """Fetch server display names for a list of discord user IDs in a guild.
    Returns mapping discord_id -> display_name (prefers nick, then global_name, then username).
    Missing users are omitted.
    """
    import aiohttp
    headers_json = {'Authorization': f'Bot {token}', 'Content-Type': 'application/json'}
    base_url = f'https://discord.com/api/v10/guilds/{guild_id}/members'
    names: Dict[str, str] = {}
    async with aiohttp.ClientSession() as session:
        for uid in discord_ids:
            if not uid:
                continue
            try:
                url = f"{base_url}/{uid}"
                async with session.get(url, headers=headers_json) as resp:
                    if resp.status != 200:
                        continue
                    data = await resp.json()
                    nick = data.get('nick')
                    user = data.get('user') or {}
                    global_name = user.get('global_name')
                    username = user.get('username')
                    display = nick or global_name or username
                    if display:
                        names[uid] = str(display)
            except Exception:
                continue
    return names


def _build_name_mapping_for_rankings(
    df_list: List[pd.DataFrame], customer_to_discord: Dict[str, str], discord_to_display: Dict[str, str]
) -> Dict[str, str]:
    """Map ranking 'user_id' (customer_id) to guild display names using users_list and fetched member names.
    Fallback to mention or raw id if missing.
    Returns mapping customer_id -> display_name.
    """
    mapping: Dict[str, str] = {}
    for df in df_list:
        if df is None or df.empty:
            continue
        for _, r in df.iterrows():
            customer_id = _normalize_user_id(str(r.get('user_id')))
            if not customer_id:
                continue
            discord_id = customer_to_discord.get(customer_id)
            display = None
            if discord_id:
                # Prefer fetched guild display, else fall back to mention
                display = discord_to_display.get(discord_id) or f"<@{discord_id}>"
            if not display:
                # As a last resort keep the original ID for traceability
                display = f"ID:{customer_id}"
            mapping[customer_id] = display
    return mapping


def _transform_ranking_df_for_names(df: pd.DataFrame, customer_to_name: Dict[str, str]) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df2 = df.copy()
    # Rename columns
    if 'rank' in df2.columns:
        df2 = df2.rename(columns={'rank': 'classement'})
    # Replace user_id with 'nom'
    df2['nom'] = df2['user_id'].astype(str).map(lambda x: customer_to_name.get(str(x), f"ID:{x}"))
    # Reorder: classement, nom, points, rest...
    cols = ['classement', 'nom'] + [c for c in df2.columns if c not in ('classement', 'nom', 'user_id')]
    df2 = df2[cols]
    return df2


def _format_leaderboard(df: pd.DataFrame, date_str: str, id_map: Dict[str, str]) -> Tuple[str, str, List[Tuple[int, str, float]]]:
    """Build leaderboard parts for Discord:
    - returns (embed_title, embed_description, rows_used)
    The congratulations text will be built later from rows.
    Only top 10 players with points > 0 are included.
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
        
        if points <= 0:
            continue
        if len(rows) >= 10:
            break
        
        discord_id = id_map.get(user_id) or id_map.get(str(int(float(user_id)))) if user_id.replace(".", "", 1).isdigit() else id_map.get(user_id)
        display = f"<@{discord_id}>" if discord_id else f"ID:{user_id}"
        rows.append((rank, display, points))

    embed_title = f"Classement {date_str}"
    desc_lines: List[str] = []
    for rank, user, pts in rows:
        if rank == 1:
            desc_lines.append(f"1er - {user} avec {pts:.0f} point(s).")
        elif rank == 2:
            desc_lines.append(f"2ème - {user} avec {pts:.0f} point(s).")
        else:
            desc_lines.append(f"{rank}ème - {user} avec {pts:.0f} point(s).")
    embed_description = "\n".join(desc_lines) if desc_lines else "Aucun joueur avec des points > 0."

    return embed_title, embed_description, rows


def _build_congratulations(rows: List[Tuple[int, str, float]]) -> str:
    if not rows:
        return ""
    lines: List[str] = []
    # First congratulations: all players ranked 1st to 5th (including ties at 5th)
    first_5_users: List[str] = []
    for rank, user, _ in rows:
        if rank <= 5:
            first_5_users.append(user)
    if first_5_users:
        first_5_users_str = ", ".join(first_5_users)
        lines.append(f"Félicitations {first_5_users_str}, vous gagnez tous un ticket pour le Main Event des KO Series 50 000€ garanti de dimanche 5 octobre.")

    # Second congratulations: all players ranked 6th to 10th (including ties)
    sixth_to_tenth_users: List[str] = []
    for rank, user, _ in rows:
        if 6 <= rank <= 10:
            sixth_to_tenth_users.append(user)
    if sixth_to_tenth_users:
        sixth_to_tenth_users_str = ", ".join(sixth_to_tenth_users)
        lines.append(f"Bien joué {sixth_to_tenth_users_str}, vous gagnez tous un ticket 20€ à utiliser dans les tournois des KO Series !")
    return "\n".join(lines)


async def _post_to_discord(token: str, guild_id: int, channel_id: int, *, banner_path: Optional[str], embed_title: str, embed_description: str, congrats_text: str, attachments: Optional[List[Tuple[str, bytes, str]]] = None) -> None:
    import aiohttp

    url = f'https://discord.com/api/v10/channels/{channel_id}/messages'
    headers = {'Authorization': f'Bot {token}'}
    
    async with aiohttp.ClientSession() as session:
        try:
            # 1) Send banner image if provided
            if banner_path and os.path.exists(banner_path):
                form_data = aiohttp.FormData()
                with open(banner_path, 'rb') as f:
                    file_content = f.read()
                form_data.add_field('file', file_content, filename=os.path.basename(banner_path))
                async with session.post(url, headers=headers, data=form_data) as response:
                    if response.status == 200:
                        print("Banner image posted successfully!")
                    else:
                        print(f"Error posting banner: {response.status} - {await response.text()}")

            # 2) Send embed with leaderboard and button, optionally with attachment(s) in same message
            headers_json = {'Authorization': f'Bot {token}', 'Content-Type': 'application/json'}
            payload = {
                'embeds': [
                    {
                        'title': embed_title,
                        'description': embed_description,
                        'color': 0x00AEEF
                    }
                ],
                'components': [
                    {
                        'type': 1,  # action row
                        'components': [
                            {
                                'type': 2,  # button
                                'style': 5,  # link style
                                'label': DISCORD_BUTTON_LABEL,
                                'url': DISCORD_BUTTON_URL
                            }
                        ]
                    }
                ]
            }
            if attachments:
                form_data = aiohttp.FormData()
                form_data.add_field('payload_json', json.dumps(payload), content_type='application/json')
                for idx, (filename, content, content_type) in enumerate(attachments):
                    form_data.add_field(f'files[{idx}]', content, filename=filename, content_type=content_type)
                async with session.post(url, headers={'Authorization': f'Bot {token}'}, data=form_data) as response:
                    if response.status == 200:
                        print("Leaderboard embed + attachment posted successfully!")
                    else:
                        print(f"Error posting embed+file: {response.status} - {await response.text()}")
            else:
                async with session.post(url, headers=headers_json, json=payload) as response:
                    if response.status == 200:
                        print("Leaderboard embed posted successfully!")
                    else:
                        print(f"Error posting embed: {response.status} - {await response.text()}")

            # 3) Send congratulations text, if any
            if congrats_text:
                data = {'content': congrats_text}
                async with session.post(url, headers=headers_json, json=data) as response:
                    if response.status == 200:
                        print("Congratulations message posted successfully!")
                    else:
                        print(f"Error posting congratulations: {response.status} - {await response.text()}")
        except Exception as e:
            print(f"Error posting to Discord: {e}")


async def _post_files_to_discord(token: str, channel_id: int, attachments: List[Tuple[str, bytes]]) -> None:
    import aiohttp
    if not attachments:
        return
    url = f'https://discord.com/api/v10/channels/{channel_id}/messages'
    headers = {'Authorization': f'Bot {token}'}
    async with aiohttp.ClientSession() as session:
        form_data = aiohttp.FormData()
        # Minimal content to avoid empty message
        form_data.add_field('payload_json', json.dumps({'content': 'Fichiers des classements:'}), content_type='application/json')
        for idx, (filename, content) in enumerate(attachments):
            form_data.add_field(f'files[{idx}]', content, filename=filename, content_type='text/csv')
        async with session.post(url, headers=headers, data=form_data) as response:
            if response.status == 200:
                print("CSV attachments posted successfully!")
            else:
                print(f"Error posting CSVs: {response.status} - {await response.text()}")


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

    # Build leaderboard and congratulations
    date_str = target_date.strftime("%d/%m/%Y")
    embed_title, embed_description, rows = _format_leaderboard(df, date_str, id_map)
    congrats_text = _build_congratulations(rows)

    # Prepare processed CSVs with display names for the two rankings
    # Read classement_koseries and classement_général for the date
    koseries_csv_path = os.path.join(os.getcwd(), f"classement_koseries_{target_date.strftime('%d-%m-%Y')}.csv")
    general_csv_path = os.path.join(os.getcwd(), f"classement_général_{target_date.strftime('%d-%m-%Y')}.csv")
    koseries_df = pd.read_csv(koseries_csv_path) if os.path.exists(koseries_csv_path) else pd.DataFrame()
    general_df = pd.read_csv(general_csv_path) if os.path.exists(general_csv_path) else pd.DataFrame()

    # Collect discord IDs to fetch names
    discord_ids_needed: List[str] = []
    for source_df in (koseries_df, general_df):
        if source_df is None or source_df.empty:
            continue
        for _, r in source_df.iterrows():
            customer_id = str(r.get('user_id'))
            did = id_map.get(customer_id)
            if did and did not in discord_ids_needed:
                discord_ids_needed.append(did)

    # Fetch display names
    token = DISCORD_BOT_TOKEN.strip()
    if not token or token == "YOUR_BOT_TOKEN_HERE":
        raise SystemExit(
            "Please replace 'YOUR_BOT_TOKEN_HERE' with your actual Discord bot token in ko_lb.py"
        )
    discord_to_display = asyncio.get_event_loop().run_until_complete(
        _fetch_display_names(token, GUILD_ID, discord_ids_needed)
    )

    # Build customer_id -> display name mapping
    customer_to_name = _build_name_mapping_for_rankings([koseries_df, general_df], id_map, discord_to_display)

    # Transform dataframes
    koseries_named_df = _transform_ranking_df_for_names(koseries_df, customer_to_name)
    general_named_df = _transform_ranking_df_for_names(general_df, customer_to_name)

    # Serialize to CSV then ZIP both into a single archive
    attachments_zip: Optional[List[Tuple[str, bytes, str]]] = None
    if not koseries_named_df.empty or not general_named_df.empty:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
            if not koseries_named_df.empty:
                zf.writestr(
                    f"classement_koseries_{target_date.strftime('%d-%m-%Y')}_noms.csv",
                    koseries_named_df.to_csv(index=False)
                )
            if not general_named_df.empty:
                zf.writestr(
                    f"classement_général_{target_date.strftime('%d-%m-%Y')}_noms.csv",
                    general_named_df.to_csv(index=False)
                )
        zip_bytes = buffer.getvalue()
        attachments_zip = [(f"classements_{target_date.strftime('%d-%m-%Y')}.zip", zip_bytes, 'application/zip')]

    # Discord
    banner_path = os.path.join(base_dir, BANNER_FILENAME) if BANNER_FILENAME else None

    # Post
    asyncio.run(_post_to_discord(token, GUILD_ID, CHANNEL_ID, banner_path=banner_path, embed_title=embed_title, embed_description=embed_description, congrats_text=congrats_text, attachments=attachments_zip))


if __name__ == "__main__":
    main()

