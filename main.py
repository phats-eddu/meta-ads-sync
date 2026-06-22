"""
Meta Ads → Lark Sheet Sync
==========================
ดึงข้อมูล Ad Level จาก Meta Marketing API
เขียนลง Lark Sheet แยกต่อ Tab
รัน 2 รอบ/วัน: 04:00 และ 12:00 (เวลาไทย)
"""

import os
import requests
import time
import logging
from datetime import datetime, timezone, timedelta

# ────────────────────────────────────────────────────────────
# CONFIG
# ────────────────────────────────────────────────────────────
META_ACCESS_TOKEN = os.environ.get("META_TOKEN", "")
META_API_VERSION  = "v21.0"
START_DATE        = "2026-06-19"

LARK_APP_ID     = "cli_aab9c00f4f78deef"
LARK_APP_SECRET = "KptQ06EE7nZ3MOcVr3uIqe3LUCfTvzo5"
LARK_BASE_URL   = "https://open.larksuite.com/open-apis"
SPREADSHEET_TOKEN = "GCHcsDFJBhCqkQtoMdQlt2XZgTg"

ACCOUNTS = [
    {"id": "2377294275911854", "sheet_name": "Raw Data - 1 Day Series"},
    {"id": "286436192517445",  "sheet_name": "Raw Data - Shortcut Series"},
    {"id": "1439851086420085", "sheet_name": "Raw Data - XMBA"},
    {"id": "1043506717150301", "sheet_name": "Raw Data - XSeries"},
]

HEADERS = [
    "Date",
    "Campaign Name", "Campaign ID",
    "Ad Set Name", "Ad Set ID",
    "Ad Name", "Ad ID",
    "Spend (THB)",
    "Impressions", "Reach",
    "Clicks", "CTR (%)", "CPC", "CPM", "Frequency",
    "Message Starts",
    "Leads",
    "Purchases",
    "Purchase Value",
    "Purchase ROAS",
    "Cost per Lead",
    "Cost per Purchase",
]

META_FIELDS = ",".join([
    "date_start",
    "campaign_name", "campaign_id",
    "adset_name", "adset_id",
    "ad_name", "ad_id",
    "spend", "impressions", "reach",
    "clicks", "ctr", "cpc", "cpm", "frequency",
    "actions", "action_values",
    "cost_per_action_type",
    "purchase_roas", "website_purchase_roas",
])

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)

# ────────────────────────────────────────────────────────────
# LARK AUTH
# ────────────────────────────────────────────────────────────
_lark_token_cache = {"token": None, "expires_at": 0}

def get_lark_token() -> str:
    now = time.time()
    if _lark_token_cache["token"] and now < _lark_token_cache["expires_at"]:
        return _lark_token_cache["token"]
    url = f"{LARK_BASE_URL}/auth/v3/tenant_access_token/internal"
    r = requests.post(url, json={"app_id": LARK_APP_ID, "app_secret": LARK_APP_SECRET})
    r.raise_for_status()
    data = r.json()
    if data.get("code") != 0:
        raise Exception(f"Lark Auth Error: {data.get('msg')}")
    _lark_token_cache["token"] = data["tenant_access_token"]
    _lark_token_cache["expires_at"] = now + data.get("expire", 7200) - 60
    log.info("✅ Lark token refreshed")
    return _lark_token_cache["token"]

def lark_headers() -> dict:
    return {
        "Authorization": f"Bearer {get_lark_token()}",
        "Content-Type": "application/json",
    }

# ────────────────────────────────────────────────────────────
# LARK SHEET HELPERS
# ────────────────────────────────────────────────────────────
def get_sheet_id(sheet_name: str) -> str | None:
    url = f"{LARK_BASE_URL}/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/query"
    r = requests.get(url, headers=lark_headers())
    r.raise_for_status()
    data = r.json()
    if data.get("code") != 0:
        raise Exception(f"Lark Sheet query error: {data.get('msg')}")
    for sheet in data["data"]["sheets"]:
        if sheet["title"] == sheet_name:
            return sheet["sheet_id"]
    return None

def create_sheet_tab(sheet_name: str) -> str:
    url = f"{LARK_BASE_URL}/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets"
    body = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    r = requests.post(url, headers=lark_headers(), json=body)
    r.raise_for_status()
    data = r.json()
    if data.get("code") != 0:
        raise Exception(f"Create sheet error: {data.get('msg')}")
    return data["data"]["replies"][0]["addSheet"]["properties"]["sheetId"]

def ensure_sheet(sheet_name: str) -> str:
    sheet_id = get_sheet_id(sheet_name)
    if not sheet_id:
        log.info(f"📋 Creating new tab: {sheet_name}")
        sheet_id = create_sheet_tab(sheet_name)
    return sheet_id

def clear_sheet(sheet_id: str):
    url = f"{LARK_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{sheet_id}!A1:A"
    r = requests.get(url, headers=lark_headers())
    data = r.json()
    rows = data.get("data", {}).get("valueRange", {}).get("values", [])
    total_rows = len(rows)
    if total_rows <= 1:
        return
    clear_range = f"{sheet_id}!A2:W{total_rows + 10}"
    url = f"{LARK_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{clear_range}"
    requests.delete(url, headers=lark_headers())
    log.info(f"🗑️ Cleared {total_rows - 1} rows")

def write_header(sheet_id: str):
    url = f"{LARK_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
    body = {"valueRange": {"range": f"{sheet_id}!A1", "values": [HEADERS]}}
    r = requests.put(url, headers=lark_headers(), json=body)
    r.raise_for_status()
    log.info("📝 Header written")

def check_has_header(sheet_id: str) -> bool:
    url = f"{LARK_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{sheet_id}!A1"
    r = requests.get(url, headers=lark_headers())
    data = r.json()
    values = data.get("data", {}).get("valueRange", {}).get("values", [])
    return bool(values and values[0])

def write_rows_batch(sheet_id: str, rows: list):
    if not rows:
        return
    BATCH_SIZE = 1000
    start_row = 2
    for i in range(0, len(rows), BATCH_SIZE):
        batch = rows[i:i + BATCH_SIZE]
        end_row = start_row + len(batch) - 1
        range_str = f"{sheet_id}!A{start_row}:W{end_row}"
        url = f"{LARK_BASE_URL}/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values"
        body = {"valueRange": {"range": range_str, "values": batch}}
        r = requests.put(url, headers=lark_headers(), json=body)
        r.raise_for_status()
        data = r.json()
        if data.get("code") != 0:
            raise Exception(f"Write error: {data.get('msg')}")
        log.info(f"  ✍️ Written rows {start_row}–{end_row} ({len(batch)} rows)")
        start_row += len(batch)
        time.sleep(0.3)

# ────────────────────────────────────────────────────────────
# META API
# ────────────────────────────────────────────────────────────
def get_action(arr: list, action_type: str) -> float:
    if not arr:
        return 0.0
    found = next((a for a in arr if a.get("action_type") == action_type), None)
    return float(found["value"]) if found else 0.0

def get_first_action(arr: list, *action_types) -> float:
    for at in action_types:
        val = get_action(arr, at)
        if val:
            return val
    return 0.0

def parse_roas(item: dict) -> float:
    for key in ("purchase_roas", "website_purchase_roas"):
        roas = item.get(key, [])
        if isinstance(roas, list) and roas:
            return float(roas[0].get("value", 0))
    spend = float(item.get("spend", 0))
    if spend == 0:
        return 0.0
    pv = get_first_action(item.get("action_values", []),
        "purchase", "offsite_conversion.fb_pixel_purchase", "omni_purchase")
    return round(pv / spend, 4) if pv else 0.0

def parse_row(item: dict) -> list:
    actions    = item.get("actions", [])
    act_values = item.get("action_values", [])
    cost_pa    = item.get("cost_per_action_type", [])

    leads = get_first_action(actions,
        "onsite_conversion.lead_grouped",
        "lead",
        "onsite_web_lead",
        "onsite_conversion.lead",
        "offsite_conversion.fb_pixel_lead",
        "leadgen_grouped",
    )
    purchases = get_first_action(actions,
        "purchase", "offsite_conversion.fb_pixel_purchase", "omni_purchase")
    purchase_value = get_first_action(act_values,
        "purchase", "offsite_conversion.fb_pixel_purchase", "omni_purchase")
    cpl = get_first_action(cost_pa,
        "onsite_conversion.lead_grouped", "lead", "onsite_web_lead",
        "onsite_conversion.lead", "offsite_conversion.fb_pixel_lead")
    cpp = get_first_action(cost_pa,
        "purchase", "offsite_conversion.fb_pixel_purchase", "omni_purchase")

    return [
        item.get("date_start", ""),
        item.get("campaign_name", ""),
        item.get("campaign_id", ""),
        item.get("adset_name", ""),
        item.get("adset_id", ""),
        item.get("ad_name", ""),
        item.get("ad_id", ""),
        float(item.get("spend", 0)),
        int(item.get("impressions", 0)),
        int(item.get("reach", 0)),
        int(item.get("clicks", 0)),
        float(item.get("ctr", 0)),
        float(item.get("cpc", 0)),
        float(item.get("cpm", 0)),
        float(item.get("frequency", 0)),
        get_first_action(actions,
            "onsite_conversion.messaging_conversation_started_7d",
            "messaging_conversation_started_7d"),
        leads,
        purchases,
        purchase_value,
        parse_roas(item),
        cpl,
        cpp,
    ]

def fetch_insights(account_id: str, start_date: str, end_date: str) -> list:
    base_url = f"https://graph.facebook.com/{META_API_VERSION}/act_{account_id}/insights"
    all_rows = []
    after = None
    page = 0
    while True:
        page += 1
        params = {
            "access_token": META_ACCESS_TOKEN,
            "fields": META_FIELDS,
            "level": "ad",
            "time_range": f'{{"since":"{start_date}","until":"{end_date}"}}',
            "time_increment": 1,
            "limit": 500,
        }
        if after:
            params["after"] = after
        r = requests.get(base_url, params=params)
        r.raise_for_status()
        data = r.json()
        if "error" in data:
            raise Exception(f"Meta API error: {data['error']['message']}")
        items = data.get("data", [])
        log.info(f"  Page {page}: {len(items)} records")
        all_rows.extend(parse_row(item) for item in items)
        next_url = data.get("paging", {}).get("next")
        if not next_url:
            break
        after = data.get("paging", {}).get("cursors", {}).get("after")
        if not after:
            break
        time.sleep(0.5)
    return all_rows

# ────────────────────────────────────────────────────────────
# MAIN
# ────────────────────────────────────────────────────────────
def get_today_thai() -> str:
    thai_tz = timezone(timedelta(hours=7))
    return datetime.now(thai_tz).strftime("%Y-%m-%d")

def sync_account(account: dict, today: str):
    name = account["sheet_name"]
    account_id = account["id"]
    log.info(f"📊 [{name}] Starting sync (act_{account_id})")
    sheet_id = ensure_sheet(name)
    if not check_has_header(sheet_id):
        write_header(sheet_id)
    clear_sheet(sheet_id)
    log.info(f"  📅 Fetching {START_DATE} → {today}")
    rows = fetch_insights(account_id, START_DATE, today)
    rows = [r for r in rows if not (r[0] >= "2026-06-19" and r[8] == 0)]
    log.info(f"  📦 Total records: {len(rows)}")
    if not rows:
        log.warning(f"  ⚠️ No data returned")
        return
    write_rows_batch(sheet_id, rows)
    log.info(f"✅ [{name}] Done — {len(rows)} rows written")

def run():
    today = get_today_thai()
    log.info(f"🚀 Meta → Lark Sheet Sync | {today}")
    log.info(f"   Accounts: {len(ACCOUNTS)} | Date range: {START_DATE} → {today}")
    for account in ACCOUNTS:
        try:
            sync_account(account, today)
        except Exception as e:
            log.error(f"❌ [{account['sheet_name']}] Error: {e}")
        time.sleep(1)
    log.info("🎉 All done!")

if __name__ == "__main__":
    run()
