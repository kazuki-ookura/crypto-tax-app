#!/usr/bin/env python3
"""
暗号資産 確定申告用 統合CSV作成スクリプト
使い方: python crypto_tax.py [--year 2025] [--binance|--no-binance]
"""

import csv
import glob
import os
import io
import sys
import time
import hmac
import hashlib
import json
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "data")
OUTPUT_DIR = SCRIPT_DIR
YEAR = 0  # main() で --year 引数またはデフォルト(前年)で設定される
JST = timezone(timedelta(hours=9))

# --- Binance API フラグ ---
# 環境変数 FETCH_BINANCE=0 でスキップ、1（デフォルト）で有効
# コマンドライン引数 --binance / --no-binance でも上書き可能
FETCH_BINANCE = os.environ.get("FETCH_BINANCE", "1") != "0"


def load_env(path=None):
    """Load .env file as dict"""
    if path is None:
        path = os.path.join(SCRIPT_DIR, ".env")
    env = {}
    if not os.path.exists(path):
        return env
    with open(path, "r") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                key, _, val = line.partition("=")
                env[key.strip()] = val.strip()
    return env

# --- Unified record format ---
# date, exchange, currency, category, subcategory, quantity, rate_jpy, amount_jpy, fee, fee_jpy, memo

def parse_gmo(rows):
    """Parse GMOコイン CSV (UTF-8 BOM)"""
    records = []
    for row in rows:
        date_str = row.get("日時", "").strip()
        if not date_str:
            continue
        dt = datetime.strptime(date_str, "%Y/%m/%d %H:%M")
        if dt.year != YEAR:
            continue

        seisan = row.get("精算区分", "").strip()
        currency = row.get("銘柄名", "").strip()

        if seisan == "日本円入出金":
            cat = "入出金"
            subcat = row.get("入出金区分", "").strip()
            amount = Decimal(row.get("入出金金額", "0").replace(",", ""))
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "GMOコイン",
                "currency": "JPY",
                "category": cat,
                "subcategory": subcat,
                "quantity": str(amount),
                "rate_jpy": "1",
                "amount_jpy": str(amount),
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif seisan in ("販売所取引", "取引所現物取引"):
            side = row.get("売買区分", "").strip()
            qty = row.get("約定数量", "0").replace(",", "")
            rate = row.get("約定レート", "0").replace(",", "")
            amount = row.get("約定金額", "0").replace(",", "")
            fee = row.get("注文手数料", "0").replace(",", "")
            jpy_amount = row.get("日本円受渡金額", "0").replace(",", "")

            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "GMOコイン",
                "currency": currency,
                "category": "売買",
                "subcategory": f"{seisan}・{side}",
                "quantity": qty,
                "rate_jpy": rate,
                "amount_jpy": jpy_amount,
                "fee": fee,
                "fee_jpy": fee,
                "memo": ""
            })
        elif seisan == "取引所現物 取引手数料返金":
            jpy_amount = row.get("日本円受渡金額", "0").replace(",", "")
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "GMOコイン",
                "currency": currency,
                "category": "手数料返金",
                "subcategory": "取引手数料返金",
                "quantity": "0",
                "rate_jpy": "0",
                "amount_jpy": jpy_amount,
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif seisan == "暗号資産預入・送付":
            direction = row.get("授受区分", "").strip()
            qty = row.get("数量", "0").replace(",", "")
            send_fee = row.get("送付手数料", "0").replace(",", "") if row.get("送付手数料", "").strip() else "0"
            dest = row.get("送付先/送付元", "").strip()

            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "GMOコイン",
                "currency": currency,
                "category": "送付・預入",
                "subcategory": direction,
                "quantity": qty,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": send_fee,
                "fee_jpy": "0",
                "memo": dest
            })
    return records


def parse_bitlending(rows):
    """Parse BitLending CSV (UTF-8 BOM)"""
    records = []
    for row in rows:
        date_str = row.get("タイムスタンプ", "").strip()
        if not date_str:
            continue
        dt = datetime.strptime(date_str, "%Y/%m/%d %H:%M:%S")
        if dt.year != YEAR:
            continue

        kind = row.get("種別", "").strip()
        currency = row.get("銘柄名", "").strip()
        qty = row.get("数量", "0").strip()
        rate = row.get("レート", "").strip()

        if kind == "貸借料付与":
            qty_d = Decimal(qty)
            rate_d = Decimal(rate) if rate else Decimal(0)
            amount_jpy = str((qty_d * rate_d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "BitLending",
                "currency": currency,
                "category": "レンディング報酬",
                "subcategory": "貸借料付与",
                "quantity": qty,
                "rate_jpy": rate,
                "amount_jpy": amount_jpy,
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif kind == "貸出開始":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "BitLending",
                "currency": currency,
                "category": "レンディング",
                "subcategory": "貸出開始",
                "quantity": qty,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
    return records


def parse_pbr_lending(rows):
    """Parse PBRレンディング CSV (Shift_JIS)"""
    records = []
    for row in rows:
        date_str = row.get("お取引日時", "").strip().strip('"')
        if not date_str:
            continue
        dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        if dt.year != YEAR:
            continue

        kind = row.get("お取引内容", "").strip()
        summary = row.get("摘要", "").strip()
        currency = row.get("通貨種別", "").strip()
        qty = row.get("数量", "0").strip()
        rate = row.get("ご参考レート", "").strip()
        memo = row.get("備考", "").strip()

        qty_d = Decimal(qty)
        rate_d = Decimal(rate) if rate else Decimal(0)
        amount_jpy = str((qty_d * rate_d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

        # Determine category
        if kind in ("紹介報酬", "紹介報酬(ｷｬﾝﾍﾟｰﾝ)"):
            cat = "紹介報酬"
        elif "利息" in summary or kind == "通常→プレミアム":
            cat = "レンディング報酬"
        else:
            cat = "その他報酬"

        records.append({
            "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
            "exchange": "PBRレンディング",
            "currency": currency,
            "category": cat,
            "subcategory": kind,
            "quantity": qty,
            "rate_jpy": rate,
            "amount_jpy": amount_jpy,
            "fee": "0",
            "fee_jpy": "0",
            "memo": memo
        })
    return records


def parse_bitflyer(rows):
    """Parse bitFlyer CSV (UTF-8 BOM)"""
    records = []
    for row in rows:
        date_str = row.get("取引日時", "").strip().strip('"')
        if not date_str:
            continue
        dt = datetime.strptime(date_str, "%Y/%m/%d %H:%M:%S")
        if dt.year != YEAR:
            continue

        kind = row.get("取引種別", "").strip().strip('"')
        currency_raw = row.get("通貨", "").strip().strip('"')
        cur1 = row.get("通貨1", "").strip().strip('"')
        qty1 = row.get("通貨1数量", "0").strip().strip('"').replace(",", "")
        fee = row.get("手数料", "0").strip().strip('"').replace(",", "")
        rate = row.get("通貨1の対円レート", "0").strip().strip('"').replace(",", "")
        cur2 = row.get("通貨2", "").strip().strip('"')
        qty2 = row.get("通貨2数量", "0").strip().strip('"').replace(",", "")
        price = row.get("取引価格", "0").strip().strip('"').replace(",", "")
        memo = row.get("備考", "").strip().strip('"')

        if kind == "受取":
            qty_d = Decimal(qty1) if qty1 else Decimal(0)
            rate_d = Decimal(rate) if rate else Decimal(0)
            amount_jpy = str((qty_d * rate_d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "bitFlyer",
                "currency": cur1,
                "category": "受取",
                "subcategory": "受取",
                "quantity": qty1,
                "rate_jpy": rate,
                "amount_jpy": amount_jpy,
                "fee": fee,
                "fee_jpy": "0",
                "memo": memo
            })
        elif kind in ("売り", "買い"):
            # Trade: currency_raw is like "BTC/JPY"
            side = kind
            qty_d = Decimal(qty1) if qty1 else Decimal(0)
            # qty1 is negative for sells
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "bitFlyer",
                "currency": cur1,
                "category": "売買",
                "subcategory": side,
                "quantity": str(abs(qty_d)),
                "rate_jpy": price,
                "amount_jpy": qty2 if qty2 else "0",
                "fee": fee,
                "fee_jpy": "0",
                "memo": ""
            })
        elif kind == "手数料":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "bitFlyer",
                "currency": cur1,
                "category": "手数料",
                "subcategory": "手数料",
                "quantity": qty1,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": qty1,
                "fee_jpy": qty1 if cur1 == "JPY" else "0",
                "memo": memo
            })
        elif kind == "出金":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "bitFlyer",
                "currency": cur1,
                "category": "入出金",
                "subcategory": "出金",
                "quantity": qty1,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": "0",
                "fee_jpy": "0",
                "memo": memo
            })
    return records


def fetch_binance_data(api_key, secret):
    """Fetch trade/income data from Binance API for YEAR (with JSON cache)"""
    cache_dir = os.path.join(DATA_DIR, "Binance")
    cache_path = os.path.join(cache_dir, f"{YEAR}_records.json")

    if os.path.exists(cache_path):
        print(f"  Binanceキャッシュを読み込み: {cache_path}")
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)

    records = []
    base_url = "https://api.binance.com"

    def signed_request(endpoint, params=None):
        if params is None:
            params = {}
        params["timestamp"] = str(int(time.time() * 1000))
        query_string = urllib.parse.urlencode(params)
        signature = hmac.new(secret.encode(), query_string.encode(), hashlib.sha256).hexdigest()
        query_string += f"&signature={signature}"
        url = f"{base_url}{endpoint}?{query_string}"
        req = urllib.request.Request(url, headers={"X-MBX-APIKEY": api_key})
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            body = e.read().decode()
            print(f"  API error for {endpoint}: HTTP {e.code} {body}")
            return None
        except Exception as e:
            print(f"  API error for {endpoint}: {e}")
            return None

    def ms_range_chunks(start_ms, end_ms, chunk_days):
        """Yield (chunk_start, chunk_end) pairs in chunks of chunk_days days"""
        chunk_ms = chunk_days * 24 * 60 * 60 * 1000
        cur = start_ms
        while cur < end_ms:
            yield cur, min(cur + chunk_ms - 1, end_ms)
            cur += chunk_ms

    # Target year time range in milliseconds
    year_start_ms = int(datetime(YEAR, 1, 1, tzinfo=timezone.utc).timestamp() * 1000)
    year_end_ms = int(datetime(YEAR + 1, 1, 1, tzinfo=timezone.utc).timestamp() * 1000)

    # Get account info to find which symbols we have trades for
    print("Fetching Binance account info...")
    account = signed_request("/api/v3/account")
    if not account:
        print("Failed to fetch Binance account info")
        return records

    # Find non-zero balances to determine which symbols to query
    assets_with_balance = []
    for bal in account.get("balances", []):
        free = Decimal(bal["free"])
        locked = Decimal(bal["locked"])
        if free + locked > 0:
            assets_with_balance.append(bal["asset"])

    print(f"  Assets with balance: {assets_with_balance}")

    # Fetch trades for each asset paired with common quote currencies
    # Strategy: get all trades with startTime=year_start (no endTime), filter by year
    # This avoids the 24-hour window restriction while minimizing API calls
    quote_currencies = ["USDT", "BTC", "JPY", "BUSD", "BNB", "ETH", "FDUSD"]
    fetched_symbols = set()

    for asset in assets_with_balance:
        if asset in ("USDT", "JPY", "BUSD", "FDUSD"):
            continue
        for quote in quote_currencies:
            if quote == asset:
                continue
            symbol = f"{asset}{quote}"
            if symbol in fetched_symbols:
                continue
            fetched_symbols.add(symbol)

            time.sleep(0.2)  # rate limit: 5 req/sec
            symbol_trades = []
            from_id = None
            while True:
                params = {
                    "symbol": symbol,
                    "startTime": str(year_start_ms),
                    "limit": "1000"
                }
                if from_id is not None:
                    params["fromId"] = str(from_id)
                    del params["startTime"]  # fromId and startTime conflict

                trades = signed_request("/api/v3/myTrades", params)
                if not trades or not isinstance(trades, list) or len(trades) == 0:
                    break

                # Filter to target year
                year_trades = [t for t in trades if datetime.fromtimestamp(t["time"] / 1000, tz=timezone.utc).year == YEAR]
                symbol_trades.extend(year_trades)

                if len(trades) < 1000:
                    break  # no more pages
                # Check if last trade is still within target year
                last_time = datetime.fromtimestamp(trades[-1]["time"] / 1000, tz=timezone.utc)
                if last_time.year > YEAR:
                    break
                from_id = trades[-1]["id"] + 1
                time.sleep(0.2)

            if symbol_trades:
                print(f"  Found {len(symbol_trades)} trades for {symbol}")
                for t in symbol_trades:
                    trade_time = datetime.fromtimestamp(t["time"] / 1000, tz=JST)
                    side = "買" if t["isBuyer"] else "売"
                    qty = t["qty"]
                    price = t["price"]
                    quote_qty = t["quoteQty"]
                    commission = t["commission"]
                    commission_asset = t["commissionAsset"]

                    records.append({
                        "date": trade_time.strftime("%Y-%m-%d %H:%M:%S"),
                        "exchange": "Binance",
                        "currency": asset,
                        "category": "売買",
                        "subcategory": f"{side}（{symbol}）",
                        "quantity": qty,
                        "rate_jpy": price if quote == "JPY" else f"{price}({quote})",
                        "amount_jpy": quote_qty if quote == "JPY" else f"{quote_qty}({quote})",
                        "fee": commission,
                        "fee_jpy": commission if commission_asset == "JPY" else f"{commission}({commission_asset})",
                        "memo": f"pair:{symbol}"
                    })

    # Fetch Simple Earn flexible rewards
    # /sapi/v1/simple-earn: max 30 days per query
    print("  Fetching Simple Earn rewards history...")
    for earn_type in ["BONUS", "REALTIME", "REWARDS"]:
        for chunk_start, chunk_end in ms_range_chunks(year_start_ms, year_end_ms, 30):
            time.sleep(0.3)
            history = signed_request("/sapi/v1/simple-earn/flexible/history/rewardsRecord", {
                "type": earn_type,
                "startTime": str(chunk_start),
                "endTime": str(chunk_end),
                "size": "100"
            })
            if history and isinstance(history, dict):
                for h in history.get("rows", []):
                    dt = datetime.fromtimestamp(h.get("time", 0) / 1000, tz=JST)
                    records.append({
                        "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                        "exchange": "Binance",
                        "currency": h.get("asset", ""),
                        "category": "ステーキング報酬",
                        "subcategory": f"SimpleEarn/{earn_type}",
                        "quantity": h.get("rewards", "0"),
                        "rate_jpy": "0",
                        "amount_jpy": "0",
                        "fee": "0",
                        "fee_jpy": "0",
                        "memo": ""
                    })

    # Fetch staking rewards (locked staking)
    # /sapi/v1/staking: max 30 days per query
    print("  Fetching staking history...")
    for product_type in ["STAKING", "F_DEFI", "L_DEFI"]:
        for chunk_start, chunk_end in ms_range_chunks(year_start_ms, year_end_ms, 30):
            time.sleep(0.3)
            history = signed_request("/sapi/v1/staking/stakingRecord", {
                "product": product_type,
                "txnType": "INTEREST",
                "startTime": str(chunk_start),
                "endTime": str(chunk_end),
                "size": "100"
            })
            if history and isinstance(history, list):
                for h in history:
                    dt = datetime.fromtimestamp(h.get("time", 0) / 1000, tz=JST)
                    records.append({
                        "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                        "exchange": "Binance",
                        "currency": h.get("asset", ""),
                        "category": "ステーキング報酬",
                        "subcategory": product_type,
                        "quantity": h.get("amount", "0"),
                        "rate_jpy": "0",
                        "amount_jpy": "0",
                        "fee": "0",
                        "fee_jpy": "0",
                        "memo": ""
                    })

    # Fetch deposit history (max 90 days per query)
    print("  Fetching deposit history...")
    for chunk_start, chunk_end in ms_range_chunks(year_start_ms, year_end_ms, 90):
        time.sleep(0.3)
        deposits = signed_request("/sapi/v1/capital/deposit/hisrec", {
            "startTime": str(chunk_start),
            "endTime": str(chunk_end),
            "status": "1"  # success only
        })
        if deposits and isinstance(deposits, list):
            for d in deposits:
                dt = datetime.fromtimestamp(d.get("insertTime", 0) / 1000, tz=JST)
                records.append({
                    "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "exchange": "Binance",
                    "currency": d.get("coin", ""),
                    "category": "入出金",
                    "subcategory": "預入",
                    "quantity": d.get("amount", "0"),
                    "rate_jpy": "0",
                    "amount_jpy": "0",
                    "fee": "0",
                    "fee_jpy": "0",
                    "memo": d.get("network", "")
                })

    # Fetch withdrawal history (max 90 days per query)
    print("  Fetching withdrawal history...")
    for chunk_start, chunk_end in ms_range_chunks(year_start_ms, year_end_ms, 90):
        time.sleep(0.3)
        withdrawals = signed_request("/sapi/v1/capital/withdraw/history", {
            "startTime": str(chunk_start),
            "endTime": str(chunk_end),
            "status": "6"  # completed
        })
        if withdrawals and isinstance(withdrawals, list):
            for w in withdrawals:
                apply_time = w.get("applyTime", "")
                try:
                    dt = datetime.strptime(apply_time, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
                records.append({
                    "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                    "exchange": "Binance",
                    "currency": w.get("coin", ""),
                    "category": "入出金",
                    "subcategory": "送付",
                    "quantity": w.get("amount", "0"),
                    "rate_jpy": "0",
                    "amount_jpy": "0",
                    "fee": w.get("transactionFee", "0"),
                    "fee_jpy": "0",
                    "memo": w.get("network", "")
                })

    # Save to cache
    os.makedirs(cache_dir, exist_ok=True)
    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    print(f"  Binanceデータをキャッシュ保存: {cache_path}")

    return records


def find_csv(directory):
    """ディレクトリ内のCSVファイルを更新日時順（新しい順）で返す。"""
    files = glob.glob(os.path.join(directory, "*.csv"))
    return sorted(files, key=os.path.getmtime, reverse=True)


def read_csv_file(filepath, encoding="utf-8-sig"):
    """Read CSV file with specified encoding"""
    with open(filepath, "r", encoding=encoding) as f:
        reader = csv.DictReader(f)
        return list(reader)


def read_csv_sjis(filepath):
    """Read Shift_JIS encoded CSV"""
    with open(filepath, "r", encoding="shift_jis") as f:
        reader = csv.DictReader(f)
        return list(reader)


def read_bitpoint_csv(filepath):
    """Read BitPOINT CSV (Shift_JIS, has header rows to skip)"""
    with open(filepath, "r", encoding="shift_jis") as f:
        lines = f.readlines()

    # Find the header row (starts with "No,")
    header_idx = None
    for i, line in enumerate(lines):
        if line.startswith("No,"):
            header_idx = i
            break

    if header_idx is None:
        return []

    csv_text = "".join(lines[header_idx:])
    reader = csv.DictReader(io.StringIO(csv_text))
    return list(reader)


def parse_bitpoint(rows):
    """Parse BitPOINT data - staking rewards for YEAR"""
    records = []
    for row in rows:
        kind = row.get("取引種類", "").strip()
        date_str = row.get("受渡日", "").strip()

        # Skip 繰り越し and 合計
        if kind in ("繰り越し", "合計", ""):
            continue

        if not date_str:
            # Some records don't have 受渡日, check 約定日時
            date_str2 = row.get("約定日時", "").strip()
            if date_str2:
                dt = datetime.strptime(date_str2, "%Y/%m/%d %H:%M:%S")
            else:
                continue
        else:
            try:
                dt = datetime.strptime(date_str, "%Y/%m/%d")
            except:
                continue

        if dt.year != YEAR:
            continue

        currency1 = row.get("通貨コード１", "").strip()
        qty = row.get("数量", "0").strip()
        price = row.get("参考単価", "0").strip()

        if "ステーキング報酬" in kind:
            qty_d = Decimal(qty) if qty else Decimal(0)
            price_d = Decimal(price) if price else Decimal(0)
            amount_jpy = str((qty_d * price_d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S") if dt.hour else dt.strftime("%Y-%m-%d 00:00:00"),
                "exchange": "BitPOINT",
                "currency": currency1,
                "category": "ステーキング報酬",
                "subcategory": kind,
                "quantity": qty,
                "rate_jpy": price,
                "amount_jpy": amount_jpy,
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif kind == "現物取引":
            side = row.get("売買", "").strip()
            qty_d = Decimal(qty) if qty else Decimal(0)
            price_d = Decimal(price) if price else Decimal(0)
            amount_jpy = str((qty_d * price_d).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S") if dt.hour else dt.strftime("%Y-%m-%d 00:00:00"),
                "exchange": "BitPOINT",
                "currency": currency1,
                "category": "売買",
                "subcategory": f"現物取引・{side}",
                "quantity": qty,
                "rate_jpy": price,
                "amount_jpy": amount_jpy,
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif kind == "入金":
            records.append({
                "date": dt.strftime("%Y-%m-%d 00:00:00"),
                "exchange": "BitPOINT",
                "currency": currency1,
                "category": "入出金",
                "subcategory": "入金",
                "quantity": qty if qty else "0",
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
    return records


def parse_coincheck(rows):
    """Parse コインチェック CSV (UTF-8 BOM, 様式1)"""
    records = []
    for row in rows:
        date_str = row.get("取引日時", "").strip()
        if not date_str:
            continue
        dt = datetime.strptime(date_str, "%Y/%m/%d %H:%M:%S")
        if dt.year != YEAR:
            continue

        kind = row.get("取引種別", "").strip()
        trade_type = row.get("取引形態", "").strip()
        pair = row.get("通貨ペア", "").strip()
        inc_currency = row.get("増加通貨名", "").strip()
        inc_qty = row.get("増加数量", "0").strip()
        dec_currency = row.get("減少通貨名", "").strip()
        dec_qty = row.get("減少数量", "0").strip()
        exec_price = row.get("約定代金", "0").strip()
        unit_price = row.get("約定価格", "0").strip()
        fee_currency = row.get("手数料通貨", "").strip()
        fee_qty = row.get("手数料数量", "0").strip()
        memo = row.get("備考", "").strip()

        if kind == "入金":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": inc_currency,
                "category": "入出金",
                "subcategory": "入金",
                "quantity": inc_qty,
                "rate_jpy": "1" if inc_currency == "JPY" else "0",
                "amount_jpy": inc_qty if inc_currency == "JPY" else "0",
                "fee": "0",
                "fee_jpy": "0",
                "memo": ""
            })
        elif kind == "出金":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": dec_currency,
                "category": "入出金",
                "subcategory": "出金",
                "quantity": dec_qty,
                "rate_jpy": "1" if dec_currency == "JPY" else "0",
                "amount_jpy": dec_qty if dec_currency == "JPY" else "0",
                "fee": fee_qty,
                "fee_jpy": fee_qty if fee_currency == "JPY" else "0",
                "memo": ""
            })
        elif kind in ("売", "買"):
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": inc_currency if kind == "買" else dec_currency,
                "category": "売買",
                "subcategory": f"{trade_type}・{kind}",
                "quantity": inc_qty if kind == "買" else dec_qty,
                "rate_jpy": unit_price if unit_price else "0",
                "amount_jpy": exec_price if exec_price else "0",
                "fee": fee_qty,
                "fee_jpy": fee_qty if fee_currency == "JPY" else "0",
                "memo": pair
            })
        elif kind == "送金":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": dec_currency,
                "category": "送付・預入",
                "subcategory": "送付",
                "quantity": dec_qty,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": fee_qty,
                "fee_jpy": fee_qty if fee_currency == "JPY" else "0",
                "memo": row.get("送付先アドレス", "").strip()
            })
        elif kind == "受取":
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": inc_currency,
                "category": "送付・預入",
                "subcategory": "預入",
                "quantity": inc_qty,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": "0",
                "fee_jpy": "0",
                "memo": row.get("送付元アドレス", "").strip()
            })
        else:
            # その他の取引種別
            records.append({
                "date": dt.strftime("%Y-%m-%d %H:%M:%S"),
                "exchange": "コインチェック",
                "currency": inc_currency or dec_currency,
                "category": "その他",
                "subcategory": kind,
                "quantity": inc_qty or dec_qty,
                "rate_jpy": "0",
                "amount_jpy": "0",
                "fee": fee_qty,
                "fee_jpy": fee_qty if fee_currency == "JPY" else "0",
                "memo": memo
            })
    return records


def main():
    global FETCH_BINANCE, YEAR
    # .env の値を環境変数に反映（シェルで未設定の場合のみ）
    for k, v in load_env().items():
        os.environ.setdefault(k, v)
    FETCH_BINANCE = os.environ.get("FETCH_BINANCE", "1") != "0"
    # --year: 指定なければ前年（確定申告対象年）
    if "--year" in sys.argv:
        idx = sys.argv.index("--year")
        YEAR = int(sys.argv[idx + 1])
    else:
        YEAR = datetime.now().year - 1
    # --binance / --no-binance
    if "--binance" in sys.argv:
        FETCH_BINANCE = True
    elif "--no-binance" in sys.argv:
        FETCH_BINANCE = False

    print(f"対象年度: {YEAR}年")
    all_records = []

    Y = str(YEAR)

    def load_single_csv(name, parse_fn, encoding="utf-8-sig"):
        """ディレクトリ内のCSVを自動検出して読み込む（1ファイル想定）"""
        d = os.path.join(DATA_DIR, name)
        files = find_csv(d)
        if not files:
            print(f"  スキップ: {d} にCSVなし")
            return []
        if len(files) > 1:
            print(f"  複数CSVあり、{os.path.basename(files[0])} を使用（最終更新）")
        path = files[0]
        if encoding == "shift_jis":
            rows = read_csv_sjis(path)
        else:
            rows = read_csv_file(path, encoding)
        records = parse_fn(rows)
        print(f"  {len(records)} records ({Y})  [{os.path.basename(path)}]")
        return records

    # 1. GMOCoin
    print("Processing GMOCoin...")
    all_records.extend(load_single_csv("GMOCoin", parse_gmo))

    # 2. BitLending
    print("Processing BitLending...")
    all_records.extend(load_single_csv("BitLending", parse_bitlending))

    # 3. PBRLending
    print("Processing PBRLending...")
    all_records.extend(load_single_csv("PBRLending", parse_pbr_lending, encoding="shift_jis"))

    # 4. bitFlyer
    print("Processing bitFlyer...")
    all_records.extend(load_single_csv("bitFlyer", parse_bitflyer))

    # 5. BitPOINT
    print("Processing BitPOINT...")
    bp_dir = os.path.join(DATA_DIR, "BitPOINT")
    bp_files = find_csv(bp_dir)
    if not bp_files:
        print(f"  スキップ: {bp_dir} にCSVなし")
    else:
        if len(bp_files) > 1:
            print(f"  複数CSVあり、{os.path.basename(bp_files[0])} を使用（最終更新）")
        bp_rows = read_bitpoint_csv(bp_files[0])
        bp_records = parse_bitpoint(bp_rows)
        print(f"  {len(bp_records)} records ({Y})  [{os.path.basename(bp_files[0])}]")
        all_records.extend(bp_records)

    # 6. Coincheck
    print("Processing Coincheck...")
    all_records.extend(load_single_csv("Coincheck", parse_coincheck))

    # 7. Binance
    if FETCH_BINANCE:
        print("Processing Binance (API)...")
        try:
            env = load_env()
            api_key = env.get("API_KEY", "")
            secret = env.get("SECRET", "")
            if api_key and secret:
                binance_records = fetch_binance_data(api_key, secret)
                print(f"  {len(binance_records)} records ({Y})")
                all_records.extend(binance_records)
            else:
                print("  No API keys found, skipping Binance")
        except Exception as e:
            print(f"  Binance API error: {e}")
    else:
        print("Skipping Binance (FETCH_BINANCE=False / --no-binance)")

    # Sort by date
    all_records.sort(key=lambda r: r["date"])

    # --- Output: Full transaction CSV ---
    output_path = os.path.join(OUTPUT_DIR, f"{YEAR}_crypto_transactions.csv")
    headers = ["日付", "取引所", "通貨", "取引分類", "取引内容", "数量", "レート(円)", "金額(円)", "手数料", "手数料(円)", "備考"]

    with open(output_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for r in all_records:
            writer.writerow([
                r["date"], r["exchange"], r["currency"],
                r["category"], r["subcategory"],
                r["quantity"], r["rate_jpy"], r["amount_jpy"],
                r["fee"], r["fee_jpy"], r["memo"]
            ])

    print(f"\n=== 統合CSV出力完了: {output_path} ===")
    print(f"    合計 {len(all_records)} 件の取引")

    # --- Summary: 雑所得計算 ---
    print(f"\n=== {YEAR}年 暗号資産 雑所得サマリー ===")

    # Category-based income summary
    income_categories = {}
    for r in all_records:
        cat = r["category"]
        if cat in ("入出金", "送付・預入", "レンディング", "手数料"):
            continue  # Skip non-income categories

        try:
            amt = Decimal(r["amount_jpy"].split("(")[0]) if r["amount_jpy"] and r["amount_jpy"] != "0" else Decimal(0)
        except:
            amt = Decimal(0)

        key = f"{r['exchange']}_{cat}"
        if key not in income_categories:
            income_categories[key] = {"exchange": r["exchange"], "category": cat, "amount_jpy": Decimal(0), "count": 0}
        income_categories[key]["amount_jpy"] += amt
        income_categories[key]["count"] += 1

    # Output summary CSV
    summary_path = os.path.join(OUTPUT_DIR, f"{YEAR}_crypto_summary.csv")
    with open(summary_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["取引所", "取引分類", "件数", "合計金額(円)"])
        total_income = Decimal(0)
        for key in sorted(income_categories.keys()):
            item = income_categories[key]
            writer.writerow([item["exchange"], item["category"], item["count"], str(item["amount_jpy"])])
            print(f"  {item['exchange']:20s} {item['category']:20s} {item['count']:3d}件  {item['amount_jpy']:>15,}円")
            total_income += item["amount_jpy"]

        writer.writerow(["", "合計", "", str(total_income)])
        print(f"\n  {'合計':42s} {total_income:>15,}円")

    print(f"\n=== サマリーCSV出力完了: {summary_path} ===")

    # Note about limitations
    print("\n【注意事項】")
    print("  ・売買損益の計算には総平均法または移動平均法での取得単価計算が必要です")
    print("  ・このCSVは取引履歴の統合であり、損益計算は含まれていません")
    print("  ・レンディング/ステーキング報酬は受取時の時価で雑所得に計上されます")
    print("  ・Binanceの取引がJPY建てでない場合、別途円換算が必要です")
    print("  ・正確な申告にはCryptact等の損益計算ツールまたは税理士への相談を推奨します")


if __name__ == "__main__":
    main()
