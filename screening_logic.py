from datetime import datetime, timedelta
import io
import json
import os
import numpy as np
import pandas as pd
import requests
import yfinance as yf
import streamlit as st

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import gspread
from google.oauth2.service_account import Credentials

# --- 定数・設定 ---
TRACKING_EXCEL_OUT = "【ガチ研専用】ストップ高追跡リスト.xlsx"
STOCK_MASTER_FILE = "銘柄データ.xlsx"
S_HIGH_SOURCE_FILE = "【ガチ研専用】ストップ高自動取得リスト.xlsx"
GSHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1uIFIdSyAKSokBobeTaDSBAmlEC3qvBc-JTMCplCsRIg/export?format=xlsx"

# ★ご自身の『【ガチ研専用】ストップ高追跡リスト』のスプレッドシートIDを設定してください
TRACKING_GSHEET_ID = "https://docs.google.com/spreadsheets/d/1UDGhiOI7Nhv2dmZ_I5VcA_0PkLFTyD6KV-7JeJk-R90/edit?usp=sharing"


def get_gspread_client():
    """Streamlit Secretsまたはローカル設定からgspreadクライアントを取得"""
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        # 改行文字（\n）のエスケープ対策
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    elif os.path.exists("service_account.json"):
        creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
        return gspread.authorize(creds)
    return None


def load_stock_master():
    name_dict, market_dict, sector_dict = {}, {}, {}
    if os.path.exists(STOCK_MASTER_FILE):
        try:
            df_ex = pd.read_excel(STOCK_MASTER_FILE, sheet_name="銘柄データ", dtype=str)
            for _, row in df_ex.iterrows():
                code_raw = str(row.get("コード", "")).replace(".T", "").strip()
                if not code_raw or code_raw.lower() == "nan" or not code_raw.isdigit():
                    continue
                code = code_raw.zfill(4)
                name = str(row.get("銘柄名", "")).strip()
                market_raw = str(row.get("市場・商品区分", "プライム"))
                sector_raw = str(row.get("33業種区分", "未分類")).strip()

                if name and name != "nan":
                    name_dict[code] = name
                if "プライム" in market_raw:
                    market_dict[code] = "プライム"
                elif "スタンダード" in market_raw:
                    market_dict[code] = "スタンダード"
                elif "グロース" in market_raw:
                    market_dict[code] = "グロース"
                else:
                    market_dict[code] = market_raw if market_raw != "nan" else "プライム"
                sector_dict[code] = sector_raw if sector_raw and sector_raw != "nan" else "未分類"
        except Exception as e:
            print(f"市場マスター読み込みエラー: {e}")
    return name_dict, market_dict, sector_dict


DAY_COLUMNS = []
for i in range(1, 11):
    DAY_COLUMNS.extend([f"{i}日目_始値", f"{i}日目_終値", f"{i}日目_騰落率(%)"])

TRACKING_COLUMNS = [
    "アラートID", "発生日", "コード", "銘柄名", "市場", "セクター", "発生日終値", "S高理由・テーマ"
] + DAY_COLUMNS


def apply_excel_styling(file_path):
    if not os.path.exists(file_path):
        return
    try:
        wb = openpyxl.load_workbook(file_path)
        header_fill = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
        header_font = Font(name="游ゴシック", size=10, bold=True, color="FFFFFF")
        thin_border = Border(
            left=Side(style="thin", color="D3D3D3"), right=Side(style="thin", color="D3D3D3"),
            top=Side(style="thin", color="D3D3D3"), bottom=Side(style="thin", color="D3D3D3")
        )
        zebra_fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
        s_high_fill = PatternFill(start_color="E5F2FF", end_color="E5F2FF", fill_type="solid")
        s_high_temp_fill = PatternFill(start_color="FFFFFFE5", end_color="FFFFFFE5", fill_type="solid")
        signal_fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
        data_font = Font(name="游ゴシック", size=9.5)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            ws.views.sheetView[0].showGridLines = True
            if ws.max_row < 1 or ws.max_column < 1:
                continue

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.row_dimensions[1].height = 26

            for row_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[row_idx].height = 20
                row_fill = zebra_fill if row_idx % 2 == 0 else None
                row_cells = [ws.cell(row=row_idx, column=c) for c in range(1, ws.max_column + 1)]
                row_str_all = " ".join([str(c.value or "") for c in row_cells])

                if "S高引け" in row_str_all:
                    row_fill = s_high_fill
                elif "一時S高" in row_str_all:
                    row_fill = s_high_temp_fill
                elif any(x in row_str_all for x in ["🚀", "🔥", "接近中"]):
                    row_fill = signal_fill

                for col_idx, cell in enumerate(row_cells, start=1):
                    cell.font = data_font
                    cell.border = thin_border
                    if row_fill:
                        cell.fill = row_fill
                    val = cell.value
                    if isinstance(val, (int, float)):
                        cell.number_format = "#,##0.00" if "率" in str(ws.cell(row=1, column=col_idx).value) else "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                    else:
                        val_str = str(val or "")
                        if (val_str.isdigit() and len(val_str) == 4) or ("-" in val_str and len(val_str) <= 10):
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")

            for col in ws.columns:
                max_len = max(sum(2 if ord(c) > 256 else 1 for c in str(cell.value or "")) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 50)

        wb.save(file_path)
    except Exception as e:
        print(f"Excel装飾適用エラー ({file_path}): {e}")


def export_multiple_dataframes_to_styled_excel(df_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_title, df in df_dict.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_title, index=False)
            else:
                pd.DataFrame([{"情報": "該当データなし"}]).to_excel(writer, sheet_name=sheet_title, index=False)
    output.seek(0)
    
    temp_filename = "temp_multiexport_styled.xlsx"
    with open(temp_filename, "wb") as f:
        f.write(output.getvalue())
    apply_excel_styling(temp_filename)
    with open(temp_filename, "rb") as f:
        styled_data = f.read()
    if os.path.exists(temp_filename):
        os.remove(temp_filename)
    return styled_data


def load_tracking_master():
    """Googleスプレッドシート（またはフォールバックとしてローカルExcel）から最新データを読み込み"""
    df_track = pd.DataFrame(columns=TRACKING_COLUMNS)
    df_sig = pd.DataFrame()
    df_ma = pd.DataFrame()
    df_brk = pd.DataFrame()

    gc = get_gspread_client()
    if gc and TRACKING_GSHEET_ID and TRACKING_GSHEET_ID != "ここにご自身のスプレッドシートIDを貼り付け":
        try:
            sh = gc.open_by_key(TRACKING_GSHEET_ID)
            
            def read_sheet(sheet_name):
                try:
                    worksheet = sh.worksheet(sheet_name)
                    records = worksheet.get_all_records()
                    return pd.DataFrame(records)
                except:
                    return pd.DataFrame()

            df_track = read_sheet("追跡管理")
            df_sig = read_sheet("本日のシグナル")
            df_ma = read_sheet("MAタッチ事前計算")
            df_brk = read_sheet("高値ブレイク事前計算")
        except Exception as e:
            print(f"Googleスプレッドシート読み込みエラー: {e}")

    if df_track.empty and os.path.exists(TRACKING_EXCEL_OUT):
        try:
            with pd.ExcelFile(TRACKING_EXCEL_OUT, engine="openpyxl") as xls:
                if "追跡管理" in xls.sheet_names:
                    df_track = pd.read_excel(xls, sheet_name="追跡管理", dtype=str)
                if "本日のシグナル" in xls.sheet_names:
                    df_sig = pd.read_excel(xls, sheet_name="本日のシグナル", dtype=str)
                if "MAタッチ事前計算" in xls.sheet_names:
                    df_ma = pd.read_excel(xls, sheet_name="MAタッチ事前計算", dtype=str)
                if "高値ブレイク事前計算" in xls.sheet_names:
                    df_brk = pd.read_excel(xls, sheet_name="高値ブレイク事前計算", dtype=str)
        except Exception as e:
            print(f"ローカルExcel読み込みエラー: {e}")

    for col in TRACKING_COLUMNS:
        if col not in df_track.columns:
            df_track[col] = None

    return df_track, df_sig, df_ma, df_brk


def save_tracking_master(df_track, df_sig, df_ma, df_brk):
    """Googleスプレッドシート（およびローカルキャッシュ）へ一括保存"""
    gc = get_gspread_client()
    if gc and TRACKING_GSHEET_ID and TRACKING_GSHEET_ID != "ここにご自身のスプレッドシートIDを貼り付け":
        try:
            sh = gc.open_by_key(TRACKING_GSHEET_ID)
            
            def update_sheet(sheet_name, df):
                if df is None or df.empty:
                    return
                try:
                    try:
                        ws = sh.worksheet(sheet_name)
                    except:
                        ws = sh.add_worksheet(title=sheet_name, rows=100, cols=30)
                    
                    df_clean = df.fillna("").astype(str)
                    ws.clear()
                    ws.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())
                except Exception as ex:
                    print(f"シート {sheet_name} 更新エラー: {ex}")

            update_sheet("追跡管理", df_track)
            update_sheet("本日のシグナル", df_sig)
            update_sheet("MAタッチ事前計算", df_ma)
            update_sheet("高値ブレイク事前計算", df_brk)
        except Exception as e:
            print(f"Googleスプレッドシート書き込みエラー: {e}")

    try:
        with pd.ExcelWriter(TRACKING_EXCEL_OUT, engine="openpyxl") as writer:
            df_track.to_excel(writer, sheet_name="追跡管理", index=False)
            df_sig.to_excel(writer, sheet_name="本日のシグナル", index=False)
            df_ma.to_excel(writer, sheet_name="MAタッチ事前計算", index=False)
            df_brk.to_excel(writer, sheet_name="高値ブレイク事前計算", index=False)
        apply_excel_styling(TRACKING_EXCEL_OUT)
    except Exception as e:
        print(f"ローカルExcel保存エラー: {e}")


def fetch_and_update_stock_data(df_track, target_indices=None):
    if df_track.empty:
        return df_track, 0, 0

    name_dict, market_dict, sector_dict = load_stock_master()
    numeric_cols = ["発生日終値"] + DAY_COLUMNS
    for col in numeric_cols:
        if col in df_track.columns:
            df_track[col] = pd.to_numeric(df_track[col], errors="coerce")

    indices_to_process = target_indices if target_indices is not None else df_track.index
    success_count, fail_count = 0, 0
    ma_calc_rows, brk_calc_rows = [], []

    for idx in indices_to_process:
        if idx not in df_track.index:
            continue
        row = df_track.loc[idx]
        code_raw = str(row["コード"]).replace(".T", "").strip()

        if not code_raw or code_raw.lower() == "nan" or not code_raw.isdigit():
            fail_count += 1
            continue

        raw_code = code_raw.zfill(4)
        yf_code = f"{raw_code}.T"
        start_date_str = str(row["発生日"]).strip().split("T")[0].split(" ")[0]

        try:
            if "-" in start_date_str:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
            else:
                start_date = datetime.strptime(f"{datetime.now().year}/{start_date_str}", "%Y/%m/%d")
        except:
            start_date = datetime.now() - timedelta(days=5)

        df_track.at[idx, "コード"] = raw_code
        stock_name = name_dict.get(raw_code, row.get("銘柄名", f"銘柄_{raw_code}"))
        market_val = market_dict.get(raw_code, "プライム")
        sector_val = sector_dict.get(raw_code, "未分類")

        df_track.at[idx, "銘柄名"] = stock_name
        df_track.at[idx, "市場"] = market_val
        df_track.at[idx, "セクター"] = sector_val

        try:
            ticker_obj = yf.Ticker(yf_code)
            hist = ticker_obj.history(start=start_date - timedelta(days=15), end=datetime.now() + timedelta(days=2), auto_adjust=True)
            if hist is None or hist.empty:
                hist = ticker_obj.history(period="max", auto_adjust=True)

            if hist is None or hist.empty:
                fail_count += 1
                continue

            if isinstance(hist.columns, pd.MultiIndex):
                hist.columns = hist.columns.get_level_values(0)
            if hist.index.tz is not None:
                hist.index = hist.index.tz_localize(None)

            future_days = hist[hist.index >= start_date]
            if len(future_days) == 0:
                future_days = hist[hist.index >= (start_date - timedelta(days=5))]
                if len(future_days) == 0:
                    fail_count += 1
                    continue

            base_close = future_days.iloc[0]["Close"]
            if isinstance(base_close, pd.Series):
                base_close = base_close.iloc[0]

            if pd.isna(base_close) or float(base_close) == 0:
                fail_count += 1
                continue

            df_track.at[idx, "発生日終値"] = float(base_close)

            for d in range(1, 11):
                if d < len(future_days):
                    day_row = future_days.iloc[d]
                    open_val = float(day_row["Open"])
                    close_val = float(day_row["Close"])
                    ret_val = ((close_val - float(base_close)) / float(base_close)) * 100
                    df_track.at[idx, f"{d}日目_始値"] = open_val
                    df_track.at[idx, f"{d}日目_終値"] = close_val
                    df_track.at[idx, f"{d}日目_騰落率(%)"] = round(ret_val, 2)

            # MA計算
            hist_ma = ticker_obj.history(period="4mo", auto_adjust=True)
            if hist_ma is not None and not hist_ma.empty:
                if isinstance(hist_ma.columns, pd.MultiIndex):
                    hist_ma.columns = hist_ma.columns.get_level_values(0)
                c_price = float(hist_ma.iloc[-1]["Close"])
                vol = int(hist_ma.iloc[-1]["Volume"])

                ma_configs = [
                    ("SMA", 5, hist_ma["Close"].rolling(window=5).mean()),
                    ("SMA", 25, hist_ma["Close"].rolling(window=25).mean()),
                    ("SMA", 75, hist_ma["Close"].rolling(window=75).mean()),
                    ("EMA", 5, hist_ma["Close"].ewm(span=5, adjust=False).mean()),
                    ("EMA", 25, hist_ma["Close"].ewm(span=25, adjust=False).mean()),
                    ("EMA", 75, hist_ma["Close"].ewm(span=75, adjust=False).mean()),
                ]
                ma_row_dict = {
                    "コード": raw_code, "銘柄名": stock_name, "市場": market_val,
                    "セクター": sector_val, "現在値": round(c_price, 1), "出来高": vol
                }
                any_target = False
                for m_type, period, series in ma_configs:
                    if len(hist_ma) >= period:
                        val_ma = float(series.iloc[-1])
                        if pd.isna(val_ma) or val_ma == 0:
                            continue
                        st_label, dev_pct = compute_ma_status_and_dev(c_price, val_ma)
                        ma_row_dict[f"{m_type}_{period}MA値"] = round(val_ma, 1)
                        ma_row_dict[f"{m_type}_{period}乖離率(%)"] = dev_pct
                        ma_row_dict[f"{m_type}_{period}位置関係"] = st_label
                        if st_label == "🎯 接近中":
                            any_target = True
                ma_row_dict["is_target"] = str(any_target)
                ma_calc_rows.append(ma_row_dict)

            # 高値ブレイク計算
            hist_brk = None
            try:
                hist_brk = ticker_obj.history(period="1yr", auto_adjust=True)
            except:
                pass
            if hist_brk is None or hist_brk.empty or len(hist_brk) < 10:
                try:
                    hist_brk = ticker_obj.history(period="max", auto_adjust=True)
                except:
                    pass

            if hist_brk is not None and not hist_brk.empty:
                if isinstance(hist_brk.columns, pd.MultiIndex):
                    hist_brk.columns = hist_brk.columns.get_level_values(0)
                if hist_brk.index.tz is not None:
                    hist_brk.index = hist_brk.index.tz_localize(None)

                c_price = float(hist_brk.iloc[-1]["Close"])
                vol = int(hist_brk.iloc[-1]["Volume"])
                periods_def = {"1ヶ月": 20, "3ヶ月": 60, "6ヶ月": 120, "1年": 250}
                brk_row = {
                    "コード": raw_code, "銘柄名": stock_name, "市場": market_val,
                    "セクター": sector_val, "現在値": round(c_price, 1), "出来高": vol
                }

                for p_name, biz_days in periods_def.items():
                    sub_hist = hist_brk.iloc[-(biz_days + 1):-1] if len(hist_brk) > biz_days else hist_brk.iloc[:-1]
                    if sub_hist.empty:
                        sub_hist = hist_brk.iloc[:-1]
                    if not sub_hist.empty and "High" in sub_hist.columns:
                        max_row = sub_hist.loc[sub_hist["High"].idxmax()]
                        high_val = float(max_row["High"])
                        high_date_str = max_row.name.strftime("%m/%d")
                    else:
                        high_val = c_price
                        high_date_str = datetime.today().strftime("%m/%d")

                    diff_pct = round((c_price - high_val) / high_val * 100, 2) if high_val > 0 else 0.0
                    if high_val * 0.97 <= c_price < high_val:
                        status = "🔥 高値ブレイク直前 (アラート)"
                    elif high_val <= c_price <= high_val * 1.04:
                        status = "🚀 高値ブレイク達成 (初動圏)"
                    else:
                        status = "📈 監視中"

                    brk_row[f"{p_name}_高値"] = round(high_val, 1)
                    brk_row[f"{p_name}_高値日付"] = high_date_str
                    brk_row[f"{p_name}_乖離率(%)"] = diff_pct
                    brk_row[f"{p_name}_ステータス"] = status
                brk_calc_rows.append(brk_row)

            success_count += 1
        except Exception as e:
            fail_count += 1
            continue

    _, df_sig, existing_ma, existing_brk = load_tracking_master()

    if ma_calc_rows:
        df_new_ma = pd.DataFrame(ma_calc_rows)
        if not existing_ma.empty:
            updated_codes = df_new_ma["コード"].tolist()
            existing_ma = existing_ma[~existing_ma["コード"].isin(updated_codes)]
            df_ma_cache = pd.concat([existing_ma, df_new_ma], ignore_index=True)
        else:
            df_ma_cache = df_new_ma
    else:
        df_ma_cache = existing_ma

    if brk_calc_rows:
        df_new_brk = pd.DataFrame(brk_calc_rows)
        if not existing_brk.empty:
            updated_codes_b = df_new_brk["コード"].tolist()
            existing_brk = existing_brk[~existing_brk["コード"].isin(updated_codes_b)]
            df_brk_cache = pd.concat([existing_brk, df_new_brk], ignore_index=True)
        else:
            df_brk_cache = df_new_brk
    else:
        df_brk_cache = existing_brk

    save_tracking_master(df_track, df_sig, df_ma_cache, df_brk_cache)
    return df_track, success_count, fail_count


def process_s_high_sources():
    name_dict, market_dict, sector_dict = load_stock_master()
    df_src = pd.DataFrame()

    try:
        response = requests.get(GSHEET_EXPORT_URL, timeout=5)
        if response.status_code == 200:
            xls_net = pd.ExcelFile(io.BytesIO(response.content))
            net_sheet = xls_net.sheet_names[0]
            df_src = pd.read_excel(io.BytesIO(response.content), sheet_name=net_sheet, dtype=str)
    except Exception as e:
        print(f"外部スプレッドシート取得エラー: {e}")

    parsed_records = []
    if not df_src.empty:
        for _, row in df_src.iterrows():
            kubun = str(row.get("区分", "")).strip()
            if "S高引け" not in kubun:
                continue
            code_raw = str(row.get("コード", "")).replace(".T", "").strip()
            if not code_raw or code_raw.lower() == "nan" or not code_raw.isdigit():
                continue

            raw_code = code_raw.zfill(4)
            date_raw = str(row.get("日付", datetime.today().strftime("%m/%d"))).strip()
            try:
                if "/" in date_raw:
                    dt_obj = datetime.strptime(f"{datetime.now().year}/{date_raw}", "%Y/%m/%d")
                    date_str = dt_obj.strftime("%Y-%m-%d")
                else:
                    date_str = date_raw.split("T")[0]
            except:
                date_str = datetime.today().strftime("%Y-%m-%d")

            parsed_records.append({
                "code": raw_code,
                "date_dt": datetime.strptime(date_str, "%Y-%m-%d"),
                "date_str": date_str,
                "name": name_dict.get(raw_code, str(row.get("銘柄", f"銘柄_{raw_code}")).strip()),
                "market": market_dict.get(raw_code, "プライム"),
                "sector": sector_dict.get(raw_code, "未分類"),
                "theme": str(row.get("材料", "不明")).strip() or "不明",
            })

    df_parsed = pd.DataFrame(parsed_records)
    latest_records = {}
    if not df_parsed.empty:
        df_parsed = df_parsed.sort_values(by=["code", "date_dt"])
        for code, group in df_parsed.groupby("code"):
            latest_records[code] = group.iloc[-1]

    df_track, _, df_ma, df_brk = load_tracking_master()
    new_rows = []
    for code, r in latest_records.items():
        date_str = r["date_str"]
        alert_id = abs(hash(f"{code}_{date_str}")) % (10**10)

        existing_idx = None
        if not df_track.empty:
            match = df_track[(df_track["コード"] == code) & (df_track["発生日"] == date_str)]
            if not match.empty:
                existing_idx = match.index[0]

        if existing_idx is not None:
            df_track.at[existing_idx, "銘柄名"] = r["name"]
            df_track.at[existing_idx, "市場"] = r["market"]
            df_track.at[existing_idx, "セクター"] = r["sector"]
            df_track.at[existing_idx, "S高理由・テーマ"] = r["theme"]
        else:
            new_record = {
                "アラートID": alert_id, "発生日": date_str, "コード": code, "銘柄名": r["name"],
                "市場": r["market"], "セクター": r["sector"], "発生日終値": 0.0, "S高理由・テーマ": r["theme"],
            }
            for col in DAY_COLUMNS:
                new_record[col] = None
            new_rows.append(new_record)

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_track = df_new if df_track.empty else pd.concat([df_track, df_new]).drop_duplicates(subset=["アラートID"], keep="first")

    save_tracking_master(df_track, pd.DataFrame(), df_ma, df_brk)
    return df_track, pd.DataFrame()


def compute_ma_status_and_dev(current_price, ma_value, threshold_pct=3.0):
    if pd.isna(ma_value) or ma_value == 0 or pd.isna(current_price):
        return "-", "-"
    dev_pct = ((current_price - ma_value) / ma_value) * 100.0
    if current_price < ma_value:
        status = "🔻 突き抜け (下抜け)"
    elif dev_pct <= threshold_pct:
        status = "🎯 接近中"
    else:
        status = "📈 上方離脱"
    return status, round(dev_pct, 2)


def filter_adjustment_signals(
    df_track, df_ma_cache, adjustment_days=2, ma_type_selected="すべて", ma_period_selected="すべて", min_dev=-30.0, max_dev=30.0,
):
    if df_track.empty:
        return pd.DataFrame()

    results = []
    latest_date_str = datetime.today().strftime("%m/%d")

    for idx, row in df_track.iterrows():
        base_close = row.get("発生日終値")
        if pd.isna(base_close) or float(base_close) == 0:
            continue
        base_close = float(base_close)
        d1_close, d2_close, d3_close = row.get("1日目_終値"), row.get("2日目_終値"), row.get("3日目_終値")
        is_signal, latest_price = False, None

        if adjustment_days == 2:
            if pd.notna(d1_close) and pd.notna(d2_close):
                if float(d1_close) < base_close and float(d2_close) < float(d1_close):
                    is_signal, latest_price = True, float(d2_close)
        elif adjustment_days == 3:
            if pd.notna(d1_close) and pd.notna(d2_close) and pd.notna(d3_close):
                if float(d1_close) < base_close and float(d2_close) < float(d1_close) and float(d3_close) < float(d2_close):
                    is_signal, latest_price = True, float(d3_close)

        if not is_signal or latest_price is None:
            continue

        code = str(row.get("コード", "")).zfill(4)
        ma_info = {}
        if not df_ma_cache.empty and "コード" in df_ma_cache.columns:
            code_ma_df = df_ma_cache[df_ma_cache["コード"] == code]
            if not code_ma_df.empty:
                ma_row = code_ma_df.iloc[0]
                for m_t in ["SMA", "EMA"]:
                    for m_p in [5, 25, 75]:
                        key = f"{m_t}_{m_p}"
                        ma_info[key] = {
                            "MA値": pd.to_numeric(ma_row.get(f"{key}MA値"), errors="coerce"),
                            "乖離率": pd.to_numeric(ma_row.get(f"{key}乖離率(%)"), errors="coerce"),
                        }

        ma_keys = [("SMA", 5, "5SMA"), ("SMA", 25, "25SMA"), ("SMA", 75, "75SMA"), ("EMA", 5, "5EMA"), ("EMA", 25, "25EMA"), ("EMA", 75, "75EMA")]
        calc_details = {}
        for m_type, m_per, label in ma_keys:
            info = ma_info.get(f"{m_type}_{m_per}", {})
            ma_val = info.get("MA値", np.nan)
            st_label, dev_val = compute_ma_status_and_dev(latest_price, 0 if pd.isna(ma_val) else ma_val)
            calc_details[f"{label}乖離率(%)"] = dev_val
            calc_details[f"{label}位置関係"] = st_label

        target_devs_to_check = []
        for m_type, m_per, label in ma_keys:
            if (ma_type_selected in ["すべて", m_type]) and (ma_period_selected in ["すべて", f"{m_per}日"]):
                dev_val = calc_details.get(f"{label}乖離率(%)")
                if isinstance(dev_val, (int, float)):
                    target_devs_to_check.append(dev_val)

        if target_devs_to_check and not any(min_dev <= d <= max_dev for d in target_devs_to_check):
            continue

        results.append({
            "発生日": row.get("発生日", ""), "コード": code, "銘柄名": row.get("銘柄名", ""),
            "市場": row.get("市場", ""), "セクター": row.get("セクター", ""), "S高発生時終値": base_close,
            "最新株価（日付）": f"{latest_price:,.1f}円 ({latest_date_str})",
            "5SMA乖離率(%)": calc_details.get("5SMA乖離率(%)"), "5SMA位置関係": calc_details.get("5SMA位置関係"),
            "25SMA乖離率(%)": calc_details.get("25SMA乖離率(%)"), "25SMA位置関係": calc_details.get("25SMA位置関係"),
            "75SMA乖離率(%)": calc_details.get("75SMA乖離率(%)"), "75SMA位置関係": calc_details.get("75SMA位置関係"),
            "5EMA乖離率(%)": calc_details.get("5EMA乖離率(%)"), "5EMA位置関係": calc_details.get("5EMA位置関係"),
            "25EMA乖離率(%)": calc_details.get("25EMA乖離率(%)"), "25EMA位置関係": calc_details.get("25EMA位置関係"),
            "75EMA乖離率(%)": calc_details.get("75EMA乖離率(%)"), "75SMA位置関係": calc_details.get("75SMA位置関係"),
            "S高理由・テーマ": row.get("S高理由・テーマ", ""),
        })

    return pd.DataFrame(results)