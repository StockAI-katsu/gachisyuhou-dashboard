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

def load_stock_master():
  """Googleスプレッドシート（CSV URL）から、コード、銘柄名、市場、セクターの辞書を作成する"""
  name_dict = {}
  market_dict = {}
  sector_dict = {}

  sheet_url = "https://docs.google.com/spreadsheets/d/11qtf0L__69eW3b_jNfXMz6L5uKcZWRZnUtKeVFuirrE/export?format=csv&gid=0"

  try:
    df_ex = pd.read_csv(sheet_url, dtype=str)

    # --- ▼ ETF・ETN をマスターの段階で完全に除外する処理 ▼ ---
    if "市場・商品区分" in df_ex.columns:
      df_ex = df_ex[
          ~df_ex["市場・商品区分"].str.contains("ETF|ETN", na=False)
      ]

    for _, row in df_ex.iterrows():
      code_raw = str(row.get("コード", "")).replace(".T", "").strip()
      if not code_raw or code_raw.lower() == "nan" or not code_raw.isdigit():
        continue
      code = code_raw.zfill(4)

      name = str(row.get("銘柄名", "")).strip()
      market_raw = str(row.get("市場・商品区分", "プライム"))
      sector_raw = str(row.get("33業種区分", "")).strip()

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

      if sector_raw and sector_raw != "nan" and sector_raw != "-":
        sector_dict[code] = sector_raw
  except Exception as e:
    print(f"市場マスター読み込みエラー: {e}")

  return name_dict, market_dict, sector_dict


DAY_COLUMNS = []
for i in range(1, 11):
  DAY_COLUMNS.extend([f"{i}日目_始値", f"{i}日目_終値", f"{i}日目_騰落率(%)"])

TRACKING_COLUMNS = [
    "アラートID",
    "発生日",
    "コード",
    "銘柄名",
    "市場",
    "セクター",
    "発生日終値",
    "S高理由・テーマ",
] + DAY_COLUMNS


def apply_excel_styling(file_path):
  """Excelファイル内の全シートに対して、見栄えの良いプロ仕様のデザイン装飾を施す"""
  if not os.path.exists(file_path):
    return

  try:
    wb = openpyxl.load_workbook(file_path)

    header_fill = PatternFill(
        start_color="333333", end_color="333333", fill_type="solid"
    )
    header_font = Font(name="游ゴシック", size=10, bold=True, color="FFFFFF")

    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )

    zebra_fill = PatternFill(
        start_color="FAFAFA", end_color="FAFAFA", fill_type="solid"
    )
    s_high_fill = PatternFill(
        start_color="E5F2FF", end_color="E5F2FF", fill_type="solid"
    )
    s_high_temp_fill = PatternFill(
        start_color="FFFFFFE5", end_color="FFFFFFE5", fill_type="solid"
    )
    signal_fill = PatternFill(
        start_color="FFF8E1", end_color="FFF8E1", fill_type="solid"
    )

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
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        ws.row_dimensions[1].height = 26

      for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 20
        is_even = row_idx % 2 == 0
        row_fill = zebra_fill if is_even else None

        row_cells = [
            ws.cell(row=row_idx, column=c) for c in range(1, ws.max_column + 1)
        ]
        row_vals = [
            str(c.value) if c.value is not None else "" for c in row_cells
        ]
        row_str_all = " ".join(row_vals)

        if "S高引け" in row_str_all:
          row_fill = s_high_fill
        elif "一時S高" in row_str_all:
          row_fill = s_high_temp_fill
        elif (
            "🚀" in row_str_all
            or "🔥" in row_str_all
            or "接近中" in row_str_all
        ):
          row_fill = signal_fill

        for col_idx, cell in enumerate(row_cells, start=1):
          cell.font = data_font
          cell.border = thin_border
          if row_fill:
            cell.fill = row_fill

          val = cell.value
          if isinstance(val, (int, float)):
            if "率" in str(ws.cell(row=1, column=col_idx).value):
              cell.number_format = "#,##0.00"
            else:
              cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")
          else:
            val_str = str(val) if val is not None else ""
            if val_str.isdigit() and len(val_str) == 4:
              cell.alignment = Alignment(
                  horizontal="center", vertical="center"
              )
            elif "-" in val_str and len(val_str) <= 10:
              cell.alignment = Alignment(
                  horizontal="center", vertical="center"
              )
            else:
              cell.alignment = Alignment(horizontal="left", vertical="center")

      for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
          val_str = str(cell.value or "")
          cell_len = sum(2 if ord(c) > 256 else 1 for c in val_str)
          if cell_len > max_len:
            max_len = cell_len
        ws.column_dimensions[col_letter].width = min(
            max(max_len + 3, 11), 50
        )

    wb.save(file_path)
  except Exception as e:
    print(f"Excel装飾適用エラー ({file_path}): {e}")


def load_tracking_master():
  """追跡管理マスターデータと各種事前計算キャッシュを読み込む"""
  df_track = pd.DataFrame(columns=TRACKING_COLUMNS)
  df_sig = pd.DataFrame()
  df_ma = pd.DataFrame()
  df_brk = pd.DataFrame()

  if os.path.exists(TRACKING_EXCEL_OUT):
    try:
      with pd.ExcelFile(TRACKING_EXCEL_OUT, engine="openpyxl") as xls:
        if "追跡管理" in xls.sheet_names:
          df_track = pd.read_excel(xls, sheet_name="追跡管理", dtype=str)
        if "本日のシグナル" in xls.sheet_names:
          df_sig = pd.read_excel(xls, sheet_name="本日のシグナル", dtype=str)
        if "MAタッチ事前計算" in xls.sheet_names:
          df_ma = pd.read_excel(
              xls, sheet_name="MAタッチ事前計算", dtype=str
          )
        if "高値ブレイク事前計算" in xls.sheet_names:
          df_brk = pd.read_excel(
              xls, sheet_name="高値ブレイク事前計算", dtype=str
          )

      if not df_track.empty and "銘柄名" in df_track.columns:
        df_track = df_track[
            ~df_track["銘柄名"].str.contains(
                "イメージ情報|テイン（予測）", na=False
            )
        ]

      for col in TRACKING_COLUMNS:
        if col not in df_track.columns:
          df_track[col] = None

      numeric_cols = ["発生日終値"] + DAY_COLUMNS
      for col in numeric_cols:
        if col in df_track.columns:
          df_track[col] = pd.to_numeric(df_track[col], errors="coerce")
    except Exception as e:
      print(f"追跡マスター読み込みエラー: {e}")

  # --- ▼ キャッシュ内の古い廃止銘柄を排除する処理 ▼ ---
  if not df_track.empty and "コード" in df_track.columns:
    valid_track_codes = set(df_track["コード"].astype(str).str.zfill(4))

    if not df_ma.empty and "コード" in df_ma.columns:
      df_ma["コード"] = df_ma["コード"].astype(str).str.zfill(4)
      df_ma = df_ma[df_ma["コード"].isin(valid_track_codes)]

    if not df_brk.empty and "コード" in df_brk.columns:
      df_brk["コード"] = df_brk["コード"].astype(str).str.zfill(4)
      df_brk = df_brk[df_brk["コード"].isin(valid_track_codes)]

  return df_track, df_sig, df_ma, df_brk


def save_tracking_master(df_track, df_sig, df_ma, df_brk):
  """各種データをExcelファイルへ一括保存し、プロ仕様の装飾を自動適用する"""
  try:
    with pd.ExcelWriter(TRACKING_EXCEL_OUT, engine="openpyxl") as writer:
      df_track.to_excel(writer, sheet_name="追跡管理", index=False)
      df_sig.to_excel(writer, sheet_name="本日のシグナル", index=False)
      df_ma.to_excel(writer, sheet_name="MAタッチ事前計算", index=False)
      df_brk.to_excel(writer, sheet_name="高値ブレイク事前計算", index=False)

    apply_excel_styling(TRACKING_EXCEL_OUT)
  except Exception as e:
    print(f"追跡マスターExcel保存エラー: {e}")


def fetch_and_update_stock_data(df_track, target_indices=None):
  """yfinanceを使用して株価データを取得し、営業日ベースで正確にデータを格納する"""
  if df_track.empty:
    return df_track, 0, 0

  name_dict, market_dict, sector_dict = load_stock_master()

  # --- ▼ スプレッドシートに存在しない廃止銘柄・ETFをループ前に一括除外（軽量化） ▼ ---
  if name_dict:
    valid_codes = set(name_dict.keys())
    df_track["コード_check"] = (
        df_track["コード"].astype(str).str.replace(".T", "").str.zfill(4)
    )
    df_track = df_track[df_track["コード_check"].isin(valid_codes)].drop(
        columns=["コード_check"]
    )
    if df_track.empty:
      return df_track, 0, 0

  numeric_cols = ["発生日終値"] + DAY_COLUMNS
  for col in numeric_cols:
    if col in df_track.columns:
      df_track[col] = df_track[col].astype(float)

  indices_to_process = (
      target_indices if target_indices is not None else df_track.index
  )
  success_count = 0
  fail_count = 0

  ma_calc_rows = []
  brk_calc_rows = []

  for idx in indices_to_process:
    if idx not in df_track.index:
      continue
    row = df_track.loc[idx]
    code_raw = str(row["コード"]).replace(".T", "").strip()

    if (
        not code_raw
        or code_raw.lower() == "nan"
        or code_raw.lower() == "none"
        or not code_raw.isdigit()
    ):
      fail_count += 1
      continue

    raw_code = code_raw.zfill(4)
    yf_code = f"{raw_code}.T"

    start_date_raw = str(row["発生日"]).strip()
    start_date_str = start_date_raw.split("T")[0].split(" ")[0]

    try:
      if "-" in start_date_str:
        start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
      else:
        current_year = datetime.now().year
        start_date = datetime.strptime(
            f"{current_year}/{start_date_str}", "%Y/%m/%d"
        )
    except:
      start_date = datetime.now() - timedelta(days=5)

    df_track.at[idx, "コード"] = raw_code

    # --- セクターの保護・紐付けロジック ---
    existing_sector = str(row.get("セクター", "")).strip()
    if (
        existing_sector
        and existing_sector != "nan"
        and existing_sector != "未分類"
        and existing_sector != "-"
    ):
      sector_val = existing_sector
    else:
      sector_val = sector_dict.get(raw_code, "未分類")

    stock_name = name_dict.get(
        raw_code, row.get("銘柄名", f"銘柄_{raw_code}")
    )
    market_val = market_dict.get(raw_code, row.get("市場", "プライム"))

    df_track.at[idx, "銘柄名"] = stock_name
    df_track.at[idx, "市場"] = market_val
    df_track.at[idx, "セクター"] = sector_val

    try:
      ticker_obj = yf.Ticker(yf_code)
      start_uni = start_date - timedelta(days=15)
      end_uni = datetime.now() + timedelta(days=2)

      hist = ticker_obj.history(start=start_uni, end=end_uni, auto_adjust=True)
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
          ret_val = (
              (close_val - float(base_close)) / float(base_close)
          ) * 100

          df_track.at[idx, f"{d}日目_始値"] = open_val
          df_track.at[idx, f"{d}日目_終値"] = close_val
          df_track.at[idx, f"{d}日目_騰落率(%)"] = round(ret_val, 2)

      # --- MAタッチ用データの生成 ---
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
            "コード": raw_code,
            "銘柄名": stock_name,
            "市場": market_val,
            "セクター": sector_val,
            "現在値": round(c_price, 1),
            "出来高": vol,
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

      # --- 高値ブレイク用データの生成 ---
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

        periods_def = {
            "1ヶ月": 20,
            "3ヶ月": 60,
            "6ヶ月": 120,
            "1年": 250,
        }

        brk_row = {
            "コード": raw_code,
            "銘柄名": stock_name,
            "市場": market_val,
            "セクター": sector_val,
            "現在値": round(c_price, 1),
            "出来高": vol,
        }

        for p_name, biz_days in periods_def.items():
          sub_hist = (
              hist_brk.iloc[-(biz_days + 1) : -1]
              if len(hist_brk) > biz_days
              else hist_brk.iloc[:-1]
          )
          if sub_hist.empty:
            sub_hist = hist_brk.iloc[:-1]

          if not sub_hist.empty and "High" in sub_hist.columns:
            max_row = sub_hist.loc[sub_hist["High"].idxmax()]
            high_val = float(max_row["High"])
            high_date_str = max_row.name.strftime("%m/%d")
          else:
            high_val = c_price
            high_date_str = datetime.today().strftime("%m/%d")

          diff_pct = (
              round((c_price - high_val) / high_val * 100, 2)
              if high_val > 0
              else 0.0
          )

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
      existing_brk = existing_brk[
          ~existing_brk["コード"].isin(updated_codes_b)
      ]
      df_brk_cache = pd.concat([existing_brk, df_new_brk], ignore_index=True)
    else:
      df_brk_cache = df_new_brk
  else:
    df_brk_cache = existing_brk

  save_tracking_master(df_track, df_sig, df_ma_cache, df_brk_cache)

  return df_track, success_count, fail_count


def process_s_high_sources():
  """『【ガチ研専用】ストップ高自動取得リスト.xlsx』および外部URLから最新データを優先取得し、蓄積マスターへ統合する"""
  name_dict, market_dict, sector_dict = load_stock_master()

  df_src = pd.DataFrame()

  try:
    response = requests.get(GSHEET_EXPORT_URL, timeout=5)
    if response.status_code == 200:
      xls_net = pd.ExcelFile(io.BytesIO(response.content))
      net_sheet = xls_net.sheet_names[0]
      df_src = pd.read_excel(
          io.BytesIO(response.content), sheet_name=net_sheet, dtype=str
      )
      with pd.ExcelWriter(S_HIGH_SOURCE_FILE, engine="openpyxl") as writer:
        df_src.to_excel(writer, sheet_name="2026 ストップ高リスト", index=False)
      apply_excel_styling(S_HIGH_SOURCE_FILE)
  except Exception as e:
    print(f"外部スプレッドシート取得エラー（ローカルフォールバック利用）: {e}")

  if df_src.empty and os.path.exists(S_HIGH_SOURCE_FILE):
    try:
      xls = pd.ExcelFile(S_HIGH_SOURCE_FILE)
      df_src = pd.read_excel(
          S_HIGH_SOURCE_FILE, sheet_name=xls.sheet_names[0], dtype=str
      )
    except Exception as e:
      print(f"ローカルS高リスト読み込みエラー: {e}")

  parsed_records = []
  if not df_src.empty:
    for _, row in df_src.iterrows():
      kubun = str(row.get("区分", "")).strip()
      if "S高引け" not in kubun:
        continue

      code_raw = str(row.get("コード", "")).replace(".T", "").strip()
      if (
          not code_raw
          or code_raw.lower() == "nan"
          or code_raw.lower() == "none"
          or not code_raw.isdigit()
      ):
        continue

      raw_code = code_raw.zfill(4)
      date_raw = str(
          row.get("日付", datetime.today().strftime("%m/%d"))
      ).strip()

      current_year = datetime.now().year
      try:
        if "/" in date_raw:
          dt_obj = datetime.strptime(f"{current_year}/{date_raw}", "%Y/%m/%d")
          date_str = dt_obj.strftime("%Y-%m-%d")
        else:
          date_str = date_raw.split("T")[0]
      except:
        date_str = datetime.today().strftime("%Y-%m-%d")

      dt_parsed = datetime.strptime(date_str, "%Y-%m-%d")

      stock_name = name_dict.get(
          raw_code,
          str(row.get("銘柄", f"銘柄_{raw_code}")).strip(),
      )
      market_val = market_dict.get(raw_code, "プライム")

      sector_src = (
          str(row.get("業種", ""))
          .replace("業", "")
          .replace("・", "")
          .strip()
      )
      if not sector_src or sector_src == "nan":
        sector_src = str(row.get("セクター", "")).strip()

      if sector_src and sector_src != "nan" and sector_src != "未分類":
        sector_val = sector_src
      else:
        sector_val = sector_dict.get(raw_code, "未分類")

      theme_val = str(row.get("材料", "不明")).strip()
      if not theme_val or theme_val == "nan":
        theme_val = "不明"

      parsed_records.append({
          "code": raw_code,
          "date_dt": dt_parsed,
          "date_str": date_str,
          "name": stock_name,
          "market": market_val,
          "sector": sector_val,
          "theme": theme_val,
      })

  df_parsed = pd.DataFrame(parsed_records)
  latest_records = {}
  if not df_parsed.empty:
    df_parsed = df_parsed.sort_values(by=["code", "date_dt"])
    for code, group in df_parsed.groupby("code"):
      latest_row = group.iloc[-1]
      latest_records[code] = latest_row

  df_track, _, df_ma, df_brk = load_tracking_master()

  new_rows = []
  for code, r in latest_records.items():
    date_str = r["date_str"]
    alert_id = abs(hash(f"{code}_{date_str}")) % (10**10)

    existing_idx = None
    if not df_track.empty:
      match = df_track[
          (df_track["コード"] == code) & (df_track["発生日"] == date_str)
      ]
      if not match.empty:
        existing_idx = match.index[0]

    if existing_idx is not None:
      df_track.at[existing_idx, "銘柄名"] = r["name"]
      df_track.at[existing_idx, "市場"] = r["market"]
      if r["sector"] and r["sector"] != "未分類":
        df_track.at[existing_idx, "セクター"] = r["sector"]
      df_track.at[existing_idx, "S高理由・テーマ"] = r["theme"]
    else:
      new_record = {
          "アラートID": alert_id,
          "発生日": date_str,
          "コード": code,
          "銘柄名": r["name"],
          "市場": r["market"],
          "セクター": r["sector"],
          "発生日終値": 0.0,
          "S高理由・テーマ": r["theme"],
      }
      for col in DAY_COLUMNS:
        new_record[col] = None
      new_rows.append(new_record)

  if new_rows:
    df_new = pd.DataFrame(new_rows)
    if df_track.empty:
      df_track = df_new
    else:
      df_track = pd.concat([df_track, df_new]).drop_duplicates(
          subset=["アラートID"], keep="first"
      )

# ▼▼▼ 本日のシグナルシートへ全3手法のシグナルを体系的にまとめて転記・保存する処理 ▼▼▼
  _, _, df_ma_cache_loaded, df_brk_cache_loaded = load_tracking_master()

  # 1. 2日・3日調整シグナルの取得
  df_adj2_sig = filter_adjustment_signals(
      df_track,
      df_ma_cache_loaded,
      adjustment_days=2,
      min_dev=-100.0,
      max_dev=100.0,
  )
  df_adj3_sig = filter_adjustment_signals(
      df_track,
      df_ma_cache_loaded,
      adjustment_days=3,
      min_dev=-100.0,
      max_dev=100.0,
  )

  # 2. MAタッチシグナルの取得（接近中などの主要データ）
  df_ma_signals = pd.DataFrame()
  if not df_ma_cache_loaded.empty:
    cond_ma = (
        (df_ma_cache_loaded.get("SMA_5位置関係") == "🎯 接近中")
        | (df_ma_cache_loaded.get("EMA_5位置関係") == "🎯 接近中")
        | (df_ma_cache_loaded.get("is_target") == "True")
    )
    if isinstance(cond_ma, pd.Series):
      df_ma_signals = df_ma_cache_loaded[cond_ma].copy()

  # 3. 高値ブレイクシグナルの取得（直前・達成など）
  df_brk_signals = pd.DataFrame()
  if not df_brk_cache_loaded.empty:
    df_brk_signals = df_brk_cache_loaded.copy()

  # 1つのシートに読みやすいようセクションごとにデータを組み立てる
  signal_rows = []

  # --- セクション1：１．２・３日調整手法シグナル ---
  signal_rows.append({
      "シグナル区分": "【 １．２・３日調整手法シグナル（発生から２～３営業日目） 】",
      "コード": "",
      "銘柄名": "",
      "市場": "",
      "セクター": "",
      "詳細・ステータス": "",
  })
  has_adj = False
  if not df_adj2_sig.empty:
    has_adj = True
    for _, r in df_adj2_sig.iterrows():
      signal_rows.append({
          "シグナル区分": "2日調整",
          "コード": r.get("コード"),
          "銘柄名": r.get("銘柄名"),
          "市場": r.get("市場"),
          "セクター": r.get("セクター"),
          "詳細・ステータス": f"発生日: {r.get('発生日')} / 最新株価: {r.get('最新株価（日付）')}",
      })
  if not df_adj3_sig.empty:
    has_adj = True
    for _, r in df_adj3_sig.iterrows():
      signal_rows.append({
          "シグナル区分": "3日調整",
          "コード": r.get("コード"),
          "銘柄名": r.get("銘柄名"),
          "市場": r.get("市場"),
          "セクター": r.get("セクター"),
          "詳細・ステータス": f"発生日: {r.get('発生日')} / 最新株価: {r.get('最新株価（日付）')}",
      })
  if not has_adj:
    signal_rows.append({
        "シグナル区分": "お知らせ",
        "コード": "",
        "銘柄名": "",
        "市場": "",
        "セクター": "",
        "詳細・ステータス": "現在、条件に合致する本日のシグナルはありません。",
    })

  # 空行
  signal_rows.append({
      "シグナル区分": "",
      "コード": "",
      "銘柄名": "",
      "市場": "",
      "セクター": "",
      "詳細・ステータス": "",
  })

  # --- セクション2：２．MAタッチ手法シグナル ---
  signal_rows.append({
      "シグナル区分": "【 ２．MAタッチ手法シグナル（特化型フィルター完備） 】",
      "コード": "",
      "銘柄名": "",
      "市場": "",
      "セクター": "",
      "詳細・ステータス": "",
  })
  if not df_ma_signals.empty:
    for _, r in df_ma_signals.head(50).iterrows():
      signal_rows.append({
          "シグナル区分": "MA接近",
          "コード": r.get("コード"),
          "銘柄名": r.get("銘柄名"),
          "市場": r.get("市場"),
          "セクター": r.get("セクター"),
          "詳細・ステータス": f"現在値: {r.get('現在値')}円 / 5日線位置: {r.get('SMA_5位置関係', '-')}",
      })
  else:
    signal_rows.append({
        "シグナル区分": "お知らせ",
        "コード": "",
        "銘柄名": "",
        "市場": "",
        "セクター": "",
        "詳細・ステータス": "現在、条件に合致するMAタッチシグナルはありません。",
    })

  # 空行
  signal_rows.append({
      "シグナル区分": "",
      "コード": "",
      "銘柄名": "",
      "市場": "",
      "セクター": "",
      "詳細・ステータス": "",
  })

  # --- セクション3：３．高値ブレイク手法シグナル ---
  signal_rows.append({
      "シグナル区分": "【 ３．高値ブレイク手法シグナル 】",
      "コード": "",
      "銘柄名": "",
      "市場": "",
      "セクター": "",
      "詳細・ステータス": "",
  })
  if not df_brk_signals.empty:
    for _, r in df_brk_signals.head(50).iterrows():
      signal_rows.append({
          "シグナル区分": "高値ブレイク監視",
          "コード": r.get("コード"),
          "銘柄名": r.get("銘柄名"),
          "市場": r.get("市場"),
          "セクター": r.get("セクター"),
          "詳細・ステータス": f"現在値: {r.get('現在値')}円 / 3ヶ月ステータス: {r.get('3ヶ月_ステータス', '-')}",
      })
  else:
    signal_rows.append({
        "シグナル区分": "お知らせ",
        "コード": "",
        "銘柄名": "",
        "市場": "",
        "セクター": "",
        "詳細・ステータス": "現在、条件に合致する高値ブレイクシグナルはありません。",
    })

  df_sig_combined = pd.DataFrame(signal_rows)

  # 追跡管理マスターとともに一括保存
  save_tracking_master(df_track, df_sig_combined, df_ma, df_brk)

  return df_track, df_sig_combined


def compute_ma_status_and_dev(current_price, ma_value, threshold_pct=3.0):
  """株価と各移動平均線（SMA/EMA）の位置関係および乖離率(%)を計算"""
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
    df_track,
    df_ma_cache,
    adjustment_days=2,
    ma_type_selected="すべて",
    ma_period_selected="すべて",
    min_dev=-30.0,
    max_dev=30.0,
):
  """【2日・3日調整手法】スクリーニング関数（ユーザー様定義の営業日・連続下落ロジック完全準拠版）"""
  if df_track.empty or "発生日" not in df_track.columns:
    return pd.DataFrame()

  df = df_track.copy()
  df["発生日_dt"] = pd.to_datetime(
      df["発生日"].astype(str), format="%Y%m%d", errors="coerce"
  )
  if df["発生日_dt"].isna().all():
    return pd.DataFrame()

  unique_dates = sorted(df["発生日_dt"].dropna().unique(), reverse=True)

  target_date = None
  if adjustment_days == 2:
    if len(unique_dates) >= 1:
      target_date = unique_dates[0]
  elif adjustment_days == 3:
    if len(unique_dates) >= 2:
      target_date = unique_dates[1]
  else:
    return pd.DataFrame()

  if target_date is None:
    return pd.DataFrame()

  df = df[df["発生日_dt"] == target_date]

  results = []
  latest_date_str = datetime.today().strftime("%m/%d")

  for idx, row in df.iterrows():
    base_close = row.get("発生日終値")
    if pd.isna(base_close) or float(base_close) == 0:
      continue

    base_close = float(base_close)
    d1_close = row.get("1日目_終値")
    d2_close = row.get("2日目_終値")

    is_signal = False
    latest_price = None

    if adjustment_days == 2:
      if pd.notna(d2_close):
        d2_close = float(d2_close)
        if d2_close < base_close:
          is_signal = True
          latest_price = d2_close
      elif pd.notna(d1_close):
        d1_close = float(d1_close)
        if d1_close < base_close:
          is_signal = True
          latest_price = d1_close

    elif adjustment_days == 3:
      if pd.notna(d1_close) and pd.notna(d2_close):
        d1_close = float(d1_close)
        d2_close = float(d2_close)
        if d1_close < base_close and d2_close < d1_close:
          is_signal = True
          latest_price = d2_close

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
                "MA値": pd.to_numeric(
                    ma_row.get(f"{key}MA値"), errors="coerce"
                ),
                "乖離率": pd.to_numeric(
                    ma_row.get(f"{key}乖離率(%)"), errors="coerce"
                ),
            }

    ma_keys = [
        ("SMA", 5, "5SMA"),
        ("SMA", 25, "25SMA"),
        ("SMA", 75, "75SMA"),
        ("EMA", 5, "5EMA"),
        ("EMA", 25, "25EMA"),
        ("EMA", 75, "75EMA"),
    ]

    calc_details = {}
    for m_type, m_per, label in ma_keys:
      info = ma_info.get(f"{m_type}_{m_per}", {})
      ma_val = info.get("MA値", np.nan)
      if pd.isna(ma_val):
        ma_val = 0
      st_label, dev_val = compute_ma_status_and_dev(latest_price, ma_val)
      calc_details[f"{label}乖離率(%)"] = dev_val
      calc_details[f"{label}位置関係"] = st_label

    target_devs_to_check = []
    for m_type, m_per, label in ma_keys:
      type_match = (ma_type_selected == "すべて") or (
          ma_type_selected == m_type
      )
      period_match = (ma_period_selected == "すべて") or (
          ma_period_selected == f"{m_per}日"
      )

      if type_match and period_match:
        dev_val = calc_details.get(f"{label}乖離率(%)")
        if isinstance(dev_val, (int, float)):
          target_devs_to_check.append(dev_val)

    if target_devs_to_check:
      is_in_range = any(min_dev <= d <= max_dev for d in target_devs_to_check)
      if not is_in_range:
        continue

    raw_date = row.get("発生日", "")
    str_date = str(raw_date).strip()
    if len(str_date) == 8 and str_date.isdigit():
      formatted_date = f"{str_date[:4]}-{str_date[4:6]}-{str_date[6:]}"
    else:
      formatted_date = str_date

    res_row = {
        "発生日": formatted_date,
        "コード": code,
        "銘柄名": row.get("銘柄名", ""),
        "市場": row.get("市場", ""),
        "セクター": row.get("セクター", ""),
        "S高発生時終値": base_close,
        "最新株価（日付）": f"{latest_price:,.1f}円 ({latest_date_str})",
        "5SMA乖離率(%)": calc_details.get("5SMA乖離率(%)"),
        "5SMA位置関係": calc_details.get("5SMA位置関係"),
        "25SMA乖離率(%)": calc_details.get("25SMA乖離率(%)"),
        "25SMA位置関係": calc_details.get("25SMA位置関係"),
        "75SMA乖離率(%)": calc_details.get("75SMA乖離率(%)"),
        "75SMA位置関係": calc_details.get("75SMA位置関係"),
        "5EMA乖離率(%)": calc_details.get("5EMA乖離率(%)"),
        "5EMA位置関係": calc_details.get("5EMA位置関係"),
        "25EMA乖離率(%)": calc_details.get("25EMA乖離率(%)"),
        "25EMA位置関係": calc_details.get("25EMA位置関係"),
        "75EMA乖離率(%)": calc_details.get("75EMA乖离率(%)"),
        "75EMA位置関係": calc_details.get("75EMA位置関係"),
        "S高理由・テーマ": row.get("S高理由・テーマ", ""),
    }
    results.append(res_row)

  if not results:
    return pd.DataFrame()

  df_res = pd.DataFrame(results)
  cols_order = [
      "発生日",
      "コード",
      "銘柄名",
      "市場",
      "セクター",
      "S高発生時終値",
      "最新株価（日付）",
      "5SMA乖離率(%)",
      "5SMA位置関係",
      "25SMA乖離率(%)",
      "25SMA位置関係",
      "75SMA乖離率(%)",
      "75SMA位置関係",
      "5EMA乖離率(%)",
      "5EMA位置関係",
      "25EMA乖離率(%)",
      "25EMA位置関係",
      "75EMA乖離率(%)",
      "75EMA位置関係",
      "S高理由・テーマ",
  ]
  existing_cols = [c for c in cols_order if c in df_res.columns]
  return df_res[existing_cols]
