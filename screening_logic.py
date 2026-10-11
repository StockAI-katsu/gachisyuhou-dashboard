from datetime import datetime, timedelta
import io
import os
import pandas as pd
import requests
import yfinance as yf

# ファイルパスの定義
TRACKING_EXCEL_OUT = "【ガチ研専用】ストップ高追跡リスト.xlsx"
STOCK_MASTER_FILE = "銘柄データ.xlsx"
S_HIGH_SOURCE_FILE = "【ガチ研専用】ストップ高自動取得リスト.xlsx"
GSHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1uIFIdSyAKSokBobeTaDSBAmlEC3qvBc-JTMCplCsRIg/export?format=xlsx"


def load_stock_master():
  """『銘柄データ.xlsx』の『銘柄データ』シートから、コード、銘柄名、市場、セクターの辞書を作成する"""
  name_dict = {}
  market_dict = {}
  sector_dict = {}

  if os.path.exists(STOCK_MASTER_FILE):
    try:
      df_ex = pd.read_excel(
          STOCK_MASTER_FILE, sheet_name="銘柄データ", dtype=str
      )
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
          market_dict[code] = (
              market_raw if market_raw != "nan" else "プライム"
          )

        sector_dict[code] = (
            sector_raw if sector_raw and sector_raw != "nan" else "未分類"
        )
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
          df_ma = pd.read_excel(xls, sheet_name="MAタッチ事前計算", dtype=str)
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

  return df_track, df_sig, df_ma, df_brk


def save_tracking_master(df_track, df_sig, df_ma, df_brk):
  """各種データをExcelファイルへ一括保存・永続化する"""
  try:
    with pd.ExcelWriter(TRACKING_EXCEL_OUT, engine="openpyxl") as writer:
      df_track.to_excel(writer, sheet_name="追跡管理", index=False)
      df_sig.to_excel(writer, sheet_name="本日のシグナル", index=False)
      df_ma.to_excel(writer, sheet_name="MAタッチ事前計算", index=False)
      df_brk.to_excel(writer, sheet_name="高値ブレイク事前計算", index=False)
  except Exception as e:
    print(f"追跡マスターExcel保存エラー: {e}")


def fetch_and_update_stock_data(df_track, target_indices=None):
  """yfinanceを使用して株価データを取得し、営業日ベースで正確にデータを格納する（キャッシュ完全保護対応）"""
  if df_track.empty:
    return df_track, 0, 0

  name_dict, market_dict, sector_dict = load_stock_master()

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
    stock_name = name_dict.get(
        raw_code, row.get("銘柄名", f"銘柄_{raw_code}")
    )
    market_val = market_dict.get(raw_code, "プライム")
    sector_val = sector_dict.get(raw_code, "未分類")

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

      # --- 【事前計算】MAタッチ用データの生成 ---
      hist_ma = ticker_obj.history(period="4mo", auto_adjust=True)
      if hist_ma is not None and not hist_ma.empty:
        if isinstance(hist_ma.columns, pd.MultiIndex):
          hist_ma.columns = hist_ma.columns.get_level_values(0)
        c_price = float(hist_ma.iloc[-1]["Close"])
        vol = int(hist_ma.iloc[-1]["Volume"])
        ma5 = (
            float(hist_ma["Close"].rolling(window=5).mean().iloc[-1])
            if len(hist_ma) >= 5
            else c_price
        )
        ma25 = (
            float(hist_ma["Close"].rolling(window=25).mean().iloc[-1])
            if len(hist_ma) >= 25
            else c_price
        )
        ma75 = (
            float(hist_ma["Close"].rolling(window=75).mean().iloc[-1])
            if len(hist_ma) >= 75
            else None
        )

        diffs = {
            "5MA": abs(c_price - ma5) / ma5 * 100,
            "25MA": abs(c_price - ma25) / ma25 * 100,
        }
        if ma75 is not None:
          diffs["75MA"] = abs(c_price - ma75) / ma75 * 100

        closest_ma = min(diffs, key=diffs.get)
        closest_diff = diffs[closest_ma]
        is_approaching = closest_diff <= 3.0
        status_str = (
            f"🚀 接近中 ({closest_ma}から{closest_diff:.1f}%圏内)"
            if is_approaching
            else f"推移中 ({closest_ma}基準)"
        )
        entry_price = round(ma5 if closest_ma == "5MA" else ma25, 1) - 5

        ma_calc_rows.append({
            "コード": raw_code,
            "銘柄名": stock_name,
            "市場": market_val,
            "現在値": round(c_price, 1),
            "出来高": vol,
            "MA乖離率(%)": round(closest_diff, 2),
            "5MA": round(ma5, 1),
            "25MA": round(ma25, 1),
            "75MA": round(ma75, 1) if ma75 is not None else 0.0,
            "接近判定": status_str,
            "IFDOCO指値目安": entry_price,
            "OCO利確目安": round(entry_price * 1.12, 1),
            "OCO損切目安": round(entry_price * 0.95, 1),
            "is_target": str(is_approaching),
        })

      # --- 【事前計算】高値ブレイク用データの生成 ---
      hist_brk = ticker_obj.history(period="1mo", auto_adjust=True)
      if hist_brk is not None and len(hist_brk) >= 10:
        if isinstance(hist_brk.columns, pd.MultiIndex):
          hist_brk.columns = hist_brk.columns.get_level_values(0)
        c_price = float(hist_brk.iloc[-1]["Close"])
        vol = int(hist_brk.iloc[-1]["Volume"])
        recent_high = float(hist_brk["High"].iloc[:-1].max())
        diff_pct = round((c_price - recent_high) / recent_high * 100, 2)

        if recent_high * 0.97 <= c_price < recent_high:
          status = "🔥 高値ブレイク直前 (アラート)"
          is_sig = True
        elif recent_high <= c_price <= recent_high * 1.04:
          status = "🚀 高値ブレイク達成 (初動圏)"
          is_sig = True
        else:
          status = "📈 監視中"
          is_sig = False

        brk_calc_rows.append({
            "コード": raw_code,
            "銘柄名": stock_name,
            "市場": market_val,
            "現在値": round(c_price, 1),
            "直近高値": round(recent_high, 1),
            "出来高": vol,
            "乖離率(%)": diff_pct,
            "ステータス": status,
            "is_signal": str(is_sig),
        })

      success_count += 1
    except Exception as e:
      print(f"yfinance取得エラー ({yf_code}): {e}")
      fail_count += 1
      continue

  # 既存キャッシュをロードし、今回更新分を安全にマージする（既存データを完全に保護）
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
  """「S高引け」のみを対象に選別し、同一銘柄の連続S高は最終S高日を起点として統合する"""
  name_dict, market_dict, sector_dict = load_stock_master()

  df_src = pd.DataFrame()
  try:
    response = requests.get(GSHEET_EXPORT_URL)
    if response.status_code == 200:
      xls_net = pd.ExcelFile(io.BytesIO(response.content))
      net_sheet = xls_net.sheet_names[0]
      df_src = pd.read_excel(
          io.BytesIO(response.content), sheet_name=net_sheet, dtype=str
      )
      with pd.ExcelWriter(S_HIGH_SOURCE_FILE, engine="openpyxl") as writer:
        df_src.to_excel(writer, sheet_name="2026 ストップ高リスト", index=False)
  except Exception as e:
    print(f"Googleスプレッドシート取得エラー: {e}")

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
      date_raw = str(row.get("日付", datetime.today().strftime("%m/%d"))).strip()

      current_year = datetime.now().year
      try:
        if "/" in date_raw:
          dt_obj = datetime.strptime(
              f"{current_year}/{date_raw}", "%Y/%m/%d"
          )
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

  save_tracking_master(df_track, pd.DataFrame(), df_ma, df_brk)
  return df_track, pd.DataFrame()