from datetime import datetime
import io
import os
import pandas as pd
import screening_logic
import streamlit as st

# ページの基本設定
st.set_page_config(
    page_title="gachisyuhou - 実戦ストップ高追跡・スクリーニング管理",
    page_icon="📈",
    layout="wide",
)

st.title(
    "🚀 株式手法・スクリーニング管理ダッシュボード (`gachisyuhou` - 完全実戦版)"
)
st.markdown("---")

# ---------------------------------------------------------
# マスターデータのロード（セクター補完処理付き）
# ---------------------------------------------------------
df_track, df_sig, df_ma_cache, df_brk_cache = (
    screening_logic.load_tracking_master()
)

for col in screening_logic.TRACKING_COLUMNS:
  if col not in df_track.columns:
    df_track[col] = None

# df_track から確実なセクター辞書（マップ）を作成
sector_map = {}
if not df_track.empty and "コード" in df_track.columns:
  for _, r in df_track.iterrows():
    c = str(r.get("コード", "")).zfill(4)
    sec = str(r.get("セクター", "")).strip()
    if sec and sec != "nan" and sec != "未分類" and sec != "-":
      sector_map[c] = sec

# キャッシュ側にもセクターを補完
if not df_ma_cache.empty and "コード" in df_ma_cache.columns:
  df_ma_cache["コード"] = df_ma_cache["コード"].astype(str).str.zfill(4)
  for idx, row in df_ma_cache.iterrows():
    c = row["コード"]
    curr_sec = str(row.get("セクター", "")).strip()
    if (
        not curr_sec
        or curr_sec == "nan"
        or curr_sec == "未分類"
        or curr_sec == "-"
    ):
      if c in sector_map:
        df_ma_cache.at[idx, "セクター"] = sector_map[c]

if not df_brk_cache.empty and "コード" in df_brk_cache.columns:
  df_brk_cache["コード"] = df_brk_cache["コード"].astype(str).str.zfill(4)
  for idx, row in df_brk_cache.iterrows():
    c = row["コード"]
    curr_sec = str(row.get("セクター", "")).strip()
    if (
        not curr_sec
        or curr_sec == "nan"
        or curr_sec == "未分類"
        or curr_sec == "-"
    ):
      if c in sector_map:
        df_brk_cache.at[idx, "セクター"] = sector_map[c]

if (df_ma_cache.empty or df_brk_cache.empty) and not df_track.empty:
  try:
    screening_logic.process_s_high_sources()
    _, _, df_ma_cache, df_brk_cache = screening_logic.load_tracking_master()
  except Exception:
    pass

# ---------------------------------------------------------
# サイドバー：操作パネル
# ---------------------------------------------------------
st.sidebar.header("🎛️ コントロールパネル")

if os.path.exists(screening_logic.TRACKING_EXCEL_OUT):
  with open(screening_logic.TRACKING_EXCEL_OUT, "rb") as f:
    st.sidebar.download_button(
        label="📦 【ガチ研専用】ストップ高追跡リストをダウンロード",
        data=f,
        file_name="【ガチ研専用】ストップ高追跡リスト.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml"
            ".spreadsheetml.sheet"
        ),
    )
else:
  st.sidebar.info("💡 追跡マスターExcelファイルはまだ生成されていません。")

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔄 更新・同期アクション")
if st.sidebar.button(
    "🔄 Googleスプレッドシート取得 ＆ 全件一括株価更新", type="primary"
):
  with st.spinner(
      "Googleスプレッドシートから最新リストを取得し、全件一括で事前計算・更新中..."
  ):
    updated_track, _ = screening_logic.process_s_high_sources()
    updated_track, succ, fail = screening_logic.fetch_and_update_stock_data(
        updated_track
    )
    screening_logic.process_s_high_sources()
    st.sidebar.success(
        f"✅ 最新データ更新完了！ (成功: {succ}件 / 失敗: {fail}件)"
    )
    st.rerun()

if not df_track.empty:
  if "failed_indices" not in st.session_state:
    st.session_state.failed_indices = set()

  raw_missing_indices = df_track[
      (df_track["発生日終値"].isna()) | (df_track["発生日終値"] == 0)
  ].index

  active_missing_indices = [
      idx
      for idx in raw_missing_indices
      if idx not in st.session_state.failed_indices
  ]

  if len(active_missing_indices) > 0:
    if st.sidebar.button(
        f"⚠️ 未取得・エラー銘柄のみ再取得 ({len(active_missing_indices)}件)"
    ):
      with st.spinner(
          "有効な未取得銘柄データを再計算し、キャッシュを再構築中..."
      ):
        try:
          updated_track, _ = screening_logic.process_s_high_sources()
          updated_track, succ, fail = (
              screening_logic.fetch_and_update_stock_data(
                  updated_track, target_indices=active_missing_indices
              )
          )

          for idx in active_missing_indices:
            if idx in updated_track.index:
              val = updated_track.loc[idx, "発生日終値"]
              if pd.isna(val) or val == 0:
                st.session_state.failed_indices.add(idx)

          screening_logic.process_s_high_sources()
        except Exception as e:
          st.sidebar.error(f"再取得中にエラーが発生しました: {e}")

        st.sidebar.success(
            f"再取得＆キャッシュ再構築完了！ (成功: {succ}件 / 失敗: {fail}件)"
        )
        st.rerun()

if st.sidebar.button("🗑️ 追跡マスター台帳をリセット（初期化）"):
  if os.path.exists(screening_logic.TRACKING_EXCEL_OUT):
    os.remove(screening_logic.TRACKING_EXCEL_OUT)
  st.sidebar.warning("追跡マスター台帳をリセットしました。")
  st.rerun()


def calculate_optimal_strategies_for_df(dataframe, method_name):
  strategies = []
  if dataframe.empty or "1日目_始値" not in dataframe.columns:
    return pd.DataFrame()

  for in_day in range(1, 4):
    for out_day in range(in_day + 1, 11):
      returns = []
      for _, row in dataframe.iterrows():
        try:
          buy_val = row[f"{in_day}日目_始値"]
          sell_val = row[f"{out_day}日目_終値"]
          if pd.notna(buy_val) and pd.notna(sell_val) and buy_val > 0:
            ret = (sell_val - buy_val) / buy_val
            returns.append(ret)
        except:
          continue

      if len(returns) >= 2:
        win_count = sum(1 for r in returns if r > 0)
        win_rate = (win_count / len(returns)) * 100
        avg_ret = sum(returns) / len(returns) * 100

        in_label = (
            f"{in_day+1}営業日終値(引成) [寄成保険: {in_day+2}営業日始値]"
            if in_day == 1
            else f"{in_day+1}営業日終値(引成)"
        )
        out_label = f"{out_day}営業日終値(引成)"

        strategies.append({
            "対象手法": method_name,
            "イン": in_label,
            "アウト": out_label,
            "勝率": round(win_rate, 1),
            "平均リターン": round(avg_ret, 2),
            "サンプル数": len(returns),
        })

  strat_df = pd.DataFrame(strategies)
  if not strat_df.empty:
    strat_df = strat_df.sort_values(by=["勝率", "平均リターン"], ascending=False)
    return strat_df.head(3)
  return pd.DataFrame()


# ---------------------------------------------------------
# 共通コンポーネント：高値ブレイク用シグナル表示関数（型安全・エラー防止・セクター保護対応）
# ---------------------------------------------------------
def render_breakout_section(key_prefix="global", show_filters=True):
  st.markdown("### 🚀 高値ブレイク手法シグナル")

  if show_filters:
    cb1, cb2, cb3, cb4, cb5 = st.columns([1.2, 1, 1, 1, 1])
    with cb1:
      brk_period = st.selectbox(
          "高値スパン選択",
          ["1ヶ月", "3ヶ月", "6ヶ月", "1年"],
          key=f"{key_prefix}_period",
      )
    with cb2:
      brk_mkt = st.selectbox(
          "対象市場",
          ["すべて", "プライム", "スタンダード", "グロース"],
          key=f"{key_prefix}_mkt",
      )
    with cb3:
      brk_min_dev = st.number_input(
          "乖離率 下限(%)",
          min_value=-50.0,
          max_value=50.0,
          value=-5.0,
          step=1.0,
          key=f"{key_prefix}_min",
      )
    with cb4:
      brk_max_dev = st.number_input(
          "乖離率 上限(%)",
          min_value=-50.0,
          max_value=50.0,
          value=5.0,
          step=1.0,
          key=f"{key_prefix}_max",
      )
    with cb5:
      brk_st_sel = st.selectbox(
          "ステータス",
          [
              "すべて",
              "🔥 高値ブレイク直前 (アラート)",
              "🚀 高値ブレイク達成 (初動圏)",
              "📈 監視中",
          ],
          key=f"{key_prefix}_st",
      )
  else:
    brk_period = "3ヶ月"
    brk_mkt = "すべて"
    brk_min_dev = -5.0
    brk_max_dev = 5.0
    brk_st_sel = "すべて"

  if not df_brk_cache.empty:
    df_b = df_brk_cache.copy()
    if brk_mkt != "すべて":
      df_b = df_b[df_b["市場"].str.contains(brk_mkt, na=False)]

    col_prefix = brk_period
    dev_col = f"{col_prefix}_乖離率(%)"
    high_col = f"{col_prefix}_高値"
    date_col = f"{col_prefix}_高値日付"
    status_col = f"{col_prefix}_ステータス"

    if dev_col in df_b.columns:
      df_b["dev_val"] = pd.to_numeric(df_b[dev_col], errors="coerce")
      df_b = df_b[df_b["dev_val"].between(brk_min_dev, brk_max_dev)]

    if brk_st_sel != "すべて" and status_col in df_b.columns:
      df_b = df_b[df_b[status_col] == brk_st_sel]

    if not df_b.empty:
      view_rows = []
      for _, r in df_b.iterrows():
        c_code = str(r.get("コード", "")).zfill(4)
        c_sec = (
            sector_map.get(c_code)
            or str(r.get("セクター", "")).strip()
            or "未分類"
        )
        if c_sec in ["nan", "-", ""]:
          c_sec = "未分類"

        view_rows.append({
            "コード": c_code,
            "銘柄名": str(r.get("銘柄名", "")),
            "市場": str(r.get("市場", "")),
            "セクター": c_sec,
            "現在値": float(r.get("現在値", 0) or 0),
            f"{brk_period}高値": float(r.get(high_col, 0) or 0),
            "最高値日付": str(r.get(date_col, "-")),
            "乖離率(%)": float(r.get(dev_col, 0.0) or 0.0),
            "ステータス": str(r.get(status_col, "-")),
            "出来高": int(
                pd.to_numeric(r.get("出来高", 0), errors="coerce") or 0
            ),
        })
      df_view = pd.DataFrame(view_rows)

      if "出来高" in df_view.columns:
        df_view["出来高"] = df_view["出来高"].apply(lambda x: f"{x:,}")

      st.dataframe(
          df_view.drop_duplicates(subset=["コード"]),
          use_container_width=True,
          hide_index=True,
      )
    else:
      st.info(
          "現在、選択された条件（高値スパン:"
          f" {brk_period} / 乖離率: {brk_min_dev}%〜{brk_max_dev}%）に合致する銘柄はありません。"
      )
  else:
    st.info("事前計算データを準備中です。サイドバーの画面更新を実行してください。")


# ---------------------------------------------------------
# メイン画面：タブ構成
# ---------------------------------------------------------
if df_track.empty:
  st.info(
      "👈 左側のサイドバーにある **「🔄 Googleスプレッドシート取得 ＆"
      " 全件一括株価更新」**"
      " ボタンを押して、最新データをロードしてください。"
  )
else:
  tab_signal, tab_tracking, tab_adjust, tab_ma, tab_break, tab_ai = st.tabs([
      "🔔 本日のシグナル",
      "📋 追跡管理一覧（10営業日値動き）",
      "📉 2日・3日調整手法",
      "📈 MAタッチ手法 ＆ 接近監視",
      "🚀 高値ブレイク ＆ アラート",
      "🤖 セクター別最適戦略",
  ])

  # ---------------------------------------------------------
  # 🔔 本日のシグナル タブ (全手法の最新勝負銘柄を総覧)
  # ---------------------------------------------------------
  with tab_signal:
    st.subheader(
        "🎯 本日の厳選エントリーシグナル総合一覧 (判定日:"
        f" {datetime.today().strftime('%Y-%m-%d')})"
    )
    st.markdown("---")

    # 1. 2日・3日調整手法シグナル
    st.markdown("### 📉 １．２・３日調整手法シグナル（発生から２～３営業日目）")
    c_sig_a1, c_sig_a2 = st.columns(2)
    with c_sig_a1:
      st.markdown("#### 【２日調整】")
      adj2_sig_df = screening_logic.filter_adjustment_signals(
          df_track, df_ma_cache, adjustment_days=2, min_dev=-100.0, max_dev=100.0
      )
      if not adj2_sig_df.empty:
        st.dataframe(
            adj2_sig_df.drop_duplicates(subset=["コード"]),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("現在、2日調整のシグナルはありません。")

    with c_sig_a2:
      st.markdown("#### 【３日調整】")
      adj3_sig_df = screening_logic.filter_adjustment_signals(
          df_track, df_ma_cache, adjustment_days=3, min_dev=-100.0, max_dev=100.0
      )
      if not adj3_sig_df.empty:
        st.dataframe(
            adj3_sig_df.drop_duplicates(subset=["コード"]),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("現在、3日調整のシグナルはありません。")

    st.markdown("---")

    # 2. MAタッチ手法シグナル
    st.markdown("### 📈 ２．MAタッチ手法シグナル（特化型フィルター完備）")
    col_sig_d, col_sig_s = st.columns(2)

    with col_sig_d:
      st.markdown("#### ⚡ 【デイトレ狙い】5日移動平均線 接近シグナル")
      cd1, cd2 = st.columns(2)
      with cd1:
        sig_d_mkt = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="sig_tab_d_mkt",
        )
      with cd2:
        sig_d_ma = st.selectbox(
            "MA種別", ["すべて", "5SMA", "5EMA"], key="sig_tab_d_ma"
        )

      if not df_ma_cache.empty:
        df_d = df_ma_cache.copy()
        if sig_d_mkt != "すべて":
          df_d = df_d[df_d["市場"].str.contains(sig_d_mkt, na=False)]

        cond_sma = (
            (df_d["SMA_5位置関係"] == "🎯 接近中")
            if "SMA_5位置関係" in df_d.columns
            else False
        )
        cond_ema = (
            (df_d["EMA_5位置関係"] == "🎯 接近中")
            if "EMA_5位置関係" in df_d.columns
            else False
        )

        if sig_d_ma == "5SMA":
          df_d_res = df_d[cond_sma]
        elif sig_d_ma == "5EMA":
          df_d_res = df_d[cond_ema]
        else:
          df_d_res = df_d[cond_sma | cond_ema]

        if not df_d_res.empty:
          cols_d = [
              "コード",
              "銘柄名",
              "市場",
              "現在値",
              "出来高",
              "SMA_5乖離率(%)",
              "SMA_5位置関係",
              "EMA_5乖離率(%)",
              "EMA_5位置関係",
          ]
          actual_d = [c for c in cols_d if c in df_d_res.columns]
          st.dataframe(
              df_d_res[actual_d].drop_duplicates(subset=["コード"]),
              use_container_width=True,
              hide_index=True,
          )
        else:
          st.info(
              "現在、デイトレ狙い（5日線接近）のシグナルはありません。"
          )
      else:
        st.info("事前計算データを準備中です。サイドバーの画面更新を実行してください。")

    with col_sig_s:
      st.markdown("#### 🌊 【スイング狙い】25日/75日線 押し目候補シグナル")
      cs1, cs2, cs3, cs4 = st.columns(4)
      with cs1:
        sig_s_mkt = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="sig_tab_s_mkt",
        )
      with cs2:
        sig_s_period = st.selectbox(
            "移動平均線",
            ["すべて", "25SMA", "75SMA"],
            key="sig_tab_s_period",
        )
      with cs3:
        sig_s_min = st.number_input(
            "乖離率 下限(%)",
            min_value=-100.0,
            max_value=500.0,
            value=-5.0,
            step=1.0,
            key="sig_tab_s_min",
        )
      with cs4:
        sig_s_max = st.number_input(
            "乖離率 上限(%)",
            min_value=-100.0,
            max_value=500.0,
            value=10.0,
            step=1.0,
            key="sig_tab_s_max",
        )

      if not df_ma_cache.empty:
        df_s = df_ma_cache.copy()
        if sig_s_mkt != "すべて":
          df_s = df_s[df_s["市場"].str.contains(sig_s_mkt, na=False)]

        cond_25 = pd.to_numeric(
            df_s.get("SMA_25乖離率(%)"), errors="coerce"
        ).between(sig_s_min, sig_s_max)
        cond_75 = pd.to_numeric(
            df_s.get("SMA_75乖離率(%)"), errors="coerce"
        ).between(sig_s_min, sig_s_max)

        if sig_s_period == "25SMA":
          df_s_res = df_s[cond_25]
        elif sig_s_period == "75SMA":
          df_s_res = df_s[cond_75]
        else:
          df_s_res = df_s[cond_25 | cond_75]

        if not df_s_res.empty:
          cols_s = [
              "コード",
              "銘柄名",
              "市場",
              "現在値",
              "出来高",
              "SMA_25乖離率(%)",
              "SMA_25位置関係",
              "SMA_75乖離率(%)",
              "SMA_75位置関係",
          ]
          actual_s = [c for c in cols_s if c in df_s_res.columns]
          st.dataframe(
              df_s_res[actual_s].drop_duplicates(subset=["コード"]),
              use_container_width=True,
              hide_index=True,
          )
        else:
          st.info(
              "現在、スイング狙いの条件に合致するシグナルはありません。"
          )
      else:
        st.info("事前計算データを準備中です。サイドバーの画面更新を実行してください。")

    st.markdown("---")

    # 3. 高値ブレイク手法シグナル
    render_breakout_section(key_prefix="sig_tab_brk", show_filters=True)

  # ---------------------------------------------------------
  # 📋 追跡管理一覧 タブ
  # ---------------------------------------------------------
  with tab_tracking:
    st.subheader("📊 追跡管理プール（全監視中銘柄・10営業日実績）")
    trk_market = st.selectbox(
        "対象市場フィルター (追跡管理一覧)",
        ["すべて", "プライム", "スタンダード", "グロース"],
        key="trk_mkt",
    )
    df_trk_filtered = (
        df_track
        if trk_market == "すべて"
        else df_track[df_track["市場"].str.contains(trk_market, na=False)]
    )
    base_cols = ["コード", "銘柄名", "市場", "セクター", "発生日", "発生日終値"]
    day_cols = [
        col
        for col in df_trk_filtered.columns
        if col not in base_cols and col != "アラートID"
    ]
    st.dataframe(
        df_trk_filtered[
            [c for c in base_cols if c in df_trk_filtered.columns] + day_cols
        ].drop_duplicates(subset=["コード"]),
        use_container_width=True,
        hide_index=True,
    )

  # ---------------------------------------------------------
  # 📉 2日・3日調整手法 タブ（3段構成）
  # ---------------------------------------------------------
  with tab_adjust:
    st.subheader("📉 2日・3日調整 手法分析 ＆ スクリーニング")

    # 【上段】本日のシグナル
    st.markdown(
        "### 🎯 【上段】本日のシグナル（営業日・連続下落厳密判定）"
    )
    c_adj_t1, c_adj_t2 = st.columns(2)
    with c_adj_t1:
      st.markdown("#### 【２日調整】")
      adj2_top_df = screening_logic.filter_adjustment_signals(
          df_track, df_ma_cache, adjustment_days=2, min_dev=-100.0, max_dev=100.0
      )
      if not adj2_top_df.empty:
        st.dataframe(
            adj2_top_df.drop_duplicates(subset=["コード"]),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("現在、2日調整のシグナルはありません。")

    with c_adj_t2:
      st.markdown("#### 【３日調整】")
      adj3_top_df = screening_logic.filter_adjustment_signals(
          df_track, df_ma_cache, adjustment_days=3, min_dev=-100.0, max_dev=100.0
      )
      if not adj3_top_df.empty:
        st.dataframe(
            adj3_top_df.drop_duplicates(subset=["コード"]),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info("現在、3日調整のシグナルはありません。")

    st.markdown("---")

    # 【中段】全監視・スクリーニング一覧
    st.markdown("### 🔍 【中段】詳細スクリーニング・全監視一覧")
    c_adj1, c_adj2, c_adj3, c_adj4, c_adj5 = st.columns([1.2, 1, 1, 1, 1])
    with c_adj1:
      adj_market = st.selectbox(
          "対象市場フィルター",
          ["すべて", "プライム", "スタンダード", "グロース"],
          key="adj_mkt",
      )
    with c_adj2:
      adj_ma_type = st.selectbox(
          "MA種別選択", ["すべて", "SMA", "EMA"], key="adj_ma_type"
      )
    with c_adj3:
      adj_ma_period = st.selectbox(
          "移動平均線選択",
          ["すべて", "5日", "25日", "75日"],
          key="adj_ma_period",
      )
    with c_adj4:
      adj_min_dev = st.number_input(
          "MA乖離率 下限 (%)",
          min_value=-100.0,
          max_value=500.0,
          value=-30.0,
          step=1.0,
          key="adj_min_dev",
      )
    with c_adj5:
      adj_max_dev = st.number_input(
          "MA乖離率 上限 (%)",
          min_value=-100.0,
          max_value=500.0,
          value=30.0,
          step=1.0,
          key="adj_max_dev",
      )

    df_adj_filtered = (
        df_track
        if adj_market == "すべて"
        else df_track[df_track["市場"].str.contains(adj_market, na=False)]
    )

    st.markdown("##### 🎯 【2日調整】厳密対象一覧")
    adj2_filtered_df = screening_logic.filter_adjustment_signals(
        df_adj_filtered,
        df_ma_cache,
        adjustment_days=2,
        ma_type_selected=adj_ma_type,
        ma_period_selected=adj_ma_period,
        min_dev=adj_min_dev,
        max_dev=adj_max_dev,
    )
    if not adj2_filtered_df.empty:
      st.dataframe(
          adj2_filtered_df.drop_duplicates(subset=["コード"]),
          use_container_width=True,
          hide_index=True,
      )
    else:
      st.info("条件を満たす2日調整銘柄はありません。")

    st.markdown("##### 🎯 【3日調整】厳密対象一覧")
    adj3_filtered_df = screening_logic.filter_adjustment_signals(
        df_adj_filtered,
        df_ma_cache,
        adjustment_days=3,
        ma_type_selected=adj_ma_type,
        ma_period_selected=adj_ma_period,
        min_dev=adj_min_dev,
        max_dev=adj_max_dev,
    )
    if not adj3_filtered_df.empty:
      st.dataframe(
          adj3_filtered_df.drop_duplicates(subset=["コード"]),
          use_container_width=True,
          hide_index=True,
      )
    else:
      st.info("条件を満たす3日調整銘柄はありません。")

    # 【下段】セクター別最適戦略
    st.markdown("---")
    st.markdown("### 🏢 【下段】セクター別最適戦略 (過去実績ベース)")
    sectors = [
        s
        for s in df_adj_filtered["セクター"].dropna().unique()
        if str(s).strip() != "" and str(s) != "nan" and str(s) != "未分類"
    ]
    if sectors:
      for sector in sectors:
        sector_df = df_adj_filtered[df_adj_filtered["セクター"] == sector]
        if len(sector_df) >= 2:
          st.markdown(
              f"**・セクター: 【 {sector} 】** (検証数: {len(sector_df)}件)"
          )
          sec_top = calculate_optimal_strategies_for_df(
              sector_df, method_name=f"調整手法 ({sector})"
          )
          if not sec_top.empty:
            c1, c2, c3 = st.columns(3)
            cols = [c1, c2, c3]
            for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
              with cols[s_idx]:
                cols[s_idx].markdown(
                    f"**{s_idx+1}位**: イン: {strat['イン']} / アウト:"
                    f" {strat['アウト']} | 勝率:**{strat['勝率']}%**"
                    f" (平均:+{strat['平均リターン']}%)"
                )

  # ---------------------------------------------------------
  # 📈 MAタッチ手法 タブ（3段構成）
  # ---------------------------------------------------------
  with tab_ma:
    st.subheader("📈 MAタッチ手法 ＆ 接近監視")

    # 【上段】本日のシグナル（特化型フィルター付き）
    st.markdown("### 🎯 【上段】本日のシグナル（特化型フィルター完備）")
    col_ma_top_d, col_ma_top_s = st.columns(2)

    with col_ma_top_d:
      st.markdown("#### ⚡ 【デイトレ狙い】5日移動平均線 接近シグナル")
      cmd1, cmd2 = st.columns(2)
      with cmd1:
        ma_top_d_mkt = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="ma_tab_d_mkt",
        )
      with cmd2:
        ma_top_d_ma = st.selectbox(
            "MA種別", ["すべて", "5SMA", "5EMA"], key="ma_tab_d_ma"
        )

      if not df_ma_cache.empty:
        df_d = df_ma_cache.copy()
        if ma_top_d_mkt != "すべて":
          df_d = df_d[df_d["市場"].str.contains(ma_top_d_mkt, na=False)]

        cond_sma = (
            (df_d["SMA_5位置関係"] == "🎯 接近中")
            if "SMA_5位置関係" in df_d.columns
            else False
        )
        cond_ema = (
            (df_d["EMA_5位置関係"] == "🎯 接近中")
            if "EMA_5位置関係" in df_d.columns
            else False
        )

        if ma_top_d_ma == "5SMA":
          df_d_res = df_d[cond_sma]
        elif ma_top_d_ma == "5EMA":
          df_d_res = df_d[cond_ema]
        else:
          df_d_res = df_d[cond_sma | cond_ema]

        if not df_d_res.empty:
          cols_d = [
              "コード",
              "銘柄名",
              "市場",
              "現在値",
              "出来高",
              "SMA_5乖離率(%)",
              "SMA_5位置関係",
              "EMA_5乖離率(%)",
              "EMA_5位置関係",
          ]
          actual_d = [c for c in cols_d if c in df_d_res.columns]
          st.dataframe(
              df_d_res[actual_d].drop_duplicates(subset=["コード"]),
              use_container_width=True,
              hide_index=True,
          )
        else:
          st.info(
              "現在、デイトレ狙い（5日線接近）のシグナルはありません。"
          )
      else:
        st.info("事前計算データを準備中です。サイドバーの画面更新を実行してください。")

    with col_ma_top_s:
      st.markdown("#### 🌊 【スイング狙い】25日/75日線 押し目候補シグナル")
      cms1, cms2, cms3, cms4 = st.columns(4)
      with cms1:
        ma_top_s_mkt = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="ma_tab_s_mkt",
        )
      with cms2:
        ma_top_s_period = st.selectbox(
            "移動平均線",
            ["すべて", "25SMA", "75SMA"],
            key="ma_tab_s_period",
        )
      with cms3:
        ma_top_s_min = st.number_input(
            "乖離率 下限(%)",
            min_value=-100.0,
            max_value=500.0,
            value=-5.0,
            step=1.0,
            key="ma_tab_s_min",
        )
      with cms4:
        ma_top_s_max = st.number_input(
            "乖離率 上限(%)",
            min_value=-100.0,
            max_value=500.0,
            value=10.0,
            step=1.0,
            key="ma_tab_s_max",
        )

      if not df_ma_cache.empty:
        df_s = df_ma_cache.copy()
        if ma_top_s_mkt != "すべて":
          df_s = df_s[df_s["市場"].str.contains(ma_top_s_mkt, na=False)]

        cond_25 = pd.to_numeric(
            df_s.get("SMA_25乖離率(%)"), errors="coerce"
        ).between(ma_top_s_min, ma_top_s_max)
        cond_75 = pd.to_numeric(
            df_s.get("SMA_75乖離率(%)"), errors="coerce"
        ).between(ma_top_s_min, ma_top_s_max)

        if ma_top_s_period == "25SMA":
          df_s_res = df_s[cond_25]
        elif ma_top_s_period == "75SMA":
          df_s_res = df_s[cond_75]
        else:
          df_s_res = df_s[cond_25 | cond_75]

        if not df_s_res.empty:
          cols_s = [
              "コード",
              "銘柄名",
              "市場",
              "現在値",
              "出来高",
              "SMA_25乖離率(%)",
              "SMA_25位置関係",
              "SMA_75乖離率(%)",
              "SMA_75位置関係",
          ]
          actual_s = [c for c in cols_s if c in df_s_res.columns]
          st.dataframe(
              df_s_res[actual_s].drop_duplicates(subset=["コード"]),
              use_container_width=True,
              hide_index=True,
          )
        else:
          st.info(
              "現在、スイング狙いの条件に合致するシグナルはありません。"
          )
      else:
        st.info("事前計算データを準備中です。サイドバーの画面更新を実行してください。")

    st.markdown("---")

    # 【中段】全監視スクリーニング（6大フィルター完備）
    st.markdown(
        "### 🔍 【中段】詳細スクリーニング（6大フィルター完備・全MA一覧）"
    )
    if not df_ma_cache.empty:
      col_m1, col_m2, col_m3, col_m4, col_m5, col_m6 = st.columns(
          [1, 1, 1, 1, 1, 1]
      )
      with col_m1:
        ma_market = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="ma_mkt",
        )
      with col_m2:
        ma_status_filter = st.selectbox(
            "接近判定",
            ["すべて", "接近中のみ（狙い目）", "推移中のみ"],
            key="ma_st",
        )
      with col_m3:
        ma_type_filter = st.selectbox(
            "MA種別選択", ["すべて", "SMA", "EMA"], key="ma_type"
        )
      with col_m4:
        ma_period_filter = st.selectbox(
            "移動平均線選択",
            ["すべて", "5日", "25日", "75日"],
            key="ma_period",
        )
      with col_m5:
        ma_d_min = st.number_input(
            "MA乖離 下限 (%)",
            min_value=-100.0,
            max_value=500.0,
            value=-30.0,
            step=1.0,
            key="ma_d_min_k",
        )
      with col_m6:
        ma_d_max = st.number_input(
            "MA乖離 上限 (%)",
            min_value=-100.0,
            max_value=500.0,
            value=30.0,
            step=1.0,
            key="ma_d_max_k",
        )

      df_ma_filtered = df_ma_cache.copy()
      if ma_market != "すべて":
        df_ma_filtered = df_ma_filtered[
            df_ma_filtered["市場"].str.contains(ma_market, na=False)
        ]

      if ma_status_filter == "接近中のみ（狙い目）":
        df_ma_filtered = df_ma_filtered[df_ma_filtered["is_target"] == "True"]
      elif ma_status_filter == "推移中のみ":
        df_ma_filtered = df_ma_filtered[df_ma_filtered["is_target"] == "False"]

      filtered_indices = []
      for idx, row in df_ma_filtered.iterrows():
        match_found = False
        for m_t in (
            ["SMA", "EMA"] if ma_type_filter == "すべて" else [ma_type_filter]
        ):
          for m_p in (
              [5, 25, 75]
              if ma_period_filter == "すべて"
              else [int(ma_period_filter.replace("日", ""))]
          ):
            dev_val = pd.to_numeric(
                row.get(f"{m_t}_{m_p}乖離率(%)"), errors="coerce"
            )
            if pd.notna(dev_val) and (ma_d_min <= dev_val <= ma_d_max):
              match_found = True
        if match_found:
          filtered_indices.append(idx)

      df_ma_view = (
          df_ma_filtered.loc[filtered_indices].copy()
          if filtered_indices
          else pd.DataFrame()
      )

      if not df_ma_view.empty:
        display_cols = [
            "コード",
            "銘柄名",
            "市場",
            "セクター",
            "現在値",
            "出来高",
            "SMA_5乖離率(%)",
            "SMA_5位置関係",
            "SMA_25乖離率(%)",
            "SMA_25位置関係",
            "SMA_75乖離率(%)",
            "SMA_75位置関係",
            "EMA_5乖離率(%)",
            "EMA_5位置関係",
            "EMA_25乖離率(%)",
            "EMA_25位置関係",
            "EMA_75乖離率(%)",
            "EMA_75位置関係",
        ]
        actual_cols = [c for c in display_cols if c in df_ma_view.columns]
        df_ma_display = df_ma_view[actual_cols].copy()
        if "出来高" in df_ma_display.columns:
          df_ma_display["出来高"] = pd.to_numeric(
              df_ma_display["出来高"], errors="coerce"
          ).apply(
              lambda x: f"{int(x):,}" if pd.notna(x) else "0"
          )

        st.dataframe(
            df_ma_display.drop_duplicates(subset=["コード"]),
            use_container_width=True,
            hide_index=True,
        )
      else:
        st.info(
            "現在、条件を満たすMAタッチ・接近中の銘柄はありません。"
        )
    else:
      st.info("💡 サイドバーの画面更新を実行してください。")

    # 【下段】セクター別最適戦略
    st.markdown("---")
    st.markdown(
        "### 🏢 【下段】セクター別最適戦略 (MAタッチ実績ベース)"
    )
    if not df_track.empty:
      sectors = [
          s
          for s in df_track["セクター"].dropna().unique()
          if str(s).strip() != "" and str(s) != "nan" and str(s) != "未分類"
      ]
      if sectors:
        for sector in sectors:
          sector_df = df_track[df_track["セクター"] == sector]
          if len(sector_df) >= 2:
            st.markdown(
                f"**・セクター: 【 {sector} 】** (検証数: {len(sector_df)}件)"
            )
            sec_top = calculate_optimal_strategies_for_df(
                sector_df, method_name=f"MAタッチ ({sector})"
            )
            if not sec_top.empty:
              c1, c2, c3 = st.columns(3)
              cols = [c1, c2, c3]
              for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
                with cols[s_idx]:
                  cols[s_idx].markdown(
                      f"**{s_idx+1}位**: イン: {strat['イン']} /"
                      f" アウト: {strat['アウト']} |"
                      f" 勝率:**{strat['勝率']}%**"
                      f" (平均:+{strat['平均リターン']}%)"
                  )

  # ---------------------------------------------------------
  # 🚀 高値ブレイク タブ（3段構成・期間・日付・上下限対応）
  # ---------------------------------------------------------
  with tab_break:
    st.subheader("🚀 高値ブレイク ＆ モメンタム手法")

    # 【上段】本日のシグナル
    render_breakout_section(key_prefix="break_tab_top", show_filters=True)
    st.markdown("---")

    # 【中段】詳細スクリーニング一覧
    st.markdown("### 🔍 【中段】詳細スクリーニング・全監視一覧")
    render_breakout_section(key_prefix="break_tab_mid", show_filters=True)

    # 【下段】セクター別最適戦略
    st.markdown("---")
    st.markdown(
        "### 🏢 【下段】セクター別最適戦略 (高値ブレイク実績ベース)"
    )
    if not df_track.empty:
      sectors = [
          s
          for s in df_track["セクター"].dropna().unique()
          if str(s).strip() != "" and str(s) != "nan" and str(s) != "未分類"
      ]
      if sectors:
        for sector in sectors:
          sector_df = df_track[df_track["セクター"] == sector]
          if len(sector_df) >= 2:
            st.markdown(
                f"**・セクター: 【 {sector} 】** (検証数: {len(sector_df)}件)"
            )
            sec_top = calculate_optimal_strategies_for_df(
                sector_df, method_name=f"高値ブレイク ({sector})"
            )
            if not sec_top.empty:
              c1, c2, c3 = st.columns(3)
              cols = [c1, c2, c3]
              for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
                with cols[s_idx]:
                  cols[s_idx].markdown(
                      f"**{s_idx+1}位**: イン: {strat['イン']} /"
                      f" アウト: {strat['アウト']} |"
                      f" 勝率:**{strat['勝率']}%**"
                      f" (平均:+{strat['平均リターン']}%)"
                  )

  # ---------------------------------------------------------
  # 🤖 セクター別最適戦略 タブ
  # ---------------------------------------------------------
  with tab_ai:
    st.subheader("🤖 市場区分別最適戦略")
    markets = ["プライム", "スタンダード", "グロース"]
    for market in markets:
      market_df = df_track[df_track["市場"].str.contains(market, na=False)]
      if len(market_df) >= 2:
        st.markdown(
            f"**・市場: 【 {market} 】** (検証数: {len(market_df)}件)"
        )
        sec_top = calculate_optimal_strategies_for_df(
            market_df, method_name=f"調整手法 ({market})"
        )
        if not sec_top.empty:
          c1, c2, c3 = st.columns(3)
          cols = [c1, c2, c3]
          for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
            with cols[s_idx]:
              cols[s_idx].markdown(
                  f"**{s_idx+1}位**: イン: {strat['イン']} / アウト:"
                  f" {strat['アウト']} | 勝率:**{strat['勝率']}%**"
                  f" (平均:+{strat['平均リターン']}%)"
              )
