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
# マスターデータのロード（ネット通信一切なしの爆速仕様）
# ---------------------------------------------------------
df_track, df_sig, df_ma_cache, df_brk_cache = (
    screening_logic.load_tracking_master()
)

for col in screening_logic.TRACKING_COLUMNS:
  if col not in df_track.columns:
    df_track[col] = None

# 【超安全防衛策】キャッシュが空の場合、自動で再生成を試みる
if (df_ma_cache.empty or df_brk_cache.empty) and not df_track.empty:
  try:
    screening_logic.process_s_high_sources()
    _, _, df_ma_cache, df_brk_cache = screening_logic.load_tracking_master()
  except Exception:
    pass

# ---------------------------------------------------------
# サイドバー：操作パネル（株価更新・クラウド同期専用）
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
      idx for idx in raw_missing_indices if idx not in st.session_state.failed_indices
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
  else:
    if len(raw_missing_indices) > 0:
      st.sidebar.info(
          f"ℹ️ 未取得の {len(raw_missing_indices)} 件は上場廃止等のため自動スキップされました。（既存データは保護されています）"
      )

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


def get_current_business_day_diff(row):
  for d in range(10, 0, -1):
    val = row.get(f"{d}日目_終値")
    if pd.notna(val) and float(val) > 0:
      return d
  return 0


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
      "📈 MAタッチ手法 ＆ IFDOCO戦略",
      "🚀 高値ブレイク ＆ アラート",
      "🤖 セクター別最適戦略",
  ])

  with tab_signal:
    st.subheader(
        f"🎯 本日のエントリーシグナル一覧 (判定日:"
        f" {datetime.today().strftime('%Y-%m-%d')})"
    )
    sig_market = st.selectbox(
        "対象市場フィルター (本日のシグナル共通)",
        ["すべて", "プライム", "スタンダード", "グロース"],
        key="sig_mkt",
    )
    df_sig_filtered = (
        df_track
        if sig_market == "すべて"
        else df_track[df_track["市場"].str.contains(sig_market, na=False)]
    )

    st.markdown("### 📉 １．２・３日調整手法シグナル（発生から２～３営業日目）")
    st.markdown("#### 【２日調整】")
    adj2_rows = []
    for _, row in df_sig_filtered.iterrows():
      if get_current_business_day_diff(row) == 2:
        adj2_rows.append({
            "コード": row["コード"],
            "銘柄名": row["銘柄名"],
            "市場": row["市場"],
            "セクター": row["セクター"],
            "経過日数": "2営業日目",
            "S高理由・テーマ": row.get("S高理由・テーマ", "不明"),
            "エントリー推奨": (
                "2営業日終値(引成) [※引成保険: 3営業日始値(寄成)]"
            ),
        })
    if adj2_rows:
      st.dataframe(
          pd.DataFrame(adj2_rows).drop_duplicates(subset=["コード"]),
          width="stretch",
      )
    else:
      st.info("現在、2日調整のシグナルはありません。")

    st.markdown("#### 【３日調整】")
    adj3_rows = []
    for _, row in df_sig_filtered.iterrows():
      if get_current_business_day_diff(row) == 3:
        adj3_rows.append({
            "コード": row["コード"],
            "銘柄名": row["銘柄名"],
            "市場": row["市場"],
            "セクター": row["セクター"],
            "経過日数": "3営業日目",
            "S高理由・テーマ": row.get("S高理由・テーマ", "不明"),
            "エントリー推奨": (
                "3営業日終値(引成) [※引成保険: 4営業日始値(寄成)]"
            ),
        })
    if adj3_rows:
      st.dataframe(
          pd.DataFrame(adj3_rows).drop_duplicates(subset=["コード"]),
          width="stretch",
      )
    else:
      st.info("現在、3日調整のシグナルはありません。")

    st.markdown("---")
    st.markdown("### 📈 ２．MAタッチ手法（接近中・翌営業日狙い目）")
    if not df_ma_cache.empty and "市場" in df_ma_cache.columns:
      df_ma_f = (
          df_ma_cache
          if sig_market == "すべて"
          else df_ma_cache[
              df_ma_cache["市場"].str.contains(sig_market, na=False)
          ]
      )
      df_ma_target = df_ma_f[df_ma_f["is_target"] == "True"]
      if not df_ma_target.empty:
        df_ma_display = df_ma_target.copy()
        if "出来高" in df_ma_display.columns:
          df_ma_display["出来高"] = pd.to_numeric(
              df_ma_display["出来高"], errors="coerce"
          ).apply(lambda x: f"{int(x):,}" if pd.notna(x) else "0")
        st.dataframe(
            df_ma_display.drop(
                columns=["is_target", "MA乖離率(%)"], errors="ignore"
            ),
            width="stretch",
        )
      else:
        st.info("現在、MAタッチ接近シグナルはありません。")

    st.markdown("---")
    st.markdown("### 🚀 ３．高値ブレイク ＆ モメンタム手法（直前・初動圏）")
    if not df_brk_cache.empty and "市場" in df_brk_cache.columns:
      df_brk_f = (
          df_brk_cache
          if sig_market == "すべて"
          else df_brk_cache[
              df_brk_cache["市場"].str.contains(sig_market, na=False)
          ]
      )
      df_brk_target = df_brk_f[df_brk_f["is_signal"] == "True"]
      if not df_brk_target.empty:
        df_brk_display = df_brk_target.copy()
        if "出来高" in df_brk_display.columns:
          df_brk_display["出来高"] = pd.to_numeric(
              df_brk_display["出来高"], errors="coerce"
          ).apply(lambda x: f"{int(x):,}" if pd.notna(x) else "0")
        st.dataframe(
            df_brk_display.drop(
                columns=["is_signal", "乖離率(%)"], errors="ignore"
            ),
            width="stretch",
        )
      else:
        st.info("現在、高値ブレイク直前・初動圏のシグナルはありません。")

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
        width="stretch",
    )

  with tab_adjust:
    st.subheader("📉 2日・3日調整 手法分析 ＆ 対象シグナル（営業日ベース）")
    adj_market = st.selectbox(
        "対象市場フィルター (2日・3日調整)",
        ["すべて", "プライム", "スタンダード", "グロース"],
        key="adj_mkt",
    )
    df_adj_filtered = (
        df_track
        if adj_market == "すべて"
        else df_track[df_track["市場"].str.contains(adj_market, na=False)]
    )

    st.markdown("##### 🎯 【2日調整】対象シグナル銘柄")
    a2_list = []
    for _, row in df_adj_filtered.iterrows():
      if get_current_business_day_diff(row) == 2:
        a2_list.append({
            "コード": row["コード"],
            "銘柄名": row["銘柄名"],
            "市場": row["市場"],
            "セクター": row["セクター"],
            "S高理由": row.get("S高理由・テーマ", "不明"),
            "推奨アクション": (
                "2営業日終値(引成) [※引成保険: 3営業日始値(寄成)]"
            ),
        })
    if a2_list:
      st.dataframe(
          pd.DataFrame(a2_list).drop_duplicates(subset=["コード"]),
          width="stretch",
      )
    else:
      st.info("現在、2日調整の対象銘柄はありません。")

    st.markdown("---")
    st.markdown("##### 🎯 【3日調整】対象シグナル銘柄")
    a3_list = []
    for _, row in df_adj_filtered.iterrows():
      if get_current_business_day_diff(row) == 3:
        a3_list.append({
            "コード": row["コード"],
            "銘柄名": row["銘柄名"],
            "市場": row["市場"],
            "セクター": row["セクター"],
            "S高理由": row.get("S高理由・テーマ", "不明"),
            "推奨アクション": (
                "3営業日終値(引成) [※引成保険: 4営業日始値(寄成)]"
            ),
        })
    if a3_list:
      st.dataframe(
          pd.DataFrame(a3_list).drop_duplicates(subset=["コード"]),
          width="stretch",
      )
    else:
      st.info("現在、3日調整のシグナルはありません。")

    st.markdown("---")
    st.markdown("##### 🏢 セクター別最適戦略 (過去実績ベース)")
    sectors = [
        s
        for s in df_adj_filtered["セクター"].dropna().unique()
        if str(s).strip() != "" and str(s) != "nan"
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
                cols[s_idx].markdown(f"""
                                  **{s_idx+1}位**: エントリー: {strat['イン']} / イグジット: {strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%, サンプル:{strat['サンプル数']}件)
                                  """)

  # ---------------------------------------------------------
  # 📈 MAタッチ手法タブ
  # ---------------------------------------------------------
  with tab_ma:
    st.subheader(
        "📈 MAタッチ手法 ＆ IFDOCO戦略 (市場レンジ確認 ＆ ピンポイント数値指定)"
    )

    # 万が一df_ma_cacheが空の場合の最終防衛フォールバック
    if df_ma_cache.empty and not df_track.empty:
      try:
        screening_logic.process_s_high_sources()
        _, _, df_ma_cache, _ = screening_logic.load_tracking_master()
      except Exception:
        pass

    if not df_ma_cache.empty:
      col_m1, col_m2 = st.columns(2)
      with col_m1:
        ma_market = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="ma_mkt",
        )
      with col_m2:
        ma_status_filter = st.selectbox(
            "接近判定フィルター",
            ["すべて", "接近中のみ（狙い目）", "推移中のみ"],
            key="ma_st",
        )

      df_ma_market_filtered = df_ma_cache.copy()
      if ma_market != "すべて":
        df_ma_market_filtered = df_ma_market_filtered[
            df_ma_market_filtered["市場"].str.contains(ma_market, na=False)
        ]
      if ma_status_filter == "接近中のみ（狙い目）":
        df_ma_market_filtered = df_ma_market_filtered[
            df_ma_market_filtered["is_target"] == "True"
        ]
      elif ma_status_filter == "推移中のみ":
        df_ma_market_filtered = df_ma_market_filtered[
            df_ma_market_filtered["is_target"] == "False"
        ]

      df_ma_market_filtered["出来高"] = pd.to_numeric(
          df_ma_market_filtered["出来高"], errors="coerce"
      ).fillna(0)
      df_ma_market_filtered["MA乖離率(%)"] = pd.to_numeric(
          df_ma_market_filtered["MA乖離率(%)"], errors="coerce"
      ).fillna(0)

      if not df_ma_market_filtered.empty:
        market_min_vol = int(df_ma_market_filtered["出来高"].min())
        market_max_vol = int(df_ma_market_filtered["出来高"].max())
        if market_max_vol <= market_min_vol:
          market_max_vol = market_min_vol + 100000

        market_min_diff = float(df_ma_market_filtered["MA乖離率(%)"].min())
        market_max_diff = float(df_ma_market_filtered["MA乖離率(%)"].max())
        market_max_diff = max(market_max_diff, 5.0)
      else:
        market_min_vol, market_max_vol = 0, 30000000
        market_min_diff, market_max_diff = 0.0, 50.0

      if (
          "ma_curr_mkt" not in st.session_state
          or st.session_state.ma_curr_mkt != ma_market
      ):
        st.session_state.ma_curr_mkt = ma_market
        st.session_state.ma_ver = st.session_state.get("ma_ver", 0) + 1

      ver = st.session_state.get("ma_ver", 0)

      st.markdown("---")
      c_v1, c_v2 = st.columns(2)

      with c_v1:
        st.markdown(
            f"##### 📊 出来高範囲: {market_min_vol:,} 株 〜"
            f" {market_max_vol:,} 株 （選択市場に合わせて変動）"
        )
        sub_v1, sub_v2 = st.columns(2)
        with sub_v1:
          ma_v_min = st.number_input(
              "出来高 下限 (株)",
              min_value=0,
              max_value=2000000000,
              value=market_min_vol,
              step=10000,
              key=f"ma_v_min_{ver}",
          )
        with sub_v2:
          ma_v_max = st.number_input(
              "出来高 上限 (株)",
              min_value=0,
              max_value=2000000000,
              value=market_max_vol,
              step=10000,
              key=f"ma_v_max_{ver}",
          )

      with c_v2:
        st.markdown(
            f"##### 📉 MA乖離率範囲: {market_min_diff:.1f}% 〜"
            f" {market_max_diff:.1f}% （選択市場に合わせて変動）"
        )
        sub_d1, sub_d2 = st.columns(2)
        with sub_d1:
          ma_d_min = st.number_input(
              "MA乖離 下限 (%)",
              min_value=-100.0,
              max_value=100.0,
              value=round(market_min_diff, 1),
              step=0.1,
              key=f"ma_d_min_{ver}",
          )
        with sub_d2:
          ma_d_max = st.number_input(
              "MA乖離 上限 (%)",
              min_value=-100.0,
              max_value=100.0,
              value=round(market_max_diff, 1),
              step=0.1,
              key=f"ma_d_max_{ver}",
          )

      df_ma_view = df_ma_market_filtered[
          (df_ma_market_filtered["出来高"] >= ma_v_min)
          & (df_ma_market_filtered["出来高"] <= ma_v_max)
          & (df_ma_market_filtered["MA乖離率(%)"] >= ma_d_min)
          & (df_ma_market_filtered["MA乖離率(%)"] <= ma_d_max)
      ]

      df_ma_display = df_ma_view.copy()
      df_ma_display["出来高"] = df_ma_display["出来高"].apply(
          lambda x: f"{int(x):,}"
      )

      st.markdown("---")
      st.markdown(
          "##### 🎯 【上段】翌営業日エントリー推奨シグナル（接近中・高勝率厳選）"
      )
      df_ma_sig_only = df_ma_display[df_ma_display["is_target"] == "True"]
      if not df_ma_sig_only.empty:
        st.dataframe(
            df_ma_sig_only.drop(
                columns=["is_target", "MA乖離率(%)"], errors="ignore"
            ),
            width="stretch",
        )
      else:
        st.info("現在、MAタッチ接近シグナルはありません。")

      st.markdown("---")
      st.markdown("##### 📋 【下段】すべての監視対象銘柄一覧")
      st.dataframe(
          df_ma_display.drop(
              columns=["is_target", "MA乖離率(%)"], errors="ignore"
          ),
          width="stretch",
      )
    else:
      st.info(
          "💡 サイドバーの画面更新を実行してください。事前計算データが表示されます。"
      )

  # ---------------------------------------------------------
  # 🚀 高値ブレイク手法タブ
  # ---------------------------------------------------------
  with tab_break:
    st.subheader(
        "🚀 高値ブレイク ＆ モメンタム手法 (市場レンジ確認 ＆ ピンポイント数値指定)"
    )

    # 万が一df_brk_cacheが空の場合の最終防衛フォールバック
    if df_brk_cache.empty and not df_track.empty:
      try:
        screening_logic.process_s_high_sources()
        _, _, _, df_brk_cache = screening_logic.load_tracking_master()
      except Exception:
        pass

    if not df_brk_cache.empty:
      col_b1, col_b2 = st.columns(2)
      with col_b1:
        brk_market = st.selectbox(
            "対象市場",
            ["すべて", "プライム", "スタンダード", "グロース"],
            key="brk_mkt",
        )
      with col_b2:
        brk_status_filter = st.selectbox(
            "ステータス選択",
            [
                "すべて",
                "🔥 高値ブレイク直前（狙い目）",
                "🚀 高値ブレイク達成 (初動圏)",
            ],
            key="brk_st",
        )

      df_brk_market_filtered = df_brk_cache.copy()
      if brk_market != "すべて":
        df_brk_market_filtered = df_brk_market_filtered[
            df_brk_market_filtered["市場"].str.contains(brk_market, na=False)
        ]
      if brk_status_filter == "🔥 高値ブレイク直前（狙い目）":
        df_brk_market_filtered = df_brk_market_filtered[
            df_brk_market_filtered["ステータス"].str.contains("直前", na=False)
        ]
      elif brk_status_filter == "🚀 高値ブレイク達成 (初動圏)":
        df_brk_market_filtered = df_brk_market_filtered[
            df_brk_market_filtered["ステータス"].str.contains("初動圏", na=False)
        ]

      df_brk_market_filtered["出来高"] = pd.to_numeric(
          df_brk_market_filtered["出来高"], errors="coerce"
      ).fillna(0)
      df_brk_market_filtered["乖離率(%)"] = pd.to_numeric(
          df_brk_market_filtered["乖離率(%)"], errors="coerce"
      ).fillna(0)

      if not df_brk_market_filtered.empty:
        market_min_vb = int(df_brk_market_filtered["出来高"].min())
        market_max_vb = int(df_brk_market_filtered["出来高"].max())
        if market_max_vb <= market_min_vb:
          market_max_vb = market_min_vb + 100000

        market_min_db = float(df_brk_market_filtered["乖離率(%)"].min())
        market_max_db = float(df_brk_market_filtered["乖離率(%)"].max())
        market_min_db_limit = min(market_min_db, -100.0)
        market_max_db_limit = max(market_max_db, 15.0)
      else:
        market_min_vb, market_max_vb = 0, 30000000
        market_min_db_limit, market_max_db_limit = -100.0, 50.0

      if (
          "brk_curr_mkt" not in st.session_state
          or st.session_state.brk_curr_mkt != brk_market
      ):
        st.session_state.brk_curr_mkt = brk_market
        st.session_state.brk_ver = st.session_state.get("brk_ver", 0) + 1

      b_ver = st.session_state.get("brk_ver", 0)

      st.markdown("---")
      cb_v1, cb_v2 = st.columns(2)

      with cb_v1:
        st.markdown(
            f"##### 📊 出来高範囲: {market_min_vb:,} 株 〜"
            f" {market_max_vb:,} 株 （選択市場に合わせて変動）"
        )
        sub_bv1, sub_bv2 = st.columns(2)
        with sub_bv1:
          brk_v_min = st.number_input(
              "出来高 下限 (株)",
              min_value=0,
              max_value=2000000000,
              value=market_min_vb,
              step=10000,
              key=f"brk_v_min_{b_ver}",
          )
        with sub_bv2:
          brk_v_max = st.number_input(
              "出来高 上限 (株)",
              min_value=0,
              max_value=2000000000,
              value=market_max_vb,
              step=10000,
              key=f"brk_v_max_{b_ver}",
          )

      with cb_v2:
        st.markdown(
            f"##### 🚀 乖離率範囲: {market_min_db_limit:.1f}% 〜"
            f" {market_max_db_limit:.1f}% （選択市場に合わせて変動）"
        )
        sub_bd1, sub_bd2 = st.columns(2)
        with sub_bd1:
          brk_d_min = st.number_input(
              "乖離率 下限 (%)",
              min_value=-200.0,
              max_value=200.0,
              value=round(market_min_db_limit, 1),
              step=0.5,
              key=f"brk_d_min_{b_ver}",
          )
        with sub_bd2:
          brk_d_max = st.number_input(
              "乖離率 上限 (%)",
              min_value=-200.0,
              max_value=200.0,
              value=round(market_max_db_limit, 1),
              step=0.5,
              key=f"brk_d_max_{b_ver}",
          )

      df_brk_view = df_brk_market_filtered[
          (df_brk_market_filtered["出来高"] >= brk_v_min)
          & (df_brk_market_filtered["出来高"] <= brk_v_max)
          & (df_brk_market_filtered["乖離率(%)"] >= brk_d_min)
          & (df_brk_market_filtered["乖離率(%)"] <= brk_d_max)
      ]

      df_brk_display = df_brk_view.copy()
      df_brk_display["出来高"] = df_brk_display["出来高"].apply(
          lambda x: f"{int(x):,}"
      )

      st.markdown("---")
      st.markdown(
          "##### 🎯 【上段】翌営業日エントリー推奨シグナル（高値ブレイク直前"
          " ＆ 初動圏）"
      )
      df_brk_sig_only = df_brk_display[df_brk_display["is_signal"] == "True"]
      if not df_brk_sig_only.empty:
        st.dataframe(
            df_brk_sig_only.drop(
                columns=["is_signal", "乖離率(%)"], errors="ignore"
            ),
            width="stretch",
        )
      else:
        st.info("現在、高値ブレイク直前・初動圏のシグナルはありません。")

      st.markdown("---")
      st.markdown("##### 📋 【下段】すべての監視対象銘柄一覧")
      st.dataframe(
          df_brk_display.drop(
              columns=["is_signal", "乖離率(%)"], errors="ignore"
          ),
          width="stretch",
      )
    else:
      st.info(
          "💡 サイドバーの画面更新を実行してください。事前計算データが表示されます。"
      )

  with tab_ai:
    st.subheader("🤖 市場区分別最適戦略")
    markets = ["プライム", "スタンダード", "グロース"]
    for market in markets:
      market_df = df_track[df_track["市場"].str.contains(market, na=False)]
      if len(market_df) >= 2:
        st.markdown(f"**・市場: 【 {market} 】** (検証数: {len(market_df)}件)")
        sec_top = calculate_optimal_strategies_for_df(
            market_df, method_name=f"調整手法 ({market})"
        )
        if not sec_top.empty:
          c1, c2, c3 = st.columns(3)
          cols = [c1, c2, c3]
          for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
            with cols[s_idx]:
              cols[s_idx].markdown(f"""
                                **{s_idx+1}位**: エントリー: {strat['イン']} / イグジット: {strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%)
                                """)