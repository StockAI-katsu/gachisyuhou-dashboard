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
# マスターデータのロード
# ---------------------------------------------------------
df_track, df_sig, df_ma_cache, df_brk_cache = (
    screening_logic.load_tracking_master()
)

for col in screening_logic.TRACKING_COLUMNS:
    if col not in df_track.columns:
        df_track[col] = None

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
            label="📦 【ガチ研専用】マスター台帳をダウンロード",
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
    # 🔔 本日のシグナル タブ
    # ---------------------------------------------------------
    with tab_signal:
        c_title_sig, c_dl_sig = st.columns([3.5, 1.5])
        with c_title_sig:
            st.subheader(
                f"🎯 本日の厳選エントリーシグナル総合一覧 (判定日:"
                f" {datetime.today().strftime('%Y-%m-%d')})"
            )
        
        # 1. 2日・3日調整シグナルデータの事前用意
        sig_adj_mkt = st.selectbox("対象市場（調整シグナル共通）", ["すべて", "プライム", "スタンダード", "グロース"], key="sig_adj_mkt")
        df_sig_track = df_track if sig_adj_mkt == "すべて" else df_track[df_track["市場"].str.contains(sig_adj_mkt, na=False)]
        
        adj2_sig_df = screening_logic.filter_adjustment_signals(
            df_sig_track, df_ma_cache, adjustment_days=2, min_dev=-100.0, max_dev=100.0
        )
        adj3_sig_df = screening_logic.filter_adjustment_signals(
            df_sig_track, df_ma_cache, adjustment_days=3, min_dev=-100.0, max_dev=100.0
        )

        # MAタッチシグナルデータの用意（5日接近）
        df_d_res = pd.DataFrame()
        if not df_ma_cache.empty:
            df_d = df_ma_cache.copy()
            cond_sma = (df_d["SMA_5位置関係"] == "🎯 接近中") if "SMA_5位置関係" in df_d.columns else False
            cond_ema = (df_d["EMA_5位置関係"] == "🎯 接近中") if "EMA_5位置関係" in df_d.columns else False
            df_d_res = df_d[cond_sma | cond_ema].drop_duplicates(subset=["コード"])

        # 高値ブレイクシグナルデータの用意（3ヶ月・デフォルト）
        df_brk_sig_res = pd.DataFrame()
        if not df_brk_cache.empty:
            df_b = df_brk_cache.copy()
            if "3ヶ月_ステータス" in df_b.columns:
                df_b["dev_val"] = pd.to_numeric(df_b["3ヶ月_乖離率(%)"], errors="coerce")
                df_b_filtered = df_b[df_b["dev_val"].between(-5.0, 5.0)]
                df_brk_sig_res = df_b_filtered.drop_duplicates(subset=["コード"])

        with c_dl_sig:
            dict_sig_dfs = {
                "2日調整シグナル": adj2_sig_df,
                "3日調整シグナル": adj3_sig_df,
                "MAタッチシグナル": df_d_res,
                "高値ブレイクシグナル": df_brk_sig_res,
            }
            styled_excel_sig = screening_logic.export_multiple_dataframes_to_styled_excel(dict_sig_dfs)
            st.download_button(
                label="📥 本日のシグナル一括ダウンロード",
                data=styled_excel_sig,
                file_name=f"本日の厳選シグナル総合一覧_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="sig_master_dl_btn"
            )

        st.markdown("---")

        # 1. 2日・3日調整手法シグナル表示
        st.markdown("### 📉 １．２・３日調整手法シグナル（発生から２～３営業日目）")
        c_sig_a1, c_sig_a2 = st.columns(2)
        with c_sig_a1:
            st.markdown("#### 【２日調整】")
            if not adj2_sig_df.empty:
                st.markdown(f"**該当件数: {len(adj2_sig_df)}件**")
                st.dataframe(adj2_sig_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
            else:
                st.info("現在、2日調整のシグナルはありません。")

        with c_sig_a2:
            st.markdown("#### 【３日調整】")
            if not adj3_sig_df.empty:
                st.markdown(f"**該当件数: {len(adj3_sig_df)}件**")
                st.dataframe(adj3_sig_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
            else:
                st.info("現在、3日調整のシグナルはありません。")

        st.markdown("---")

        # 2. MAタッチ手法シグナル表示
        st.markdown("### 📈 ２．MAタッチ手法シグナル（特化型フィルター完備）")
        col_sig_d, col_sig_s = st.columns(2)

        with col_sig_d:
            st.markdown("#### ⚡ 【デイトレ狙い】5日移動平均線 接近シグナル")
            if not df_d_res.empty:
                cols_d = ["コード", "銘柄名", "市場", "現在値", "出来高", "SMA_5乖離率(%)", "SMA_5位置関係", "EMA_5乖離率(%)", "EMA_5位置関係"]
                actual_d = [c for c in cols_d if c in df_d_res.columns]
                st.dataframe(df_d_res[actual_d], use_container_width=True, hide_index=True)
            else:
                st.info("現在、デイトレ狙い（5日線接近）のシグナルはありません。")

        with col_sig_s:
            st.markdown("#### 🌊 【スイング狙い】25日/75日線 押し目候補シグナル")
            cs1, cs2, cs3, cs4 = st.columns(4)
            with cs1:
                sig_s_mkt = st.selectbox("対象市場", ["すべて", "プライム", "スタンダード", "グロース"], key="sig_tab_s_mkt")
            with cs2:
                sig_s_period = st.selectbox("移動平均線", ["すべて", "25SMA", "75SMA"], key="sig_tab_s_period")
            with cs3:
                sig_s_min = st.number_input("乖離率 下限(%)", min_value=-100.0, max_value=500.0, value=-5.0, step=1.0, key="sig_tab_s_min")
            with cs4:
                sig_s_max = st.number_input("乖離率 上限(%)", min_value=-100.0, max_value=500.0, value=10.0, step=1.0, key="sig_tab_s_max")

            if not df_ma_cache.empty:
                df_s = df_ma_cache.copy()
                if sig_s_mkt != "すべて":
                    df_s = df_s[df_s["市場"].str.contains(sig_s_mkt, na=False)]
                cond_25 = pd.to_numeric(df_s.get("SMA_25乖離率(%)"), errors="coerce").between(sig_s_min, sig_s_max)
                cond_75 = pd.to_numeric(df_s.get("SMA_75乖離率(%)"), errors="coerce").between(sig_s_min, sig_s_max)
                if sig_s_period == "25SMA":
                    df_s_res = df_s[cond_25]
                elif sig_s_period == "75SMA":
                    df_s_res = df_s[cond_75]
                else:
                    df_s_res = df_s[cond_25 | cond_75]

                if not df_s_res.empty:
                    cols_s = ["コード", "銘柄名", "市場", "現在値", "出来高", "SMA_25乖離率(%)", "SMA_25位置関係", "SMA_75乖離率(%)", "SMA_75位置関係"]
                    actual_s = [c for c in cols_s if c in df_s_res.columns]
                    st.dataframe(df_s_res[actual_s].drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
                else:
                    st.info("現在、スイング狙いの条件に合致するシグナルはありません。")

        st.markdown("---")

        # 3. 高値ブレイク手法シグナル表示
        st.markdown("### 🚀 ３．高値ブレイク手法シグナル")
        if not df_brk_sig_res.empty:
            st.dataframe(df_brk_sig_res.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
        else:
            st.info("現在、条件を満たす高値ブレイクシグナルはありません。")

    # ---------------------------------------------------------
    # 📋 追跡管理一覧 タブ
    # ---------------------------------------------------------
    with tab_tracking:
        c_title_t, c_dl_t = st.columns([3.5, 1.5])
        with c_title_t:
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
        df_trk_view = df_trk_filtered[[c for c in base_cols if c in df_trk_filtered.columns] + day_cols].drop_duplicates(subset=["コード"])

        with c_dl_t:
            dict_trk_dfs = {"追跡管理一覧": df_trk_view}
            styled_excel_t = screening_logic.export_multiple_dataframes_to_styled_excel(dict_trk_dfs)
            st.download_button(
                label="📥 追跡管理一覧 一括ダウンロード",
                data=styled_excel_t,
                file_name=f"追跡管理一覧_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="trk_master_dl_btn"
            )

        st.dataframe(df_trk_view, use_container_width=True, hide_index=True)

    # ---------------------------------------------------------
    # 📉 2日・3日調整手法 タブ
    # ---------------------------------------------------------
    with tab_adjust:
        c_title_adj, c_dl_adj = st.columns([3.5, 1.5])
        with c_title_adj:
            st.subheader("📉 2日・3日調整 手法分析 ＆ スクリーニング")

        # フィルタリング条件の取得用
        c_adj1, c_adj2, c_adj3, c_adj4, c_adj5 = st.columns([1.2, 1, 1, 1, 1])
        with c_adj1:
            adj_market = st.selectbox("対象市場フィルター", ["すべて", "プライム", "スタンダード", "グロース"], key="adj_mkt")
        with c_adj2:
            adj_ma_type = st.selectbox("MA種別選択", ["すべて", "SMA", "EMA"], key="adj_ma_type")
        with c_adj3:
            adj_ma_period = st.selectbox("移動平均線選択", ["すべて", "5日", "25日", "75日"], key="adj_ma_period")
        with c_adj4:
            adj_min_dev = st.number_input("MA乖離率 下限 (%)", min_value=-100.0, max_value=500.0, value=-30.0, step=1.0, key="adj_min_dev")
        with c_adj5:
            adj_max_dev = st.number_input("MA乖離率 上限 (%)", min_value=-100.0, max_value=500.0, value=30.0, step=1.0, key="adj_max_dev")

        df_adj_filtered = (
            df_track
            if adj_market == "すべて"
            else df_track[df_track["市場"].str.contains(adj_market, na=False)]
        )

        adj2_top_df = screening_logic.filter_adjustment_signals(df_track, df_ma_cache, adjustment_days=2, min_dev=-100.0, max_dev=100.0)
        adj3_top_df = screening_logic.filter_adjustment_signals(df_track, df_ma_cache, adjustment_days=3, min_dev=-100.0, max_dev=100.0)
        adj2_filtered_df = screening_logic.filter_adjustment_signals(df_adj_filtered, df_ma_cache, adjustment_days=2, ma_type_selected=adj_ma_type, ma_period_selected=adj_ma_period, min_dev=adj_min_dev, max_dev=adj_max_dev)
        adj3_filtered_df = screening_logic.filter_adjustment_signals(df_adj_filtered, df_ma_cache, adjustment_days=3, ma_type_selected=adj_ma_type, ma_period_selected=adj_ma_period, min_dev=adj_min_dev, max_dev=adj_max_dev)

        with c_dl_adj:
            dict_adj_dfs = {
                "2日調整シグナル": adj2_top_df.drop_duplicates(subset=["コード"]) if not adj2_top_df.empty else pd.DataFrame(),
                "3日調整シグナル": adj3_top_df.drop_duplicates(subset=["コード"]) if not adj3_top_df.empty else pd.DataFrame(),
                "2日調整スクリーニング": adj2_filtered_df.drop_duplicates(subset=["コード"]) if not adj2_filtered_df.empty else pd.DataFrame(),
                "3日調整スクリーニング": adj3_filtered_df.drop_duplicates(subset=["コード"]) if not adj3_filtered_df.empty else pd.DataFrame(),
            }
            styled_excel_adj = screening_logic.export_multiple_dataframes_to_styled_excel(dict_adj_dfs)
            st.download_button(
                label="📥 2日・3日調整一括ダウンロード",
                data=styled_excel_adj,
                file_name=f"2日3日調整手法分析_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="adj_master_dl_btn"
            )

        st.markdown("---")
        st.markdown("### 🎯 【上段】本日のシグナル（営業日・連続下落厳密判定）")
        c_adj_t1, c_adj_t2 = st.columns(2)
        with c_adj_t1:
            st.markdown("#### 【２日調整】")
            if not adj2_top_df.empty:
                st.dataframe(adj2_top_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
            else:
                st.info("現在、2日調整のシグナルはありません。")
        with c_adj_t2:
            st.markdown("#### 【３日調整】")
            if not adj3_top_df.empty:
                st.dataframe(adj3_top_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
            else:
                st.info("現在、3日調整のシグナルはありません。")

        st.markdown("---")
        st.markdown("### 🔍 【中段】詳細スクリーニング・全監視一覧")
        st.markdown("##### 🎯 【2日調整】厳密対象一覧")
        if not adj2_filtered_df.empty:
            st.dataframe(adj2_filtered_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
        else:
            st.info("条件を満たす2日調整銘柄はありません。")

        st.markdown("##### 🎯 【3日調整】厳密対象一覧")
        if not adj3_filtered_df.empty:
            st.dataframe(adj3_filtered_df.drop_duplicates(subset=["コード"]), use_container_width=True, hide_index=True)
        else:
            st.info("条件を満たす3日調整銘柄はありません。")

    # ---------------------------------------------------------
    # 📈 MAタッチ手法 タブ
    # ---------------------------------------------------------
    with tab_ma:
        c_title_ma, c_dl_ma = st.columns([3.5, 1.5])
        with c_title_ma:
            st.subheader("📈 MAタッチ手法 ＆ 接近監視")

        # フィルタリング条件取得
        col_m1, col_m2, col_m3, col_m4, col_m5, col_m6 = st.columns([1, 1, 1, 1, 1, 1])
        with col_m1:
            ma_market = st.selectbox("対象市場", ["すべて", "プライム", "スタンダード", "グロース"], key="ma_mkt")
        with col_m2:
            ma_status_filter = st.selectbox("接近判定", ["すべて", "接近中のみ（狙い目）", "推移中のみ"], key="ma_st")
        with col_m3:
            ma_type_filter = st.selectbox("MA種別選択", ["すべて", "SMA", "EMA"], key="ma_type")
        with col_m4:
            ma_period_filter = st.selectbox("移動平均線選択", ["すべて", "5日", "25日", "75日"], key="ma_period")
        with col_m5:
            ma_d_min = st.number_input("MA乖離 下限 (%)", min_value=-100.0, max_value=500.0, value=-30.0, step=1.0, key="ma_d_min_k")
        with col_m6:
            ma_d_max = st.number_input("MA乖離 上限 (%)", min_value=-100.0, max_value=500.0, value=30.0, step=1.0, key="ma_d_max_k")

        # スクリーニング結果の作成
        df_ma_filtered = df_ma_cache.copy() if not df_ma_cache.empty else pd.DataFrame()
        df_ma_view = pd.DataFrame()
        if not df_ma_filtered.empty:
            if ma_market != "すべて":
                df_ma_filtered = df_ma_filtered[df_ma_filtered["市場"].str.contains(ma_market, na=False)]
            if ma_status_filter == "接近中のみ（狙い目）":
                df_ma_filtered = df_ma_filtered[df_ma_filtered["is_target"] == "True"]
            elif ma_status_filter == "推移中のみ":
                df_ma_filtered = df_ma_filtered[df_ma_filtered["is_target"] == "False"]

            filtered_indices = []
            for idx, row in df_ma_filtered.iterrows():
                match_found = False
                for m_t in (["SMA", "EMA"] if ma_type_filter == "すべて" else [ma_type_filter]):
                    for m_p in ([5, 25, 75] if ma_period_filter == "すべて" else [int(ma_period_filter.replace("日", ""))]):
                        dev_val = pd.to_numeric(row.get(f"{m_t}_{m_p}乖離率(%)"), errors="coerce")
                        if pd.notna(dev_val) and (ma_d_min <= dev_val <= ma_d_max):
                            match_found = True
                if match_found:
                    filtered_indices.append(idx)
            if filtered_indices:
                df_ma_view = df_ma_filtered.loc[filtered_indices].drop_duplicates(subset=["コード"])

        with c_dl_ma:
            dict_ma_dfs = {"MAタッチスクリーニング": df_ma_view}
            styled_excel_ma = screening_logic.export_multiple_dataframes_to_styled_excel(dict_ma_dfs)
            st.download_button(
                label="📥 MAタッチ一括ダウンロード",
                data=styled_excel_ma,
                file_name=f"MAタッチ手法分析_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="ma_master_dl_btn"
            )

        st.markdown("---")
        st.markdown("### 🔍 【中段】詳細スクリーニング（6大フィルター完備・全MA一覧）")
        if not df_ma_view.empty:
            display_cols = [
                "コード", "銘柄名", "市場", "セクター", "現在値", "出来高",
                "SMA_5乖離率(%)", "SMA_5位置関係",
                "SMA_25乖離率(%)", "SMA_25位置関係",
                "SMA_75乖離率(%)", "SMA_75位置関係",
                "EMA_5乖離率(%)", "EMA_5位置関係",
                "EMA_25乖離率(%)", "EMA_25位置関係",
                "EMA_75乖離率(%)", "EMA_75位置関係"
            ]
            actual_cols = [c for c in display_cols if c in df_ma_view.columns]
            st.dataframe(df_ma_view[actual_cols], use_container_width=True, hide_index=True)
        else:
            st.info("現在、条件を満たすMAタッチ・接近中の銘柄はありません。")

    # ---------------------------------------------------------
    # 🚀 高値ブレイク タブ
    # ---------------------------------------------------------
    with tab_break:
        c_title_brk, c_dl_brk = st.columns([3.5, 1.5])
        with c_title_brk:
            st.subheader("🚀 高値ブレイク ＆ モメンタム手法")

        # 抽出用パラメータ
        cb1, cb2, cb3, cb4, cb5 = st.columns([1.2, 1, 1, 1, 1])
        with cb1:
            brk_period = st.selectbox("高値スパン選択", ["1ヶ月", "3ヶ月", "6ヶ月", "1年"], key="break_tab_top_period")
        with cb2:
            brk_mkt = st.selectbox("対象市場", ["すべて", "プライム", "スタンダード", "グロース"], key="break_tab_top_mkt")
        with cb3:
            brk_min_dev = st.number_input("乖離率 下限(%)", min_value=-50.0, max_value=50.0, value=-5.0, step=1.0, key="break_tab_top_min")
        with cb4:
            brk_max_dev = st.number_input("乖離率 上限(%)", min_value=-50.0, max_value=50.0, value=5.0, step=1.0, key="break_tab_top_max")
        with cb5:
            brk_st_sel = st.selectbox("ステータス", ["すべて", "🔥 高値ブレイク直前 (アラート)", "🚀 高値ブレイク達成 (初動圏)", "📈 監視中"], key="break_tab_top_st")

        df_brk_view = pd.DataFrame()
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
                    view_rows.append({
                        "コード": str(r.get("コード", "")).zfill(4),
                        "銘柄名": str(r.get("銘柄名", "")),
                        "市場": str(r.get("市場", "")),
                        "セクター": str(r.get("セクター", "")),
                        "現在値": float(r.get("現在値", 0) or 0),
                        f"{brk_period}高値": float(r.get(high_col, 0) or 0),
                        "最高値日付": str(r.get(date_col, "-")),
                        "乖離率(%)": float(r.get(dev_col, 0.0) or 0.0),
                        "ステータス": str(r.get(status_col, "-")),
                        "出来高": int(pd.to_numeric(r.get("出来高", 0), errors="coerce") or 0),
                    })
                df_brk_view = pd.DataFrame(view_rows).drop_duplicates(subset=["コード"])

        with c_dl_brk:
            dict_brk_dfs = {f"高値ブレイク_{brk_period}": df_brk_view}
            styled_excel_brk = screening_logic.export_multiple_dataframes_to_styled_excel(dict_brk_dfs)
            st.download_button(
                label="📥 高値ブレイク一括ダウンロード",
                data=styled_excel_brk,
                file_name=f"高値ブレイク手法_{brk_period}_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="brk_master_dl_btn"
            )

        st.markdown("---")
        st.markdown("### 🚀 高値ブレイク手法シグナル抽出結果")
        if not df_brk_view.empty:
            st.dataframe(df_brk_view, use_container_width=True, hide_index=True)
        else:
            st.info(f"現在、選択された条件（高値スパン: {brk_period} / 乖離率: {brk_min_dev}%〜{brk_max_dev}%）に合致する銘柄はありません。")

    # ---------------------------------------------------------
    # 🤖 セクター別最適戦略 タブ
    # ---------------------------------------------------------
    with tab_ai:
        st.subheader("🤖 市場区分別最適戦略")
        markets = ["プライム", "スタンダード", "グロース"]
        for market in markets:
            market_df = df_track[df_track["市場"].str.contains(market, na=False)]
            if len(market_df) >= 2:
                st.markdown(f"**・市場: 【 {market} 】** (検証数: {len(market_df)}件)")
                sec_top = calculate_optimal_strategies_for_df(market_df, method_name=f"調整手法 ({market})")
                if not sec_top.empty:
                    c1, c2, c3 = st.columns(3)
                    cols = [c1, c2, c3]
                    for s_idx, (_, strat) in enumerate(sec_top.iterrows()):
                        with cols[s_idx]:
                            cols[s_idx].markdown(f"**{s_idx+1}位**: イン: {strat['イン']} / アウト: {strat['アウト']} | 勝率:**{strat['勝率']}%** (平均:+{strat['平均リターン']}%)")