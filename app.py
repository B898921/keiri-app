import streamlit as st
import time

# ページの設定
st.set_page_config(page_title="経理アシスタント（レシート自動入力）", layout="centered")

st.title("🧾 経理レシート自動入力アプリ")
st.write("レシート画像とマスターエクセルをアップロードすると、AIが読み取ってエクセルに自動追記します。")

# --- 入力フォーム ---
st.subheader("1. 基本情報の入力")
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("精算者氏名", placeholder="例：二井 龍彦")
with col2:
    allowance = st.number_input("今月支給額", min_value=0, step=1000, value=225000)

st.subheader("2. ファイルのアップロード")
# レシート画像のアップロード（複数枚対応）
uploaded_images = st.file_uploader("レシート画像または動画をアップロード（複数可）", type=["png", "jpg", "jpeg", "mp4", "mov"], accept_multiple_files=True)
# マスターエクセルのアップロード
uploaded_excel = st.file_uploader("マスターエクセル（追記先）をアップロード", type=["xlsx"])

# --- 実行ボタン ---
st.subheader("3. 処理の実行")
if st.button("データ抽出＆エクセル追記を開始"):
    if not name or not uploaded_images or not uploaded_excel:
        st.warning("氏名、レシート画像、マスターエクセルのすべてを入力・アップロードしてください。")
    else:
        with st.spinner("AIがレシートを読み取り、エクセルを処理しています..."):
            
            # ---------------------------------------------------------
            # 【ここに実際の処理コードを書きます】
            # 1. 画像をAI (Gemini API等) に送ってJSONでデータをもらう
            # 2. openpyxlで uploaded_excel を読み込み、データを追記
            # 3. output.xlsx として保存
            # ---------------------------------------------------------
            time.sleep(3) # 処理中のダミーアニメーション
            
            st.success("処理が完了しました！")
            
            # --- ダウンロードボタンの表示 ---
            # ※実際は処理されたExcelファイルの中身を渡します
            st.download_button(
                label="📥 完成したエクセルをダウンロード",
                data=uploaded_excel.getvalue(), # ※ダミーとして元のファイルを返しています
                file_name=f"処理済み_精算書_{name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
