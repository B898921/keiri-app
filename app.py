import streamlit as st
import time

# ページの設定
st.set_page_config(page_title="小口現金 自動精算アプリ", layout="centered")

st.title("🧾 小口現金 自動精算アプリ")
st.write("レシート画像とマスターエクセルをアップロードすると、AIが読み取ってエクセルに自動追記します。")

# --- 入力フォーム ---
st.subheader("1. 基本情報の入力")
col1, col2 = st.columns(2)
with col1:
    # プレースホルダー（例）をベラジオ太郎に変更
    name = st.text_input("精算者氏名", placeholder="例：ベラジオ太郎")
with col2:
    # デフォルト値を画像に合わせて300,000に変更
    allowance = st.number_input("今月支給額", min_value=0, step=1000, value=300000)

st.subheader("2. ファイルのアップロード")
# レシート画像のアップロード（大文字の拡張子やHEIC形式も追加して表示されない問題を解決）
uploaded_images = st.file_uploader(
    "レシート画像または動画をアップロード（複数可）", 
    accept_multiple_files=True
)

# マスターエクセルのアップロード
uploaded_excel = st.file_uploader("マスターエクセル（追記先）をアップロード", type=["xlsx", "XLSX"])

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
            # ---------------------------------------------------------
            time.sleep(3) # 処理中のダミーアニメーション
            
            st.success("処理が完了しました！")
            
            # --- ダウンロードボタンの表示 ---
            st.download_button(
                label="📥 完成したエクセルをダウンロード",
                data=uploaded_excel.getvalue(),
                file_name=f"処理済み_精算書_{name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
