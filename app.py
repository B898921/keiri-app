import streamlit as st
import time

# ページの設定
st.set_page_config(page_title="小口現金 自動精算アプリ", layout="centered")

# --- UIカスタマイズ（ボタンを隠す） ---
st.markdown("""
    <style>
    /* ファイルアップロード枠内のボタンを非表示にする */
    [data-testid="stFileUploader"] section button {
        display: none;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🧾 小口現金 自動精算アプリ")
st.write("レシート画像とマスターエクセルをドラッグ＆ドロップすると、AIが読み取ってエクセルに自動追記します。")

# --- 入力フォーム ---
st.subheader("1. 基本情報の入力")
col1, col2 = st.columns(2)
with col1:
    name = st.text_input("精算者氏名", placeholder="例：ベラジオ太郎")
with col2:
    allowance = st.number_input("今月支給額", min_value=0, step=1000, value=300000)

st.subheader("2. ファイルのアップロード")
# レシート画像のアップロード（種類の制限をなくし、文言を変更）
uploaded_images = st.file_uploader(
    "レシート画像または動画をドラッグ＆ドロップ（複数可）", 
    accept_multiple_files=True
)

# マスターエクセルのアップロード（種類の制限をなくし、文言を変更）
uploaded_excel = st.file_uploader("マスターエクセル（追記先）をドラッグ＆ドロップ")

# --- 実行ボタン ---
st.subheader("3. 処理の実行")
if st.button("データ抽出＆エクセル追記を開始"):
    if not name or not uploaded_images or not uploaded_excel:
        st.warning("氏名、レシート画像（動画）、マスターエクセルのすべてを入力・ドラッグ＆ドロップしてください。")
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
