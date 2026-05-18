import streamlit as st
import pandas as pd
import openpyxl
import os
import win32com.client
from datetime import datetime

st.set_page_config(page_title="自動発注システム", layout="wide")
st.title("自動発注・連絡書システム")

# =========================================================
# パスワード認証機能
# =========================================================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
        
    if not st.session_state["password_correct"]:
        st.info("システムを利用するにはパスワードを入力してください。")
        pwd = st.text_input("パスワード", type="password")
        if pwd:
            if pwd == "sk0123":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("😕 パスワードが間違っています。")
        return False
    return True

if not check_password():
    st.stop()

# =========================================================
# 【新機能】ファイルアップローダー（GitHubを開かずに更新）
# =========================================================
st.sidebar.subheader("マスターファイルの更新")
uploaded_file = st.sidebar.file_uploader(
    "エクセルファイルをアップロードしてください（自動的にマスターファイルとして上書きされます）", 
    type=["xlsx"]
)

# --- Excel出力関数の定義 ---
def convert_excel_to_pdf(excel_file_path, pdf_file_path):
    abs_excel = os.path.abspath(excel_file_path)
    abs_pdf = os.path.abspath(pdf_file_path)
    
    # 保存先フォルダの作成
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
    
    excel = None
    wb = None
    try:
        # win32comでExcelをバックグラウンド起動
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # ワークブックを開く
        wb = excel.Workbooks.Open(abs_excel)
        
        # 全てのシートを選択状態にする（これで複数シートを一括でPDFにできる）
        # Select multiple sheets: wb.Sheets(list_of_sheet_names).Select()
        # 今回は発注書と連絡書のみ残っているのでそのままExportAsFixedFormatでOK
        
        # PDFとして保存 (0 = xlTypePDF)
        wb.ExportAsFixedFormat(0, abs_pdf)
    finally:
        if wb is not None:
            wb.Close(False)
        if excel is not None:
            excel.Quit()

def generate_order_excel(order_df, selected_contractor, filename="拾い出し表.xlsx"):
    wb = openpyxl.load_workbook(filename)
    if "発注書" in wb.sheetnames:
        ws = wb["発注書"]
    else:
        ws = wb.create_sheet("発注書")
        
    def safe_set(row, col, value):
        cell = ws.cell(row=row, column=col)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            cell.value = value

    max_row = ws.max_row
    if max_row >= 2:
        for r in range(2, max_row + 1):
            for c in range(1, 17):
                safe_set(r, c, None)
                
    today_str = datetime.now().strftime("%Y/%m/%d")
    
    for idx, row in enumerate(order_df.itertuples(), start=2):
        safe_set(idx, 1, getattr(row, "材料コード", None))
        safe_set(idx, 2, today_str)
        safe_set(idx, 5, getattr(row, "発注先", None))
        safe_set(idx, 6, getattr(row, "担当者", None))
        safe_set(idx, 7, getattr(row, "名称", None))
        safe_set(idx, 8, getattr(row, "規格", None))
        safe_set(idx, 9, getattr(row, "規格_1", None))
        safe_set(idx, 10, getattr(row, "規格_2", None))
        
        def clean_val(val):
            if pd.isna(val) or val == "NaN":
                return ""
            return val
            
        safe_set(idx, 11, clean_val(getattr(row, "階", None)))
        safe_set(idx, 12, clean_val(getattr(row, "発注", None)))
        safe_set(idx, 13, clean_val(getattr(row, "単位", None)))
        
        ndate = getattr(row, "納品日", None)
        if pd.notna(ndate) and ndate != "NaN":
            if isinstance(ndate, datetime):
                safe_set(idx, 14, ndate.strftime("%Y/%m/%d"))
            else:
                val_str = str(ndate).split(" ")[0]
                safe_set(idx, 14, val_str)
        else:
            safe_set(idx, 14, "")
            
        safe_set(idx, 15, clean_val(getattr(row, "納品場所", None)))
        safe_set(idx, 16, clean_val(getattr(row, "納品備考", None)))
        
    sheets_to_keep = ["発注書", "連絡書"]
    for sheet in list(wb.sheetnames):
        if sheet not in sheets_to_keep:
            wb.remove(wb[sheet])
            
    if wb.sheetnames:
        wb.active = 0
    for view in wb.views:
        view.activeTab = 0
        view.firstSheet = 0
            
    output_filename = f"発注書_連絡書_{selected_contractor}.xlsx"
    wb.save(output_filename)
    wb.close()
    return output_filename

# --- 1. Excelデータの読み込みとクレンジング ---
@st.cache_data
def load_data(file_source="拾い出し表.xlsx"):
    df = pd.read_excel(file_source, sheet_name="拾い出し")
    
    new_columns = {}
    for col in df.columns:
        col_str = str(col).strip()
        if "名称" in col_str:
            new_columns[col] = "名称"
        elif "材料" in col_str and ("コード" in col_str or "ｺｰﾄﾞ" in col_str):
            new_columns[col] = "材料コード"
        elif "発注業者" in col_str or "発注先" in col_str:
            new_columns[col] = "発注先"
        elif col_str == "担当者":
            new_columns[col] = "担当者"
        elif "規格" in col_str:
            if "14.5" in col_str:
                new_columns[col] = "規格"
            elif "8" in col_str:
                new_columns[col] = "規格_1"
            elif "6.75" in col_str or "入数" in col_str:
                new_columns[col] = "規格_2"
        elif "階" in col_str:
            new_columns[col] = "階"
        elif col_str.startswith("発注") and "予定日" not in col_str:
            new_columns[col] = "発注"
        elif "単位" in col_str:
            new_columns[col] = "単位"
        elif "納品日" in col_str:
            new_columns[col] = "納品日"
        elif "納品場所" in col_str:
            new_columns[col] = "納品場所"
        elif "納品備考" in col_str:
            new_columns[col] = "納品備考"
            
    df = df.rename(columns=new_columns)
    
    required_cols = ["材料コード", "名称", "発注先", "担当者", "規格", "規格_1", "規格_2", "階", "発注", "単位", "納品日", "納品場所", "納品備考"]
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""
            
    df = df.dropna(subset=['名称'])
    df = df[df['名称'].astype(str).str.strip() != ""]
    
    df.insert(0, '発注対象', False) 
    return df

# =========================================================
# アップロード時の処理
# =========================================================
if uploaded_file is not None:
    # 新しいファイルがアップロードされた時だけ処理を実行する（毎回の再ランでチェック状態がリセットされるのを防ぐため）
    if "last_uploaded_file_id" not in st.session_state or st.session_state.last_uploaded_file_id != uploaded_file.file_id:
        try:
            # アップロードされたファイルの名前が何であっても、サーバー上では「拾い出し表.xlsx」として上書き保存する
            with open("拾い出し表.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
                
            # 再読み込み
            st.session_state.raw_df = load_data("拾い出し表.xlsx")
            st.session_state.display_df = st.session_state.raw_df.copy()
            
            # 処理済みのファイルIDを記録
            st.session_state.last_uploaded_file_id = uploaded_file.file_id
            
            st.sidebar.success(f"✅ 「{uploaded_file.name}」をマスターファイルとして更新しました！")
        except Exception as e:
            st.sidebar.error(f"ファイルの読み込みに失敗しました: {e}")

# 1. データの読み込み（初回のみ）
if "raw_df" not in st.session_state:
    try:
        st.session_state.raw_df = load_data("拾い出し表.xlsx")
    except Exception as e:
        st.warning("現在、データが読み込まれていません。サイドバーから「拾い出し表.xlsx」をアップロードしてください。")
        st.session_state.raw_df = pd.DataFrame()

# メイン処理（データが正常に読み込まれている場合のみ実行）
if not st.session_state.raw_df.empty:

    # フィルタリングされた表示用データの初期化
    if "display_df" not in st.session_state:
        st.session_state.display_df = st.session_state.raw_df.copy()

    # データエディタでチェックが押された直後に、原本（raw_df）を更新するコールバック関数
    def handle_editor_change(editor_key):
        editor_state = st.session_state.get(editor_key, {})
        edited_rows = editor_state.get("edited_rows", {})
        for pos, edits in edited_rows.items():
            if "発注対象" in edits:
                # 表示用データフレームの行番号から、元の本当の行番号を特定
                real_idx = st.session_state.display_df.index[int(pos)]
                st.session_state.raw_df.at[real_idx, "発注対象"] = edits["発注対象"]

    # --- 2. 拾い出し表の表示と操作（チェックボックス） ---
    st.subheader("1. 発注する材料にチェックを入れてください")

    # 外出しのフィルター検索ボックス
    search_query = st.text_input("🔍 絞り込み検索（業者名や材料名などを入力すると、それ以外は非表示になります）", "")

    # 検索ワードに基づいて表示用データフレーム（display_df）を更新します
    if search_query:
        mask = (
            st.session_state.raw_df['発注先'].astype(str).str.contains(search_query, case=False, na=False) |
            st.session_state.raw_df['名称'].astype(str).str.contains(search_query, case=False, na=False) |
            st.session_state.raw_df['材料コード'].astype(str).str.contains(search_query, case=False, na=False) |
            st.session_state.raw_df['規格'].astype(str).str.contains(search_query, case=False, na=False)
        )
        st.session_state.display_df = st.session_state.raw_df[mask].copy()
    else:
        st.session_state.display_df = st.session_state.raw_df.copy()

    display_columns = ['発注対象', '材料コード', '名称', '規格', '発注', '単位', '発注先']

    # 検索条件ごとに一意のキーを発行します
    current_editor_key = f"material_editor_{search_query}"

    # 3. データエディタの表示
    st.data_editor(
        st.session_state.display_df,
        column_order=display_columns,
        hide_index=True,
        use_container_width=True,
        disabled=["材料コード", "名称", "規格", "発注", "単位", "発注先"],
        key=current_editor_key,
        on_change=handle_editor_change,
        kwargs={"editor_key": current_editor_key}
    )

    # --- 3. 業者の選択プルダウン ---
    st.subheader("2. 発注先を選択してください")
    contractors = st.session_state.raw_df['発注先'].dropna().unique().tolist()
    selected_contractor = st.selectbox("発注先業者", ["--選択してください--"] + contractors)

    # --- 4. 発注書の自動生成ビュー ---
    st.subheader("3. 発注書プレビュー")
    if selected_contractor != "--選択してください--":
        # チェックが入っている材料データをすべて抽出します（原本 raw_df から抽出）
        order_df = st.session_state.raw_df[st.session_state.raw_df['発注対象'] == True].copy()
        
        # 抽出したすべての材料の「発注先」を、選択した業者名に上書きします
        order_df['発注先'] = selected_contractor
        
        if not order_df.empty:
            # 材料コードの昇順で並び替え
            order_df = order_df.sort_values(by=['材料コード'], ascending=True)
            
            st.success(f"{selected_contractor} 宛ての発注データが指定の順序で抽出されました。")
            # プレビュー表示（印刷に不要な列を隠す）
            preview_columns = ['材料コード', '名称', '規格', '発注', '単位']
            st.dataframe(order_df[preview_columns], hide_index=True, use_container_width=True)
            
            col1, col2 = st.columns(2)
            
            # 従来のエクセル発行ボタン
            with col1:
                if st.button(f"この内容で {selected_contractor} へ発注書と連絡書(Excel)を発行する", type="secondary", use_container_width=True):
                    output_file = generate_order_excel(order_df, selected_contractor)
                    
                    st.toast("発注書と連絡書を発行しました！", icon="🎉")
                    
                    st.success(f"✅ **{selected_contractor} 宛てのエクセルを新しく作成しました！**")
                    st.info("※ このファイルには『発注書』シートと『連絡書』シートのみが含まれています。")
                    
                    with open(output_file, "rb") as f:
                        excel_data = f.read()
                    st.download_button(
                        label=f"📥 作成された「{output_file}」をダウンロードする",
                        data=excel_data,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            
            # PDFダウンロードボタン
            with col2:
                if st.button(f"📄 発注書と連絡書(PDF)を作成する", type="primary", use_container_width=True):
                    with st.spinner("PDFを作成中..."):
                        # 1. まずExcelを作成する
                        output_excel = generate_order_excel(order_df, selected_contractor)
                        
                        # 2. ファイル名の生成
                        mat_code = str(order_df.iloc[0]['材料コード'])
                        vendor_name = selected_contractor
                        now_str = datetime.now().strftime("%Y%m%d_%H%M")
                        file_name = f"{mat_code}_{vendor_name}_{now_str}.pdf"
                        
                        # カレントディレクトリに一時保存
                        save_full_path = os.path.join(os.getcwd(), file_name)
                        
                        try:
                            # 3. 作成したExcelからPDFに変換して保存
                            convert_excel_to_pdf(output_excel, save_full_path)
                            st.success(f"✅ **PDFの作成が完了しました！**")
                            st.toast("PDFを作成しました！", icon="📄")
                            
                            # ダウンロードボタンを表示
                            with open(save_full_path, "rb") as f:
                                pdf_data = f.read()
                            st.download_button(
                                label=f"📥 作成された「{file_name}」をダウンロードする",
                                data=pdf_data,
                                file_name=file_name,
                                mime="application/pdf",
                                use_container_width=True
                            )
                        except Exception as e:
                            st.error(f"PDFの作成に失敗しました。Excelがインストールされているか確認してください。\nエラー詳細: {e}")
        else:
            st.warning("チェックされた材料はありません。")
    else:
        st.info("上のリストから業者を選択してください。")
