import streamlit as st
import pandas as pd
import openpyxl
import os
import platform
import subprocess
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
            if pwd == st.secrets["app_password"]:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("😕 パスワードが間違っています。")
        return False
    return True

if not check_password():
    st.stop()

# =========================================================
# マスターファイルの更新
# =========================================================
st.sidebar.subheader("マスターファイルの更新")
uploaded_file = st.sidebar.file_uploader(
    "エクセルファイルをアップロードしてください（自動的にマスターファイルとして上書きされます）", 
    type=["xlsx"]
)

# --- 🎯 PDF変換ロジック（シートを消さずに隠す安全仕様） ---
def convert_excel_to_pdf(excel_file_path, pdf_file_path):
    abs_excel = os.path.abspath(excel_file_path)
    abs_pdf = os.path.abspath(pdf_file_path)
    
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
    
    # PDF専用の一時ファイルを作成
    temp_pdf_excel = abs_excel.replace(".xlsx", "_temp_pdf.xlsx")
    wb_temp = openpyxl.load_workbook(abs_excel)
    
    # 「連絡書」シートを特定
    target_sheet_name = None
    for sheet_name in wb_temp.sheetnames:
        if "連絡書" in sheet_name.strip():
            target_sheet_name = sheet_name
            break
    if target_sheet_name is None and len(wb_temp.sheetnames) > 0:
        target_sheet_name = wb_temp.sheetnames[0]
        
    # 🎯【重要】シートを削除せず「非表示」にするだけ。これで関数は壊れません。
    for sheet in list(wb_temp.sheetnames):
        if sheet != target_sheet_name:
            wb_temp[sheet].sheet_state = 'hidden'
            
    # 2枚目の白紙が出ないよう印刷範囲を再度固定
    ws_renraku = wb_temp[target_sheet_name]
    ws_renraku.print_area = 'A1:H20'
    ws_renraku.sheet_properties.pageSetUpPr.fitToPage = True
    ws_renraku.page_setup.fitToWidth = 1
    ws_renraku.page_setup.fitToHeight = 1
    
    wb_temp.save(temp_pdf_excel)
    wb_temp.close()
    
    if platform.system() == "Windows":
        try:
            import win32com.client
        except ImportError:
            raise Exception("Windows環境ですが win32com がインストールされていません。")
        excel = None
        wb = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(temp_pdf_excel)
            wb.ExportAsFixedFormat(0, abs_pdf)
        finally:
            if wb is not None: wb.Close(False)
            if excel is not None: excel.Quit()
            if os.path.exists(temp_pdf_excel): os.remove(temp_pdf_excel)
    else:
        try:
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "pdf", 
                temp_pdf_excel, "--outdir", os.path.dirname(abs_pdf)
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            base_name = os.path.splitext(os.path.basename(temp_pdf_excel))[0]
            lo_pdf_path = os.path.join(os.path.dirname(abs_pdf), f"{base_name}.pdf")
            
            if lo_pdf_path != abs_pdf:
                if os.path.exists(abs_pdf): os.remove(abs_pdf)
                os.rename(lo_pdf_path, abs_pdf)
        except subprocess.CalledProcessError as e:
            raise Exception(f"LibreOfficeでのPDF変換に失敗: {e.stderr.decode('utf-8', errors='ignore')}")
        except Exception as e:
            raise Exception(f"PDF変換エラー: {e}")
        finally:
            if os.path.exists(temp_pdf_excel):
                try: os.remove(temp_pdf_excel)
                except: pass

# --- Excel出力ロジック ---
def generate_order_excel(order_df, selected_contractor, project_title="", staff_name="", filename="拾い出し表.xlsx"):
    wb = openpyxl.load_workbook(filename)
    
    if "発注書" in wb.sheetnames:
        ws = wb["発注書"]
    else:
        ws = wb.create_sheet("発注書")
        
    def safe_set(sheet_obj, row, col, value):
        cell = sheet_obj.cell(row=row, column=col)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            cell.value = value
            
    def clean_val(val):
        if pd.isna(val) or val == "NaN": return ""
        return val

    # 1. 【発注書】シートへのデータ流し込み（ここは今まで通り安全に実行）
    max_row = ws.max_row
    if max_row >= 2:
        for r in range(2, max_row + 1):
            for c in range(1, 17):
                safe_set(ws, r, c, None)
                
    today_str = datetime.now().strftime("%Y/%m/%d")
    
    for idx, row in enumerate(order_df.itertuples(), start=2):
        safe_set(ws, idx, 1, getattr(row, "材料コード", None))
        safe_set(ws, idx, 2, today_str)
        safe_set(ws, idx, 5, getattr(row, "発注先", None))
        safe_set(ws, idx, 6, getattr(row, "担当者", None))
        safe_set(ws, idx, 7, getattr(row, "名称", None))
        safe_set(ws, idx, 8, getattr(row, "規格", None))
        safe_set(ws, idx, 9, getattr(row, "規格_1", None))
        safe_set(ws, idx, 10, getattr(row, "規格_2", None))
        safe_set(ws, idx, 11, clean_val(getattr(row, "階", None)))
        safe_set(ws, idx, 12, clean_val(getattr(row, "発注", None)))
        safe_set(ws, idx, 13, clean_val(getattr(row, "単位", None)))
        
        ndate = getattr(row, "納品日", None)
        if pd.notna(ndate) and ndate != "NaN":
            if isinstance(ndate, datetime):
                safe_set(ws, idx, 14, ndate.strftime("%Y/%m/%d"))
            else:
                safe_set(ws, idx, 14, str(ndate).split(" ")[0])
        else:
            safe_set(ws, idx, 14, "")
            
        safe_set(ws, idx, 15, clean_val(getattr(row, "納品場所", None)))
        safe_set(ws, idx, 16, clean_val(getattr(row, "納品備考", None)))
        
    # 2. 【連絡書】シートの更新（破壊的ロジックを排除）
    target_sheet_name = None
    for sheet_name in wb.sheetnames:
        if "連絡書" in sheet_name.strip():
            target_sheet_name = sheet_name
            break
            
    if target_sheet_name is None and len(wb.sheetnames) > 0:
        target_sheet_name = wb.sheetnames[0]
            
    if target_sheet_name in wb.sheetnames:
        ws_renraku = wb[target_sheet_name]
        
        # 🎯 宛先・件名だけをピンポイントで上書き。明細行はいじらないのでレイアウト崩れなし。
        if staff_name.strip():
            ws_renraku.cell(row=2, column=3).value = staff_name.strip()
            
        if project_title.strip():
            # 件名を入れるセル。位置がズレている場合は row=3, column=2 などに変更してください。
            ws_renraku.cell(row=4, column=2).value = project_title.strip() 
        
        # 印刷設定の固定
        ws_renraku.print_area = 'A1:H20'
        ws_renraku.sheet_properties.pageSetUpPr.fitToPage = True
        ws_renraku.page_setup.fitToWidth = 1
        ws_renraku.page_setup.fitToHeight = 1
                
    # PDF出力のために「発注書」は消さない。そのまま保存します。
    output_filename = f"連絡書_{selected_contractor}.xlsx"
    wb.save(output_filename)
    wb.close()
    return output_filename

# --- Excelデータの読み込み ---
@st.cache_data
def load_data(file_source="拾い出し表.xlsx"):
    df = pd.read_excel(file_source, sheet_name="拾い出し")
    new_columns = {}
    for col in df.columns:
        col_str = str(col).strip()
        if "名称" in col_str: new_columns[col] = "名称"
        elif "材料" in col_str and ("コード" in col_str or "ｺｰﾄﾞ" in col_str): new_columns[col] = "材料コード"
        elif "発注業者" in col_str or "発注先" in col_str: new_columns[col] = "発注先"
        elif col_str == "担当者": new_columns[col] = "担当者"
        elif "規格" in col_str:
            if "14.5" in col_str: new_columns[col] = "規格"
            elif "8" in col_str: new_columns[col] = "規格_1"
            elif "6.75" in col_str or "入数" in col_str: new_columns[col] = "規格_2"
        elif "階" in col_str: new_columns[col] = "階"
        elif col_str.startswith("発注") and "予定日" not in col_str: new_columns[col] = "発注"
        elif "単位" in col_str: new_columns[col] = "単位"
        elif "納品日" in col_str: new_columns[col] = "納品日"
        elif "納品場所" in col_str: new_columns[col] = "納品場所"
        elif "納品備考" in col_str: new_columns[col] = "納品備考"
            
    df = df.rename(columns=new_columns)
    required_cols = ["材料コード", "名称", "発注先", "担当者", "規格", "規格_1", "規格_2", "階", "発注", "単位", "納品日", "納品場所", "納品備考"]
    for col in required_cols:
        if col not in df.columns: df[col] = ""
            
    df = df.dropna(subset=['名称'])
    df = df[df['名称'].astype(str).str.strip() != ""]
    df.insert(0, '発注対象', False) 
    return df

if uploaded_file is not None:
    if "last_uploaded_file_id" not in st.session_state or st.session_state.last_uploaded_file_id != uploaded_file.file_id:
        try:
            with open("拾い出し表.xlsx", "wb") as f:
                f.write(uploaded_file.getbuffer())
            st.session_state.raw_df = load_data("拾い出し表.xlsx")
            st.session_state.display_df = st.session_state.raw_df.copy()
            st.session_state.last_uploaded_file_id = uploaded_file.file_id
            st.sidebar.success(f"✅ マスターファイルを更新しました！")
        except Exception as e:
            st.sidebar.error(f"ファイルの読み込みに失敗しました: {e}")

if "raw_df" not in st.session_state:
    try: st.session_state.raw_df = load_data("拾い出し表.xlsx")
    except:
        st.warning("現在、データが読み込まれていません。サイドバーからファイルをアップロードしてください。")
        st.session_state.raw_df = pd.DataFrame()

if not st.session_state.raw_df.empty:
    if "display_df" not in st.session_state:
        st.session_state.display_df = st.session_state.raw_df.copy()

    def handle_editor_change(editor_key):
        editor_state = st.session_state.get(editor_key, {})
        edited_rows = editor_state.get("edited_rows", {})
        for pos, edits in edited_rows.items():
            if "発注対象" in edits:
                real_idx = st.session_state.display_df.index[int(pos)]
                st.session_state.raw_df.at[real_idx, "発注対象"] = edits["発注対象"]

    st.subheader("1. 発注する材料にチェックを入れてください")
    search_query = st.text_input("🔍 絞り込み検索", "")

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
    current_editor_key = f"material_editor_{search_query}"

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

    st.subheader("2. 発注先および宛先情報を指定してください")
    
    meta_col1, meta_col2 = st.columns(2)
    with meta_col1:
        contractors = st.session_state.raw_df['発注先'].dropna().unique().tolist()
        selected_contractor = st.selectbox("発注先業者", ["--選択してください--"] + contractors)
        project_title = st.text_input("件名（現場名など）", value="白石送電事務所 新築工事")
    with meta_col2:
        staff_name = st.text_input("宛先担当者名（苗字のみで可）", value="", placeholder="例: satou")

    st.subheader("3. 発注書プレビュー")
    if selected_contractor != "--選択してください--":
        order_df = st.session_state.raw_df[st.session_state.raw_df['発注対象'] == True].copy()
        order_df['発注先'] = selected_contractor
        
        if not order_df.empty:
            order_df = order_df.sort_values(by=['材料コード'], ascending=True)
            st.success(f"{selected_contractor} 宛ての発注データが抽出されました。")
            preview_columns = ['材料コード', '名称', '規格', '発注', '単位']
            st.dataframe(order_df[preview_columns], hide_index=True, use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button(f"① Excelを発行する", type="primary", use_container_width=True):
                    output_file = generate_order_excel(order_df, selected_contractor, project_title=project_title, staff_name=staff_name)
                    st.session_state["generated_excel"] = output_file
                    st.toast("Excelを発行しました！", icon="🎉")
            
            if "generated_excel" in st.session_state and st.session_state["generated_excel"]:
                excel_file_path = st.session_state["generated_excel"]
                with col1:
                    with open(excel_file_path, "rb") as f: excel_data = f.read()
                    st.download_button(
                        label=f"📥 Excelをダウンロード",
                        data=excel_data,
                        file_name=excel_file_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    st.markdown("### PDFへ変換")
                    mat_code = str(order_df.iloc[0]['材料コード'])
                    vendor_name = selected_contractor
                    now_str = datetime.now().strftime("%Y%m%d_%H%M")
                    default_pdf_name = f"連絡書_{vendor_name}_{now_str}.pdf"
                    
                    pdf_filename_input = st.text_input("保存するPDFのファイル名", value=default_pdf_name)
                    pdf_file_path = os.path.join(os.getcwd(), pdf_filename_input)
                    
                    if st.button("② この名前でPDFを作成する", type="secondary", use_container_width=True):
                        with st.spinner("ExcelからPDFへ変換中..."):
                            try:
                                convert_excel_to_pdf(excel_file_path, pdf_file_path)
                                st.session_state["generated_pdf"] = pdf_file_path
                                st.toast("PDF変換が完了しました！", icon="📄")
                            except Exception as e:
                                st.error(f"PDF変換エラー: {e}")
                    
                    if "generated_pdf" in st.session_state and st.session_state["generated_pdf"] == pdf_file_path:
                        if os.path.exists(pdf_file_path):
                            with open(pdf_file_path, "rb") as f: pdf_data = f.read()
                            st.download_button(
                                label=f"📥 PDFをダウンロード",
                                data=pdf_data,
                                file_name=pdf_filename_input,
                                mime="application/pdf",
                                use_container_width=True
                            )
        else:
            st.warning("チェックされた材料はありません。")
    else:
        st.info("上のリストから業者を選択してください。")
