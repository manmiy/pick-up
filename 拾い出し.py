import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
import os
import platform
import subprocess
import math
from datetime import datetime
from copy import copy

# ページ設定
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
            # st.secrets["app_password"] を利用（設定されていない場合は暫定パスワード）
            try:
                correct_password = st.secrets["app_password"]
            except:
                correct_password = "admin"  # secrets未設定時のフォールバック
                
            if pwd == correct_password:
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("😕 パスワードが間違っています。")
        return False
    return True

if not check_password():
    st.stop()

# =========================================================
# マスターファイルの更新（サイドバー）
# =========================================================
st.sidebar.subheader("マスターファイルの更新")
uploaded_file = st.sidebar.file_uploader(
    "エクセルファイルをアップロードしてください（自動的にマスターファイルとして上書きされます）", 
    type=["xlsx"]
)

# --- PDF変換ロジック ---
def convert_excel_to_pdf(excel_file_path, pdf_file_path):
    abs_excel = os.path.abspath(excel_file_path)
    abs_pdf = os.path.abspath(pdf_file_path)
    
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)
    temp_pdf_excel = abs_excel.replace(".xlsx", "_temp_pdf.xlsx")
    wb_temp = openpyxl.load_workbook(abs_excel)
    
    target_sheet_name = None
    for sheet_name in wb_temp.sheetnames:
        if "連絡書" in sheet_name.strip():
            target_sheet_name = sheet_name
            break
    if target_sheet_name is None and len(wb_temp.sheetnames) > 0:
        target_sheet_name = wb_temp.sheetnames[0]
        
    for sheet in list(wb_temp.sheetnames):
        if sheet != target_sheet_name:
            wb_temp[sheet].sheet_state = 'veryHidden'
            wb_temp[sheet].print_area = None
            
    wb_temp.active = wb_temp.sheetnames.index(target_sheet_name)
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

# --- スタイルの複製用ヘルパー関数 ---
def copy_cell_style(src_cell, dest_cell):
    if src_cell.has_style:
        dest_cell.font = copy(src_cell.font)
        dest_cell.border = copy(src_cell.border)
        dest_cell.fill = copy(src_cell.fill)
        dest_cell.number_format = copy(src_cell.number_format)
        dest_cell.protection = copy(src_cell.protection)
        dest_cell.alignment = copy(src_cell.alignment)

# --- Excel出力ロジック（自動拡張・ページ数自動カウント版） ---
# --- Excel出力ロジック（自動拡張・ページ数自動カウント・結合セル完全対応版） ---
def generate_order_excel(order_df, selected_contractor, project_title="", staff_name="", order_date="", filename="拾い出し表.xlsx"):
    wb = openpyxl.load_workbook(filename)
    
    if "発注書" in wb.sheetnames:
        ws = wb["発注書"]
    else:
        ws = wb.create_sheet("発注書")
        
    def safe_set(sheet_obj, row, col, value):
        """結合セルの裏側に書き込んでクラッシュするのを防ぐ保護関数"""
        cell = sheet_obj.cell(row=row, column=col)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell):
            cell.value = value
            
    def clean_val(val):
        if pd.isna(val) or val == "NaN": return ""
        return val

    # 発注書シートの初期化
    for r in range(2, 500): 
        for c in range(1, 17):
            safe_set(ws, r, c, None)
                
    # 発注書シートへの書き込み
    for idx, row in enumerate(order_df.itertuples(), start=2):
        safe_set(ws, idx, 1, getattr(row, "材料コード", None))
        safe_set(ws, idx, 2, order_date)
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
        
    target_sheet_name = None
    for sheet_name in wb.sheetnames:
        if "連絡書" in sheet_name.strip():
            target_sheet_name = sheet_name
            break
    if target_sheet_name is None and len(wb.sheetnames) > 0:
        target_sheet_name = wb.sheetnames[0]
            
    if target_sheet_name in wb.sheetnames:
        ws_renraku = wb[target_sheet_name]
        
        num_items = len(order_df)
        total_pages = max(1, math.ceil(num_items / 14))
        existing_merged_ranges = list(ws_renraku.merged_cells.ranges)
        
        # 2ページ目以降のフォーマット複製
        for p in range(1, total_pages):
            row_offset = p * 20
            
            for src_row in range(1, 21):
                dest_row = src_row + row_offset
                ws_renraku.row_dimensions[dest_row].height = ws_renraku.row_dimensions[src_row].height
                
                for col in range(1, 9): 
                    src_cell = ws_renraku.cell(row=src_row, column=col)
                    dest_cell = ws_renraku.cell(row=dest_row, column=col)
                    
                    # 結合セルの場合は値を直接コピーしない
                    if not isinstance(src_cell, openpyxl.cell.cell.MergedCell):
                        dest_cell.value = src_cell.value
                    copy_cell_style(src_cell, dest_cell)
            
            # 結合状態の複製
            for merged_range in existing_merged_ranges:
                if merged_range.bounds[1] <= 20: 
                    min_col, min_row, max_col, max_row = merged_range.bounds
                    ws_renraku.merge_cells(
                        start_row=min_row + row_offset, start_column=min_col,
                        end_row=max_row + row_offset, end_column=max_col
                    )
        
        # 連絡書シートへのデータ書き込み（全て safe_set を経由させる）
        for p in range(total_pages):
            row_offset = p * 20
            
            if order_date.strip():
                safe_set(ws_renraku, 1 + row_offset, 4, order_date.strip())
                
            safe_set(ws_renraku, 1 + row_offset, 6, f"{p + 1} / {total_pages}")
            safe_set(ws_renraku, 1 + row_offset, 7, None)
            safe_set(ws_renraku, 1 + row_offset, 8, None)
                
            if staff_name.strip():
                safe_set(ws_renraku, 2 + row_offset, 3, staff_name.strip())
                
            if project_title.strip():
                title_cell = ws_renraku.cell(row=3 + row_offset, column=2)
                if not isinstance(title_cell, openpyxl.cell.cell.MergedCell):
                    title_cell.value = project_title.strip()
                    title_cell.alignment = Alignment(shrinkToFit=True, vertical='center', horizontal='left')
                safe_set(ws_renraku, 3 + row_offset, 3, None)
                safe_set(ws_renraku, 3 + row_offset, 4, None)
                safe_set(ws_renraku, 3 + row_offset, 5, None)

            for i in range(14):
                item_idx = p * 14 + i
                target_row = 5 + i + row_offset
                
                if item_idx < num_items:
                    item = order_df.iloc[item_idx]
                    safe_set(ws_renraku, target_row, 1, "□")
                    safe_set(ws_renraku, target_row, 2, clean_val(item.get("名称", "")))
                    safe_set(ws_renraku, target_row, 3, clean_val(item.get("規格", "")))
                    safe_set(ws_renraku, target_row, 4, clean_val(item.get("規格_1", "")))
                    safe_set(ws_renraku, target_row, 5, clean_val(item.get("規格_2", "")))
                    safe_set(ws_renraku, target_row, 6, clean_val(item.get("発注", "")))
                    safe_set(ws_renraku, target_row, 7, clean_val(item.get("単位", "")))
                else:
                    for col_idx in range(1, 8):
                        safe_set(ws_renraku, target_row, col_idx, None)
        
        max_print_row = 20 * total_pages
        ws_renraku.print_area = f'A1:H{max_print_row}'
        ws_renraku.sheet_properties.pageSetUpPr.fitToPage = True
        ws_renraku.page_setup.fitToWidth = 1
        ws_renraku.page_setup.fitToHeight = total_pages
                
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

# =========================================================
# アプリケーション実行部
# =========================================================

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
    try: 
        st.session_state.raw_df = load_data("拾い出し表.xlsx")
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
    st.caption("💡 行（材料名や左端の余白）をクリックすると、対応する広告・参考資料シートが下部に自動プレビューされます。")
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

    # 行選択イベントを有効化 (on_select="rerun")
    st.data_editor(
        st.session_state.display_df,
        column_order=display_columns,
        hide_index=True,
        use_container_width=True,
        disabled=["材料コード", "名称", "規格", "発注", "単位", "発注先"],
        key=current_editor_key,
        on_change=handle_editor_change,
        kwargs={"editor_key": current_editor_key},
        on_select="rerun",
        selection_mode="single"
    )
    
    # =========================================================
    # 選択された材料に応じた「広告シート」の自動プレビュー表示
    # =========================================================
    editor_state = st.session_state.get(current_editor_key, {})
    selection = editor_state.get("selection", {})
    selected_rows_indices = selection.get("rows", [])
    
    if selected_rows_indices:
        selected_idx = selected_rows_indices[0]
        clicked_material_name = str(st.session_state.display_df.iloc[selected_idx]["名称"])
        clicked_vendor = str(st.session_state.display_df.iloc[selected_idx]["発注先"])
        
        st.markdown("---")
        st.subheader(f"📖 『{clicked_material_name}』の参考資料")
        
        try:
            # 既存のシート一覧を取得
            all_sheets = pd.ExcelFile("拾い出し表.xlsx").sheet_names
            exclude = ["拾い出し", "発注書", "連絡書", "並べ替えビュー"]
            ad_sheets = [s for s in all_sheets if not any(x in s for x in exclude)]
            
            target_ad_sheet = None
            
            # 名称や業者名から、該当するシートを推測
            if "シーリング" in clicked_material_name or "シール" in clicked_vendor:
                matches = [s for s in ad_sheets if "シーリング" in s or "シール" in s]
                if matches: target_ad_sheet = matches[0]
            elif "外壁" in clicked_material_name or "サイディング" in clicked_material_name or "サイディング" in clicked_vendor:
                matches = [s for s in ad_sheets if "外壁" in s or "手間" in s or "サイディング" in s]
                if matches: target_ad_sheet = matches[0]
                
            if target_ad_sheet:
                # 広告シートをプレビュー表示
                ad_df = pd.read_excel("拾い出し表.xlsx", sheet_name=target_ad_sheet).fillna("")
                st.info(f"💡 マスターファイル内の「{target_ad_sheet}」シートを自動表示しています。")
                st.dataframe(ad_df, use_container_width=True, hide_index=True)
            else:
                st.info("この材料に対応する個別の広告・参考シートは見つかりませんでした。")
                
        except Exception as e:
            st.error(f"参考資料の読み込み中にエラーが発生しました: {e}")
            
    st.markdown("---")

    st.subheader("2. 発注先および宛先情報を指定してください")
    
    col_top1, col_top2, col_top3 = st.columns([4, 3, 3])
    with col_top1:
        contractors = st.session_state.raw_df['発注先'].dropna().unique().tolist()
        selected_contractor = st.selectbox("発注先業者", ["--選択してください--"] + contractors)
    with col_top2:
        staff_name = st.text_input("宛先担当者名（苗字のみで可）", value="", placeholder="例: satou")
    with col_top3:
        today_str_ui = datetime.now().strftime("%Y/%m/%d")
        order_date = st.text_input("日付", value=today_str_ui)
        
    project_title = st.text_input("件名（現場名など）", value="白石送電事務所 新築工事")

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
                    output_file = generate_order_excel(
                        order_df, 
                        selected_contractor, 
                        project_title=project_title, 
                        staff_name=staff_name,
                        order_date=order_date
                    )
                    st.session_state["generated_excel"] = output_file
                    st.toast("Excelを発行しました！", icon="🎉")
            
            if "generated_excel" in st.session_state and st.session_state["generated_excel"]:
                excel_file_path = st.session_state["generated_excel"]
                with col1:
                    with open(excel_file_path, "rb") as f: 
                        excel_data = f.read()
                    st.download_button(
                        label=f"📥 Excelをダウンロード",
                        data=excel_data,
                        file_name=excel_file_path,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col2:
                    st.markdown("### PDFへ変換")
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
                            with open(pdf_file_path, "rb") as f: 
                                pdf_data = f.read()
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
