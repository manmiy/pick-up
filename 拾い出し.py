import streamlit as st
import pandas as pd
import os
import platform
from datetime import datetime

# =========================================================
# PDF描画用ライブラリ (ReportLab) のインポート
# =========================================================
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
            try:
                correct_password = st.secrets["app_password"]
            except:
                correct_password = "admin"  
                
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
# 🎯 ファイル読み込み先の設定（エラーの根本解決）
# =========================================================
# 毎回アップロードしなくても、このファイルがフォルダにあれば自動で読み込みます
DEFAULT_FILE = "000_拾い出し表_発注書（テスト）.xlsx"

st.sidebar.subheader("マスターファイルの更新")
uploaded_file = st.sidebar.file_uploader(
    "エクセルデータをアップロード（※フォルダにファイルがあれば不要です）", 
    type=["xlsx"]
)

# どのエクセルファイルを使うか決定するロジック
target_file = None
if uploaded_file is not None:
    # アップロードされたら「拾い出し表.xlsx」として保存して使う
    with open("拾い出し表.xlsx", "wb") as f:
        f.write(uploaded_file.getbuffer())
    target_file = "拾い出し表.xlsx"
    if "last_uploaded_file_id" not in st.session_state or st.session_state.last_uploaded_file_id != uploaded_file.file_id:
        st.session_state.last_uploaded_file_id = uploaded_file.file_id
        if "raw_df" in st.session_state:
            del st.session_state.raw_df
        st.sidebar.success("✅ マスターデータを更新しました！")
elif os.path.exists(DEFAULT_FILE):
    # アップロードされていなくても、デフォルトファイルがあればそれを使う
    target_file = DEFAULT_FILE
elif os.path.exists("拾い出し表.xlsx"):
    # 過去にアップロードしたファイルが残っていればそれを使う
    target_file = "拾い出し表.xlsx"

# =========================================================
# 🎯 Pythonネイティブ PDF直接生成ロジック (Excel非依存)
# =========================================================
def generate_pdf_report(order_df, pdf_path, selected_contractor, project_title, staff_name, order_date):
    font_name = "JapaneseFont"
    try:
        if platform.system() == "Windows":
            pdfmetrics.registerFont(TTFont(font_name, 'C:\\Windows\\Fonts\\msgothic.ttc'))
        else:
            pdfmetrics.registerFont(TTFont(font_name, 'ipaexg.ttf'))
    except Exception:
        st.warning("日本語フォントが見つからないため、PDFの文字化けが発生する可能性があります。")
        font_name = "Helvetica" 

    doc = SimpleDocTemplate(
        pdf_path, 
        pagesize=A4, 
        rightMargin=30, 
        leftMargin=30, 
        topMargin=30, 
        bottomMargin=30
    )
    
    elements = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='TitleStyle', fontName=font_name, fontSize=18, alignment=1, spaceAfter=20)
    info_style = ParagraphStyle(name='InfoStyle', fontName=font_name, fontSize=12, spaceAfter=10)
    right_style = ParagraphStyle(name='RightStyle', fontName=font_name, fontSize=10, alignment=2)

    elements.append(Paragraph(f"Date: {order_date}", right_style))
    elements.append(Paragraph("連 絡 書", title_style))
    
    elements.append(Paragraph(f"<b>{selected_contractor}</b>", info_style))
    if staff_name:
        elements.append(Paragraph(f"担当: {staff_name} 様", info_style))
    
    elements.append(Spacer(1, 10))
    elements.append(Paragraph(f"<b>件名:</b> {project_title}", info_style))
    elements.append(Paragraph("※下記の内容にて手配をお願い致します。", ParagraphStyle(name='Note', fontName=font_name, fontSize=10)))
    elements.append(Spacer(1, 15))

    table_data = [["確 認", "名 称", "規 格", "階・納品場所", "備 考", "発注数", "単位"]]
    
    def clean_val(val):
        if pd.isna(val) or str(val) == "NaN": return ""
        return str(val)

    for idx, row in order_df.iterrows():
        place_str = f"{clean_val(row.get('階', ''))} {clean_val(row.get('納品場所', ''))}".strip()
        table_data.append([
            "□",
            clean_val(row.get("名称")),
            clean_val(row.get("規格")),
            place_str,
            clean_val(row.get("納品備考")),
            clean_val(row.get("発注")),
            clean_val(row.get("単位"))
        ])

    col_widths = [40, 120, 100, 100, 100, 45, 35] 
    t = Table(table_data, colWidths=col_widths, repeatRows=1) 
    
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),      
        ('ALIGN', (0,0), (-1,0), 'CENTER'),           
        ('ALIGN', (0,1), (0,-1), 'CENTER'),           
        ('ALIGN', (5,1), (6,-1), 'RIGHT'),            
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),         
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),  
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), 
        ('FONTSIZE', (0,0), (-1,-1), 9),              
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    
    elements.append(t)
    
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont(font_name, 9)
        page_num_text = f"Page {doc.page}"
        canvas.drawRightString(A4[0] - 30, 20, page_num_text)
        canvas.restoreState()

    doc.build(elements, onFirstPage=add_page_number, onLaterPages=add_page_number)


# --- データの読み込み ---
@st.cache_data
def load_data(file_source):
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
# アプリケーション UI 構築部
# =========================================================

# 🎯 データが読み込めるかチェック
if target_file is None:
    st.warning("📂 エクセルデータが見つかりません。プログラムと同じフォルダにエクセルファイルを置くか、左のメニューからアップロードしてください。")
else:
    if "raw_df" not in st.session_state:
        try: 
            st.session_state.raw_df = load_data(target_file)
        except Exception as e:
            st.error(f"データの読み込みに失敗しました: {e}")
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

# 🎯 広告シートの自動プレビュー機能
editor_state = st.session_state.get(current_editor_key, {})
selection = editor_state.get("selection", {})
selected_rows_indices = selection.get("rows", [])

if selected_rows_indices:
        
        if selected_rows_indices:
            selected_idx = selected_rows_indices[0]
            clicked_material_name = str(st.session_state.display_df.iloc[selected_idx]["名称"])
            clicked_vendor = str(st.session_state.display_df.iloc[selected_idx]["発注先"])
            
            st.markdown("---")
            st.subheader(f"📖 『{clicked_material_name}』の参考資料")
            
            try:
                # target_file を使ってシートを読み込む
                all_sheets = pd.ExcelFile(target_file).sheet_names
                exclude = ["拾い出し", "発注書", "連絡書", "並べ替えビュー"]
                ad_sheets = [s for s in all_sheets if not any(x in s for x in exclude)]
                
                target_ad_sheet = None
                if "シーリング" in clicked_material_name or "シール" in clicked_vendor:
                    matches = [s for s in ad_sheets if "シーリング" in s or "シール" in s]
                    if matches: target_ad_sheet = matches[0]
                elif "外壁" in clicked_material_name or "サイディング" in clicked_material_name or "サイディング" in clicked_vendor:
                    matches = [s for s in ad_sheets if "外壁" in s or "手間" in s or "サイディング" in s]
                    if matches: target_ad_sheet = matches[0]
                    
                if target_ad_sheet:
                    ad_df = pd.read_excel(target_file, sheet_name=target_ad_sheet).fillna("")
                    st.info(f"💡 マスターデータ内の「{target_ad_sheet}」シートを表示しています。")
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

        st.subheader("3. 発注書プレビュー・発行")
        if selected_contractor != "--選択してください--":
            order_df = st.session_state.raw_df[st.session_state.raw_df['発注対象'] == True].copy()
            
            if not order_df.empty:
                order_df = order_df.sort_values(by=['材料コード'], ascending=True)
                st.success(f"{selected_contractor} 宛ての発注データが抽出されました。")
                
                preview_columns = ['材料コード', '名称', '規格', '発注', '単位']
                st.dataframe(order_df[preview_columns], hide_index=True, use_container_width=True)
                
                now_str = datetime.now().strftime("%Y%m%d_%H%M")
                default_pdf_name = f"連絡書_{selected_contractor}_{now_str}.pdf"
                pdf_file_path = os.path.join(os.getcwd(), default_pdf_name)
                
                if st.button("📄 美しいPDFの連絡書を作成する", type="primary", use_container_width=True):
                    with st.spinner("PythonでPDFを直接描画中..."):
                        try:
                            generate_pdf_report(
                                order_df, 
                                pdf_file_path,
                                selected_contractor, 
                                project_title, 
                                staff_name,
                                order_date
                            )
                            st.session_state["generated_pdf"] = pdf_file_path
                            st.toast("PDFの生成が完了しました！", icon="🎉")
                        except Exception as e:
                            st.error(f"PDF生成エラー: {e}")
                
                if "generated_pdf" in st.session_state and st.session_state["generated_pdf"] == pdf_file_path:
                    if os.path.exists(pdf_file_path):
                        with open(pdf_file_path, "rb") as f: 
                            pdf_data = f.read()
                        st.download_button(
                            label=f"📥 作成したPDFをダウンロード",
                            data=pdf_data,
                            file_name=default_pdf_name,
                            mime="application/pdf",
                            use_container_width=True
                        )
            else:
                st.warning("チェックされた材料はありません。")
        else:
            st.info("上のリストから業者を選択してください。")
