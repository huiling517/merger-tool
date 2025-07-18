import streamlit as st
import pandas as pd
import io
from functools import reduce
import base64

# --- 通用輔助函數 ---
def to_excel(df):
    """將 DataFrame 轉換為可供下載的 Excel Bytes 物件"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='合併結果')
    processed_data = output.getvalue()
    return processed_data

def read_and_clean_sheet(file_obj, sheet_name, header_index=0):
    """讀取指定的 Excel 工作表並進行基本清理，包含日期格式處理"""
    file_obj.seek(0)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=header_index)
    
    df.columns = [str(col) for col in df.columns]

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d').replace('NaT', '')
        elif df[col].dtype == 'object':
            df[col] = df[col].astype(str)
            
    return df

def highlight_duplicated_keys(row, duplicated_keys_set, key_column):
    """Pandas Styler 函數：如果行的索引鍵在重複集合中，則標註背景色"""
    color = 'background-color: #fff9c4' # 淡黃色
    default_color = ''
    if row[key_column] in duplicated_keys_set:
        return [color] * len(row)
    return [default_color] * len(row)

# --- Streamlit 應用程式介面 ---
st.set_page_config(page_title="Excel 全能合併工具", page_icon="🧩", layout="wide")

st.title("🧩 Excel 全能合併工具")

# --- 模式選擇 ---
app_mode = st.radio(
    "請選擇您要使用的工具模式：",
    ('雙檔查找合併 (VLOOKUP)', '多檔合併 (縱向/橫向)'),
    horizontal=True,
)
st.divider()

# 初始化 session_state
if 'final_df' not in st.session_state:
    st.session_state.final_df = None
if 'unmatched_df' not in st.session_state:
    st.session_state.unmatched_df = None
if 'duplication_warning_keys' not in st.session_state:
    st.session_state.duplication_warning_keys = []
if 'merge_key' not in st.session_state:
    st.session_state.merge_key = None
    
# ==============================================================================
# ======================== 模式一：雙檔查找合併 (VLOOKUP) =========================
# ==============================================================================
if app_mode == '雙檔查找合併 (VLOOKUP)':
    st.header("模式：雙檔查找合併 (VLOOKUP)")
    st.markdown("此模式會以**左表**為基礎，從**右表**中查找符合條件的資料，並將指定欄位新增至左表。")
    
    st.subheader("步驟一：上傳檔案並選擇工作表")
    col1, col2 = st.columns(2)
    df_left, df_right = None, None

    with col1:
        st.markdown("##### 👈 主要檔案 (左表)")
        uploaded_file_left = st.file_uploader("這是您要保留所有資料的檔案", type=["xlsx", "xls"], key="uploader_left")
        if uploaded_file_left:
            try:
                file_buffer_left = io.BytesIO(uploaded_file_left.getvalue())
                sheet_names_left = pd.ExcelFile(file_buffer_left).sheet_names
                left_sheet_name = st.selectbox("選擇主要工作表", sheet_names_left, key="sheet_left")
                header_left = st.number_input("左表標頭在第幾列?", min_value=1, value=1, key="header_left")
                if left_sheet_name:
                    df_left = read_and_clean_sheet(file_buffer_left, left_sheet_name, header_left - 1)
                    st.write("左表預覽：")
                    st.dataframe(df_left.head(), height=200)
            except Exception as e:
                st.error(f"讀取左表失敗: {e}")

    with col2:
        st.markdown("##### 👉 查找資料檔案 (右表)")
        uploaded_file_right = st.file_uploader("這是您要從中提取資料的檔案", type=["xlsx", "xls"], key="uploader_right")
        if uploaded_file_right:
            try:
                file_buffer_right = io.BytesIO(uploaded_file_right.getvalue())
                sheet_names_right = pd.ExcelFile(file_buffer_right).sheet_names
                right_sheet_name = st.selectbox("選擇查找資料的工作表", sheet_names_right, key="sheet_right")
                header_right = st.number_input("右表標頭在第幾列?", min_value=1, value=1, key="header_right")
                if right_sheet_name:
                    df_right = read_and_clean_sheet(file_buffer_right, right_sheet_name, header_right - 1)
                    st.write("右表預覽：")
                    st.dataframe(df_right.head(), height=200)
            except Exception as e:
                st.error(f"讀取右表失敗: {e}")

    if df_left is not None and df_right is not None:
        st.divider()
        st.subheader("步驟二：設定合併條件並執行")
        common_columns = list(set(df_left.columns) & set(df_right.columns))
        
        if not common_columns:
            st.error("錯誤：兩個工作表之間沒有任何共同的欄位名稱，無法進行合併。")
        else:
            with st.form("vlookup_form"):
                merge_key = st.selectbox("選擇用來對應的欄位 (共同索引鍵)", common_columns)
                available_cols_from_right = [col for col in df_right.columns if col != merge_key]
                cols_to_merge = st.multiselect("選擇要從右表加入到左表的欄位", available_cols_from_right, default=available_cols_from_right)
                
                submitted_vlookup = st.form_submit_button("🚀 執行查找合併", type="primary")

            if submitted_vlookup:
                st.session_state.final_df = None
                st.session_state.unmatched_df = None
                st.session_state.duplication_warning_keys = []
                st.session_state.merge_key = merge_key
                
                if not merge_key:
                    st.warning("請選擇一個用來對應的共同索引鍵。")
                else:
                    with st.spinner("正在合併資料並進行分析..."):
                        try:
                            duplicated_rows = df_right[df_right.duplicated(subset=[merge_key], keep=False)]
                            if not duplicated_rows.empty:
                                st.session_state.duplication_warning_keys = duplicated_rows[merge_key].unique().tolist()
                            
                            df_right_selected = df_right[[merge_key] + cols_to_merge]
                            merged_df = pd.merge(df_left, df_right_selected, on=merge_key, how='left')
                            
                            duplicated_keys = st.session_state.get('duplication_warning_keys', [])
                            if duplicated_keys:
                                merged_df['備註'] = ''
                                condition = merged_df[merge_key].isin(duplicated_keys)
                                merged_df.loc[condition, '備註'] = '一對多關係提醒'
                            
                            st.session_state.final_df = merged_df
                            st.success("🎉 查找合併成功！")

                            unmatched_df = df_right[~df_right[merge_key].isin(df_left[merge_key])]
                            st.session_state.unmatched_df = unmatched_df

                        except Exception as e:
                            st.error(f"合併失敗: {e}")

# ==============================================================================
# ======================== 模式二：多檔合併 (縱向/橫向) =========================
# ==============================================================================
elif app_mode == '多檔合併 (縱向/橫向)':
    st.header("模式：多檔合併 (縱向/橫向)")
    st.markdown("此模式可合併多個檔案，或單一檔案內的多個工作表。")

    uploaded_files = st.file_uploader(
        "請上傳您所有要處理的 Excel 檔案",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="multi_file_uploader"
    )

    if uploaded_files:
        with st.form("multi_merge_form"):
            st.subheader("1. 合併模式設定")
            merge_type = st.radio("請選擇合併方式：", ('縱向合併 (上下堆疊)', '橫向合併 (左右拼接)'), horizontal=True)
            header_row_from_user = st.number_input("所有檔案的標頭 (Header) 都在第幾列？", min_value=1, value=1)
            
            file_configs = {}
            for uploaded_file in uploaded_files:
                try:
                    file_buffer = io.BytesIO(uploaded_file.getvalue())
                    xls = pd.ExcelFile(file_buffer)
                    sheet_names = xls.sheet_names
                    file_configs[uploaded_file.name] = {"file_object": file_buffer, "sheet_names": sheet_names}
                except Exception as e:
                    st.error(f"讀取檔案 '{uploaded_file.name}' 的工作表列表失敗: {e}")
            
            join_key, join_how = "", "inner"

            st.divider()
            st.subheader("2. 檔案與工作表設定")
            all_selected_sheets_info = []
            for filename, config in file_configs.items():
                selected_sheets = st.multiselect(
                    f"檔案: `{filename}` - 請勾選要合併的工作表",
                    options=config["sheet_names"],
                    default=config["sheet_names"][0] if config["sheet_names"] else None,
                    key=f"sheets_{filename}"
                )
                config["selected_sheets"] = selected_sheets
                for sheet in selected_sheets:
                    all_selected_sheets_info.append((config["file_object"], sheet))
            
            if merge_type == '橫向合併 (左右拼接)':
                st.divider()
                st.subheader("3. 橫向合併專用設定")
                
                common_columns_for_key = []
                if all_selected_sheets_info:
                    try:
                        dfs_for_cols = [read_and_clean_sheet(f[0], f[1], header_row_from_user-1) for f in all_selected_sheets_info]
                        if dfs_for_cols:
                            column_sets = [set(df.columns) for df in dfs_for_cols]
                            common_columns_for_key = list(set.intersection(*column_sets))
                    except Exception as e:
                        st.warning(f"計算共同欄位時發生錯誤: {e}")

                if not common_columns_for_key:
                     st.warning("您目前選擇的工作表之間沒有共同欄位，無法進行橫向合併。")
                     join_key = st.text_input("或請手動輸入用來對齊的「共同欄位」名稱 (Key)", disabled=True)
                else:
                    join_key = st.selectbox("請選擇用來對齊的「共同欄位」(Key)", common_columns_for_key)
                
                merge_options_display = {
                    "內連接 (Inner Join) - 只保留所有表中共有的資料": "inner",
                    "外連接 (Outer Join) - 保留所有表中出現過的資料": "outer",
                    "左連接 (Left Join) - 以第一個選擇的表為基礎": "left",
                }
                selected_display = st.selectbox(
                    "選擇合併類型", 
                    options=list(merge_options_display.keys()),
                    help="決定如何處理在不同表中無法對應的資料。"
                )
                join_how = merge_options_display[selected_display]
            
            st.divider()
            st.subheader("4. 其他設定")
            add_source_col = st.checkbox("新增「來源檔案/工作表」欄位 (僅在縱向合併時有效)", value=True)
            submitted = st.form_submit_button("🚀 執行多檔合併", type="primary")

        if submitted:
            st.session_state.final_df = None
            st.session_state.unmatched_df = None
            st.session_state.duplication_warning_keys = []
            st.session_state.merge_key = None
            
            all_dfs_to_merge = []
            with st.spinner('正在讀取所有選定的工作表...'):
                for filename, config in file_configs.items():
                    if "selected_sheets" in config:
                        for sheet_name in config["selected_sheets"]:
                            df = read_and_clean_sheet(config["file_object"], sheet_name, header_row_from_user - 1)
                            if add_source_col and merge_type == '縱向合併 (上下堆疊)':
                                df['來源檔案'] = filename
                                df['來源工作表'] = sheet_name
                            all_dfs_to_merge.append(df)
            
            if not all_dfs_to_merge:
                st.warning("未成功讀取任何工作表。")
            else:
                merged_df = None
                with st.spinner('正在執行合併...'):
                    try:
                        if merge_type == '縱向合併 (上下堆疊)':
                            merged_df = pd.concat(all_dfs_to_merge, ignore_index=True)
                        else: # 橫向合併
                            if not join_key: 
                                st.error("橫向合併錯誤：必須提供「共同欄位」。")
                            elif len(all_dfs_to_merge) < 2: 
                                st.warning("橫向合併至少需要兩個工作表。")
                            else:
                                st.session_state.merge_key = join_key
                                
                                all_duplicated_keys = set()
                                for df in all_dfs_to_merge:
                                    if join_key in df.columns:
                                        duplicates_in_df = df[df.duplicated(subset=[join_key], keep=False)]
                                        if not duplicates_in_df.empty:
                                            all_duplicated_keys.update(duplicates_in_df[join_key].unique())
                                if all_duplicated_keys:
                                    st.session_state.duplication_warning_keys = list(all_duplicated_keys)

                                merged_df = reduce(lambda left, right: pd.merge(left, right, on=join_key, how=join_how), all_dfs_to_merge)
                                
                                duplicated_keys = st.session_state.get('duplication_warning_keys', [])
                                if duplicated_keys:
                                    merged_df['備註'] = ''
                                    condition = merged_df[join_key].isin(duplicated_keys)
                                    merged_df.loc[condition, '備註'] = '一對多關係提醒'

                        if merged_df is not None:
                            st.session_state.final_df = merged_df
                            st.success("🎉 合併成功！")
                    except Exception as e:
                        st.error(f"合併失敗: {e}")

# --- 通用結果顯示區 (已修改) ---
if 'final_df' in st.session_state and st.session_state.final_df is not None:
    st.divider()
    
    # 注入 CSS 樣式
    st.markdown("""
    <style>
        .dataframe-container table {
            width: auto !important;
            border-collapse: collapse;
            text-align: left;
        }
        .dataframe-container th, .dataframe-container td {
            padding: 8px;
            border: 1px solid #ddd;
        }
        .dataframe-container th {
            background-color: #f2f2f2;
        }
    </style>
    """, unsafe_allow_html=True)

    keys_list = st.session_state.get('duplication_warning_keys', [])
    if keys_list:
        st.warning(f"⚠️ **注意：一對多關係提醒**")
        st.info(f"在您選擇的工作表中，以下 {len(keys_list)} 個鍵值存在重複，這可能導致最終結果的行數增加。重複的資料列在下方表格中已用**淡黃色標註**，並在**「備註」**欄位中加以說明。")
        with st.expander("點此查看重複的鍵值列表"):
            st.code("\n".join(map(str, keys_list)))
    
    st.header("✅ 主要合併結果預覽與下載")
    final_df = st.session_state.final_df
    st.info(f"合併結果：共 {final_df.shape[0]} 筆資料，{final_df.shape[1]} 個欄位。")

    # 準備 Styler 物件
    if keys_list:
        duplicated_keys_set = set(keys_list)
        merge_key_for_style = st.session_state.get('merge_key')
        if merge_key_for_style and merge_key_for_style in final_df.columns:
            styler = final_df.style.apply(
                highlight_duplicated_keys, 
                axis=1, 
                duplicated_keys_set=duplicated_keys_set,
                key_column=merge_key_for_style
            )
        else:
            styler = final_df.style
    else:
        styler = final_df.style

    # 將 Styler 物件轉換為 HTML 並顯示
    html = styler.to_html(index=False)
    st.markdown(f'<div class="dataframe-container">{html}</div>', unsafe_allow_html=True)

    excel_data = to_excel(final_df)
    st.download_button(
        label="📥 下載主要合併結果",
        data=excel_data,
        file_name="合併結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_main"
    )

if 'unmatched_df' in st.session_state and st.session_state.unmatched_df is not None:
    unmatched_df = st.session_state.unmatched_df
    if not unmatched_df.empty:
        st.divider()
        expander_title = f"⚠️ 右表中未能比對的資料 ({len(unmatched_df)} 筆) - 點此展開查看"
        with st.expander(expander_title):
            merge_key_for_unmatched = st.session_state.get('merge_key', '共同索引鍵')
            st.warning(f"以下是來自右表的 {len(unmatched_df)} 筆資料，因為它們的「{merge_key_for_unmatched}」在左表中找不到對應項目，所以未能合併。")
            
            # --- START: 同樣修改未比對資料的顯示方式 ---
            unmatched_html = unmatched_df.to_html(index=False)
            st.markdown(f'<div class="dataframe-container">{unmatched_html}</div>', unsafe_allow_html=True)
            # --- END: 同樣修改未比對資料的顯示方式 ---

            unmatched_excel_data = to_excel(unmatched_df)
            st.download_button(
                label="📥 下載這份未比對的資料",
                data=unmatched_excel_data,
                file_name="未能比對的資料.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_unmatched"
            )
