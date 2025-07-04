import streamlit as st
import pandas as pd
import io


# --- 通用輔助函數 ---

def to_excel(df):
    """將 DataFrame 轉換為可供下載的 Excel Bytes 物件"""
    output = io.BytesIO()
    # 使用 openpyxl 引擎，對中文和多樣格式支援較好
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='合併結果')
    processed_data = output.getvalue()
    return processed_data


def read_and_clean_sheet(file_obj, sheet_name, header_index=0):
    """讀取指定的 Excel 工作表並進行基本清理，防止類型錯誤"""
    # 讀取前重置檔案指標，這對於 Streamlit 的運作模式很重要
    file_obj.seek(0)
    df = pd.read_excel(
        file_obj,
        sheet_name=sheet_name,
        header=header_index
    )
    # 將 object 類型欄位轉換為字串，避免 PyArrow 序列化錯誤
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].astype(str)
    return df


# --- Streamlit 應用程式介面 ---

st.set_page_config(page_title="Excel 全能合併工具", page_icon="🧩", layout="wide")

st.title("🧩 Excel 全能合併工具")

# --- 模式選擇 ---
app_mode = st.radio(
    "請選擇您要使用的工具模式：",
    ('多檔合併 (縱向/橫向)', '雙檔查找合併 (VLOOKUP)'),
    horizontal=True,
    label_visibility="collapsed"  # 隱藏標籤，讓介面更簡潔
)

st.divider()

# ==============================================================================
# ======================== 模式一：多檔合併 (縱向/橫向) =========================
# ==============================================================================
if app_mode == '多檔合併 (縱向/橫向)':

    st.header("模式：多檔合併 (縱向/橫向)")
    st.markdown("此模式可合併多個檔案或單一檔案內的多個工作表。")

    uploaded_files = st.file_uploader(
        "請上傳您所有要處理的 Excel 檔案",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key="multi_file_uploader"
    )

    if uploaded_files:
        with st.form("multi_merge_form"):
            st.subheader("1. 合併模式設定")
            merge_type = st.radio(
                "請選擇合併方式：",
                ('縱向合併 (上下堆疊)', '橫向合併 (左右拼接)'),
                horizontal=True
            )

            join_key, join_how = "", "inner"
            if merge_type == '橫向合併 (左右拼接)':
                st.info("橫向合併會將您選定的所有工作表，根據一個「共同欄位」左右拼接。")
                join_key = st.text_input("請輸入用來對齊的「共同欄位」名稱 (Key)")
                join_how = st.selectbox("選擇合併類型", ['inner', 'outer', 'left', 'right'])

            st.divider()

            st.subheader("2. 檔案與工作表設定")
            file_configs = {}
            for uploaded_file in uploaded_files:
                try:
                    # 使用 BytesIO 來避免重複讀取檔案
                    file_buffer = io.BytesIO(uploaded_file.getvalue())
                    xls = pd.ExcelFile(file_buffer)
                    sheet_names = xls.sheet_names

                    selected_sheets = st.multiselect(
                        f"檔案: `{uploaded_file.name}` - 請勾選要合併的工作表 (可多選)",
                        options=sheet_names,
                        default=sheet_names[0] if sheet_names else None,
                        key=f"sheets_{uploaded_file.name}"
                    )

                    if selected_sheets:
                        file_configs[uploaded_file.name] = {
                            "file_object": file_buffer,
                            "selected_sheets": selected_sheets
                        }
                except Exception as e:
                    st.error(f"讀取檔案 '{uploaded_file.name}' 失敗: {e}")

            st.divider()

            st.subheader("3. 通用設定")
            header_row_from_user = st.number_input("標頭 (Header) 在第幾列？", min_value=1, value=1)
            add_source_col = st.checkbox("新增「來源檔案/工作表」欄位 (僅在縱向合併時有效)", value=True)

            submitted = st.form_submit_button("🚀 執行多檔合併", type="primary")

        if submitted:
            actual_header_index = header_row_from_user - 1
            all_dfs_to_merge = []
            error_messages = []

            with st.spinner('正在讀取所有選定的工作表...'):
                for filename, config in file_configs.items():
                    for sheet_name in config["selected_sheets"]:
                        try:
                            df = read_and_clean_sheet(config["file_object"], sheet_name, actual_header_index)
                            if add_source_col and merge_type == '縱向合併 (上下堆疊)':
                                df['來源檔案'] = filename
                                df['來源工作表'] = sheet_name
                            all_dfs_to_merge.append(df)
                        except Exception as e:
                            error_messages.append(f"讀取 '{filename}' 的 '{sheet_name}' 失敗: {e}")

            if error_messages:
                for error in error_messages: st.error(error)

            if not all_dfs_to_merge:
                st.warning("未成功讀取任何工作表。")
            else:
                merged_df = None
                with st.spinner('正在執行合併...'):
                    try:
                        if merge_type == '縱向合併 (上下堆疊)':
                            merged_df = pd.concat(all_dfs_to_merge, ignore_index=True)
                            st.success("🎉 縱向合併成功！")
                        else:  # 橫向合併
                            if not join_key:
                                st.error("橫向合併錯誤：必須提供「共同欄位」。")
                            elif len(all_dfs_to_merge) < 2:
                                st.warning("橫向合併至少需要兩個工作表。")
                            else:
                                merged_df = all_dfs_to_merge[0]
                                for i in range(1, len(all_dfs_to_merge)):
                                    merged_df = pd.merge(merged_df, all_dfs_to_merge[i], on=join_key, how=join_how)
                                st.success("🎉 橫向合併成功！")

                        if merged_df is not None:
                            st.session_state.final_df = merged_df
                    except Exception as e:
                        st.error(f"合併失敗: {e}")

# ==============================================================================
# ===================== 模式二：雙檔查找合併 (VLOOKUP) ========================
# ==============================================================================
elif app_mode == '雙檔查找合併 (VLOOKUP)':

    st.header("模式：雙檔查找合併 (VLOOKUP)")
    st.markdown("""
    此模式會以**左表**為基礎，從**右表**中查找符合條件的資料，並將指定欄位新增至左表。
    """)

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
                cols_to_merge = st.multiselect("選擇要從右表加入到左表的欄位", available_cols_from_right,
                                               default=available_cols_from_right)

                submitted_vlookup = st.form_submit_button("🚀 執行查找合併", type="primary")

            if submitted_vlookup:
                if not merge_key or not cols_to_merge:
                    st.warning("請至少選擇一個索引鍵和一個要合併的欄位。")
                else:
                    with st.spinner("正在合併資料..."):
                        try:
                            df_right_selected = df_right[[merge_key] + cols_to_merge]
                            # 使用 left join，保留左表所有資料
                            merged_df = pd.merge(df_left, df_right_selected, on=merge_key, how='left')
                            st.session_state.final_df = merged_df
                            st.success("🎉 查找合併成功！")
                        except Exception as e:
                            st.error(f"合併失敗: {e}")

# --- 通用結果顯示區 ---
if 'final_df' in st.session_state and st.session_state.final_df is not None:
    st.divider()
    st.header("最終結果預覽與下載")
    final_df = st.session_state.final_df
    st.info(f"合併結果：共 {final_df.shape[0]} 筆資料，{final_df.shape[1]} 個欄位。")
    st.dataframe(final_df)

    excel_data = to_excel(final_df)
    st.download_button(
        label="📥 下載合併後的 Excel 檔案",
        data=excel_data,
        file_name="合併結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
