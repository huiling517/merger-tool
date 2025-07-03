import streamlit as st
import pandas as pd
import io


# --- 核心功能函數 ---

@st.cache_data
def load_excel_sheets(uploaded_file):
    """讀取上傳的 Excel 檔案並返回所有工作表的字典和名稱列表"""
    try:
        file_content = uploaded_file.getvalue()
        xls = pd.ExcelFile(io.BytesIO(file_content))

        all_sheets = {}
        sheet_names = xls.sheet_names

        for sheet_name in sheet_names:
            # 讀取單一工作表
            df = pd.read_excel(xls, sheet_name=sheet_name)

            # --- START: 錯誤修正 ---
            # 為了避免 PyArrow 類型錯誤，將所有 object 類型的欄位轉換為字串。
            # 這是導致您看到錯誤的根本原因。
            for col in df.columns:
                if df[col].dtype == 'object':
                    # 將該欄位中的所有值都轉換成字串格式
                    df[col] = df[col].astype(str)
            # --- END: 錯誤修正 ---

            all_sheets[sheet_name] = df

        return all_sheets, sheet_names
    except Exception as e:
        st.error(f"讀取檔案 {uploaded_file.name} 時發生錯誤：{e}")
        return None, None


# (後續的程式碼保持不變...)
def convert_df_to_excel(df):
    """將 DataFrame 轉換為可供下載的 Excel 檔案（in-memory）"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='合併結果')
    processed_data = output.getvalue()
    return processed_data


# --- Streamlit 介面 (保持不變) ---

st.set_page_config(page_title="Excel 跨檔案合併工具", layout="wide")

st.title("📊 Excel 跨檔案智慧合併工具")
st.markdown("""
這個工具可以幫助您合併**兩個不同 Excel 檔案**中的工作表。
1.  在左側上傳您的**主要檔案**，並選擇要使用的工作表。
2.  在右側上傳您要用來**查找資料的檔案**，並選擇對應的工作表。
3.  設定合併條件後，點擊按鈕即可預覽及下載結果。
""")

if 'merged_df' not in st.session_state:
    st.session_state.merged_df = None

st.header("步驟一：上傳檔案並選擇工作表")

col1, col2 = st.columns(2)

df_left = None
df_right = None
sheet_names_left = []
sheet_names_right = []

with col1:
    st.subheader("主要檔案 (左表)")
    uploaded_file_left = st.file_uploader(
        "請選擇一個 .xlsx 檔案",
        type=["xlsx"],
        key="uploader_left",
        help="這是您要保留所有資料，並新增欄位進來的工作表所在的檔案。"
    )
    if uploaded_file_left:
        all_sheets_left, sheet_names_left = load_excel_sheets(uploaded_file_left)
        if all_sheets_left:
            left_sheet_name = st.selectbox(
                "選擇主要工作表",
                sheet_names_left,
                key="sheet_left"
            )
            df_left = all_sheets_left[left_sheet_name]
            st.write("左表預覽：")
            st.dataframe(df_left.head(), height=200)

with col2:
    st.subheader("查找資料檔案 (右表)")
    uploaded_file_right = st.file_uploader(
        "請選擇一個 .xlsx 檔案",
        type=["xlsx"],
        key="uploader_right",
        help="這是您要從中提取資料，並加入到左表的檔案。"
    )
    if uploaded_file_right:
        all_sheets_right, sheet_names_right = load_excel_sheets(uploaded_file_right)
        if all_sheets_right:
            right_sheet_name = st.selectbox(
                "選擇查找資料的工作表",
                sheet_names_right,
                key="sheet_right"
            )
            df_right = all_sheets_right[right_sheet_name]
            st.write("右表預覽：")
            st.dataframe(df_right.head(), height=200)

if df_left is not None and df_right is not None:
    st.header("步驟二：設定合併條件並執行")

    common_columns = list(set(df_left.columns) & set(df_right.columns))

    if not common_columns:
        st.error("兩個選擇的工作表之間沒有任何共同的欄位名稱，無法進行合併。請檢查欄位名稱是否一致（例如，兩邊都有「職工編號」）。")
    else:
        merge_key = st.selectbox(
            "選擇用來對應的欄位 (共同索引鍵)",
            common_columns,
            help="例如：兩個工作表都有的「職工編號」或「訂單ID」。"
        )

        available_cols_from_right = [col for col in df_right.columns if col != merge_key]
        cols_to_merge = st.multiselect(
            "選擇要從右表加入到左表的欄位",
            available_cols_from_right,
            default=available_cols_from_right,
            help="勾選您想要新增到主要工作表的欄位。"
        )

        if st.button("🚀 執行合併", type="primary"):
            if not merge_key or not cols_to_merge:
                st.warning("請至少選擇一個索引鍵和一個要合併的欄位。")
            else:
                with st.spinner("正在合併資料..."):
                    df_right_selected = df_right[[merge_key] + cols_to_merge]
                    merged_df = pd.merge(
                        df_left,
                        df_right_selected,
                        on=merge_key,
                        how='left'
                    )
                    st.session_state.merged_df = merged_df
                    st.success("合併成功！")

if st.session_state.merged_df is not None:
    st.header("步驟三：預覽與下載結果")
    st.info(f"合併結果：共 {st.session_state.merged_df.shape[0]} 列，{st.session_state.merged_df.shape[1]} 欄。")
    st.dataframe(st.session_state.merged_df)

    excel_data = convert_df_to_excel(st.session_state.merged_df)

    st.download_button(
        label="📥 下載合併後的 Excel 檔案",
        data=excel_data,
        file_name="合併結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
