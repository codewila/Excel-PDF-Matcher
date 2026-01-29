import streamlit as st
import pandas as pd
import pdfplumber
import PyPDF2
from io import BytesIO
import re
from typing import Dict
from difflib import SequenceMatcher

# Set page configuration
st.set_page_config(
    page_title="Excel-PDF Color Matcher",
    page_icon="🎨",
    layout="wide"
)

# --- CSS for better visibility ---
st.markdown("""
<style>
    .stDataFrame { font-size: 14px; }
    .main { background-color: #f5f7f9; }
</style>
""", unsafe_allow_html=True)

def clean_string_for_comparison(text):
    """
    Comparison ko asan banane ke liye text se spaces aur special characters hatata hai.
    """
    if not text or pd.isna(text):
        return ""
    # Sab lowercase karein, alphanumeric ke ilawa sab hatayein, aur spaces khatam karein
    text = str(text).lower().strip()
    text = re.sub(r'[^a-z0-9]', '', text) 
    return text

def extract_text_from_pdf(pdf_file) -> Dict[int, str]:
    """Extract text from PDF pages"""
    text_dict = {}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    text_dict[i + 1] = text
    except Exception as e:
        st.error(f"Error extracting text with pdfplumber: {e}")
        try:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            for i, page in enumerate(pdf_reader.pages):
                text = page.extract_text()
                if text:
                    text_dict[i + 1] = text
        except Exception as e2:
            st.error(f"Error extracting text: {e2}")
    return text_dict

def check_value_in_pdf(value, pdf_text_dict, threshold=0.8):
    """
    Check if value exists in PDF with flexible matching.
    """
    original_val = str(value).strip()
    
    # Empty values handle karein
    if not original_val or original_val.lower() in ['nan', 'none', '', 'nat', 'n.a', 'na']:
        return False, 'Empty'

    search_lower = original_val.lower()
    search_super_clean = clean_string_for_comparison(original_val)

    for page_num, text in pdf_text_dict.items():
        pdf_text_lower = text.lower()
        pdf_text_super_clean = clean_string_for_comparison(text)
        
        # 1. Exact Match (with spaces)
        if search_lower in pdf_text_lower:
            return True, 'Exact'
        
        # 2. Clean Match (No spaces/punctuation) - Addresses ke liye best hai
        if search_super_clean in pdf_text_super_clean:
            return True, 'Exact'
            
        # 3. Fuzzy Matching (Agar spelling mein thoda farq ho)
        if len(search_lower) > 3:
            # PDF text ko words mein split karein
            words = re.findall(r'\b\w+\b', pdf_text_lower)
            for word in words:
                if len(word) > 3:
                    similarity = SequenceMatcher(None, search_lower, word).ratio()
                    if similarity >= threshold:
                        return True, 'Fuzzy'
                        
    return False, 'No Match'

def generate_excel_with_colors(df, status_df):
    """Generates an Excel file where cells are colored based on status"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Analysis_Result', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Analysis_Result']
        
        # Formats definition
        green_format = workbook.add_format({'bg_color': '#90EE90', 'font_color': '#000000', 'border': 1})
        red_format = workbook.add_format({'bg_color': '#FFB6C1', 'font_color': '#000000', 'border': 1})
        default_format = workbook.add_format({'border': 1})

        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                col_name = df.columns[col_idx]
                status = status_df.at[row_idx, col_name]
                
                # Excel rows are +1 due to header
                if status in ['Exact', 'Fuzzy']:
                    worksheet.write(row_idx + 1, col_idx, value, green_format)
                elif status == 'No Match':
                    worksheet.write(row_idx + 1, col_idx, value, red_format)
                else:
                    val_to_write = "" if pd.isna(value) else value
                    worksheet.write(row_idx + 1, col_idx, val_to_write, default_format)
                        
    return output.getvalue()

def main():
    st.title("📊 Excel-PDF Color Matcher (Robust Version)")
    st.markdown("""
    - 🟢 **Green**: Match Mil Gaya (Exact ya Clean match)
    - 🔴 **Red**: Match Nahi Mila
    - *Tip: Agar sahi data Red ho raha hai, toh sidebar se 'Match Accuracy' thoda kam karein.*
    """)

    with st.sidebar:
        st.header("⚙️ Settings")
        excel_file = st.file_uploader("Upload Excel", type=['xlsx', 'xls', 'csv'])
        pdf_file = st.file_uploader("Upload PDF", type=['pdf'])
        threshold = st.slider("Fuzzy Match Accuracy", 0.5, 1.0, 0.80)
        file_type = st.radio("Excel File Type", ['Excel', 'CSV'])

    if excel_file and pdf_file:
        if st.button("🔍 Start Matching", type="primary", use_container_width=True):
            try:
                if file_type == 'Excel':
                    df = pd.read_excel(excel_file)
                else:
                    df = pd.read_csv(excel_file)

                with st.spinner("PDF Scan ho raha hai..."):
                    pdf_text = extract_text_from_pdf(pdf_file)
                    if not pdf_text:
                        st.error("PDF se text nahi nikal paye. Kya ye scanned image hai?")
                        st.stop()

                    status_df = pd.DataFrame(index=df.index, columns=df.columns)
                    
                    progress_bar = st.progress(0)
                    total_cells = df.size
                    done = 0
                    
                    for col in df.columns:
                        for idx in df.index:
                            val = df.at[idx, col]
                            _, status = check_value_in_pdf(val, pdf_text, threshold)
                            status_df.at[idx, col] = status
                            done += 1
                            if done % 20 == 0:
                                progress_bar.progress(min(done/total_cells, 1.0))
                    
                    progress_bar.progress(1.0)

                    # Preview Table logic
                    def highlight_cells(data):
                        df_colors = pd.DataFrame('', index=data.index, columns=data.columns)
                        for col in data.columns:
                            for idx in data.index:
                                s = status_df.at[idx, col]
                                if s in ['Exact', 'Fuzzy']:
                                    df_colors.at[idx, col] = 'background-color: #90EE90'
                                elif s == 'No Match':
                                    df_colors.at[idx, col] = 'background-color: #FFB6C1'
                        return df_colors

                    st.subheader("👀 Preview Results")
                    st.dataframe(df.style.apply(highlight_cells, axis=None), use_container_width=True)

                    # Download
                    excel_data = generate_excel_with_colors(df, status_df)
                    st.success("Matching Complete!")
                    st.download_button(
                        label="📥 Download Colored Excel Report",
                        data=excel_data,
                        file_name="Matching_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Error: {e}")

if __name__ == "__main__":
    main()
