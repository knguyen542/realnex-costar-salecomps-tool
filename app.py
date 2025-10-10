import pandas as pd
import re
import io
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ----------- Hide Streamlit Default Menu / GitHub Link ----------
st.set_page_config(page_title="RealNex CoStar Sale Comps Tool", layout="centered", initial_sidebar_state="collapsed")
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# -------- Helper functions --------
def clean_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^A-Za-z0-9 ]+", "", str(value))

def split_name(full_name):
    if pd.isna(full_name):
        return "", ""
    name = clean_text(full_name).strip()
    parts = name.split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], parts[-1]

def safe_fullname(first, last):
    first = clean_text(first or "")
    last  = clean_text(last  or "")
    return f"{first} {last}".strip() if first and last else (first or last)

# -------- Load Static Reference Files --------
@st.cache_data
def load_reference_files():
    template_df = pd.read_excel("RealNex_Template.xlsx", engine="openpyxl")
    mapping_df = pd.read_excel("Template_CoStar_Alignment_ByData.xlsx", engine="openpyxl")
    mapping_df.columns = [str(c).strip() for c in mapping_df.columns]
    return template_df, mapping_df

template_df, mapping_df = load_reference_files()

# -------- Streamlit UI --------
st.title("📘 RealNex CoStar Sale Comps Import Tool")

st.write("""
Upload your **CoStar Sale Comps export (.xlsx)** below.  
The tool will automatically align your data with RealNex’s standard import format and generate 3 files for download.
""")

costar_file = st.file_uploader("📂 Upload CoStar Sale Comps Export", type=["xlsx"])

if st.button("🚀 Process File"):
    if not costar_file:
        st.error("Please upload your CoStar Sale Comps file.")
    else:
        costar_df = pd.read_excel(costar_file, sheet_name=0, engine="openpyxl")
        final_df = pd.DataFrame(index=costar_df.index)

        # Mapping logic
        for _, row in mapping_df.iterrows():
            tpl_col = str(row["Template Header"]).strip()
            src_val = row["CoStar data header"]

            if pd.notna(src_val):
                src_expr = str(src_val).strip()
                if "+" in src_expr:
                    parts = [p.strip() for p in src_expr.split("+")]
                    valid = [p for p in parts if p in costar_df.columns]
                    if valid:
                        combined = costar_df[valid[0]].fillna("").astype(str)
                        for add_col in valid[1:]:
                            combined = combined.str.cat(costar_df[add_col].fillna("").astype(str), sep=" ")
                        combined = combined.str.replace(r"\s+", " ", regex=True).str.strip()
                        combined = combined.map(clean_text).where(lambda s: s != "", "")
                        final_df[tpl_col] = combined
                    else:
                        final_df[tpl_col] = ""
                elif src_expr in costar_df.columns:
                    final_df[tpl_col] = costar_df[src_expr]
                else:
                    final_df[tpl_col] = ""
            else:
                final_df[tpl_col] = ""

        # --- Generate Outputs ---
        aligned_io = io.BytesIO()
        audit_io = io.BytesIO()
        report_text = io.StringIO()

        final_df.to_excel(aligned_io, index=False, engine="openpyxl")
        aligned_io.seek(0)

        mapping_df.to_excel(audit_io, index=False, engine="openpyxl")
        audit_io.seek(0)

        report_text.write("RealNex CoStar Import – Run Report\n")
        report_text.write("==================================\n\n")
        report_text.write("Processed successfully using built-in RealNex Template and Mapping Sheet.\n")

        # Store in session to keep downloads visible
        st.session_state['aligned'] = aligned_io.getvalue()
        st.session_state['audit'] = audit_io.getvalue()
        st.session_state['report'] = report_text.getvalue()

        st.success("✅ Processing complete! Scroll down to download your files.")

# --- Always show download buttons if session data exists ---
if 'aligned' in st.session_state:
    st.download_button("⬇️ Download Aligned File", st.session_state['aligned'], file_name="aligned.xlsx")
if 'audit' in st.session_state:
    st.download_button("⬇️ Download Mapping Audit", st.session_state['audit'], file_name="mapping_audit.xlsx")
if 'report' in st.session_state:
    st.download_button("⬇️ Download Run Report", st.session_state['report'], file_name="run_report.txt")
