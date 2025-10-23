import pandas as pd
import re
import io
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ----- Page setup -----
st.set_page_config(page_title="RealNex CoStar Sale Comps Tool", layout="wide", page_icon="🏙️")

# ----- Custom CSS with Skyline Background -----
bg_url = "https://raw.githubusercontent.com/YOUR_USERNAME/realnex-streamlit-tool/main/skyline.jpg"  # use your local image stored in the repo

st.markdown(f"""
    <style>
        body {{
            background: url('{bg_url}') no-repeat center center fixed;
            background-size: cover;
            color: white;
        }}
        .main {{
            background-color: rgba(0, 0, 0, 0.65);
            padding: 2rem;
            border-radius: 15px;
            box-shadow: 0 4px 25px rgba(0,0,0,0.4);
            margin: 3rem auto;
            max-width: 750px;
        }}
        h1 {{
            text-align: center;
            color: #FFFFFF;
            font-size: 2.4rem;
            font-weight: 700;
        }}
        .stButton>button {{
            background-color: #0066CC;  /* new contrasting blue */
            color: white;
            border-radius: 8px;
            font-size: 1.1rem;
            padding: 0.6rem 1.2rem;
            border: none;
        }}
        .stButton>button:hover {{
            background-color: #1E88E5;
            color: #fff;
        }}
        .download-btn button {{
            background-color: #F15A24 !important;  /* RealNex orange for downloads */
            color: white !important;
            border-radius: 8px !important;
        }}
        .download-btn button:hover {{
            background-color: #ff784e !important;
        }}
    </style>
""", unsafe_allow_html=True)

# ----- Helper functions -----
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

# ----- Load static references -----
@st.cache_data
def load_reference_files():
    template_df = pd.read_excel("RealNex_Template.xlsx", engine="openpyxl")
    mapping_df = pd.read_excel("Template_CoStar_Alignment_ByData.xlsx", engine="openpyxl")
    mapping_df.columns = [str(c).strip() for c in mapping_df.columns]
    return template_df, mapping_df

template_df, mapping_df = load_reference_files()

# ----- UI layout -----
with st.container():
    st.markdown("<div class='main'>", unsafe_allow_html=True)
    st.title("🏙️ RealNex CoStar Sale Comps Import Tool")
    st.markdown("""
        Upload your **CoStar Sale Comps export (.xlsx)** below.  
        The tool will automatically align your data to the RealNex import template and generate three ready-to-download files.
    """)

    costar_file = st.file_uploader("📂 Upload CoStar Sale Comps Export", type=["xlsx"])

    if st.button("🚀 Process File"):
        if not costar_file:
            st.error("Please upload your CoStar Sale Comps file.")
        else:
            costar_df = pd.read_excel(costar_file, sheet_name=0, engine="openpyxl")
            final_df = pd.DataFrame(index=costar_df.index)

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

            # Output files
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

            st.session_state['aligned'] = aligned_io.getvalue()
            st.session_state['audit'] = audit_io.getvalue()
            st.session_state['report'] = report_text.getvalue()
            st.success("✅ Processing complete! Scroll down to download your files.")

    if 'aligned' in st.session_state:
        st.markdown("<br><hr><h3 style='text-align:center;'>📎 Download Results</h3>", unsafe_allow_html=True)
        st.download_button("⬇️ Download Aligned File", st.session_state['aligned'], file_name="aligned.xlsx")
        st.download_button("⬇️ Download Mapping Audit", st.session_state['audit'], file_name="mapping_audit.xlsx")
        st.download_button("⬇️ Download Run Report", st.session_state['report'], file_name="run_report.txt")
    st.markdown("</div>", unsafe_allow_html=True)

