import pandas as pd
import re
import io
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

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
    last  = clean_text(last or "")
    return f"{first} {last}".strip() if first and last else (first or last)

def highlight_blanks(audit_file):
    wb = load_workbook(audit_file)
    ws = wb.active
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    header_row = [cell.value for cell in ws[1]]
    if "status" in header_row:
        status_idx = header_row.index("status") + 1
        for r in range(2, ws.max_row + 1):
            if ws.cell(row=r, column=status_idx).value == "BLANK (no match)":
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = red_fill
    wb.save(audit_file)

# -------- Streamlit UI --------
st.set_page_config(page_title="RealNex CoStar Sale Comps Tool", layout="centered")
st.title("📘 RealNex CoStar Sale Comps Import Tool")

st.write("""
Upload your **CoStar Sale Comps export**, the **RealNex template**, and the **mapping sheet**.  
The tool will align columns and generate 3 outputs you can download instantly.
""")

costar_file = st.file_uploader("📂 Upload CoStar Sale Comps Export (.xlsx)", type=["xlsx"])
template_file = st.file_uploader("📂 Upload RealNex Template (.xlsx)", type=["xlsx"])
mapping_file = st.file_uploader("📂 Upload Mapping Sheet (.xlsx)", type=["xlsx"])

if st.button("🚀 Process Files"):

    if not (costar_file and template_file and mapping_file):
        st.error("Please upload all 3 files before processing.")
    else:
        # Read files
        costar_df = pd.read_excel(costar_file, sheet_name=0, engine="openpyxl")
        mapping_df = pd.read_excel(mapping_file, sheet_name=0, engine="openpyxl")

        # Normalize headers
        mapping_df.columns = [str(c).strip() for c in mapping_df.columns]

        final_df = pd.DataFrame(index=costar_df.index)
        audit_rows = []

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

        # Auto name logic
        name_cols = [c for c in final_df.columns if "Name" in c]
        for col in name_cols:
            if "Full Name" in col:
                continue
            if "First" not in col and "Last" not in col:
                firsts, lasts = zip(*final_df[col].map(split_name))
                final_df[col.replace("Name", "First name")] = firsts
                final_df[col.replace("Name", "Last name")]  = lasts
                final_df[col.replace("Name", "Full Name")]  = [f"{f} {l}".strip() for f, l in zip(firsts, lasts)]
            elif "First" in col:
                base = col.replace("First name", "").strip()
                last_col = base + "Last name"
                if last_col in final_df.columns:
                    final_df[base + "Full Name"] = [
                        " ".join(x for x in [clean_text(f or ""), clean_text(l or "")] if x).strip()
                        for f, l in zip(final_df[col], final_df[last_col])
                    ]

        if "Buyers Broker Agent First Name" in final_df.columns and "Buyers Broker Agent Last Name" in final_df.columns:
            final_df["Proc Agent.Name"] = [
                safe_fullname(f, l)
                for f, l in zip(final_df["Buyers Broker Agent First Name"], final_df["Buyers Broker Agent Last Name"])
            ]
        if "Listing Broker Agent First Name" in final_df.columns and "Listing Broker Agent Last Name" in final_df.columns:
            final_df["List Agent.Name"] = [
                safe_fullname(f, l)
                for f, l in zip(final_df["Listing Broker Agent First Name"], final_df["Listing Broker Agent Last Name"])
            ]

        for c in final_df.columns:
            if "Company" in c:
                final_df[c] = final_df[c].map(clean_text)

        # --------- Create output files in memory ---------
        aligned_io = io.BytesIO()
        audit_io = io.BytesIO()
        report_io = io.StringIO()

        final_df.to_excel(aligned_io, index=False, engine="openpyxl")
        aligned_io.seek(0)

        pd.DataFrame(audit_rows).to_excel(audit_io, index=False, engine="openpyxl")
        audit_io.seek(0)

        report_io.write("RealNex CoStar Import – Run Report\n")
        report_io.write("==================================\n\n")
        report_io.write("Processed successfully!\n")

        # --------- Download buttons ---------
        st.success("✅ Processing complete! Download your files below:")

        st.download_button("⬇️ Download Aligned File", aligned_io, file_name="aligned.xlsx")
        st.download_button("⬇️ Download Mapping Audit", audit_io, file_name="mapping_audit.xlsx")
        st.download_button("⬇️ Download Run Report", report_io.getvalue(), file_name="run_report.txt")
