import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="CTS Billable Hours",
    page_icon="🕐",
    layout="wide"
)

st.title("🕐 Therapist Billable Hours")
st.caption("Upload the **Charges Entered By User** report to calculate monthly billable hours per therapist.")

# ── Settings ──────────────────────────────────────────────────────────────────
with st.expander("⚙️ Calculation Settings", expanded=False):
    col1, col2 = st.columns(2)
    with col1:
        rate = st.number_input("Hourly therapy rate ($)", value=180, min_value=1)
    with col2:
        eval_amounts_input = st.text_input("Eval charge amounts (comma-separated)", value="200, 280")

eval_amounts = []
for v in eval_amounts_input.split(","):
    try:
        eval_amounts.append(float(v.strip()))
    except ValueError:
        pass

# ── Helpers ───────────────────────────────────────────────────────────────────
def is_cpt_code(val):
    """Match standard 5-digit CPT codes (97530) and alphanumeric codes (S9152)."""
    if val is None:
        return False
    return bool(re.match(r'^\s*[A-Z]?\d{4,5}\s*$', str(val)))

def visit_minutes_to_hours(mins):
    """Apply 8-minute rule per individual visit."""
    try:
        mins = float(mins)
    except (ValueError, TypeError):
        return 0.0
    units = int(mins // 15) + (1 if mins % 15 >= 8 else 0)
    return units * 0.25

def parse_report(df):
    """Parse the raw XLS report into per-therapist totals."""
    therapist_data = {}
    report_period = ""

    for i, row in df.iterrows():
        row_vals = row.tolist()

        # Grab report period
        for cell in row_vals:
            cell_str = str(cell) if cell is not None else ""
            if "Service Date Range:" in cell_str:
                m = re.search(r'(\d+/\d+/\d+\s*-\s*\d+/\d+/\d+)', cell_str)
                if m:
                    report_period = m.group(1).strip()

        # Find CPT column
        cpt_col = None
        for c, val in enumerate(row_vals):
            if is_cpt_code(val):
                cpt_col = c
                break
        if cpt_col is None:
            continue

        # Column layout relative to CPT:
        # +2 = minutes, +7 = treating provider, +10 = charge amount
        try:
            treating = str(row_vals[cpt_col + 7] or "").strip()
            charge_raw = str(row_vals[cpt_col + 10] or "")
            mins_raw = str(row_vals[cpt_col + 2] or "")

            charge = float(re.sub(r'[^0-9.\-]', '', charge_raw))
            mins = float(re.sub(r'[^0-9.\-]', '', mins_raw)) if mins_raw.strip() else 0.0
        except (IndexError, ValueError):
            continue

        if not treating or charge <= 0:
            continue

        if treating not in therapist_data:
            therapist_data[treating] = {
                "total_billed": 0.0,
                "eval_total": 0.0,
                "eval_count": 0,
                "eval_hours": 0.0,
            }

        d = therapist_data[treating]
        d["total_billed"] += charge

        if charge in eval_amounts:
            d["eval_total"] += charge
            d["eval_count"] += 1
            d["eval_hours"] += visit_minutes_to_hours(mins)

    results = []
    for name, d in therapist_data.items():
        therapy_hours = (d["total_billed"] - d["eval_total"]) / rate
        total_hours = therapy_hours + d["eval_hours"]
        results.append({
            "Therapist": name,
            "Total Billed": d["total_billed"],
            "Eval $": d["eval_total"],
            "Evals": d["eval_count"],
            "Therapy Hrs": round(therapy_hours, 2),
            "Eval Hrs": round(d["eval_hours"], 2),
            "Total Hrs": round(total_hours, 2),
        })

    results.sort(key=lambda x: x["Total Hrs"], reverse=True)
    return results, report_period

# ── File Upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload report (.xls or .xlsx)",
    type=["xls", "xlsx"],
    help="Export the Charges Entered By User report from your EMR system"
)

if uploaded_file:
    try:
        engine = "xlrd" if uploaded_file.name.endswith(".xls") else "openpyxl"
        df = pd.read_excel(uploaded_file, engine=engine, header=None)
        results, report_period = parse_report(df)

        if not results:
            st.error("No charge data found. Please check that this is the Charges Entered By User report.")
        else:
            # ── Header ────────────────────────────────────────────────────────
            if report_period:
                st.subheader(f"Results — {report_period}")
            else:
                st.subheader("Results")

            # ── Summary metrics ───────────────────────────────────────────────
            total_hrs = sum(r["Total Hrs"] for r in results)
            avg_hrs = total_hrs / len(results) if results else 0

            m1, m2, m3 = st.columns(3)
            m1.metric("Therapists", len(results))
            m2.metric("Total Hours", f"{total_hrs:.1f}")
            m3.metric("Average Hours", f"{avg_hrs:.1f}")

            st.divider()

            # ── Results table ─────────────────────────────────────────────────
            results_df = pd.DataFrame(results)

            # Format display
            display_df = results_df.copy()
            display_df["Total Billed"] = display_df["Total Billed"].apply(lambda x: f"${x:,.2f}")
            display_df["Eval $"] = display_df["Eval $"].apply(lambda x: f"${x:,.2f}" if x > 0 else "—")
            display_df["Evals"] = display_df["Evals"].apply(lambda x: str(x) if x > 0 else "—")
            display_df["Eval Hrs"] = display_df["Eval Hrs"].apply(lambda x: f"{x:.2f}" if x > 0 else "—")
            display_df["Therapy Hrs"] = display_df["Therapy Hrs"].apply(lambda x: f"{x:.2f}")
            display_df["Total Hrs"] = display_df["Total Hrs"].apply(lambda x: f"{x:.2f}")

            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Total Hrs": st.column_config.TextColumn("Total Hrs", width="small"),
                }
            )

            st.caption("Eval hours apply the 8-minute rule per individual visit. Therapy hours = (total billed − eval $) ÷ rate.")

            # ── Export ────────────────────────────────────────────────────────
            csv = results_df.to_csv(index=False)
            filename = f"billable_hours_{report_period.replace('/', '-').replace(' ', '') if report_period else 'export'}.csv"
            st.download_button(
                label="⬇️ Download CSV",
                data=csv,
                file_name=filename,
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"Could not read file: {e}")

else:
    st.info("Upload a report above to get started.")
