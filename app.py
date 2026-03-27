"""
Monarch Investment — Vendor Aging Property Matcher
Browser-based desktop application (Streamlit)
"""

import io
import json
import sys
import tempfile
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))
from vendor_matcher_core import run_matcher, CURATED_LOOKUP, save_custom_lookup, load_custom_lookup

_local_path = Path(__file__).parent / "custom_lookup.json"
_cloud_path = Path(tempfile.gettempdir()) / "monarch_custom_lookup.json"
CUSTOM_LOOKUP_PATH = _local_path if _local_path.parent.exists() else _cloud_path

# ── page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Monarch — Vendor Aging Matcher",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] {
    background: #f5f7fa;
    font-family: 'Segoe UI', Arial, sans-serif;
}
.monarch-header {
    background: linear-gradient(135deg, #1F3864 0%, #2E5FA3 100%);
    border-radius: 12px;
    padding: 24px 32px;
    margin-bottom: 24px;
    color: white;
}
.monarch-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
.monarch-header p  { margin: 4px 0 0 0; opacity: .8; font-size: .95rem; }
.stat-row { display: flex; gap: 16px; margin: 20px 0; }
.stat-card {
    flex: 1; border-radius: 10px; padding: 20px 24px;
    text-align: center; color: white; font-weight: 600;
}
.stat-card .num { font-size: 2.4rem; line-height: 1; }
.stat-card .lbl { font-size: .8rem; opacity: .88; margin-top: 4px;
                  text-transform: uppercase; letter-spacing: .5px; }
.card-total   { background: #1F3864; }
.card-matched { background: #2D7D46; }
.card-review  { background: #C07000; }
[data-testid="stFileUploader"] {
    border: 2px dashed #C8D3E8; border-radius: 10px;
    padding: 8px; background: white;
}
div.stButton > button {
    background: #1F3864; color: white; border: none;
    border-radius: 8px; padding: 12px 32px;
    font-size: 1rem; font-weight: 600; width: 100%;
}
div.stButton > button:hover { background: #2E5FA3; }
[data-testid="stDownloadButton"] > button {
    background: #2D7D46 !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    width: 100% !important; font-weight: 600 !important;
    font-size: 1rem !important; padding: 12px !important;
}
.banner { border-radius: 8px; padding: 14px 20px; margin: 12px 0; font-weight: 500; }
.banner-success { background: #E8F5E9; border-left: 4px solid #2D7D46; color: #1B5E20; }
.banner-warning { background: #FFF8E1; border-left: 4px solid #F9A825; color: #7B4F00; }
.banner-saved   { background: #E3F2FD; border-left: 4px solid #1565C0; color: #0D47A1; }
</style>
""", unsafe_allow_html=True)

# ── header ───────────────────────────────────────────────────────────────────
custom_lookup = load_custom_lookup(CUSTOM_LOOKUP_PATH)
total_known = len(CURATED_LOOKUP) + len(custom_lookup)

st.markdown("""
<div class="monarch-header">
  <h1>🏢 Vendor Aging — Property Matcher</h1>
  <p>Upload a vendor aging report and the Monarch property list to automatically match property names to Yardi pcodes.</p>
</div>
""", unsafe_allow_html=True)

# ── sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    st.markdown("---")
    st.markdown("**Property Master List**")
    st.caption("Upload an updated PropertyList.xlsx whenever Monarch acquires new properties.")
    prop_file = st.file_uploader("PropertyList.xlsx", type=["xlsx"],
                                 key="prop_upload", label_visibility="collapsed")
    st.markdown("---")
    st.markdown("**Match Confidence Threshold**")
    threshold = st.slider("Fuzzy match score (0–100)", min_value=50,
                          max_value=100, value=75, step=5,
                          help="Matches below this score are flagged. 75 recommended.")
    st.markdown("---")
    st.markdown("**Restore Saved Pcodes**")
    st.caption("Upload a previously downloaded lookup table to restore your confirmed pcodes across sessions.")
    uploaded_lookup = st.file_uploader(
        "monarch_custom_lookup.json", type=["json"],
        key="lookup_upload", label_visibility="collapsed"
    )
    if uploaded_lookup is not None:
        try:
            restored = json.loads(uploaded_lookup.read())
            save_custom_lookup(restored, CUSTOM_LOOKUP_PATH)
            st.session_state["custom_lookup_json"] = json.dumps(restored, indent=2)
            st.success(f"✅ Restored {len(restored)} confirmed pcodes.")
            custom_lookup = restored
        except Exception as e:
            st.error(f"Could not read lookup file: {e}")
    st.markdown("---")
    st.markdown(
        f"<small style='color:#888'>Monarch Investment Management<br>"
        f"Curated lookup: <b>{len(CURATED_LOOKUP)}</b> built-in entries<br>"
        f"Custom confirmed: <b>{len(custom_lookup)}</b> saved entries</small>",
        unsafe_allow_html=True,
    )

# ── upload ────────────────────────────────────────────────────────────────────
col_upload, col_info = st.columns([3, 2], gap="large")
with col_upload:
    st.markdown("### 📂 Step 1 — Upload Vendor Aging Report")
    vendor_file = st.file_uploader("Drop your vendor aging .xlsx here", type=["xlsx"],
                                   key="vendor_upload", label_visibility="collapsed")
with col_info:
    st.markdown("### ℹ️ How it works")
    st.markdown("""
1. **Upload** the vendor's aging report
2. **Upload** PropertyList.xlsx in the sidebar
3. Click **Run Matching**
4. Confirm any flagged pcodes and **save** them to the lookup table
5. **Download** the matched Excel
""")

st.markdown("---")

# ── run ───────────────────────────────────────────────────────────────────────
st.markdown("### ▶️ Step 2 — Run Matching")
run_col, _ = st.columns([2, 3])
with run_col:
    run_btn = st.button("🔍  Match Properties", use_container_width=True)

if run_btn:
    if vendor_file is None:
        st.error("Please upload a vendor aging report first.")
    elif prop_file is None:
        st.error("Please upload the PropertyList.xlsx in the sidebar.")
    else:
        with st.spinner("Matching property names…"):
            try:
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as vf:
                    vf.write(vendor_file.read())
                    vendor_path = vf.name
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as pf:
                    pf.write(prop_file.read())
                    prop_path = pf.name

                output_buf = io.BytesIO()
                df_result, review_df, n_total, n_review = run_matcher(
                    vendor_path, prop_path, output_buf,
                    fuzzy_threshold=threshold,
                    custom_lookup_path=CUSTOM_LOOKUP_PATH,
                )
                output_buf.seek(0)

                st.session_state["result_df"]   = df_result
                st.session_state["review_df"]   = review_df
                st.session_state["output_buf"]  = output_buf.read()
                st.session_state["n_total"]     = n_total
                st.session_state["n_review"]    = n_review
                st.session_state["vendor_name"] = Path(vendor_file.name).stem
                st.session_state["ran"]         = True
                st.session_state.pop("saved_msg", None)

            except Exception as e:
                st.error(f"Something went wrong: {e}")
                st.code(traceback.format_exc())
                st.session_state["ran"] = False

# ── results ───────────────────────────────────────────────────────────────────
if st.session_state.get("ran"):
    n_total     = st.session_state["n_total"]
    n_review    = st.session_state["n_review"]
    n_matched   = n_total - n_review
    vendor_name = st.session_state["vendor_name"]
    review_df   = st.session_state["review_df"]

    st.markdown("---")
    st.markdown("### ✅ Step 3 — Results")

    # KPI cards — n_review is the TRUE row count (all flagged invoice rows)
    n_unique_flagged = review_df["Customer"].nunique() if not review_df.empty else 0
    st.markdown(f"""
<div class="stat-row">
  <div class="stat-card card-total">
    <div class="num">{n_total}</div>
    <div class="lbl">Total Invoice Rows</div>
  </div>
  <div class="stat-card card-matched">
    <div class="num">{n_matched}</div>
    <div class="lbl">Confirmed Matches</div>
  </div>
  <div class="stat-card card-review">
    <div class="num">{n_review}</div>
    <div class="lbl">Rows Needing Review<br><span style="font-size:.75rem;opacity:.8">({n_unique_flagged} unique property names)</span></div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── Needs Review section ─────────────────────────────────────────────────
    if n_review > 0:
        st.markdown(f"""
<div class="banner banner-warning">
⚠️ <b>{n_review} invoice rows</b> across <b>{n_unique_flagged} unique property names</b> could not be confidently matched.
Type the correct Yardi pcode in the <b>Confirmed Pcode</b> column below, then click <b>Save to Lookup Table</b>.
Saved entries are permanent — they will match automatically on every future run.
</div>
""", unsafe_allow_html=True)

        # Build editable table — ALL flagged rows with Invoice #
        # Identify the invoice column name
        result_df = st.session_state["result_df"]
        invoice_col = next((c for c in result_df.columns
                            if "invoice" in c.lower()), None)
        customer_col = next((c for c in result_df.columns
                             if "customer" in c.lower()), None)

        # Pull all flagged rows from the full result (not deduplicated)
        flagged_rows = result_df[result_df["Needs Review"] == True].copy()

        # Build display df
        display_cols = []
        if invoice_col:
            display_cols.append(invoice_col)
        display_cols += [customer_col, "Matched Pcode", "Matched Property Name",
                         "Match Confidence"]
        display_df = flagged_rows[display_cols].copy().reset_index(drop=True)

        # Rename for cleaner headers
        rename_map = {
            invoice_col:            "Invoice #",
            customer_col:           "Vendor Property Name",
            "Matched Pcode":        "Best Guess Pcode",
            "Matched Property Name":"Best Guess Property Name",
            "Match Confidence":     "Confidence",
        }
        display_df = display_df.rename(columns={k: v for k, v in rename_map.items() if k in display_df.columns})

        # Add editable confirmed pcode column (pre-fill with best guess)
        display_df["Confirmed Pcode"] = display_df["Best Guess Pcode"]

        # Editable table
        edited_df = st.data_editor(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Invoice #":             st.column_config.TextColumn("Invoice #",        width="medium",  disabled=True),
                "Vendor Property Name":  st.column_config.TextColumn("Vendor Name",      width="large",   disabled=True),
                "Best Guess Pcode":      st.column_config.TextColumn("Best Guess Pcode", width="small",   disabled=True),
                "Best Guess Property Name": st.column_config.TextColumn("Best Guess Match", width="large", disabled=True),
                "Confidence":            st.column_config.TextColumn("Confidence",       width="small",   disabled=True),
                "Confirmed Pcode":       st.column_config.TextColumn(
                    "✏️ Confirmed Pcode",
                    width="medium",
                    help="Type the correct Yardi pcode here. Leave blank to skip.",
                ),
            },
            key="review_editor",
        )

        # Save button
        save_col, dl_col2 = st.columns([2, 3])
        with save_col:
            if st.button("💾  Save Confirmed Pcodes to Lookup Table", use_container_width=True):
                # Read edits directly from Streamlit's session state for the data_editor.
                # This is more reliable than using the edited_df return value, which can
                # miss edits when the user hasn't tabbed/clicked out of the cell first.
                editor_state = st.session_state.get("review_editor", {})
                edited_rows = editor_state.get("edited_rows", {})  # {str(row_idx): {col: val}}

                existing = load_custom_lookup(CUSTOM_LOOKUP_PATH)
                new_entries = 0

                if not edited_rows:
                    st.warning("No edits detected. Type the correct pcode in the '✏️ Confirmed Pcode' column, press Tab to commit the cell, then click Save.")
                else:
                    for row_idx_str, changes in edited_rows.items():
                        if "Confirmed Pcode" not in changes:
                            continue
                        pcode = str(changes["Confirmed Pcode"]).strip()
                        if not pcode or pcode.lower() in ("unknown", "nan", "none", ""):
                            continue
                        try:
                            row_idx = int(row_idx_str)
                            vendor_name = display_df.iloc[row_idx]["Vendor Property Name"].strip()
                            best_guess_name = display_df.iloc[row_idx].get("Best Guess Property Name", "")
                        except (IndexError, KeyError):
                            continue
                        vendor_name_key = vendor_name.upper()
                        existing[vendor_name_key] = {
                            "pcode": pcode,
                            "official_name": best_guess_name,
                            "confidence": "HIGH",
                            "needs_review": False,
                            "confirmed_by": "user",
                            "confirmed_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        }
                        new_entries += 1

                    if new_entries == 0:
                        st.warning("Pcodes were blank or 'UNKNOWN' — enter the correct Yardi pcode (e.g. hrmo) and try again.")
                    else:
                        save_custom_lookup(existing, CUSTOM_LOOKUP_PATH)
                        st.session_state["custom_lookup_json"] = json.dumps(existing, indent=2)
                        st.session_state["saved_msg"] = f"✅ {new_entries} pcode(s) saved to the lookup table. They will match automatically on all future runs in this session. Use the download button to keep them permanently."
                        st.rerun()

        if st.session_state.get("saved_msg"):
            st.markdown(f"""
<div class="banner banner-saved">
{st.session_state['saved_msg']}
</div>
""", unsafe_allow_html=True)
            # Offer download of the lookup table JSON for persistence across sessions
            lookup_json = st.session_state.get("custom_lookup_json") or json.dumps(
                load_custom_lookup(CUSTOM_LOOKUP_PATH), indent=2
            )
            if lookup_json and lookup_json != "{}":
                st.download_button(
                    label="⬇  Download Lookup Table (to restore on next session)",
                    data=lookup_json,
                    file_name="monarch_custom_lookup.json",
                    mime="application/json",
                    help="Save this file and re-upload it in the sidebar next time to restore your confirmed pcodes.",
                    key="dl_lookup",
                )

    else:
        st.markdown("""
<div class="banner banner-success">
✅  All property names matched with high confidence — no manual review needed!
</div>
""", unsafe_allow_html=True)

    # Preview
    with st.expander("📋 Preview matched data (first 20 rows)", expanded=False):
        result_df = st.session_state["result_df"]
        preview_cols = [c for c in ["Invoice #", customer_col if "customer_col" in dir() else "Customer",
                        "Matched Pcode", "Matched Property Name",
                        "Match Confidence", "Needs Review", "Grand Total"]
                        if c in result_df.columns]
        st.dataframe(result_df[preview_cols].head(20),
                     use_container_width=True, hide_index=True)

    # Download
    st.markdown("---")
    st.markdown("### 💾 Step 4 — Download")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    dl_col, _ = st.columns([2, 3])
    with dl_col:
        st.download_button(
            label=f"⬇  Download  {vendor_name}_Matched.xlsx",
            data=st.session_state["output_buf"],
            file_name=f"{vendor_name}_Matched_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
