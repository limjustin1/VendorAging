"""
Monarch Investment — Vendor Aging Matcher  (Streamlit UI)
"""

import io
import json
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from vendor_matcher_core import (
    run_matcher,
    load_custom_lookup,
    save_custom_lookup,
    normalize,
    VENDOR_NAMES,
    get_vendor_lookup_filename,
)

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
BASE_DIR         = Path(__file__).parent
PROP_LIST_PATH   = BASE_DIR / "PropertyList (9).xlsx"
CUSTOM_LOOKUP_DIR = BASE_DIR / "custom_lookups"
CUSTOM_LOOKUP_DIR.mkdir(exist_ok=True)


def get_lookup_path(vendor_name: str) -> Path:
    """Return the filesystem path for a vendor's custom lookup JSON."""
    return CUSTOM_LOOKUP_DIR / get_vendor_lookup_filename(vendor_name)


# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Monarch Vendor Aging Matcher",
    page_icon="🏢",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Sidebar — vendor lookup management
# ---------------------------------------------------------------------------
with st.sidebar:
    st.image("https://i.imgur.com/0y9nGCN.png", width=180) if False else None
    st.markdown("## 🏢 Monarch Vendor Aging")
    st.markdown("---")

    st.markdown("### Vendor Lookup Tables")
    st.caption(
        "Each vendor has its own lookup table. Download after confirming pcodes "
        "so you can restore them on the next session."
    )

    sidebar_vendor = st.selectbox(
        "Manage lookup for vendor:",
        VENDOR_NAMES,
        key="sidebar_vendor",
    )

    sidebar_lookup_path = get_lookup_path(sidebar_vendor)
    sidebar_lookup      = load_custom_lookup(sidebar_lookup_path)

    st.caption(f"Stored entries: **{len(sidebar_lookup)}**")

    # Download button
    if sidebar_lookup:
        lookup_json = json.dumps(sidebar_lookup, indent=2, ensure_ascii=False)
        st.download_button(
            label=f"⬇  Download {sidebar_vendor} lookup",
            data=lookup_json,
            file_name=get_vendor_lookup_filename(sidebar_vendor),
            mime="application/json",
            key="dl_sidebar_lookup",
        )

    # Upload / restore
    st.markdown("**Restore saved lookup:**")
    uploaded_lookup = st.file_uploader(
        f"Upload {get_vendor_lookup_filename(sidebar_vendor)}",
        type=["json"],
        key="lookup_upload",
        label_visibility="collapsed",
    )
    if uploaded_lookup is not None:
        try:
            restored = json.loads(uploaded_lookup.read())
            save_custom_lookup(restored, sidebar_lookup_path)
            st.success(f"✅ Restored {len(restored)} pcodes for {sidebar_vendor}.")
        except Exception as e:
            st.error(f"Failed to restore: {e}")

    st.markdown("---")
    st.markdown("### Settings")
    fuzzy_threshold = st.slider("Fuzzy match threshold", 60, 95, 75, 5,
                                help="Lower = more matches auto-confirmed; Higher = stricter.")

# ---------------------------------------------------------------------------
# Main layout
# ---------------------------------------------------------------------------
st.title("Vendor Aging → Property Code Matcher")
st.markdown(
    "Upload a vendor aging report and the Monarch property list to automatically "
    "map each property name to its Yardi property code."
)

# ---- Step 1: File uploads + vendor selection --------------------------------
st.markdown("---")
col_left, col_right = st.columns([1, 1])

with col_left:
    st.subheader("Step 1 — Select Vendor & Upload Files")

    vendor_name = st.selectbox(
        "Which vendor is this report from?",
        VENDOR_NAMES,
        key="vendor_select",
        help="Selecting the correct vendor loads its saved lookup table and uses the "
             "correct column mapping for that file format.",
    )

    vendor_file = st.file_uploader(
        f"Upload {vendor_name} aging report (.xlsx)",
        type=["xlsx"],
        key="vendor_upload",
    )

    prop_file = st.file_uploader(
        "Upload Monarch property list (.xlsx) — or leave blank to use saved copy",
        type=["xlsx"],
        key="prop_upload",
    )

with col_right:
    st.subheader("About the selected vendor")
    from vendor_matcher_core import VENDOR_CONFIGS
    cfg = VENDOR_CONFIGS.get(vendor_name, {})
    if cfg:
        st.markdown(
            f"**Sheet:** `{cfg.get('sheet', 'auto-detect')}`  \n"
            f"**Property column:** `{cfg.get('prop_col', 'auto-detect')}`  \n"
            f"**Invoice column:** `{cfg.get('invoice_col', 'auto-detect')}`"
        )
    lookup_path = get_lookup_path(vendor_name)
    current_lookup = load_custom_lookup(lookup_path)
    st.metric("Confirmed pcodes stored", len(current_lookup))
    if current_lookup:
        st.caption("Top stored entries:")
        for k, v in list(current_lookup.items())[:5]:
            st.caption(f"• {k} → `{v['pcode']}`")


# ---- Step 2: Run matching --------------------------------------------------
st.markdown("---")
run_btn = st.button("▶  Run Matching", type="primary", disabled=(vendor_file is None))

if run_btn and vendor_file is not None:
    # Resolve property list path
    if prop_file is not None:
        prop_source = prop_file
    elif PROP_LIST_PATH.exists():
        prop_source = str(PROP_LIST_PATH)
    else:
        st.error("No property list found. Please upload one.")
        st.stop()

    with st.spinner(f"Matching {vendor_name} report…"):
        try:
            output_buf = io.BytesIO()
            df_result, review_df, n_total, n_review = run_matcher(
                vendor_file,
                prop_source,
                output_buf,
                fuzzy_threshold=fuzzy_threshold,
                custom_lookup_path=str(lookup_path),
                vendor_name=vendor_name,
            )
            output_buf.seek(0)
            st.session_state["match_result"]    = df_result
            st.session_state["review_df"]       = review_df
            st.session_state["output_bytes"]    = output_buf.read()
            st.session_state["n_total"]         = n_total
            st.session_state["n_review"]        = n_review
            st.session_state["vendor_name"]     = vendor_name
            st.session_state["lookup_path"]     = str(lookup_path)
            # Identify prop_col used
            from vendor_matcher_core import VENDOR_CONFIGS
            vcfg = VENDOR_CONFIGS.get(vendor_name, {})
            st.session_state["prop_col"] = vcfg.get("prop_col", "Customer")
        except Exception as e:
            st.error(f"Matching failed: {e}")
            st.stop()

# ---- Show results ----------------------------------------------------------
if "match_result" in st.session_state:
    df_result  = st.session_state["match_result"]
    review_df  = st.session_state["review_df"]
    n_total    = st.session_state["n_total"]
    n_review   = st.session_state["n_review"]
    active_vendor = st.session_state.get("vendor_name", vendor_name)
    active_lookup_path = st.session_state.get("lookup_path", str(lookup_path))
    prop_col   = st.session_state.get("prop_col", "Customer")

    st.markdown("---")
    st.subheader(f"Results — {active_vendor}")

    m1, m2, m3 = st.columns(3)
    m1.metric("Total rows", n_total)
    m2.metric("Matched (no review needed)", n_total - n_review)
    m3.metric("Flagged for review", n_review,
              delta=f"{n_review / max(n_total, 1) * 100:.1f}%",
              delta_color="inverse")

    # Download matched Excel
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    dl_fname = f"{active_vendor}_Matched_{ts}.xlsx"
    st.download_button(
        label="⬇  Download Matched Excel",
        data=st.session_state["output_bytes"],
        file_name=dl_fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_excel",
    )

    # ---- Step 3: Review flagged rows ---------------------------------------
    if n_review > 0:
        st.markdown("---")
        st.subheader(f"Step 3 — Review & Confirm Flagged Properties ({n_review} rows)")
        st.caption(
            "Edit the **Confirmed Pcode** column for any property you can identify, "
            "then click **Save to Lookup Table**. Confirmed pcodes are stored in the "
            f"**{active_vendor}** lookup and will auto-match on future runs."
        )

        # Build display DataFrame for the editor
        # Deduplicate by vendor property name — show one row per unique name
        if prop_col not in review_df.columns:
            # fallback: find a likely column
            prop_col = review_df.columns[0]

        display_df = (
            review_df
            .drop_duplicates(subset=[prop_col])
            .reset_index(drop=True)
            .rename(columns={
                prop_col: "Vendor Property Name",
                "Matched Pcode": "Best Guess Pcode",
                "Matched Property Name": "Best Guess Property Name",
            })
        )

        display_df["Confirmed Pcode"] = ""

        editable_cols = {
            "Vendor Property Name":   st.column_config.TextColumn(disabled=True),
            "Best Guess Pcode":       st.column_config.TextColumn(disabled=True),
            "Best Guess Property Name": st.column_config.TextColumn(disabled=True),
            "Match Confidence":       st.column_config.TextColumn(disabled=True),
            "Match Method":           st.column_config.TextColumn(disabled=True),
            "Confirmed Pcode":        st.column_config.TextColumn(
                                          help="Type the correct Yardi pcode here",
                                          required=False),
        }

        keep_cols = [c for c in ["Vendor Property Name", "Best Guess Pcode",
                                  "Best Guess Property Name", "Match Confidence",
                                  "Match Method", "Confirmed Pcode"]
                     if c in display_df.columns or c == "Confirmed Pcode"]

        st.data_editor(
            display_df[keep_cols],
            column_config=editable_cols,
            use_container_width=True,
            num_rows="fixed",
            key="review_editor",
        )

        if st.button("💾  Save to Lookup Table", type="secondary"):
            editor_state = st.session_state.get("review_editor", {})
            edited_rows  = editor_state.get("edited_rows", {})

            existing = load_custom_lookup(active_lookup_path)
            saved_count = 0

            for row_idx_str, changes in edited_rows.items():
                if "Confirmed Pcode" not in changes:
                    continue
                pcode = str(changes["Confirmed Pcode"]).strip()
                if not pcode or pcode.lower() in ("unknown", "nan", "none", ""):
                    continue

                row_idx   = int(row_idx_str)
                vend_name = display_df.iloc[row_idx]["Vendor Property Name"].strip()
                best_name = display_df.iloc[row_idx].get("Best Guess Property Name", "")
                norm_key  = normalize(vend_name)

                existing[norm_key] = {
                    "pcode":        pcode,
                    "official_name": best_name,
                    "confidence":   "HIGH",
                    "needs_review": False,
                    "confirmed_by": "user",
                    "confirmed_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "vendor":       active_vendor,
                }
                saved_count += 1

            if saved_count:
                save_custom_lookup(existing, active_lookup_path)
                st.success(
                 f"✅ {saved_count} pcode(s) saved to **{active_vendor}** lookup table."
                )
                # Offer download of the updated lookup
                lookup_json = json.dumps(existing, indent=2, ensure_ascii=False)
                st.download_button(
                    label=f"⬇  Download updated {active_vendor} lookup (save for next session)",
                    data=lookup_json,
                    file_name=get_vendor_lookup_filename(active_vendor),
                    mime="application/json",
                    key="dl_lookup_after_save",
                )
            else:
                st.warning(
                    "0 pcodes saved. Make sure you typed a pcode in the "
                    "**Confirmed Pcode** column and pressed Tab or Enter before clicking Save."
                )

    else:
        st.success("✅ All properties matched with high confidence — nothing to review!")
