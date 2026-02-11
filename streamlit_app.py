import requests
import base64
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, numbers, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import uuid



st.set_page_config(layout="wide")

POWER_AUTOMATE_URL = st.secrets["power_automate"]["url"]

# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------
if "df" not in st.session_state:
    st.session_state.df = None
if "editor_df" not in st.session_state:
    st.session_state.editor_df = None
if "editor_initialized" not in st.session_state:
    st.session_state.editor_initialized = False
# -------------------------------------------------
# HEADER
# -------------------------------------------------
st.title("📊 Interactive Table Review & Price Analysis")


# -------------------------------------------------
# PDF UPLOAD → POWER AUTOMATE (HIGHEST PRIORITY)
# -------------------------------------------------
st.subheader("📄 Upload 3 Quote PDFs (Power Automate)")

pdfs = st.file_uploader(
    "Upload 1 or more PDF quotes",
    type=["pdf"],
    accept_multiple_files=True
)

if pdfs:  # at least one file uploaded
    if st.button("🚀 Process PDFs via Power Automate"):
        with st.spinner("Sending PDFs to Power Automate…"):

            files_payload = []
            for pdf in pdfs:
                encoded = base64.b64encode(pdf.read()).decode("ascii")
                files_payload.append({
                    "name": pdf.name,
                    "content": encoded
                })

            response = requests.post(
                POWER_AUTOMATE_URL,
                json={"files": files_payload},
                headers={
                    "Content-Type": "application/json"
                },
                timeout=600
            )

        if response.status_code != 200:
            st.error("Power Automate failed to process PDFs")
            st.stop()

        # Expecting base64 CSV back
        csv_bytes = base64.b64decode(response.json()["csv"])
        df = pd.read_csv(io.BytesIO(csv_bytes))

        # 🔑 HANDOFF POINT — everything else already works
        for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        st.session_state.df = df
        st.session_state.editor_df = None
        st.session_state.editor_initialized = False

        st.session_state.current_job_path = None
        st.session_state.job_loaded_from_queue = False
        
        # 🔥 Store CSV bytes for download
        st.session_state.csv_bytes = csv_bytes

        st.success(f"✅ CSV generated from {len(pdfs)} PDF(s) and loaded")
        st.rerun()
else:
    st.info("Upload 1 or more PDFs to start processing")

# 🔥 AUTO-DOWNLOAD CSV BUTTON (appears after processing)
if "csv_bytes" in st.session_state and st.session_state.csv_bytes is not None:
    st.download_button(
        label="📥 Download Raw CSV",
        data=st.session_state.csv_bytes,
        file_name=f"quotes_raw_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv"
    )

# -------------------------------------------------
# UPLOAD FILE (MANUAL OVERRIDE)
# -------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload CSV or Excel (manual override)",
    type=["csv", "xlsx"]
)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    for col in ["type", "supplier", "brand", "code", "description", "Power Type"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # 🔹 Override queue state
    st.session_state.df = df.copy()
    st.session_state.editor_df = None
    st.session_state.editor_initialized = False
    st.session_state.current_job_path = None
    st.session_state.job_loaded_from_queue = False

    st.success("📤 Manual file loaded (queue overridden)")

# ==============================================================
# STEP 3: ADD THESE HELPER FUNCTIONS (before the editor section)
# ==============================================================

def init_editor_structure(df):
    """
    Convert flat CSV DataFrame into parent-child structure.
    Assigns id, parentId, order to each row.
    Items (type='item') get parentId=None.
    Subitems (type='subitem') get parentId = the preceding item's id.
    """
    edf = df.copy()
    edf["id"] = [str(uuid.uuid4())[:8] for _ in range(len(edf))]
    edf["parentId"] = None
    edf["order"] = 0

    current_parent_id = None
    child_counter = {}

    for idx in edf.index:
        row_type = str(edf.at[idx, "type"]).strip().lower()

        if row_type == "item":
            current_parent_id = edf.at[idx, "id"]
            edf.at[idx, "parentId"] = None
            edf.at[idx, "order"] = 0
            child_counter[current_parent_id] = 0
        elif row_type == "subitem":
            if current_parent_id is not None:
                edf.at[idx, "parentId"] = current_parent_id
                edf.at[idx, "order"] = child_counter.get(current_parent_id, 0)
                child_counter[current_parent_id] = child_counter.get(current_parent_id, 0) + 1
            else:
                # Orphan subitem — treat as item
                edf.at[idx, "type"] = "item"
                edf.at[idx, "parentId"] = None
                edf.at[idx, "order"] = 0

    return edf


def get_product_groups(edf):
    """Group items by code + Power Type."""
    items = edf[edf["type"] == "item"]
    groups = items.groupby(["code", "Power Type"], dropna=False)
    return groups


def reorder_row(edf, row_id, direction):
    """Move a row up or down among its siblings."""
    row = edf[edf["id"] == row_id].iloc[0]
    parent = row["parentId"]
    current_order = row["order"]

    if row["type"] == "item":
        # Siblings = items with same code + Power Type
        siblings = edf[
            (edf["type"] == "item") &
            (edf["code"] == row["code"]) &
            (edf["Power Type"] == row["Power Type"])
        ].sort_values("order")
    else:
        # Siblings = subitems with same parentId
        siblings = edf[
            (edf["parentId"] == parent) &
            (edf["type"] == "subitem")
        ].sort_values("order")

    orders = siblings["order"].tolist()
    ids = siblings["id"].tolist()
    pos = ids.index(row_id)

    if direction == "up" and pos > 0:
        swap_id = ids[pos - 1]
        swap_order = edf.loc[edf["id"] == swap_id, "order"].iloc[0]
        edf.loc[edf["id"] == row_id, "order"] = swap_order
        edf.loc[edf["id"] == swap_id, "order"] = current_order
    elif direction == "down" and pos < len(ids) - 1:
        swap_id = ids[pos + 1]
        swap_order = edf.loc[edf["id"] == swap_id, "order"].iloc[0]
        edf.loc[edf["id"] == row_id, "order"] = swap_order
        edf.loc[edf["id"] == swap_id, "order"] = current_order

    return edf


def convert_type(edf, row_id):
    """Toggle item <-> subitem with orphan handling."""
    row = edf[edf["id"] == row_id].iloc[0]

    if row["type"] == "subitem":
        # Subitem → Item: clear parent, set order 0
        edf.loc[edf["id"] == row_id, "type"] = "item"
        edf.loc[edf["id"] == row_id, "parentId"] = None
        edf.loc[edf["id"] == row_id, "order"] = 0
    else:
        # Item → Subitem: promote orphaned children to items first
        children = edf[edf["parentId"] == row_id]
        for cidx in children.index:
            edf.at[cidx, "type"] = "item"
            edf.at[cidx, "parentId"] = None
            edf.at[cidx, "order"] = 0

        # Now this row needs a new parent — find first item in same product group
        same_product = edf[
            (edf["type"] == "item") &
            (edf["code"] == row["code"]) &
            (edf["Power Type"] == row["Power Type"]) &
            (edf["id"] != row_id)
        ]
        if not same_product.empty:
            new_parent = same_product.iloc[0]["id"]
            existing_children = edf[edf["parentId"] == new_parent]
            next_order = existing_children["order"].max() + 1 if not existing_children.empty else 0
            edf.loc[edf["id"] == row_id, "type"] = "subitem"
            edf.loc[edf["id"] == row_id, "parentId"] = new_parent
            edf.loc[edf["id"] == row_id, "order"] = next_order

    return edf


def spread_row(edf, row_id, target_item_ids):
    """Copy a row as subitem under each target item."""
    source = edf[edf["id"] == row_id].iloc[0]

    new_rows = []
    for target_id in target_item_ids:
        target_item = edf[edf["id"] == target_id].iloc[0]
        existing = edf[edf["parentId"] == target_id]
        next_order = int(existing["order"].max() + 1) if not existing.empty else 0

        new_row = source.copy()
        new_row["id"] = str(uuid.uuid4())[:8]
        new_row["parentId"] = target_id
        new_row["type"] = "subitem"
        new_row["order"] = next_order
        new_row["supplier"] = target_item["supplier"]
        new_rows.append(new_row)

    if new_rows:
        edf = pd.concat([edf, pd.DataFrame(new_rows)], ignore_index=True)
    return edf


def delete_row(edf, row_id):
    """Delete row. If item, cascade-delete children."""
    row = edf[edf["id"] == row_id].iloc[0]
    if row["type"] == "item":
        edf = edf[edf["parentId"] != row_id]  # delete children
    edf = edf[edf["id"] != row_id]  # delete self
    return edf


def editor_to_flat_df(edf):
    """
    Convert editor DataFrame back to flat format compatible
    with your existing HTML preview and Excel generator.
    Strips editor-only columns (id, parentId, order).
    Preserves the parent-child ordering.
    """
    # Sort: items first (by code, Power Type, order), then their children
    result_rows = []
    items = edf[edf["type"] == "item"].sort_values(["code", "Power Type", "order"])

    for _, item in items.iterrows():
        result_rows.append(item)
        children = edf[edf["parentId"] == item["id"]].sort_values("order")
        for _, child in children.iterrows():
            result_rows.append(child)

    # Also include any orphaned subitems (shouldn't exist, but safety)
    all_ids = set(r["id"] for r in result_rows)
    for _, row in edf.iterrows():
        if row["id"] not in all_ids:
            result_rows.append(row)

    flat = pd.DataFrame(result_rows)
    # Drop editor columns — your preview/Excel don't need them
    flat = flat.drop(columns=["id", "parentId", "order"], errors="ignore")
    flat = flat.reset_index(drop=True)
    return flat
    
# editor dynamic

if st.session_state.df is not None:

    # --- Initialize editor structure from flat df ---
    if not st.session_state.editor_initialized or st.session_state.editor_df is None:
        st.session_state.editor_df = init_editor_structure(st.session_state.df)
        st.session_state.editor_initialized = True

    edf = st.session_state.editor_df

    st.subheader("✏️ Quote Editor")

    # === RAW TABLE EDITOR (for direct cell edits: price, description, etc.) ===
    with st.expander("📝 Edit Raw Data (click cells to edit)", expanded=False):
        # Show editable columns only (hide id/parentId/order from user)
        display_cols = [c for c in edf.columns if c not in ["id", "parentId", "order"]]
        edited = st.data_editor(
            edf[display_cols],
            use_container_width=True,
            num_rows="dynamic",
            key="raw_editor"
        )
        # Sync edits back (preserve id/parentId/order)
        for col in display_cols:
            edf[col] = edited[col].values[:len(edf)]
        st.session_state.editor_df = edf

    # === STRUCTURED EDITOR (grouped view with actions) ===
    st.markdown("---")
    st.markdown("#### 📦 Product Groups")

    items_only = edf[edf["type"] == "item"]
    product_keys = items_only[["code", "Power Type"]].drop_duplicates().values.tolist()

    for code, power_type in product_keys:
        group_items = edf[
            (edf["type"] == "item") &
            (edf["code"] == code) &
            (edf["Power Type"] == power_type)
        ].sort_values("order")

        with st.expander(f"🔹 {code} — {power_type} ({len(group_items)} suppliers)", expanded=False):
            for _, item in group_items.iterrows():
                item_id = item["id"]
                children = edf[edf["parentId"] == item_id].sort_values("order")
                subtotal = item["price"] if pd.notna(item["price"]) else 0
                for _, ch in children.iterrows():
                    subtotal += ch["price"] if pd.notna(ch["price"]) else 0

                # --- ITEM HEADER ---
                st.markdown(
                    f"**🏢 {item['supplier']}** — {item['description']} "
                    f"— `${float(item['price']):,.2f}` "
                    f"| Subtotal: `${float(subtotal):,.2f}`"
                )

                # --- ITEM ACTION BUTTONS ---
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    if st.button("⬆️", key=f"up_{item_id}", help="Move up"):
                        st.session_state.editor_df = reorder_row(edf, item_id, "up")
                        st.rerun()
                with col2:
                    if st.button("⬇️", key=f"down_{item_id}", help="Move down"):
                        st.session_state.editor_df = reorder_row(edf, item_id, "down")
                        st.rerun()
                with col3:
                    if st.button("🔄 → Sub", key=f"conv_{item_id}", help="Convert to subitem"):
                        st.session_state.editor_df = convert_type(edf, item_id)
                        st.rerun()
                with col4:
                    if st.button("🗑️", key=f"del_{item_id}", help="Delete item + children"):
                        child_count = len(children)
                        st.session_state.editor_df = delete_row(edf, item_id)
                        st.rerun()

                # --- SUBITEMS ---
                if not children.empty:
                    for _, child in children.iterrows():
                        child_id = child["id"]
                        st.markdown(
                            f"&nbsp;&nbsp;&nbsp;&nbsp;↳ {child['description']} "
                            f"— `${float(child['price']):,.2f}`"
                        )
                        c1, c2, c3, c4 = st.columns(4)
                        with c1:
                            if st.button("⬆️", key=f"up_{child_id}"):
                                st.session_state.editor_df = reorder_row(edf, child_id, "up")
                                st.rerun()
                        with c2:
                            if st.button("⬇️", key=f"down_{child_id}"):
                                st.session_state.editor_df = reorder_row(edf, child_id, "down")
                                st.rerun()
                        with c3:
                            if st.button("🔄 → Item", key=f"conv_{child_id}"):
                                st.session_state.editor_df = convert_type(edf, child_id)
                                st.rerun()
                        with c4:
                            if st.button("🗑️", key=f"del_{child_id}"):
                                st.session_state.editor_df = delete_row(edf, child_id)
                                st.rerun()

                st.markdown("---")

    # === SPREAD (COPY) TOOL ===
    st.markdown("#### 📋 Spread Row to Multiple Items")

    all_rows = edf[["id", "type", "supplier", "description"]].copy()
    all_rows["label"] = all_rows.apply(
        lambda r: f"[{r['type']}] {r['supplier']} — {r['description']}", axis=1
    )
    row_options = dict(zip(all_rows["label"], all_rows["id"]))

    source_label = st.selectbox(
        "Select row to copy",
        options=list(row_options.keys()),
        key="spread_source"
    )

    # Target items (only items)
    item_rows = edf[edf["type"] == "item"][["id", "supplier", "code", "description"]].copy()
    item_rows["label"] = item_rows.apply(
        lambda r: f"{r['supplier']} — {r['code']} — {r['description']}", axis=1
    )
    target_options = dict(zip(item_rows["label"], item_rows["id"]))

    target_labels = st.multiselect(
        "Spread as subitem to these items",
        options=list(target_options.keys()),
        key="spread_targets"
    )

    if st.button("📋 Spread Now", key="spread_btn"):
        if source_label and target_labels:
            source_id = row_options[source_label]
            target_ids = [target_options[t] for t in target_labels]
            st.session_state.editor_df = spread_row(edf, source_id, target_ids)
            st.success(f"Spread to {len(target_ids)} items!")
            st.rerun()
        else:
            st.warning("Select a source row and at least one target.")

    # === ADD NEW ROW ===
    st.markdown("#### ➕ Add New Row")
    with st.expander("Add a new item or subitem", expanded=False):
        new_type = st.radio("Type", ["item", "subitem"], horizontal=True, key="new_type")
        new_supplier = st.text_input("Supplier", key="new_supplier")
        new_brand = st.text_input("Brand", key="new_brand")
        new_code = st.text_input("Code", key="new_code")
        new_desc = st.text_input("Description", key="new_desc")
        new_power = st.text_input("Power Type", key="new_power")
        new_price = st.number_input("Price", min_value=0.0, value=0.0, key="new_price")

        parent_id = None
        if new_type == "subitem":
            parent_label = st.selectbox(
                "Parent Item",
                options=list(target_options.keys()),
                key="new_parent"
            )
            if parent_label:
                parent_id = target_options[parent_label]

        if st.button("➕ Add Row", key="add_row_btn"):
            new_id = str(uuid.uuid4())[:8]
            if new_type == "subitem" and parent_id:
                existing = edf[edf["parentId"] == parent_id]
                new_order = int(existing["order"].max() + 1) if not existing.empty else 0
            else:
                new_order = 0
                parent_id = None

            new_row = {
                "id": new_id,
                "parentId": parent_id,
                "order": new_order,
                "type": new_type,
                "supplier": new_supplier,
                "brand": new_brand,
                "code": new_code,
                "description": new_desc,
                "Power Type": new_power,
                "price": new_price,
            }
            st.session_state.editor_df = pd.concat(
                [edf, pd.DataFrame([new_row])], ignore_index=True
            )
            st.success(f"Added new {new_type}: {new_desc}")
            st.rerun()

    # === SYNC BACK TO FLAT DF (feeds your preview + Excel) ===
    st.session_state.df = editor_to_flat_df(st.session_state.editor_df)


# -------------------------------------------------
# TAX INPUT
# -------------------------------------------------
st.subheader("💲 Tax Settings")
tax_percent = st.number_input("Tax Percentage", min_value=0.0, value=12.0)

# -------------------------------------------------
# HTML PREVIEW (EXCEL-STYLE)
# -------------------------------------------------
st.subheader("👀 Price Analysis Preview (HTML Table)")

def generate_html_table(df, tax_percent):
    tax_rate = tax_percent / 100

    html = """
    <div style="overflow-x:auto;">
    <style>
        table {
            border-collapse: collapse !important;
            width: 100%;
            margin-bottom: 40px;
            font-family: Arial, sans-serif;
            background-color: #ffffff;
            color: #000000;
            border: 1px solid #bfbfbf;
        }

        th, td {
            border: 1px solid #bfbfbf !important;
            padding: 6px 8px;
            vertical-align: middle;
            text-align: left;
            background-clip: padding-box;
        }

        th {
            background-color: #dae9f8;
            font-weight: 600;
        }

        .total-row td {
            background-color: #fce4d6;
            font-weight: bold;
        }
    </style>
    """

    main_items = df[
        (df["type"] == "item") &
        df["Power Type"].notna() &
        (df["Power Type"] != "")
    ]

    for code, power_type in main_items[["code", "Power Type"]].drop_duplicates().values:

        items_for_code = df[
            (df["code"] == code) &
            (
                (df["Power Type"] == power_type) |
                (df["Power Type"].isna()) |
                (df["Power Type"] == "")
            ) &
            (df["type"].isin(["item", "subitem"]))
        ]

        suppliers = items_for_code["supplier"].unique()
        brand = items_for_code[items_for_code["type"] == "item"].iloc[0]["brand"]
        descriptions = items_for_code["description"].unique()

        body_rows = len(descriptions) + 2  # items + tax + total

        html += "<table>"

        # HEADER
        html += "<tr>"
        html += "<th>Details</th><th></th><th>QTY</th><th>Items</th>"
        for s in suppliers:
            html += f"<th>{s}</th>"
        html += "</tr>"

        totals = {s: 0 for s in suppliers}

        # FIRST ITEM ROW (with DETAILS)
        first_desc = descriptions[0]

        html += "<tr>"
        html += f"""
            <td rowspan="{body_rows}">
                <b>Brand</b><br>{brand}<br><br>
                <b>Code</b><br>{code}<br><br>
                <b>Power Type</b><br>{power_type}
            </td>
            <td rowspan="{body_rows}"></td>
            <td>1</td>
            <td>{first_desc}</td>
        """

        for s in suppliers:
            row = items_for_code[
                (items_for_code["supplier"] == s) &
                (items_for_code["description"] == first_desc)
            ]
            price = float(row["price"].iloc[0]) if not row.empty else 0
            totals[s] += price
            html += f"<td>${price:,.2f}</td>"

        html += "</tr>"

        # REMAINING ITEM ROWS
        for desc in descriptions[1:]:
            html += "<tr>"
            html += f"<td>1</td><td>{desc}</td>"

            for s in suppliers:
                row = items_for_code[
                    (items_for_code["supplier"] == s) &
                    (items_for_code["description"] == desc)
                ]
                price = float(row["price"].iloc[0]) if not row.empty else 0
                totals[s] += price
                html += f"<td>${price:,.2f}</td>"

            html += "</tr>"

        # TAX ROW
        html += "<tr>"
        html += "<td></td><td><b>Tax</b></td>"
        for _ in suppliers:
            html += f"<td>{tax_percent:.2f}%</td>"
        html += "</tr>"

        # TOTAL ROW
        html += "<tr class='total-row'>"
        html += "<td></td><td>Total</td>"
        for s in suppliers:
            total = totals[s] * (1 + tax_rate)
            html += f"<td>${total:,.2f}</td>"
        html += "</tr>"

        html += "</table>"

    html += "</div>"
    return html


# 🔥 RENDER HTML (LIVE, REACTIVE)
if (
    "df" in st.session_state
    and st.session_state.df is not None
    and not st.session_state.df.empty
):
    html = generate_html_table(st.session_state.df, tax_percent)
    st.markdown(html, unsafe_allow_html=True)
else:
    st.info("⬆️ Upload or generate data to see the price analysis preview.")

# -------------------------------------------------
# GENERATE FINAL EXCEL (MINIMALIST FORMATTING - RFQ STYLE)
# -------------------------------------------------
st.subheader("📥 Generate Final Excel")

if st.button("Generate Excel File"):
    df = st.session_state.df
    tax_rate = tax_percent / 100

    wb = Workbook()
    ws = wb.active
    ws.title = "Price Analysis"
    ws.sheet_view.showGridLines = False  # Clean minimalist look
    
    # --- MINIMALIST DESIGN TOKENS (MATCHING RFQ STYLE) ---
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    
    # Color palette matching RFQ Details
    HEADER_BLUE = PatternFill(start_color="288AD6", end_color="288AD6", fill_type="solid")
    DETAILS_BG = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    COLUMN_HEADER_BG = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    WINNER_BG = PatternFill(start_color="F2FAF2", end_color="F2FAF2", fill_type="solid")
    SPECS_BG = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    
    # Border styles
    SUBTLE_BORDER = Border(bottom=Side(style='thin', color="F0F0F0"))
    COLUMN_HEADER_BORDER = Border(bottom=Side(style='medium', color="E5E5E5"))
    TOTAL_BORDER = Border(top=Side(style='medium', color="288AD6"))
    
    # Text colors
    TEXT_PRIMARY = "1D1D1F"
    TEXT_SECONDARY = "86868B"
    WHITE = "FFFFFF"

    current_row = 1

    main_items = df[
        (df["type"] == "item") & 
        df["Power Type"].notna() & 
        (df["Power Type"] != "")
    ]

    for opt_idx, (code, power_type) in enumerate(
        main_items[["code", "Power Type"]].drop_duplicates().values, 1
    ):
        # Get all items for this code/power type combination
        items_for_code = df[
            (df["code"] == code) &
            (
                (df["Power Type"] == power_type) |
                (df["Power Type"].isna()) |
                (df["Power Type"] == "")
            ) &
            (df["type"].isin(["item", "subitem"]))
        ]

        suppliers = list(items_for_code["supplier"].unique())
        brand = items_for_code[items_for_code["type"] == "item"].iloc[0]["brand"]
        descriptions = list(items_for_code["description"].unique())

        # --- DETERMINE WINNER (LOWEST TOTAL PRICE) ---
        winner_supplier = ""
        min_total = float('inf')
        for sup in suppliers:
            sup_items = items_for_code[items_for_code["supplier"] == sup]
            total = sup_items["price"].sum() * (1 + tax_rate)
            if total < min_total:
                min_total = total
                winner_supplier = sup

        # === 1. OPTION TITLE (BLUE HEADER) ===
        title_row = current_row
        ws.row_dimensions[title_row].height = 40
        
        # Merge across all columns
        last_col = 6 + len(suppliers) - 1
        ws.merge_cells(
            start_row=title_row,
            start_column=2,
            end_row=title_row,
            end_column=last_col
        )
        
        title_cell = ws.cell(row=title_row, column=2, value=f"Option {opt_idx:02d}")
        title_cell.font = Font(name='Segoe UI', bold=True, size=14, color=WHITE)
        title_cell.fill = HEADER_BLUE
        title_cell.alignment = Alignment(horizontal="left", vertical="center", indent=2)

        # === 2. COLUMN HEADERS ===
        header_row = title_row + 1
        ws.row_dimensions[header_row].height = 28

        headers = ["DETAILS", "IMAGE", "QTY", "LINE ITEM"] + suppliers
        for i, h in enumerate(headers):
            col_idx = i + 2
            cell = ws.cell(row=header_row, column=col_idx, value=h.upper())
            cell.font = Font(name='Segoe UI', size=9, bold=True, color=TEXT_SECONDARY)
            cell.fill = WINNER_BG if h == winner_supplier else COLUMN_HEADER_BG
            cell.alignment = Alignment(
                horizontal="center" if col_idx >= 6 else "left",
                vertical="center"
            )
            cell.border = COLUMN_HEADER_BORDER

        # === 3. CONTENT ROWS ===
        start_data_row = header_row + 1
        num_item_rows = len(descriptions)
        num_body_rows = num_item_rows + 3  # +3 for total before tax, tax, and total rows

        # DETAILS column (merged vertically) - STOP BEFORE TOTAL ROW
        detail_val = f"BRAND\n{brand}\n\nCODE\n{code}\n\nPOWER\n{power_type}"
        d_cell = ws.cell(row=start_data_row, column=2, value=detail_val)
        d_cell.alignment = Alignment(
            wrap_text=True, 
            vertical="top", 
            horizontal="left",
            indent=1
        )
        d_cell.font = Font(name='Segoe UI', size=10, color=TEXT_SECONDARY)
        d_cell.fill = DETAILS_BG
        ws.merge_cells(
            start_row=start_data_row,
            start_column=2,
            end_row=start_data_row + num_item_rows + 1,  # Only through tax row
            end_column=2
        )
        
        # IMAGE placeholder (merged vertically) - STOP BEFORE TOTAL ROW
        img_cell = ws.cell(row=start_data_row, column=3, value="[ PHOTO ]")
        img_cell.alignment = Alignment(horizontal="center", vertical="center")
        img_cell.font = Font(name='Segoe UI', size=10, color="CCCCCC", italic=True)
        img_cell.fill = DETAILS_BG
        ws.merge_cells(
            start_row=start_data_row,
            start_column=3,
            end_row=start_data_row + num_item_rows + 1,  # Only through tax row
            end_column=3
        )

        # ITEM ROWS
        for idx, desc in enumerate(descriptions):
            r_num = start_data_row + idx
            ws.row_dimensions[r_num].height = 32

            # QTY column
            qty_cell = ws.cell(row=r_num, column=4, value=1)
            qty_cell.alignment = Alignment(horizontal="center", vertical="center")
            qty_cell.font = Font(name='Segoe UI', size=11, color=TEXT_PRIMARY, bold=True)
            qty_cell.border = SUBTLE_BORDER

            # LINE ITEM (description)
            desc_cell = ws.cell(row=r_num, column=5, value=desc)
            desc_cell.font = Font(name='Segoe UI', size=11, color=TEXT_PRIMARY)
            desc_cell.alignment = Alignment(vertical="center", horizontal="left")
            desc_cell.border = SUBTLE_BORDER

            # SUPPLIER PRICES (with formulas tied to QTY)
            qty_letter = get_column_letter(4)
            for s_idx, sup in enumerate(suppliers):
                col = 6 + s_idx

                # Get price for this supplier/description combo
                price_row = items_for_code[
                    (items_for_code["supplier"] == sup) &
                    (items_for_code["description"] == desc)
                ]
                price = float(price_row["price"].iloc[0]) if not price_row.empty else 0

                # Formula: =QTY * price
                cell = ws.cell(row=r_num, column=col, value=f"={qty_letter}{r_num}*{price}")
                cell.number_format = '$#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = SUBTLE_BORDER
                cell.font = Font(name='Segoe UI', size=11, color=TEXT_PRIMARY)

                if sup == winner_supplier:
                    cell.fill = WINNER_BG

        # === 4. TOTAL BEFORE TAX ROW ===
        total_before_tax_row = start_data_row + num_item_rows
        ws.row_dimensions[total_before_tax_row].height = 32

        # Merge QTY and LINE ITEM columns for label
        ws.merge_cells(
            start_row=total_before_tax_row,
            start_column=4,
            end_row=total_before_tax_row,
            end_column=5
        )
        
        total_before_tax_label = ws.cell(row=total_before_tax_row, column=4, value="Total Before Tax")
        total_before_tax_label.font = Font(name='Segoe UI', size=11, bold=True, color=TEXT_PRIMARY)
        total_before_tax_label.alignment = Alignment(vertical="center", horizontal="left")
        total_before_tax_label.border = SUBTLE_BORDER

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Total Before Tax formula: SUM(items)
            tbt_cell = ws.cell(
                row=total_before_tax_row,
                column=col,
                value=f"=SUM({col_letter}{start_data_row}:{col_letter}{total_before_tax_row-1})"
            )
            tbt_cell.number_format = '$#,##0.00'
            tbt_cell.alignment = Alignment(horizontal="right", vertical="center")
            tbt_cell.font = Font(name='Segoe UI', size=11, bold=True, color=TEXT_PRIMARY)
            tbt_cell.border = SUBTLE_BORDER

            if sup == winner_supplier:
                tbt_cell.fill = WINNER_BG

        # === 5. TAX ROW ===
        tax_row = total_before_tax_row + 1
        ws.row_dimensions[tax_row].height = 28

        # Merge QTY and LINE ITEM columns for label
        ws.merge_cells(
            start_row=tax_row,
            start_column=4,
            end_row=tax_row,
            end_column=5
        )
        
        tax_label = ws.cell(row=tax_row, column=4, value=f"Tax ({int(tax_rate*100)}%)")
        tax_label.font = Font(name='Segoe UI', size=10, color=TEXT_SECONDARY)
        tax_label.alignment = Alignment(vertical="center", horizontal="left")
        tax_label.border = SUBTLE_BORDER

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Tax formula: Subtotal * tax_rate
            t_cell = ws.cell(
                row=tax_row,
                column=col,
                value=f"={col_letter}{total_before_tax_row}*{tax_rate}"
            )
            t_cell.number_format = '$#,##0.00'
            t_cell.alignment = Alignment(horizontal="right", vertical="center")
            t_cell.font = Font(name='Segoe UI', size=10, color=TEXT_SECONDARY)
            t_cell.border = SUBTLE_BORDER

            if sup == winner_supplier:
                t_cell.fill = WINNER_BG

        # === 6. FINAL TOTAL ROW ===
        total_row = tax_row + 1
        ws.row_dimensions[total_row].height = 40
        
        # Merge DETAILS, IMAGE, QTY, and LINE ITEM columns (leave empty)
        ws.merge_cells(
            start_row=total_row,
            start_column=2,
            end_row=total_row,
            end_column=5
        )
        empty_cell = ws.cell(row=total_row, column=2, value="")
        empty_cell.fill = DETAILS_BG
        empty_cell.border = TOTAL_BORDER

        for s_idx, sup in enumerate(suppliers):
            col = 6 + s_idx
            col_letter = get_column_letter(col)

            # Total formula: Subtotal + Tax
            tot_cell = ws.cell(
                row=total_row,
                column=col,
                value=f"={col_letter}{total_before_tax_row}+{col_letter}{tax_row}"
            )
            tot_cell.font = Font(name='Segoe UI', bold=True, size=13, color=TEXT_PRIMARY)
            tot_cell.number_format = '$#,##0.00'
            tot_cell.alignment = Alignment(horizontal="right", vertical="center")
            tot_cell.border = TOTAL_BORDER

            if sup == winner_supplier:
                tot_cell.fill = WINNER_BG

        # === 7. SPECS & DESCRIPTION BLOCK ===
        specs_row = total_row + 1
        ws.row_dimensions[specs_row].height = 60

        # Merge across all columns
        ws.merge_cells(
            start_row=specs_row,
            start_column=2,
            end_row=specs_row,
            end_column=last_col
        )

        specs_content = ws.cell(
            row=specs_row,
            column=2,
            value="SPECS & DESCRIPTION\n\nEnter item specifications, dimensions, and technical details here..."
        )
        specs_content.font = Font(name='Segoe UI', size=10, color=TEXT_SECONDARY)
        specs_content.alignment = Alignment(
            wrap_text=True, 
            vertical="top",
            horizontal="left",
            indent=2
        )
        specs_content.fill = SPECS_BG
        specs_content.border = Border(bottom=Side(style='thin', color="F0F0F0"))

        # Move to next table (with spacing)
        current_row = specs_row + 3

    # === COLUMN WIDTH ADJUSTMENTS ===
    ws.column_dimensions['A'].width = 2   # Left margin
    ws.column_dimensions['B'].width = 18  # Details
    ws.column_dimensions['C'].width = 12  # Image
    ws.column_dimensions['D'].width = 8   # QTY
    ws.column_dimensions['E'].width = 35  # Line Item

    for i in range(len(suppliers)):
        ws.column_dimensions[get_column_letter(6+i)].width = 16

    # Save to downloadable file
    output = io.BytesIO()
    wb.save(output)

    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name=f"price_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )