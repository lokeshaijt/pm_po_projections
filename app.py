"""
PM PO Projection Mailer — Streamlit app

Upload a Pending PO export and a PM Requirement (Shortfall) workbook,
generate the PO Issued / PO Projection report, browse CFC/CTN/TRAY
category sections with supplier averaging, preview supplier-wise
projections, and email suppliers their category-wise projection.

All parsing/business logic lives in logic.py (no streamlit dependency,
so it's independently testable); this file is the UI layer only.
"""

import smtplib
from datetime import date
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

import logic as L

st.set_page_config(page_title="PM PO Projection Mailer", layout="wide")


def send_email(smtp_host, smtp_port, username, password, to_email, subject, body):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = username
    msg["To"] = to_email
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(username, password)
        server.sendmail(username, [to_email], msg.as_string())


st.title("PM PO Projection Mailer")
st.caption(
    "Upload the Pending PO and PM Requirement files to generate the report, "
    "browse category-wise projections, and email suppliers directly."
)

with st.sidebar:
    st.header("1. Upload files")
    po_file = st.file_uploader("Pending PO file (.xlsx)", type="xlsx", key="po_file")
    req_file = st.file_uploader("PM Requirement file (.xlsx)", type="xlsx", key="req_file")
    master_file = st.file_uploader(
        "Item Category Master (.xlsx)", type="xlsx", key="master_file",
        help="Columns: ITEM NAME, Item Type, ITEM CATEGORY (any order).",
    )
    today = st.date_input("As-of date", value=date.today())
    window_days = st.number_input("Delivery window (days)", min_value=1, max_value=90, value=14)
    generate = st.button("Generate report", type="primary")

if generate:
    if not (po_file and req_file and master_file):
        st.error("Please upload the Pending PO, PM Requirement, and Item Category Master files.")
    else:
        with st.spinner("Processing..."):
            po_df = L.parse_pending_po(po_file)
            req_data = L.parse_requirement_workbook(req_file)
            name_to_cat, name_to_type = L.load_category_master(master_file)
            item_to_cat, item_to_type = L.build_item_category_map(req_data, name_to_cat, name_to_type)
            item_week_proj = L.build_item_week_projections(req_data)
            po_issued = L.build_po_issued_df(po_df, today, window_days)

            category_order = sorted(set(name_to_cat.values()))
            category_df = L.build_category_table(
                item_week_proj, item_to_cat, item_to_type, category_order, req_data["projection_weeks"]
            )

            st.session_state["po_issued"] = po_issued
            st.session_state["category_df"] = category_df
            st.session_state["projection_weeks"] = req_data["projection_weeks"]
            st.session_state["report_bytes"] = L.generate_report_workbook(
                po_issued, category_df, req_data["projection_weeks"]
            )
        st.success(
            f"Report generated: {len(po_issued)} PO Issued rows, "
            f"{(category_df['Total'] > 0).sum()} categories with a projection."
        )

if "category_df" in st.session_state:
    category_df = st.session_state["category_df"]
    projection_weeks = st.session_state["projection_weeks"]
    po_issued = st.session_state["po_issued"]

    st.divider()
    st.header("2. Download report")
    st.download_button(
        "Download PM PO Report (.xlsx)",
        data=st.session_state["report_bytes"],
        file_name=f"PM_PO_Report_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.header("3. Category sections")
    tabs = st.tabs(L.ITEM_TYPES)
    for item_type, tab in zip(L.ITEM_TYPES, tabs):
        with tab:
            sub = category_df[category_df["Item Type"] == item_type]
            sub_nonzero = sub[sub["Total"] > 0]
            st.subheader(f"{item_type} — categories with a projection")
            st.dataframe(sub_nonzero, use_container_width=True, hide_index=True)

            suppliers = L.SUPPLIERS_BY_TYPE.get(item_type, [])
            selected = st.multiselect(f"Select {item_type} suppliers", suppliers, key=f"suppliers_{item_type}")
            if len(selected) >= 1:
                st.write(f"**Average projection per selected supplier ({len(selected)} selected)**")
                avg_row = {"Item Category": "TOTAL / AVG"}
                for w in projection_weeks:
                    total_w = sub[f"Wk #{w}"].sum()
                    avg_row[f"Wk #{w}"] = total_w / len(selected)
                avg_df = pd.DataFrame([avg_row])
                st.dataframe(avg_df, use_container_width=True, hide_index=True)

    st.divider()
    st.header("4. Supplier-wise preview")
    all_suppliers = L.SUPPLIERS_BY_TYPE["CFC"] + L.SUPPLIERS_BY_TYPE["CTN"]
    all_suppliers = list(dict.fromkeys(all_suppliers))  # dedupe, keep order
    preview_supplier = st.selectbox("Select a supplier", all_suppliers)
    preview_type = None
    for t, sups in L.SUPPLIERS_BY_TYPE.items():
        if preview_supplier in sups:
            preview_type = t
            break
    if preview_type:
        preview_df = category_df[(category_df["Item Type"] == preview_type) & (category_df["Total"] > 0)]
        st.write(f"**{preview_supplier}** ({preview_type} supplier)")
        st.dataframe(preview_df, use_container_width=True, hide_index=True)

    st.divider()
    st.header("5. Email suppliers")
    st.caption(
        "Requires SMTP credentials in Streamlit secrets: "
        "`smtp_host`, `smtp_port`, `smtp_username`, `smtp_password`, "
        "and a mapping of supplier name -> email address under `supplier_emails`."
    )
    email_supplier = st.selectbox("Send projection to", all_suppliers, key="email_supplier")
    email_type = next((t for t, s in L.SUPPLIERS_BY_TYPE.items() if email_supplier in s), None)
    supplier_cat_df = category_df[(category_df["Item Type"] == email_type) & (category_df["Total"] > 0)] if email_type else pd.DataFrame()
    body_preview = L.build_supplier_email_body(email_supplier, supplier_cat_df, projection_weeks)
    st.text_area("Email preview", body_preview, height=250)

    if st.button("Send email"):
        try:
            to_email = st.secrets["supplier_emails"][email_supplier]
            send_email(
                st.secrets["smtp_host"],
                int(st.secrets["smtp_port"]),
                st.secrets["smtp_username"],
                st.secrets["smtp_password"],
                to_email,
                f"Item Category-wise Projection — {email_supplier}",
                body_preview,
            )
            st.success(f"Email sent to {email_supplier} ({to_email}).")
        except KeyError as e:
            st.error(f"Missing secret: {e}. Configure SMTP + supplier_emails in Streamlit secrets.")
        except Exception as e:
            st.error(f"Failed to send email: {e}")
else:
    st.info("Upload the three files in the sidebar and click **Generate report** to get started.")
