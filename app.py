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
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

import logic as L

st.set_page_config(page_title="PM PO Projection Mailer", layout="wide")


def send_email(smtp_host, smtp_port, username, password, to_email, subject, text_body, html_body):
    msg = MIMEMultipart("alternative")
    msg.attach(MIMEText(text_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))
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
            # New id per generated report so supplier picks from a previous
            # report don't carry over onto different data.
            st.session_state["report_id"] = st.session_state.get("report_id", 0) + 1
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
    st.header("3. Assign suppliers")
    st.caption(
        "Tap a **suppliers** cell next to a week's projection and pick one or more suppliers. "
        "The projection is split evenly between the suppliers you pick "
        "(e.g. 1000 with 4 suppliers = 250 each)."
    )
    report_id = st.session_state.get("report_id", 0)
    selections = {}
    ignored = 0
    tabs = st.tabs(L.ITEM_TYPES)
    for item_type, tab in zip(L.ITEM_TYPES, tabs):
        with tab:
            sub = category_df[(category_df["Item Type"] == item_type) & (category_df["Total"] > 0)]
            if sub.empty:
                st.info(f"No {item_type} categories with a projection.")
                continue
            suppliers = L.SUPPLIERS_BY_TYPE.get(item_type, [])
            editor_df = pd.DataFrame({"Item Category": sub["Item Category"].values})
            column_config = {"Item Category": st.column_config.TextColumn(disabled=True, pinned=True)}
            # Only weeks where this type has some projection get a column pair.
            tab_weeks = [w for w in projection_weeks if (sub[f"Wk #{w}"] > 0).any()]
            for w in tab_weeks:
                qty_col, sup_col = f"Wk #{w}", f"Wk #{w} suppliers"
                editor_df[qty_col] = [f"{v:.0f}" if v > 0 else "" for v in sub[qty_col].values]
                editor_df[sup_col] = [[] for _ in range(len(sub))]
                column_config[qty_col] = st.column_config.TextColumn(L.week_label(w), disabled=True)
                column_config[sup_col] = st.column_config.MultiselectColumn(
                    f"{L.week_label(w)} suppliers", options=suppliers, width="medium",
                    help=f"Suppliers for the {L.week_label(w)} projection; it is split evenly between them.",
                )
            edited = st.data_editor(
                editor_df,
                column_config=column_config,
                hide_index=True,
                use_container_width=True,
                num_rows="fixed",
                key=f"alloc_{item_type}_{report_id}",
            )
            for _, row in edited.iterrows():
                for w in tab_weeks:
                    picked = list(row[f"Wk #{w} suppliers"] or [])
                    if not picked:
                        continue
                    if not row[f"Wk #{w}"]:
                        ignored += 1
                        continue
                    selections[(row["Item Category"], w)] = picked
    if ignored:
        st.warning(f"{ignored} supplier pick(s) are on weeks with no projection and were ignored.")

    alloc_df = L.allocate_to_suppliers(category_df, selections, projection_weeks)

    st.divider()
    st.header("4. Supplier-wise preview")
    if alloc_df.empty:
        st.info("Assign suppliers in section 3 to see the supplier-wise projection here.")
    else:
        st.markdown(L.build_allocation_html(alloc_df, projection_weeks), unsafe_allow_html=True)

    st.divider()
    st.header("5. Email suppliers")
    st.caption(
        "Requires SMTP credentials in Streamlit secrets: "
        "`smtp_host`, `smtp_port`, `smtp_username`, `smtp_password`, "
        "and a mapping of supplier name -> email address under `supplier_emails`."
    )
    allocated_suppliers = list(dict.fromkeys(alloc_df["Supplier"]))
    if not allocated_suppliers:
        st.info("No supplier has an assigned projection yet.")
    else:
        def email_parts(supplier):
            s_df = alloc_df[alloc_df["Supplier"] == supplier]
            return (
                L.build_supplier_allocation_text(supplier, s_df, projection_weeks),
                L.build_supplier_email_html(supplier, s_df, projection_weeks),
            )

        def send_to(supplier):
            text_body, html_body = email_parts(supplier)
            to_email = st.secrets["supplier_emails"][supplier]
            send_email(
                st.secrets["smtp_host"],
                int(st.secrets["smtp_port"]),
                st.secrets["smtp_username"],
                st.secrets["smtp_password"],
                to_email,
                f"Item Category-wise Projection — {supplier}",
                text_body,
                html_body,
            )
            return to_email

        email_supplier = st.selectbox("Send projection to", allocated_suppliers, key="email_supplier")
        st.write("**Email preview**")
        with st.container(border=True):
            st.markdown(email_parts(email_supplier)[1], unsafe_allow_html=True)

        col_one, col_all = st.columns(2)
        targets = []
        if col_one.button("Send email", type="primary"):
            targets = [email_supplier]
        if col_all.button(f"Send to all assigned suppliers ({len(allocated_suppliers)})"):
            targets = allocated_suppliers
        for supplier in targets:
            try:
                to_email = send_to(supplier)
                st.success(f"Email sent to {supplier} ({to_email}).")
            except KeyError as e:
                st.error(f"{supplier}: missing secret {e}. Configure SMTP + supplier_emails in Streamlit secrets.")
            except Exception as e:
                st.error(f"{supplier}: failed to send email: {e}")
else:
    st.info("Upload the three files in the sidebar and click **Generate report** to get started.")
