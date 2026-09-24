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
from pathlib import Path
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import streamlit as st

import logic as L
import ui

st.set_page_config(page_title="PM PO Projection Mailer", page_icon=str(ui.LOGO_PATH), layout="wide")
ui.inject_css()


EMAILS_PATH = Path(__file__).parent / "supplier_emails.json"


def get_supplier_emails() -> dict:
    if "supplier_emails" not in st.session_state:
        try:
            seed = dict(st.secrets.get("supplier_emails", {}))
        except Exception:  # no secrets.toml at all
            seed = {}
        st.session_state["supplier_emails"] = L.load_supplier_emails(EMAILS_PATH, seed)
    return st.session_state["supplier_emails"]


def update_supplier_emails(emails: dict) -> None:
    st.session_state["supplier_emails"] = emails
    L.save_supplier_emails(EMAILS_PATH, emails)


def smtp_settings() -> dict:
    """SMTP settings from secrets, either as an [smtp] table
    (server, port, username, password, optional sender) or as flat
    smtp_host / smtp_port / smtp_username / smtp_password keys."""
    if "smtp" in st.secrets:
        cfg = st.secrets["smtp"]
        host, port, user, pwd = cfg["server"], cfg["port"], cfg["username"], cfg["password"]
        sender = cfg.get("sender", user)
    else:
        host, port = st.secrets["smtp_host"], st.secrets["smtp_port"]
        user, pwd = st.secrets["smtp_username"], st.secrets["smtp_password"]
        sender = st.secrets.get("smtp_sender", user)
    return {"host": host, "port": int(port), "username": user, "password": pwd, "sender": sender}


def send_email(to_emails, subject, text_body, html_body):
    cfg = smtp_settings()
    msg = MIMEMultipart("alternative")
    msg.attach(MIMEText(text_body, "plain"))
    msg.attach(MIMEText(html_body, "html"))
    msg["Subject"] = subject
    msg["From"] = cfg["sender"]
    msg["To"] = ", ".join(to_emails)
    with smtplib.SMTP(cfg["host"], cfg["port"]) as server:
        server.starttls()
        server.login(cfg["username"], cfg["password"])
        server.sendmail(cfg["sender"], to_emails, msg.as_string())


ui.hero(
    "PM PO Projection Mailer",
    "Generate the PO report, assign week-wise projections to suppliers, and email them directly.",
)

with st.sidebar:
    st.logo(str(ui.LOGO_PATH), size="large")
    ui.sidebar_title("① Upload files")
    po_file = st.file_uploader("Pending PO file (.xlsx)", type="xlsx", key="po_file")
    req_file = st.file_uploader("PM Requirement file (.xlsx)", type="xlsx", key="req_file")
    master_file = st.file_uploader(
        "Item Category Master (.xlsx)", type="xlsx", key="master_file",
        help="Columns: ITEM NAME, Item Type, ITEM CATEGORY (any order).",
    )
    today = st.date_input("As-of date", value=date.today())
    generate = st.button("Generate report", type="primary", use_container_width=True)
    ui.sidebar_note(f"PO Issued covers deliveries due in the next {L.PO_WINDOW_DAYS} days, plus overdue.")

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
            po_issued = L.build_po_issued_df(po_df, today, L.PO_WINDOW_DAYS)

            category_order = sorted(set(name_to_cat.values()))
            category_df = L.build_category_table(
                item_week_proj, item_to_cat, item_to_type, category_order, req_data["projection_weeks"]
            )

            st.session_state["unmatched_df"] = L.build_unmatched_items_df(req_data, item_week_proj, item_to_cat)
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

    unmatched_df = st.session_state.get("unmatched_df")
    overdue = int((po_issued["No. of days to arrive"] < 0).sum()) if len(po_issued) else 0
    ui.cards([
        ("PO Issued rows", f"{len(po_issued):,}", False),
        ("Overdue POs", f"{overdue:,}", overdue > 0),
        ("Categories with projection", f"{int((category_df['Total'] > 0).sum()):,}", False),
        ("Projection weeks", f"Wk {projection_weeks[0]}–{projection_weeks[-1]}" if projection_weeks else "—", False),
        ("Items left out", f"{0 if unmatched_df is None else len(unmatched_df):,}",
         unmatched_df is not None and not unmatched_df.empty),
    ])
    if unmatched_df is not None and not unmatched_df.empty:
        st.warning(
            f"{len(unmatched_df)} item(s) with a projection are not in the Item Category Master, "
            f"so {unmatched_df['Total'].sum():,.0f} units are left out of PO Projection. "
            "Add them to the master, or map them in `CATEGORY_OVERRIDES` in logic.py."
        )
        with st.expander("Show items left out"):
            st.dataframe(unmatched_df, hide_index=True, use_container_width=True)

    ui.section(2, "Download report", "PO Issued and PO Projection sheets.")
    st.download_button(
        "Download PM PO Report (.xlsx)",
        data=st.session_state["report_bytes"],
        file_name=f"PM_PO_Report_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    ui.section(
        3, "Assign suppliers",
        "Tap a <b>suppliers</b> cell next to a week's projection and pick one or more suppliers. "
        "The projection is split evenly between them (e.g. 1000 with 4 suppliers = 250 each).",
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
            column_config = {"Item Category": st.column_config.TextColumn(disabled=True)}
            # Only weeks where this type has some projection get a column pair.
            tab_weeks = [w for w in projection_weeks if (sub[f"Wk #{w}"] > 0).any()]
            for w in tab_weeks:
                qty_col, sup_col = f"Wk #{w}", f"Wk #{w} suppliers"
                editor_df[qty_col] = [f"{v:.0f}" if v > 0 else "" for v in sub[qty_col].values]
                editor_df[sup_col] = [[] for _ in range(len(sub))]
                column_config[qty_col] = st.column_config.TextColumn(L.week_label(w), disabled=True)
                column_config[sup_col] = st.column_config.MultiselectColumn(
                    f"{L.week_label(w)} suppliers", options=suppliers, width="medium", color=ui.GOLD_LIGHT,
                    help=f"Suppliers for the {L.week_label(w)} projection; it is split evenly between them.",
                )
            # Read-only columns are drawn faded by default; force dark text.
            text_cols = ["Item Category"] + [f"Wk #{w}" for w in tab_weeks]
            edited = st.data_editor(
                editor_df.style.set_properties(subset=text_cols, color="#111111"),
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

    ui.section(4, "Supplier-wise preview", "What each supplier will receive, week by week.")
    if alloc_df.empty:
        st.info("Assign suppliers in section 3 to see the supplier-wise projection here.")
    else:
        st.markdown(L.build_allocation_html(alloc_df, projection_weeks), unsafe_allow_html=True)
    st.download_button(
        "Download full report (.xlsx)",
        data=L.generate_report_workbook(po_issued, category_df, projection_weeks, alloc_df),
        file_name=f"PM_PO_Full_Report_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        help="PO Issued, PO Projection and Supplier-wise Projection sheets.",
    )

    ui.section(5, "Email suppliers")
    st.caption(
        "Requires SMTP credentials in Streamlit secrets: an `[smtp]` table with "
        "`server`, `port`, `username`, `password` (optional `sender`). "
        "Supplier email IDs are managed in the **Supplier email IDs** section below."
    )
    # Pending POs (PO Issued window) grouped under their roster supplier name.
    po_mail = po_issued.copy()
    po_mail["Roster Supplier"] = po_mail["Customer Name"].map(L.match_roster_supplier)
    unmatched = sorted({str(n) for n in po_mail.loc[po_mail["Roster Supplier"].isna(), "Customer Name"].dropna()})
    if unmatched:
        with st.expander(f"{len(unmatched)} PO supplier name(s) not on the supplier list — their POs are not emailed"):
            st.write(", ".join(unmatched))

    roster = list(dict.fromkeys(s for sups in L.SUPPLIERS_BY_TYPE.values() for s in sups))
    mail_suppliers = [
        s for s in roster
        if s in set(alloc_df["Supplier"]) or s in set(po_mail["Roster Supplier"].dropna())
    ]
    if not mail_suppliers:
        st.info("No supplier has an assigned projection or a pending PO yet.")
    else:
        def email_parts(supplier):
            s_df = alloc_df[alloc_df["Supplier"] == supplier]
            po_df = po_mail[po_mail["Roster Supplier"] == supplier]
            return (
                L.build_supplier_allocation_text(supplier, s_df, projection_weeks, po_df),
                L.build_supplier_email_html(supplier, s_df, projection_weeks, po_df),
            )

        def send_to(supplier):
            text_body, html_body = email_parts(supplier)
            to_emails = get_supplier_emails().get(supplier, [])
            if not to_emails:
                raise ValueError("no email ID saved — add one under Supplier email IDs")
            send_email(
                to_emails,
                f"PO Issued & Item Category-wise Projection — {supplier}",
                text_body,
                html_body,
            )
            return ", ".join(to_emails)

        email_supplier = st.selectbox("Send email to", mail_suppliers, key="email_supplier")
        st.write("**Email preview**")
        recipients = get_supplier_emails().get(email_supplier, [])
        st.caption("To: " + (", ".join(recipients) if recipients else "— no email ID saved for this supplier —"))
        with st.container(border=True):
            st.markdown(email_parts(email_supplier)[1], unsafe_allow_html=True)

        col_one, col_all = st.columns(2)
        targets = []
        if col_one.button("Send email", type="primary"):
            targets = [email_supplier]
        if col_all.button(f"Send to all listed suppliers ({len(mail_suppliers)})"):
            targets = mail_suppliers
        for supplier in targets:
            try:
                to_email = send_to(supplier)
                st.success(f"Email sent to {supplier} ({to_email}).")
            except KeyError as e:
                st.error(f"{supplier}: missing secret {e}. Configure SMTP in Streamlit secrets.")
            except Exception as e:
                st.error(f"{supplier}: failed to send email: {e}")
else:
    st.info("Upload the three files in the sidebar and click **Generate report** to get started.")
    ui.getting_started()

ui.section("@", "Supplier email IDs", "Emails go to every ID saved for the supplier.")
emails = get_supplier_emails()
for supplier in L.all_suppliers():
    saved = emails.get(supplier, [])
    types = "/".join(t for t, sups in L.SUPPLIERS_BY_TYPE.items() if supplier in sups)
    label = f"{supplier} ({types})" + (f" — {', '.join(saved)}" if saved else " — no email ID")
    with st.expander(label, expanded=st.session_state.get("email_open") == supplier):
        for addr in saved:
            c_addr, c_rm = st.columns([5, 1])
            c_addr.write(addr)
            if c_rm.button("Remove", key=f"rm_{supplier}_{addr}"):
                st.session_state["email_open"] = supplier
                update_supplier_emails({**emails, supplier: [e for e in saved if e != addr]})
                st.rerun()

        adding_key = f"adding_{supplier}"
        if not st.session_state.get(adding_key):
            if st.button("Add", key=f"add_{supplier}"):
                st.session_state["email_open"] = supplier
                st.session_state[adding_key] = True
                st.rerun()
        else:
            with st.form(f"form_{supplier}", clear_on_submit=True, border=False):
                new_addr = st.text_input("Email ID", placeholder="name@company.com")
                c_save, c_cancel = st.columns(2)
                save = c_save.form_submit_button("Save", type="primary")
                cancel = c_cancel.form_submit_button("Cancel")
            if cancel:
                st.session_state[adding_key] = False
                st.rerun()
            if save:
                new_addr = new_addr.strip()
                if not L.is_valid_email(new_addr):
                    st.error("Enter a valid email ID, e.g. name@company.com.")
                elif new_addr.lower() in (e.lower() for e in saved):
                    st.warning(f"{new_addr} is already saved for {supplier}.")
                else:
                    update_supplier_emails({**emails, supplier: saved + [new_addr]})
                    st.session_state[adding_key] = False
                    st.rerun()
