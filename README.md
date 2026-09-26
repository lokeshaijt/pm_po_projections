# PM PO Projection Mailer (Streamlit)

## What it does
1. Upload a **Pending PO** export, a **PM Requirement** workbook (Shortfall sheet),
   and the **Item Category Master** (ITEM NAME / Item Type / ITEM CATEGORY).
2. Generates a two-sheet report: **PO Issued** (flat PO list, 2-week delivery
   window incl. overdue) and **PO Projection** (one row per Item Category,
   week-wise MOQ-adjusted projection, categories with a projection only).
3. **Assign suppliers** — three tabs (**CFC / CTN / TRAY**), each a grid of
   that type's categories with a projection. Next to every week's quantity
   is a **suppliers** cell: tap it and pick one or more suppliers from the
   dropdown. The quantity is split evenly between the picked suppliers
   (1000 with 4 suppliers = 250 each; rounded to whole units).
4. A **supplier-wise projection** table: supplier name merged down its rows,
   then Item category, Item type, and one column per week. **Download full
   report** gives the PO Issued + PO Projection sheets plus this table as a
   **Supplier-wise Projection** sheet. All sheets use Calibri.
5. **Email a supplier** (or all listed suppliers). Each email has two
   tables: **PO Issued** (that supplier's pending POs in the delivery window,
   overdue included, negative days in red) and **PO Projection** (their
   allocation, using the fixed template in `logic.EMAIL_TEMPLATE_INTRO`).
   A section is left out when the supplier has nothing for it. Sent as HTML
   with a plain-text fallback. PO supplier names ("Customer Name" in the
   Pending PO file) are matched to the roster ignoring case, punctuation and
   Pvt/Ltd-style suffixes (`logic.match_roster_supplier`); names that don't
   match are listed in the app and their POs are not emailed.

## Files
- `app.py` — Streamlit UI only.
- `storage.py` — permanent settings store (GitHub branch, or local file).
- `ui.py` — branding/CSS helpers (header, section headings, summary cards);
  `.streamlit/config.toml` holds the JAY black & gold theme, `assets/` the logo.
- `logic.py` — all parsing/business logic, no Streamlit dependency
  (importable/testable on its own — see the smoke test in chat history).
- `requirements.txt` — `streamlit` (1.50+ for the dropdown cells), `pandas`, `openpyxl`.
- `.streamlit/secrets.toml.example` — copy to `.streamlit/secrets.toml`
  (local) or paste into the app's Secrets box (Streamlit Community Cloud).

## Supplier email IDs and Cc (permanent storage)
- **Supplier email IDs** section: **Add** → type the ID → **Save**, and
  **Remove** next to each saved ID. Several IDs per supplier; emails go to all.
- **Cc** (in the Email suppliers section): same Add/Save/Remove; saved Cc IDs
  are copied on every email, including Send to all.

Both are stored by `storage.py`. With a `[github]` table in secrets (see
`.streamlit/secrets.toml.example`) they are saved as `settings.json` on the
repo's **app-data** branch, so they survive restarts and redeploys, and
saving doesn't trigger a redeploy (only pushes to `main` do). The token
needs "Contents: Read and write" on this repo only. Without it the app
falls back to a local `settings.json` and shows a warning that the IDs are
not permanent. On first run the list is seeded from the optional
`[supplier_emails]` secrets table and `cc` in `[smtp]`.

## Business rules encoded in `logic.py`
- **MOQ tiers** (`TIERS`): CFC/TRAY [250,500,1000,3000,5000,10000],
  CTN [5000,...,200000], plus single-tier categories (T-SHIRT, ANGLE, BOPP,
  SLIP SHEET, THREAD, AL-WIRE = 1 tier each). Laminated-sheet items always
  use tier [300]; POUCH items with "GARANT" in the name use
  [5000,10000,25000,50000]. If the shortfall exceeds every tier, the
  Suggested MOQ is `shortfall * 1.02` (rounded).
- **Projection weeks**: the Requirement file's 8-week window is auto-detected
  (`Shortfall` sheet, "Week NN" headers); the **first 2 weeks are dropped**
  and the remaining 6 are used for projections (confirmed rule: "starting 2
  weeks" are excluded).
- **Category matching**: item description → Item Category Master (exact,
  normalized match). Items that don't match anything in the master are
  **excluded from PO Projection** (current scope is CFC/CTN/TRAY only —
  confirmed with the user; non-tea-bag items like ANGLE/BOPP/THREAD/
  T-SHIRT/LAMINATED-ROLLS have no home in the master and are dropped from
  this sheet, though they still appear in PO Issued).
- **Known naming mismatches**: `CATEGORY_OVERRIDES` in `logic.py` is an
  empty dict ready for item-code → category corrections (e.g. items whose
  live description says "SC ENV" where the master only has a "DC ENV"
  variant recorded). Add entries there as they're identified — do **not**
  build a blind SC/DC substitution rule, it was tested and does not
  generalize safely across flavors.
- **Suppliers** (`SUPPLIERS_BY_TYPE`): CFC and TRAY share the same 5
  suppliers (Abirami Packaging, Arr Print Pack, Mayura Packaging,
  SRI RANGA INDUSTRIES, Sri Venkateswara Packings); CTN has 3 (Lovely
  Offset Printers (P) Ltd, Salem Print Pack, Shraddha Saburi Printers
  Private Limited). Update this dict if the roster changes.
- **PO Issued window**: fixed at 14 days (`PO_WINDOW_DAYS` in `logic.py`),
  **including already-overdue** deliveries (not just future ones) — a
  pending PO with a past delivery date is still open and still urgent.
- **Delay styling**: "No. of days to arrive" shown in bold red with a red
  fill when negative; Delivery date itself stays plain (no red/"(Delay)"
  marking there — that was explicitly requested to be removed).

## Deploying to Streamlit Community Cloud
1. Create a new GitHub repo, push these files (`app.py`, `logic.py`,
   `requirements.txt`, this README). **Do not commit** a real
   `secrets.toml` — only the `.example` file.
2. Go to https://share.streamlit.io, "New app", point it at the repo,
   branch, and `app.py` as the entry point.
3. In the app's Settings → Secrets, paste the `[smtp]` table from
   `.streamlit/secrets.toml.example` filled in with real SMTP credentials
   (e.g. a Gmail address + app password, or your company's SMTP relay) and
   real supplier email addresses.
4. Deploy. Every time you push to the branch, Streamlit Community Cloud
   auto-redeploys.

## Known open items (as of last handoff)
- 5 CTN items (CTN GV Garden Strawberry / Mint / Orange / Royal Lemon /
  Spicy Ginger 25 SC ENV TBGS) are not in the Item Category Master and are
  mapped to "25TB DC ENV (AFRICA)" by item name in `CATEGORY_OVERRIDES`. If
  the Requirement file spells them differently, they show up under "Show
  items left out" in the app. Adding them to the master is the lasting fix.
- Whether ANGLE/BOPP/THREAD/T-SHIRT/LAMINATED-ROLLS items should get their
  own rows in PO Projection (under coarse labels) instead of being
  excluded is still an open question for the user.
- Email sending uses plain SMTP (`smtplib`) with Streamlit secrets; no
  OAuth/Graph API integration yet.
