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
4. A **supplier-wise preview** table: supplier name merged down its rows,
   then Item category, Item type, and one column per week.
5. **Email a supplier** (or all assigned suppliers) their allocation, using
   the fixed template (see `logic.EMAIL_TEMPLATE_INTRO`) with the same table
   as an HTML email (plain-text fallback included).

## Files
- `app.py` — Streamlit UI only.
- `logic.py` — all parsing/business logic, no Streamlit dependency
  (importable/testable on its own — see the smoke test in chat history).
- `requirements.txt` — `streamlit` (1.50+ for the dropdown cells), `pandas`, `openpyxl`.
- `.streamlit/secrets.toml.example` — copy to `.streamlit/secrets.toml`
  (local) or paste into the app's Secrets box (Streamlit Community Cloud).

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
- **PO Issued window**: 2 weeks by default (configurable in the sidebar),
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
3. In the app's Settings → Secrets, paste the contents of
   `.streamlit/secrets.toml.example` filled in with real SMTP credentials
   (e.g. a Gmail address + app password, or your company's SMTP relay) and
   real supplier email addresses.
4. Deploy. Every time you push to the branch, Streamlit Community Cloud
   auto-redeploys.

## Known open items (as of last handoff)
- 5 CTN items ("...25 SC ENV TBGS" flavors: Garden Strawberry, Mint,
  Orange, Royal Lemon, Spicy Ginger) were manually confirmed to map to
  "25TB DC ENV (AFRICA)" in chat, but that override is **not yet wired
  into `CATEGORY_OVERRIDES`** — add their item codes there when known.
- Whether ANGLE/BOPP/THREAD/T-SHIRT/LAMINATED-ROLLS items should get their
  own rows in PO Projection (under coarse labels) instead of being
  excluded is still an open question for the user.
- Email sending uses plain SMTP (`smtplib`) with Streamlit secrets; no
  OAuth/Graph API integration yet.
