// ============================================================
// CRM MULTI-FILE SYSTEM — Google Apps Script
// VERSION 14.0 — Order Ledger / Client Summary / Client Dues / Dashboard
// ============================================================
// WHAT'S NEW IN v14.0:
//
//   1. Four new Hub sheets created by setupHubSheet():
//      • Order Ledger      — one row per order (detailed ledger)
//      • Client Summary    — one row per client (aggregated totals)
//      • Client Dues       — auto-populated unpaid/partial orders
//      • Dashboard         — all clients ranked by due balance
//
//   2. CRM Spoke: Client Dues sheet added (filtered view for each CRM).
//
//   3. Client Sub-Tab: Due Date column added (col P / SUB_COL_DUE_DATE).
//      Due Date lock: cannot be changed within 7 days of being set.
//
//   4. Sales Manager role merged into Sales (single 'Sales' role).
//      All former Sales Manager capabilities now live under Sales.
//      getCrmConfig_() backward-compat alias: salesManagerEmails → salesEmails.
//
//   5. Client Credit sheet restructured:
//      Removed: Request New Limit, Request Reason, Request Status, Approved By.
//      Added: Client Email (col F). UUID scratch col moved to col H.
//
//   6. onCRMSubTabEdit v14.0:
//      • Syncs Due Date (sub-tab col P) to CRM Main + Hub Order Ledger
//      • Two-way sync: Inv No + Inv Date now flow from CRM sub-tab → Hub
//      • Duplicate PI No check (per client tab)
//      • Duplicate Inv No check (across all tabs in this CRM spoke)
//      • Row background colour applied after every edit
//
//   7. onMasterHubEdit v14.0:
//      • Watches 'Order Ledger' sheet (not 'Master Hub')
//      • Due Date entered by Sales → pushed to CRM sub-tab + Client Dues
//      • Inv No / Inv Date entered by Logistics → pushed to CRM spoke
//
//   8. New automation (Session 3):
//      • Daily payment reminders (7-day warning, due today, overdue escalation)
//      • Order Summary Gmail Draft builder
//      • Toggle ON/OFF from ⚡ Automation menu
//
//   9. refreshClientDues_() — populates Client Dues on Hub + all CRM spokes.
//  10. refreshDashboard_() — populates Dashboard ranked by total due balance.
//  11. updateClientSummary_() — aggregates per-client totals to Client Summary.
//
//   All v13.0 features retained (role-based protection, credit checks,
//   payment/dispatch/approval notifications, activity log, etc.)
// ============================================================

// ============================================================
// SECTION 1 — CONSTANTS & CONFIGURATION (v14.0)
// ============================================================

// ============================================================
// === SESSION 1 OUTPUT ===
// CRM MULTI-FILE SYSTEM — Google Apps Script
// Changes: Constants, Column Definitions, Setup Functions
// Version: v14.0 — Order Ledger / Client Summary / Dues / Dashboard
// ============================================================
// SCOPE OF THIS FILE:
//   1. All updated constants (sheet names + column indices)
//   2. Updated setupHubSheet()   — creates 4 new Hub sheets
//   3. Updated setupCRMSheet()   — renames Main → "Order Ledger CRM",
//                                   adds Client Dues tab,
//                                   updates Client Sub-Tab headers
//   4. Updated setupCreditSheet() — removes cols F–I, adds col F = Client Email
//   5. Merged Sales Manager + Sales → single role 'Sales'
//      (role constants + config arrays + dropdown + protections)
//
// NOT IN THIS FILE:
//   • No trigger code
//   • No sync / onEdit logic
//   • No email / notification helpers
//   • No credit-check / refresh helpers
// ============================================================


// ── THE ONLY HARDCODED VALUE IN THE ENTIRE SYSTEM ────────────
const MASTER_HUB_ID = '17CtXl662TlzXjjVR8VXJhJbZGvrlGNN61ryspaBmjbU'; // TODO: replace with your Master Hub Sheet ID

// ── Config Sheet Name ─────────────────────────────────────────
const MEMBERS_SHEET_NAME  = 'Team Members';

// ── CRM Members sheet column indices (1-based) ────────────────
// Col A: Gmail      — the person's Google account email
// Col B: File ID    — their CRM spoke spreadsheet ID (blank for hub roles)
// Col C: Role       — CRM | Logistics | Accounts | Sales | Admin
//         NOTE v14.0: 'Sales Manager' and 'Sales' are now merged into 'Sales'.
//         Any existing 'Sales Manager' rows should be updated to 'Sales'.
// Col D: Name       — display name
// Col E: Active     — TRUE / FALSE (inactive rows are ignored by scripts)
const MC_COL_GMAIL    = 1;
const MC_COL_FILE_ID  = 2;
const MC_COL_ROLE     = 3;
const MC_COL_NAME     = 4;
const MC_COL_ACTIVE   = 5;
const MC_TOTAL_COLS   = 5;

// ── Sheet Name Constants — Hub ────────────────────────────────
const HUB_MAIN_SHEET_NAME         = 'Master Hub';
const HUB_ORDER_LEDGER_SHEET      = 'Order Ledger';       // NEW v14.0 — one row per order
const HUB_CLIENT_SUMMARY_SHEET    = 'Client Summary';     // NEW v14.0 — one row per client
const HUB_DUES_SHEET              = 'Client Dues';        // NEW v14.0 — auto-populated
const HUB_DASHBOARD_SHEET         = 'Dashboard';          // NEW v14.0 — all clients by due balance
const CREDIT_SHEET_NAME           = 'Credit List';
const LOG_SHEET_NAME              = 'Activity Log';

// ── Sheet Name Constants — CRM Spoke ─────────────────────────
const MAIN_SHEET_NAME             = 'Main';               // legacy name kept for sync logic compatibility
const CRM_ORDER_LEDGER_SHEET      = 'Order Ledger CRM';   // NEW v14.0 — replaces 'Main' display name
const CRM_DUES_SHEET              = 'Client Dues';        // NEW v14.0 — CRM-filtered dues view

// ── Helper: get CRM main sheet by either old or new name ─────
// Always use this instead of getSheetByName(MAIN_SHEET_NAME)
// so both old 'Main' and new 'Order Ledger CRM' tab names work.
function getCRMMainSheet_(ss) {
  return ss.getSheetByName(CRM_ORDER_LEDGER_SHEET)
      || ss.getSheetByName('Main');
}

// ── Aliases for Session 2/3 generated code ───────────────────
// Session 2 and 3 used different constant names than Session 1.
// These aliases map Session 2/3 names → Session 1 constants so
// all code compiles without changes.
const ORDER_LEDGER_SHEET_NAME   = HUB_ORDER_LEDGER_SHEET;   // 'Order Ledger'
const ORDER_LEDGER_CRM_NAME     = CRM_ORDER_LEDGER_SHEET;   // 'Order Ledger CRM'
const CLIENT_DUES_SHEET_NAME    = HUB_DUES_SHEET;           // 'Client Dues'
const CLIENT_SUMMARY_SHEET_NAME = HUB_CLIENT_SUMMARY_SHEET; // 'Client Summary'
const DASHBOARD_SHEET_NAME      = HUB_DASHBOARD_SHEET;      // 'Dashboard'

// ── Legacy Credit Sheet column aliases ───────────────────────
// v14.0 removed CC_COL_REQ_INCREASE/REASON/STATUS/APPROVED_BY from the
// Credit sheet. Some older functions still reference them.
// Map them to col 6 (Client Email — the new col F) as a safe fallback
// so those functions don't crash. They will no-op on write since the
// columns no longer exist in the sheet.
const CC_COL_REQ_INCREASE  = 6;   // was Request New Limit — now Client Email
const CC_COL_REQ_REASON    = 6;   // was Request Reason    — now Client Email
const CC_COL_REQ_STATUS    = 6;   // was Request Status    — now Client Email
const CC_COL_APPROVED_BY_CC = 6;  // was Approved By (CC)  — now Client Email

// ── Master Hub "Master Hub" Sheet Column Indices (1-based) ────
// Unchanged from v13 — kept for full backward compatibility with all sync/trigger code.
const HUB_COL_UUID            = 1;   // A
const HUB_COL_CUSTOMER        = 2;   // B
const HUB_COL_CRM_OWNER       = 3;   // C
const HUB_COL_PI_NO           = 4;   // D
const HUB_COL_INV_NO          = 5;   // E
const HUB_COL_DATE            = 6;   // F  Inv Date
const HUB_COL_INV_VALUE       = 7;   // G
const HUB_COL_PAY_STATUS      = 8;   // H  formula
const HUB_COL_DUE_DATE        = 9;   // I
const HUB_COL_PAY_RECEIVED    = 10;  // J  Accounts-owned
const HUB_COL_PAY_DATE        = 11;  // K  Accounts-owned
const HUB_COL_DISPATCH_STATUS = 12;  // L  formula
const HUB_COL_DUE_BALANCE     = 13;  // M  formula
const HUB_COL_APPROVAL_REQ    = 14;  // N
const HUB_COL_APPROVED_BY     = 15;  // O  Sales types here
const HUB_COL_REMARKS         = 16;  // P
const HUB_COL_AWB_NO          = 17;  // Q  Logistics-owned
const HUB_COL_DISP_DATE       = 18;  // R  Logistics-owned
// v15.2 FIX: was 18 (a leftover from the old 18-column "Master Hub" layout).
// Order Ledger actually has 23 columns now (20 original + Approval Required,
// Approved By, Remarks). At 18, every rowData read in onMasterHubEdit silently
// truncated before reaching Due Date/Due Balance/the new approval columns.
const HUB_TOTAL_COLS          = 24;

// ── Order Ledger Sheet Column Indices (Hub) (1-based) ─────────
// NEW v14.0 — "Order Ledger" sheet, one row per order
const OL_COL_UUID             = 1;   // A  UUID
const OL_COL_CLIENT_NAME      = 2;   // B  Client Name
const OL_COL_CRM_OWNER        = 3;   // C  CRM Owner
const OL_COL_PI_NO            = 4;   // D  PI No
const OL_COL_INV_NO           = 5;   // E  Inv No
const OL_COL_INV_DATE         = 6;   // F  Inv Date
const OL_COL_ORDER_TYPE       = 7;   // G  Order Type
const OL_COL_ITEM_VALUE       = 8;   // H  Item Value
const OL_COL_DEV_CHARGES      = 9;   // I  Dev Charges
const OL_COL_FREIGHT          = 10;  // J  Freight
const OL_COL_DISCOUNT         = 11;  // K  Discount
const OL_COL_FINAL_INVOICE    = 12;  // L  Final Invoice
const OL_COL_AWB_NO           = 13;  // M  AWB No
const OL_COL_DISPATCH_DATE    = 14;  // N  Dispatch Date
const OL_COL_DISPATCH_STATUS  = 15;  // O  Dispatch Status
const OL_COL_PAY_RECEIVED     = 16;  // P  Payment Received
const OL_COL_PAY_DATE         = 17;  // Q  Payment Date
const OL_COL_PAY_STATUS       = 18;  // R  Payment Status
const OL_COL_DUE_DATE         = 19;  // S  Due Date
const OL_COL_DUE_BALANCE      = 20;  // T  Due Balance
// v15.2 NEW — requested columns at the end of Order Ledger
const OL_COL_APPROVAL_REQ     = 21;  // U  Approval Required — script-computed
const OL_COL_APPROVED_BY      = 22;  // V  Approved By — Admin/Sales dropdown
const OL_COL_REMARKS          = 23;  // W  Remarks — free text
const OL_COL_DUE_DATE_SET_ON  = 24;  // X  hidden — tracks 7-day Due Date lock
const OL_TOTAL_COLS           = 20;

// ── Client Summary Sheet Column Indices (Hub) (1-based) ───────
// NEW v14.0 — "Client Summary" sheet, one row per client
const CS_COL_CLIENT_NAME      = 1;   // A  Client Name
const CS_COL_CRM_OWNER        = 2;   // B  CRM Owner
const CS_COL_TOTAL_ORDERS     = 3;   // C  Total Orders
const CS_COL_TOTAL_INVOICED   = 4;   // D  Total Invoiced
const CS_COL_TOTAL_PAID       = 5;   // E  Total Paid
const CS_COL_TOTAL_DUE        = 6;   // F  Total Due
const CS_COL_CREDIT_LIMIT     = 7;   // G  Credit Limit
const CS_COL_USED_CREDIT      = 8;   // H  Used Credit
const CS_COL_AVAIL_CREDIT     = 9;   // I  Available Credit
const CS_COL_CREDIT_STATUS    = 10;  // J  Credit Status
const CS_COL_LAST_ORDER_DATE  = 11;  // K  Last Order Date
const CS_TOTAL_COLS           = 11;

// ── Client Dues Sheet Column Indices (Hub) (1-based) ──────────
// NEW v14.0 — "Client Dues" sheet, auto-populated (no manual entry)
// Mirrors Order Ledger but filtered to unpaid/partial orders only.
// Sorted by Due Date ascending.
const DUE_COL_CLIENT_NAME     = 1;   // A  Client Name
const DUE_COL_CRM_OWNER       = 2;   // B  CRM Owner
const DUE_COL_PI_NO           = 3;   // C  PI No
const DUE_COL_INV_NO          = 4;   // D  Inv No
const DUE_COL_FINAL_INVOICE   = 5;   // E  Final Invoice
const DUE_COL_PAY_RECEIVED    = 6;   // F  Payment Received
const DUE_COL_DUE_BALANCE     = 7;   // G  Due Balance
const DUE_COL_DUE_DATE        = 8;   // H  Due Date
const DUE_COL_PAY_STATUS      = 9;   // I  Payment Status
const DUE_TOTAL_COLS          = 9;

// ── Dashboard Sheet Column Indices (Hub) (1-based) ────────────
// NEW v14.0 — "Dashboard" sheet, all clients sorted by due balance desc
const DB_COL_CLIENT_NAME      = 1;   // A  Client Name
const DB_COL_CRM_OWNER        = 2;   // B  CRM Owner
const DB_COL_TOTAL_DUE        = 3;   // C  Total Due Balance
const DB_COL_TOTAL_INVOICED   = 4;   // D  Total Invoiced
const DB_COL_TOTAL_PAID       = 5;   // E  Total Paid
const DB_COL_OVERDUE_COUNT    = 6;   // F  Overdue Orders (past Due Date, unpaid)
const DB_COL_CREDIT_STATUS    = 7;   // G  Credit Status
const DB_COL_LAST_ORDER_DATE  = 8;   // H  Last Order Date
const DB_TOTAL_COLS           = 8;

// ── Client Credit Sheet Column Indices (1-based) ─────────────
// v14.0 CHANGE: Removed old cols F (Req New Limit), G (Req Reason),
//   H (Req Status), I (Approved By). Added new col F = Client Email.
//   Col G = Last Updated (was col J).  UUID scratch col is now col H (hidden).
const CC_COL_CLIENT_NAME  = 1;   // A
const CC_COL_CRM_OWNER    = 2;   // B
const CC_COL_CREDIT_LIMIT = 3;   // C
const CC_COL_USED_CREDIT  = 4;   // D  — script-maintained, do not edit
const CC_COL_AVAIL_CREDIT = 5;   // E  — script-maintained, do not edit
const CC_COL_CLIENT_EMAIL = 6;   // F  NEW v14.0
const CC_COL_LAST_UPDATED = 7;   // G  (was col J in v13)
const CC_COL_UUID         = 8;   // H  hidden internal scratch col (was col K)
const CC_TOTAL_COLS       = 7;   // Visible cols A–G (UUID col H is appended separately)

// ── CRM Spoke "Order Ledger CRM" Sheet Column Indices (1-based) ─
// v14.0 CHANGE: 'Main' sheet is now labelled/created as 'Order Ledger CRM'.
//   CRM_COL_* constants keep the SAME numbers so all sync/trigger code
//   that references CRM_COL_* requires zero changes.
//   The only structural change: CRM_COL_CRM_OWNER (col C) is REMOVED from
//   the CRM spoke sheet header — the owner is implicit (it IS the CRM user).
//   The column still exists in constants for Hub-mapping purposes.
const CRM_COL_UUID            = 1;   // A  hidden
const CRM_COL_CUSTOMER        = 2;   // B  Client Name
const CRM_COL_CRM_OWNER       = 3;   // C  HIDDEN in CRM spoke (owner is implicit)
const CRM_COL_PI_NO           = 4;   // D
const CRM_COL_INV_NO          = 5;   // E
const CRM_COL_DATE            = 6;   // F  Inv Date
const CRM_COL_INV_VALUE       = 7;   // G
const CRM_COL_PAY_STATUS      = 8;   // H
const CRM_COL_DUE_DATE        = 9;   // I
const CRM_COL_PAY_RECEIVED    = 10;  // J  filled by Hub script
const CRM_COL_PAY_DATE        = 11;  // K  filled by Hub script
const CRM_COL_DISPATCH_STATUS = 12;  // L  formula
const CRM_COL_DUE_BALANCE     = 13;  // M  formula
const CRM_COL_APPROVAL_REQ    = 14;  // N
const CRM_COL_APPROVED_BY     = 15;  // O
const CRM_COL_REMARKS         = 16;  // P
// v14.0: AWB No and Dispatch Date pushed from Hub → Order Ledger CRM
const CRM_COL_AWB_NO          = 17;  // Q  AWB No (from Logistics via Hub)
const CRM_COL_DISP_DATE       = 18;  // R  Dispatch Date (from Logistics via Hub)

// ── Client Sub-Tab Column Indices (1-based) ───────────────────
// v14.0 ADDITION: SUB_COL_DUE_DATE added at col P (16).
const SUB_COL_PI_NO           = 1;
const SUB_COL_INV_NO          = 2;
const SUB_COL_INV_DATE        = 3;   // Inv Date
const SUB_COL_ORDER_TYPE      = 4;
const SUB_COL_AWB_NO          = 5;
const SUB_COL_DISP_DATE       = 6;
const SUB_COL_ITEM_VALUE_USD  = 7;
const SUB_COL_DEV_CHARGES_USD = 8;
const SUB_COL_FOB_VALUE       = 9;
const SUB_COL_FREIGHT         = 10;
const SUB_COL_CIF_VALUE       = 11;
const SUB_COL_DISCOUNT        = 12;
const SUB_COL_FINAL_INV       = 13;
const SUB_COL_PAY_RECEIVED    = 14;
const SUB_COL_PAY_DATE        = 15;
const SUB_COL_DUE_DATE        = 16;  // NEW v14.0 — col P

// ── Column maps ───────────────────────────────────────────────
const HUB_ACCOUNTS_COLS     = [HUB_COL_PAY_RECEIVED, HUB_COL_PAY_DATE];
const HUB_LOGISTICS_COLS    = [HUB_COL_AWB_NO, HUB_COL_DISP_DATE];
// HUB_ALLOWED_PUSH_COLS: Order Ledger columns that trigger Hub→CRM push
// v14.0: uses OL_COL_* since onMasterHubEdit now watches Order Ledger sheet
const HUB_ALLOWED_PUSH_COLS = [
  OL_COL_PI_NO,           // D  v15.3 NEW — Admin corrects PI No
  OL_COL_PAY_RECEIVED,    // P  Accounts enters Payment Received
  OL_COL_PAY_DATE,        // Q  Accounts enters Payment Date
  OL_COL_INV_NO,          // E  Logistics or Accounts enters Inv No
  OL_COL_INV_DATE,        // F  Logistics or Accounts enters Inv Date
  OL_COL_ORDER_TYPE,      // G  v15.3 NEW — Admin corrects Order Type
  OL_COL_ITEM_VALUE,      // H  v15.3 NEW — Admin corrects Item Value
  OL_COL_DEV_CHARGES,     // I  v15.3 NEW — Admin corrects Dev Charges
  OL_COL_FREIGHT,         // J  v15.3 NEW — Admin corrects Freight
  OL_COL_DISCOUNT,        // K  v15.3 NEW — Admin corrects Discount
  OL_COL_FINAL_INVOICE,   // L  Final Invoice value → syncs to CRM col G
  OL_COL_AWB_NO,          // M  Logistics enters AWB No
  OL_COL_DISPATCH_DATE,   // N  Logistics enters Dispatch Date
  OL_COL_DISPATCH_STATUS, // O  Script sets Dispatch Status
  OL_COL_DUE_DATE,        // S  Accounts enters Due Date
  OL_COL_APPROVAL_REQ,    // U  v15.3 NEW — safety net for manual Admin edits
                          //    (normally script-only via reEvaluateAllCreditFlags_)
  OL_COL_APPROVED_BY,     // V  v15.2 NEW — Admin/Sales sets Approved By
  OL_COL_REMARKS,         // W  v15.2 NEW — notes field
];

// HUB_TO_CRM_MAIN: Order Ledger col → CRM Order Ledger CRM col (push back to CRM)
// When Hub/Logistics/Accounts edit these in Order Ledger → sync to CRM spoke
// v15.2 FIX: removed AWB No / Dispatch Date — per request, these should only
// ever flow into the individual client sub-tab (HUB_TO_SUB_TAB below), not
// into Order Ledger CRM's own Q/R columns. Item Value/Dev Charges/Freight/
// Discount/Final Invoice/Payment Status/Due Balance are deliberately NOT
// mapped here — Order Ledger CRM has no individual columns for the first four
// (only a combined Inv Value/G), and the last two are LIVE FORMULAS on the
// CRM side (col H, col M) — pushing a static Hub value into a formula cell
// would silently break that formula.
const HUB_TO_CRM_MAIN = {
  [OL_COL_PI_NO]:          CRM_COL_PI_NO,          // v15.3 NEW
  [OL_COL_PAY_RECEIVED]:   CRM_COL_PAY_RECEIVED,
  [OL_COL_PAY_DATE]:       CRM_COL_PAY_DATE,
  [OL_COL_INV_NO]:         CRM_COL_INV_NO,
  [OL_COL_INV_DATE]:       CRM_COL_DATE,
  [OL_COL_FINAL_INVOICE]:  CRM_COL_INV_VALUE,      // Final Invoice → CRM col G
  [OL_COL_DISPATCH_STATUS]:CRM_COL_DISPATCH_STATUS,
  [OL_COL_DUE_DATE]:       CRM_COL_DUE_DATE,
  [OL_COL_APPROVAL_REQ]:   CRM_COL_APPROVAL_REQ,   // v15.3 NEW — safety net
  [OL_COL_APPROVED_BY]:    CRM_COL_APPROVED_BY,    // v15.2 NEW
  [OL_COL_REMARKS]:        CRM_COL_REMARKS,        // v15.2 NEW
};

// HUB_TO_SUB_TAB: Order Ledger col → CRM sub-tab col (push back to sub-tab too)
// v15.3 NEW: PI No, Order Type, Item Value, Dev Charges, Freight, Discount —
// these are the fields CRM normally enters themselves, but if Admin corrects
// one directly on the Hub, it should flow back down so CRM's sub-tab reflects
// the correction. Final Invoice/Due Balance are NOT pushed here — they're
// live formulas in the sub-tab (col M, computed from G/H/J/L), and overwriting
// a formula cell with a static value would break it on that row permanently.
const HUB_TO_SUB_TAB = {
  [OL_COL_PI_NO]:          SUB_COL_PI_NO,           // v15.3 NEW
  [OL_COL_PAY_RECEIVED]:   SUB_COL_PAY_RECEIVED,
  [OL_COL_PAY_DATE]:       SUB_COL_PAY_DATE,
  [OL_COL_INV_NO]:         SUB_COL_INV_NO,
  [OL_COL_INV_DATE]:       SUB_COL_INV_DATE,
  [OL_COL_ORDER_TYPE]:     SUB_COL_ORDER_TYPE,      // v15.3 NEW
  [OL_COL_ITEM_VALUE]:     SUB_COL_ITEM_VALUE_USD,  // v15.3 NEW
  [OL_COL_DEV_CHARGES]:    SUB_COL_DEV_CHARGES_USD, // v15.3 NEW
  [OL_COL_FREIGHT]:        SUB_COL_FREIGHT,         // v15.3 NEW
  [OL_COL_DISCOUNT]:       SUB_COL_DISCOUNT,        // v15.3 NEW
  [OL_COL_AWB_NO]:         SUB_COL_AWB_NO,
  [OL_COL_DISPATCH_DATE]:  SUB_COL_DISP_DATE,
  [OL_COL_DUE_DATE]:       SUB_COL_DUE_DATE,
};

// SUB_TO_HUB: maps sub-tab column → Master Hub column
// These columns sync CRM sub-tab → Master Hub (legacy sheet)
// AND Master Hub → CRM sub-tab (two-way)
const SUB_TO_HUB = {
  [SUB_COL_PI_NO]:    HUB_COL_PI_NO,      // A → D  PI No
  [SUB_COL_INV_NO]:   HUB_COL_INV_NO,     // B → E  Inv No (CRM or Logistics)
  [SUB_COL_INV_DATE]: HUB_COL_DATE,       // C → F  Inv Date (CRM or Logistics)
  [SUB_COL_AWB_NO]:   HUB_COL_AWB_NO,     // E → Q  AWB No (Logistics or CRM)
  [SUB_COL_DISP_DATE]:HUB_COL_DISP_DATE,  // F → R  Dispatch Date (Logistics or CRM)
};

// SUB_TO_ORDER_LEDGER: maps sub-tab column → Order Ledger column
// These are the columns that also write to the new Order Ledger sheet
const SUB_TO_ORDER_LEDGER = {
  [SUB_COL_PI_NO]:    OL_COL_PI_NO,
  [SUB_COL_INV_NO]:   OL_COL_INV_NO,
  [SUB_COL_INV_DATE]: OL_COL_INV_DATE,
  [SUB_COL_AWB_NO]:   OL_COL_AWB_NO,
  [SUB_COL_DISP_DATE]:OL_COL_DISPATCH_DATE,
};

// ── Internal write loop-guard key (PropertiesService) ────────
const PROP_SKIP_NEXT_EDIT = 'skip_next_crm_edit';

// ── Role name constant (merged Sales Manager + Sales → 'Sales') ─
// v14.0: 'Sales Manager' role is retired. All members formerly in
// 'Sales Manager' should be updated to 'Sales' in the CRM Members sheet.
// The single 'Sales' role retains all capabilities of both old roles:
//   • Add new CRM Members
//   • Add/edit clients in Client Credit (all columns)
//   • Approve credit-exceeded orders (col O in Hub)
const ROLE_CRM        = 'CRM';
const ROLE_LOGISTICS  = 'Logistics';
const ROLE_ACCOUNTS   = 'Accounts';
const ROLE_SALES      = 'Sales';          // covers former Sales + Sales Manager
const ROLE_ADMIN      = 'Admin';
const VALID_ROLES     = [ROLE_CRM, ROLE_LOGISTICS, ROLE_ACCOUNTS, ROLE_SALES, ROLE_ADMIN];


// ============================================================
// getCrmConfig_()  — v14.0: Sales Manager merged into Sales
// ─────────────────────────────────────────────────────────────
// Returns:
// {
//   crmFileIds:    { 'gmail': 'fileId', ... },
//   salesEmails:   ['email', ...],   // covers former Sales + Sales Manager
//   adminEmails:   ['email', ...],
//   logisticsEmails: [...],
//   accountsEmails:  [...],
//   allMembers:    [ { gmail, fileId, role, name, active }, ... ]
// }
// NOTE: salesManagerEmails is kept as an ALIAS to salesEmails for backward
// compatibility with any existing code that references it. Both point to the
// same array. Remove the alias once all callers are updated.
// ============================================================


// ============================================================
// setupMembersSheet()  — v14.0: dropdown updated (Sales Manager removed)
// ============================================================


// ============================================================
// setupHubSheet()  — v14.0
// ─────────────────────────────────────────────────────────────
// Creates / refreshes ALL Hub sheets:
//   1. Master Hub   (unchanged core sheet)
//   2. Order Ledger (one row per order — detailed ledger)
//   3. Client Summary (one row per client — aggregate view)
//   4. Client Dues  (auto-populated — unpaid/partial orders)
//   5. Dashboard    (all clients sorted by due balance desc)
//   6. Client Credit (via setupCreditSheet())
// ============================================================
function setupHubSheet() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = Session.getEffectiveUser().getEmail();
    const config     = getCrmConfig_();
    const crmSpokeEmails  = Object.keys(config.crmFileIds);
    const adminEditors     = [...new Set([ownerEmail, ...config.adminEmails,     ...crmSpokeEmails])];
    const logisticsEditors = [...new Set([ownerEmail, ...config.logisticsEmails, ...crmSpokeEmails])];
    const accountsEditors  = [...new Set([ownerEmail, ...config.accountsEmails,  ...crmSpokeEmails])];
    const approvalEditors  = [...new Set([ownerEmail, ...config.adminEmails,
      ...config.salesEmails, ...crmSpokeEmails])];

    // Master Hub removed v15.0 — delete if still exists
    const _legacyHub1 = ss.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (_legacyHub1) ss.deleteSheet(_legacyHub1);

    // ── 2. Order Ledger sheet ─────────────────────────────────
    setupHubOrderLedgerSheet_(ss, adminEditors, logisticsEditors, accountsEditors);

    // ── 3. Client Summary sheet ───────────────────────────────
    setupHubClientSummarySheet_(ss, adminEditors);

    // ── 4. Client Dues sheet ──────────────────────────────────
    setupHubDuesSheet_(ss, adminEditors);

    // ── 5. Dashboard sheet ────────────────────────────────────
    setupHubDashboardSheet_(ss, adminEditors);

    // ── 6. Client Credit sheet ────────────────────────────────
    setupCreditSheet();

    logActivity_(ss, 'HUB SETUP', 'setupHubSheet() v14.0', Session.getActiveUser().getEmail());
    alertIfUi_(
      '✅ Master Hub setup complete (v14.0).\n\n' +
      'Sheets created/updated:\n' +
      '• Master Hub\n• Order Ledger\n• Client Summary\n• Client Dues\n• Dashboard\n• Client Credit'
    );
  } catch (err) {
    Logger.log('ERROR in setupHubSheet: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}


// ============================================================
// setupHubOrderLedgerSheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Order Ledger" sheet (one row per order).
// Columns: UUID | Client Name | CRM Owner | PI No | Inv No |
//          Inv Date | Order Type | Item Value | Dev Charges |
//          Freight | Discount | Final Invoice | AWB No |
//          Dispatch Date | Dispatch Status | Payment Received |
//          Payment Date | Payment Status | Due Date | Due Balance
// ============================================================


// ============================================================
// setupHubClientSummarySheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Client Summary" sheet (one row per client).
// Columns: Client Name | CRM Owner | Total Orders | Total Invoiced |
//          Total Paid | Total Due | Credit Limit | Used Credit |
//          Available Credit | Credit Status | Last Order Date
// ============================================================


// ============================================================
// setupHubDuesSheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Client Dues" sheet (auto-populated, no manual entry).
// Shows only unpaid/partial orders, sorted by Due Date ascending.
// ============================================================


// ============================================================
// setupHubDashboardSheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Dashboard" sheet (all clients, sorted by due balance desc).
// ============================================================


// ============================================================
// setupCreditSheet()  — v14.0
// ─────────────────────────────────────────────────────────────
// CHANGES vs v13:
//   • Removed cols F (Req New Limit), G (Req Reason),
//     H (Req Status), I (Approved By)
//   • Added col F = Client Email
//   • Col G = Last Updated   (was col J)
//   • UUID scratch col is now col H — hidden (was col K)
//   • CC_TOTAL_COLS updated to 7 (visible cols A–G)
//   • Sales Manager role references replaced by Sales
// ============================================================


// ============================================================
// setupCRMSheet()  — v14.0
// ─────────────────────────────────────────────────────────────
// CHANGES vs v13:
//   1. Sheet is still stored internally as 'Main' (MAIN_SHEET_NAME)
//      for sync compatibility, but the tab is RENAMED to
//      CRM_ORDER_LEDGER_SHEET ('Order Ledger CRM') via setName().
//   2. CRM_COL_CRM_OWNER column (col C) is now HIDDEN in the CRM spoke.
//      The owner is implicit — the spoke belongs to one CRM user.
//   3. New "Client Dues" sheet created in the spoke (CRM-filtered dues).
//   4. createClientTab() updated to add "Due Date" column (col P) to
//      all new client sub-tab headers.
// ============================================================


// ============================================================
// setupCRMDuesSheet_()
// NEW v14.0 — Internal helper called by setupCRMSheet()
// Creates the "Client Dues" sheet in the CRM spoke.
// This is a filtered/read-only view showing only this CRM user's
// orders that have outstanding balances.
// It is auto-populated by the Hub sync scripts — no manual entry.
// ============================================================


// ============================================================
// createClientTab()  — v14.0
// ─────────────────────────────────────────────────────────────
// CHANGE vs v13: Added "Due Date" header at col P (SUB_COL_DUE_DATE = 16).
// All column index references are unchanged (formulas use G:L which are
// Item Value through Discount — col indices unchanged).
// ============================================================


// ============================================================
// autoAddClientToCredit_()  — v14.0
// ─────────────────────────────────────────────────────────────
// Updated to use new CC_TOTAL_COLS (7) and CC_COL_UUID (8).
// Client Email (col F) defaults to blank — must be set manually
// or via a future onboarding flow.
// ============================================================


// ============================================================
// setCreditLimit_()  — v14.0
// Updated to use new CC_TOTAL_COLS (7) and CC_COL_UUID (8).
// Credit-request columns (old F–I) removed — no longer written.
// ============================================================


// ============================================================
// invalidateConfigCache_()
// ============================================================


// ============================================================
// ONBOARDING ROLE INSTRUCTIONS — v14.0
// Updated for merged Sales role (used in notifyAndShareMember_)
// ============================================================
const ROLE_INSTRUCTIONS_V14 = {
  'crm': `Your personal CRM sheet has been shared with you. You can:\n` +
         `  • Add new customer orders in the "Main" tab (Order Ledger CRM)\n` +
         `  • Track payment status, dispatch, and due balances\n` +
         `  • View client sub-tabs for order details (inc. Due Date)\n` +
         `  • Monitor your outstanding balances in the "Client Dues" tab\n\n` +
         `You also have read-only access to the Master Hub to monitor overall order flow.`,

  'logistics': `You have been added as Logistics on the Master Hub. You can:\n` +
               `  • Update AWB No (col Q) and Dispatch Date (col R) for any order\n` +
               `  • All other columns are view-only for your role\n\n` +
               `Please update dispatch details promptly once shipments go out.`,

  'accounts': `You have been added as Accounts on the Master Hub. You can:\n` +
              `  • Update Payment Received (col J) and Payment Date (col K)\n` +
              `  • All other columns are view-only for your role\n\n` +
              `Please keep payment records up to date for accurate balance tracking.`,

  'sales': `You have been added as Sales on the Master Hub. You can:\n` +
           `  • Add new CRM Members in the CRM Members tab\n` +
           `  • Enter Client Name and assign CRM in the Client Credit tab\n` +
           `  • Edit all columns in the Client Credit tab (inc. Client Email)\n` +
           `  • Approve credit-exceeded orders (col O in Master Hub)\n` +
           `  • Monitor all CRM activity and order statuses\n\n` +
           `Note: "Sales" role now covers all former Sales + Sales Manager capabilities.`,

  'admin': `You have been added as Admin on the Master Hub. You have full access to:\n` +
           `  • All sheets: Master Hub, Order Ledger, Client Summary, Client Dues,\n` +
           `    Dashboard, CRM Members, Client Credit, Activity Log\n` +
           `  • Add or deactivate team members in the CRM Members tab\n` +
           `  • Run all menu functions (Export, Reminders, Setup)\n\n` +
           `Please do not share this file with anyone not listed in CRM Members.`,
};
// Backward-compat alias for old 'sales manager' key
ROLE_INSTRUCTIONS_V14['sales manager'] = ROLE_INSTRUCTIONS_V14['sales'];

//


function getCrmConfig_() {
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'crm_config_v14';
  const cached   = cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  try {
    const hubSS     = SpreadsheetApp.openById(MASTER_HUB_ID);
    const membSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);

    if (!membSheet) {
      Logger.log('getCrmConfig_: "CRM Members" sheet not found. Run Setup -> Setup Members Sheet first.');
      return {
        crmFileIds: {}, salesEmails: [], salesManagerEmails: [],
        adminEmails: [], logisticsEmails: [], accountsEmails: [], allMembers: []
      };
    }

    const data   = membSheet.getDataRange().getValues();
    const config = {
      crmFileIds:         {},
      salesEmails:        [],   // merged Sales + Sales Manager
      adminEmails:        [],
      logisticsEmails:    [],
      accountsEmails:     [],
      allMembers:         []
    };

    for (let i = 1; i < data.length; i++) {
      const row    = data[i];
      const gmail  = (row[MC_COL_GMAIL   - 1] || '').toString().trim().toLowerCase();
      const fileId = (row[MC_COL_FILE_ID - 1] || '').toString().trim();
      const role   = (row[MC_COL_ROLE    - 1] || '').toString().trim();
      const name   = (row[MC_COL_NAME    - 1] || '').toString().trim();
      const active = row[MC_COL_ACTIVE   - 1];

      if (!gmail || !role) continue;
      if (active === false || active === 'FALSE' || active === 'false') continue;

      config.allMembers.push({ gmail, fileId, role, name });

      switch (role.toLowerCase()) {
        case 'sales':
        case 'sales manager':   // backward compat — treat legacy rows as Sales
          config.salesEmails.push(gmail);
          break;
        case 'crm':
          if (fileId) config.crmFileIds[gmail] = fileId;
          break;
        case 'admin':
          config.adminEmails.push(gmail);
          break;
        case 'logistics':
          config.logisticsEmails.push(gmail);
          break;
        case 'accounts':
          config.accountsEmails.push(gmail);
          break;
      }
    }

    // Backward-compat alias — callers using salesManagerEmails still work
    config.salesManagerEmails = config.salesEmails;

    try { cache.put(cacheKey, JSON.stringify(config), 360); } catch (_) {}
    return config;

  } catch (err) {
    Logger.log('ERROR in getCrmConfig_: ' + err.message);
    return {
      crmFileIds: {}, salesEmails: [], salesManagerEmails: [],
      adminEmails: [], logisticsEmails: [], accountsEmails: [], allMembers: []
    };
  }
}

function invalidateConfigCache_() {
  try {
    CacheService.getScriptCache().remove('crm_config_v14');
    CacheService.getScriptCache().remove('crm_config_v3'); // remove old key too
    Logger.log('invalidateConfigCache_: cache cleared.');
  } catch (_) {}
}

function setupMembersSheet() {
  try {
    const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
    let membSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
    if (!membSheet) membSheet = hubSS.insertSheet(MEMBERS_SHEET_NAME);

    const headers = ['Gmail', 'File ID (CRM Spoke)', 'Role', 'Display Name', 'Active'];
    const hdrRange = membSheet.getRange(1, 1, 1, MC_TOTAL_COLS);
    hdrRange.setValues([headers]);
    hdrRange.setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF').setFontSize(11);
    membSheet.setFrozenRows(1);
    membSheet.setColumnWidth(MC_COL_GMAIL,   260);
    membSheet.setColumnWidth(MC_COL_FILE_ID, 300);
    membSheet.setColumnWidth(MC_COL_ROLE,    140);
    membSheet.setColumnWidth(MC_COL_NAME,    160);
    membSheet.setColumnWidth(MC_COL_ACTIVE,   70);

    // v14.0: 'Sales Manager' removed from dropdown — merged into 'Sales'
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(VALID_ROLES, true)
      .setAllowInvalid(false)
      .build();
    membSheet.getRange(2, MC_COL_ROLE, 999, 1).setDataValidation(roleRule);

    membSheet.getRange(2, MC_COL_ACTIVE, 999, 1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());

    membSheet.getRange(1, MC_COL_FILE_ID)
      .setNote('Get this from the CRM member\'s spreadsheet URL:\ndocs.google.com/spreadsheets/d/[FILE_ID_IS_HERE]/edit\nLeave blank for non-CRM roles.');

    membSheet.getRange(1, MC_COL_ROLE)
      .setNote(
        'CRM = has their own spoke file\n' +
        'Logistics = edits AWB/dispatch cols in Master Hub\n' +
        'Accounts = edits Payment cols in Master Hub\n' +
        'Sales = approves credit-exceeded orders + can add new clients/members + edit all Client Credit cols\n' +
        '        (formerly "Sales" and "Sales Manager" — now merged)\n' +
        'Admin = full access'
      );

    const ownerEmail = Session.getEffectiveUser().getEmail();
    membSheet.getRange(2, 1, 1, MC_TOTAL_COLS).setValues([[
      ownerEmail, '', ROLE_ADMIN, 'Admin (You)', true
    ]]);

    const config2     = getCrmConfig_();
    const membEditors = [...new Set([
      ownerEmail,
      ...config2.adminEmails,
      ...config2.salesEmails
    ])];

    membSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    membSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    const membProt = membSheet.getRange(2, 1, 999, MC_TOTAL_COLS)
      .protect().setDescription('Admin/Sales only — CRM Members data');
    membProt.removeEditors(membProt.getEditors());
    membProt.addEditors(membEditors);
    membProt.setWarningOnly(false);

    membSheet.hideSheet();

    const cfGray = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=E2=FALSE')
      .setFontColor('#9E9E9E')
      .setBackground('#F5F5F5')
      .setRanges([membSheet.getRange(2, 1, 999, MC_TOTAL_COLS)])
      .build();
    membSheet.setConditionalFormatRules([cfGray]);

    invalidateConfigCache_();
    logActivity_(hubSS, 'MEMBERS SHEET SETUP', 'setupMembersSheet()', ownerEmail);

    alertIfUi_(
      '✅ CRM Members sheet created!\n\n' +
      'Roles available:\n' +
      '• CRM — own spoke file, manages client orders\n' +
      '• Logistics — edits AWB/dispatch columns\n' +
      '• Accounts — edits Payment columns\n' +
      '• Sales — add members, add clients, edit credit, approve orders\n' +
      '          (replaces former "Sales" + "Sales Manager" roles)\n' +
      '• Admin — full access\n\n' +
      'To deactivate a member: uncheck the Active box.\n' +
      'No code changes ever needed.'
    );
  } catch (err) {
    Logger.log('ERROR in setupMembersSheet: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function setupHubSheet() {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail      = Session.getEffectiveUser().getEmail();
    const config          = getCrmConfig_();
    const crmSpokeEmails  = Object.keys(config.crmFileIds);
    const adminEditors     = [...new Set([ownerEmail, ...config.adminEmails,     ...crmSpokeEmails])];
    // v15.1: Logistics- and Accounts-exclusive columns no longer include CRM spoke
    // emails — CRM has no Hub access at all, and these fields now flow ONE-WAY
    // Hub → Spoke (via onMasterHubEdit), so there's no longer a legitimate script
    // push from the CRM side into these specific ranges.
    const logisticsEditors = [...new Set([ownerEmail, ...config.logisticsEmails])];
    const accountsEditors  = [...new Set([ownerEmail, ...config.accountsEmails])];
    // Shared columns (Dispatch Status, Freight, Dev Charges) are still legitimately
    // pushed UP from the CRM sub-tab (Item Value/Dev Charges/Freight entry), so CRM
    // spoke emails stay here.
    const logisticsCrmSharedEditors = [...new Set([ownerEmail, ...config.logisticsEmails, ...crmSpokeEmails])];
    const approvalEditors  = [...new Set([ownerEmail, ...config.adminEmails,
      ...config.salesEmails, ...crmSpokeEmails])];

    // Master Hub removed v15.0 — delete if still exists
    const _legacyHub2 = ss.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (_legacyHub2) ss.deleteSheet(_legacyHub2);

    // ── 2. Order Ledger sheet ─────────────────────────────────
    setupHubOrderLedgerSheet_(ss, adminEditors, logisticsEditors, accountsEditors, logisticsCrmSharedEditors);

    // ── 3. Client Summary sheet ───────────────────────────────
    setupHubClientSummarySheet_(ss, adminEditors);

    // ── 4. Client Dues sheet ──────────────────────────────────
    setupHubDuesSheet_(ss, adminEditors);

    // ── 5. Dashboard sheet ────────────────────────────────────
    setupHubDashboardSheet_(ss, adminEditors);

    // ── 6. Client Credit sheet ────────────────────────────────
    setupCreditSheet();

    logActivity_(ss, 'HUB SETUP', 'setupHubSheet() v14.0', Session.getActiveUser().getEmail());
    alertIfUi_(
      '✅ Master Hub setup complete (v14.0).\n\n' +
      'Sheets created/updated:\n' +
      '• Master Hub\n• Order Ledger\n• Client Summary\n• Client Dues\n• Dashboard\n• Client Credit'
    );
  } catch (err) {
    Logger.log('ERROR in setupHubSheet: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function setupHubOrderLedgerSheet_(ss, adminEditors, logisticsEditors, accountsEditors, logisticsCrmSharedEditors) {
  let sheet = ss.getSheetByName(HUB_ORDER_LEDGER_SHEET);
  if (!sheet) sheet = ss.insertSheet(HUB_ORDER_LEDGER_SHEET);

  const headers = [
    'UUID',             // A  OL_COL_UUID
    'Client Name',      // B  OL_COL_CLIENT_NAME
    'CRM Owner',        // C  OL_COL_CRM_OWNER
    'PI No',            // D  OL_COL_PI_NO
    'Inv No',           // E  OL_COL_INV_NO
    'Inv Date',         // F  OL_COL_INV_DATE
    'Order Type',       // G  OL_COL_ORDER_TYPE
    'Item Value',       // H  OL_COL_ITEM_VALUE
    'Dev Charges',      // I  OL_COL_DEV_CHARGES
    'Freight',          // J  OL_COL_FREIGHT
    'Discount',         // K  OL_COL_DISCOUNT
    'Final Invoice',    // L  OL_COL_FINAL_INVOICE
    'AWB No',           // M  OL_COL_AWB_NO
    'Dispatch Date',    // N  OL_COL_DISPATCH_DATE
    'Dispatch Status',  // O  OL_COL_DISPATCH_STATUS
    'Payment Received', // P  OL_COL_PAY_RECEIVED
    'Payment Date',     // Q  OL_COL_PAY_DATE
    'Payment Status',   // R  OL_COL_PAY_STATUS
    'Due Date',         // S  OL_COL_DUE_DATE
    'Due Balance',      // T  OL_COL_DUE_BALANCE
    'Approval Required',// U  OL_COL_APPROVAL_REQ   — v15.2 NEW, script-computed
    'Approved By',      // V  OL_COL_APPROVED_BY    — v15.2 NEW, Admin/Sales dropdown
    'Remarks',          // W  OL_COL_REMARKS        — v15.2 NEW, free text
    'Due Date Set On'   // X  OL_COL_DUE_DATE_SET_ON — v15.2 NEW, hidden, 7-day lock tracking
  ];

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF')
          .setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);

  // Hide UUID column (internal)
  sheet.hideColumns(OL_COL_UUID);

  const colWidths = {
    [OL_COL_CLIENT_NAME]:     180, [OL_COL_CRM_OWNER]:       180,
    [OL_COL_PI_NO]:           100, [OL_COL_INV_NO]:           80,
    [OL_COL_INV_DATE]:        100, [OL_COL_ORDER_TYPE]:       110,
    [OL_COL_ITEM_VALUE]:      110, [OL_COL_DEV_CHARGES]:      110,
    [OL_COL_FREIGHT]:          90, [OL_COL_DISCOUNT]:          90,
    [OL_COL_FINAL_INVOICE]:   120, [OL_COL_AWB_NO]:           110,
    [OL_COL_DISPATCH_DATE]:   110, [OL_COL_DISPATCH_STATUS]:  130,
    [OL_COL_PAY_RECEIVED]:    130, [OL_COL_PAY_DATE]:         110,
    [OL_COL_PAY_STATUS]:      120, [OL_COL_DUE_DATE]:         110,
    [OL_COL_DUE_BALANCE]:     120
  };
  Object.entries(colWidths).forEach(([c, w]) => sheet.setColumnWidth(Number(c), w));

  // Dropdowns
  sheet.getRange(2, OL_COL_ORDER_TYPE, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Production', 'Sample', 'Showroom'], true)
      .setAllowInvalid(false).build()
  );
  sheet.getRange(2, OL_COL_DISPATCH_STATUS, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Dispatched', 'Delivered'], true)
      .setAllowInvalid(false).build()
  );
  // v15.1: Payment Status is no longer a manual dropdown — it's now an
  // auto-calculated ARRAYFORMULA (see below), driven off Final Invoice (L)
  // and Payment Received (P). Nobody, including Accounts, sets this by hand.

  // Date number formats
  sheet.getRange(2, OL_COL_INV_DATE,      999, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, OL_COL_DISPATCH_DATE, 999, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, OL_COL_PAY_DATE,      999, 1).setNumberFormat('dd/mm/yyyy');
  sheet.getRange(2, OL_COL_DUE_DATE,      999, 1).setNumberFormat('dd/mm/yyyy');

  // Currency number formats
  const currencyCols = [OL_COL_ITEM_VALUE, OL_COL_DEV_CHARGES, OL_COL_FREIGHT,
                        OL_COL_DISCOUNT, OL_COL_FINAL_INVOICE, OL_COL_PAY_RECEIVED,
                        OL_COL_DUE_BALANCE];
  currencyCols.forEach(c => sheet.getRange(2, c, 999, 1).setNumberFormat('#,##0.00'));

  // v15.1 FIX (correction): the live ARRAYFORMULA approach below was removed.
  // It spilled "" results down to row 1000+ for every blank row, which made
  // getLastRow() return 1000+ even on a brand-new empty sheet — causing
  // upsertOrderLedgerRow_'s "new row = getLastRow()+1" logic to place new
  // orders at row 1001 instead of row 2. Payment Status and Due Balance are
  // now script-computed instead (same pattern as every other column on this
  // sheet) — written by upsertOrderLedgerRow_ for new/updated orders, and by
  // onMasterHubEdit when Accounts edits Payment Received/Date directly.
  // Clear any formula left over from a previous setup run:
  sheet.getRange(2, OL_COL_PAY_STATUS,   999, 1).clearContent();
  sheet.getRange(2, OL_COL_DUE_BALANCE,  999, 1).clearContent();

  // Conditional formatting
  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    // Overdue & unpaid — red. v15.1 FIX: added $B2<>"" guard — previously blank
    // rows matched too (blank Due Date < TODAY() and blank <> "Received" are both
    // TRUE), which painted every empty row in the sheet red.
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($B2<>"",S2<>"",S2<TODAY(),R2<>"Received")`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),
    // Partial payment — orange
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($B2<>"",R2="Partial")`)
      .setBackground('#FFE0B2').setRanges([cfRange]).build(),
    // Fully received — green
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($B2<>"",R2="Received")`)
      .setBackground('#E8F5E9').setRanges([cfRange]).build(),
  ]);

  // Protections: Order Ledger is admin-managed (sync-populated)
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  // v15.1: full column-by-column protection matrix.
  // Logistics-exclusive: Inv No, Inv Date (E–F)
  const invProt = sheet.getRange(2, OL_COL_INV_NO, 1000, 2)
    .protect().setDescription('Logistics — Inv No & Inv Date (E–F)');
  invProt.removeEditors(invProt.getEditors());
  invProt.addEditors([...new Set([...adminEditors, ...logisticsEditors])]);
  invProt.setWarningOnly(false);

  // Logistics-exclusive: AWB No, Dispatch Date (M–N)
  const dispProt = sheet.getRange(2, OL_COL_AWB_NO, 1000, 2)
    .protect().setDescription('Logistics — AWB No & Dispatch Date (M–N)');
  dispProt.removeEditors(dispProt.getEditors());
  dispProt.addEditors([...new Set([...adminEditors, ...logisticsEditors])]);
  dispProt.setWarningOnly(false);

  // Logistics + CRM shared: Dispatch Status (O)
  const dsProt = sheet.getRange(2, OL_COL_DISPATCH_STATUS, 1000, 1)
    .protect().setDescription('Logistics + CRM — Dispatch Status (O)');
  dsProt.removeEditors(dsProt.getEditors());
  dsProt.addEditors([...new Set([...adminEditors, ...logisticsCrmSharedEditors])]);
  dsProt.setWarningOnly(false);

  // Logistics + CRM shared: Dev Charges, Freight (I–J)
  const devFrtProt = sheet.getRange(2, OL_COL_DEV_CHARGES, 1000, 2)
    .protect().setDescription('Logistics + CRM — Dev Charges & Freight (I–J)');
  devFrtProt.removeEditors(devFrtProt.getEditors());
  devFrtProt.addEditors([...new Set([...adminEditors, ...logisticsCrmSharedEditors])]);
  devFrtProt.setWarningOnly(false);

  // Accounts-exclusive: Payment Received, Payment Date (P–Q)
  const payProt = sheet.getRange(2, OL_COL_PAY_RECEIVED, 1000, 2)
    .protect().setDescription('Accounts — Payment Received & Date (P–Q)');
  payProt.removeEditors(payProt.getEditors());
  payProt.addEditors([...new Set([...adminEditors, ...accountsEditors])]);
  payProt.setWarningOnly(false);

  // Accounts-exclusive: Due Date (S)
  const dueDateProt = sheet.getRange(2, OL_COL_DUE_DATE, 1000, 1)
    .protect().setDescription('Accounts — Due Date (S)');
  dueDateProt.removeEditors(dueDateProt.getEditors());
  dueDateProt.addEditors([...new Set([...adminEditors, ...accountsEditors])]);
  dueDateProt.setWarningOnly(false);

  // Nobody manually — Payment Status & Due Balance are script-computed (R, T).
  // v15.1 FIX: editors include accountsEditors too — the script writes the
  // computed value here under ACCOUNTS' authorization when they edit Payment
  // Received/Date (see onMasterHubEdit), and under CRM's authorization when a
  // new order is upserted (adminEditors already includes crmSpokeEmails).
  const autoCalcProt = sheet.getRange(2, OL_COL_PAY_STATUS, 1000, 1)
    .protect().setDescription('Auto-calculated — Payment Status (R), script only');
  autoCalcProt.removeEditors(autoCalcProt.getEditors());
  autoCalcProt.addEditors([...new Set([...adminEditors, ...accountsEditors])]);
  autoCalcProt.setWarningOnly(false);

  const dueBalProt = sheet.getRange(2, OL_COL_DUE_BALANCE, 1000, 1)
    .protect().setDescription('Auto-calculated — Due Balance (T), script only');
  dueBalProt.removeEditors(dueBalProt.getEditors());
  dueBalProt.addEditors([...new Set([...adminEditors, ...accountsEditors])]);
  dueBalProt.setWarningOnly(false);

  // Remaining cols: Admin only (Client Name, CRM Owner, PI No, Order Type,
  // Item Value, Discount, Final Invoice)
  const adminColRanges = [
    sheet.getRange(2, OL_COL_CLIENT_NAME, 1000, OL_COL_PI_NO - OL_COL_CLIENT_NAME + 1), // B–D
    sheet.getRange(2, OL_COL_ORDER_TYPE,  1000, 2),                                      // G–H (Order Type, Item Value)
    sheet.getRange(2, OL_COL_DISCOUNT,    1000, 2),                                      // K–L (Discount, Final Invoice)
  ];
  adminColRanges.forEach((r, idx) => {
    const p = r.protect().setDescription('Admin only — core order data (' + idx + ')');
    p.removeEditors(p.getEditors());
    p.addEditors(adminEditors);
    p.setWarningOnly(false);
  });

  // v15.2 NEW: Approval Required (U), Approved By (V), Remarks (W)
  const config       = getCrmConfig_();
  const ownerEmailHS  = Session.getEffectiveUser().getEmail();
  const adminSalesEditors = [...new Set([ownerEmailHS, ...config.adminEmails, ...config.salesEmails])];
  const remarksEditors    = [...new Set([
    ownerEmailHS, ...config.adminEmails, ...config.salesEmails,
    ...config.logisticsEmails, ...config.accountsEmails
  ])];

  // Approval Required — nobody manually, script-computed by reEvaluateAllCreditFlags_
  const apprReqProt = sheet.getRange(2, OL_COL_APPROVAL_REQ, 1000, 1)
    .protect().setDescription('Auto-calculated — Approval Required (U), script only');
  apprReqProt.removeEditors(apprReqProt.getEditors());
  apprReqProt.addEditors(adminEditors);
  apprReqProt.setWarningOnly(false);

  // Approved By — Admin + Sales only, dropdown of their names/emails
  const apprByProt = sheet.getRange(2, OL_COL_APPROVED_BY, 1000, 1)
    .protect().setDescription('Admin/Sales — Approved By (V)');
  apprByProt.removeEditors(apprByProt.getEditors());
  apprByProt.addEditors(adminSalesEditors);
  apprByProt.setWarningOnly(false);
  const approverList = [...new Set([...config.adminEmails, ...config.salesEmails])];
  sheet.getRange(2, OL_COL_APPROVED_BY, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['', ...approverList], true)
      .setAllowInvalid(true)
      .build()
  );

  // Remarks — Admin/Sales/Logistics/Accounts, free text
  const remarksProt = sheet.getRange(2, OL_COL_REMARKS, 1000, 1)
    .protect().setDescription('Admin/Sales/Logistics/Accounts — Remarks (W)');
  remarksProt.removeEditors(remarksProt.getEditors());
  remarksProt.addEditors(remarksEditors);
  remarksProt.setWarningOnly(false);

  // Due Date Set On (X) — hidden, script-only, tracks the 7-day Due Date lock
  const dueSetOnProt = sheet.getRange(2, OL_COL_DUE_DATE_SET_ON, 1000, 1)
    .protect().setDescription('Auto-calculated — Due Date Set On (X), script only, hidden');
  dueSetOnProt.removeEditors(dueSetOnProt.getEditors());
  dueSetOnProt.addEditors(adminEditors);
  dueSetOnProt.setWarningOnly(false);
  sheet.hideColumns(OL_COL_DUE_DATE_SET_ON);

  // Note on usage
  sheet.getRange(1, OL_COL_UUID)
    .setNote('Order Ledger — one row per order. Auto-populated by sync scripts. Do not add rows manually.');

  Logger.log('setupHubOrderLedgerSheet_: "' + HUB_ORDER_LEDGER_SHEET + '" ready (v15.2 protection matrix).');
}

function setupHubClientSummarySheet_(ss, adminEditors) {
  let sheet = ss.getSheetByName(HUB_CLIENT_SUMMARY_SHEET);
  if (!sheet) sheet = ss.insertSheet(HUB_CLIENT_SUMMARY_SHEET);

  const headers = [
    'Client Name',       // A  CS_COL_CLIENT_NAME
    'CRM Owner',         // B  CS_COL_CRM_OWNER
    'Total Orders',      // C  CS_COL_TOTAL_ORDERS
    'Total Invoiced',    // D  CS_COL_TOTAL_INVOICED
    'Total Paid',        // E  CS_COL_TOTAL_PAID
    'Total Due',         // F  CS_COL_TOTAL_DUE
    'Credit Limit',      // G  CS_COL_CREDIT_LIMIT
    'Used Credit',       // H  CS_COL_USED_CREDIT
    'Available Credit',  // I  CS_COL_AVAIL_CREDIT
    'Credit Status',     // J  CS_COL_CREDIT_STATUS
    'Last Order Date'    // K  CS_COL_LAST_ORDER_DATE
  ];

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF')
          .setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);

  const colWidths = {
    [CS_COL_CLIENT_NAME]:    200, [CS_COL_CRM_OWNER]:     180,
    [CS_COL_TOTAL_ORDERS]:   110, [CS_COL_TOTAL_INVOICED]: 130,
    [CS_COL_TOTAL_PAID]:     120, [CS_COL_TOTAL_DUE]:      120,
    [CS_COL_CREDIT_LIMIT]:   120, [CS_COL_USED_CREDIT]:    120,
    [CS_COL_AVAIL_CREDIT]:   130, [CS_COL_CREDIT_STATUS]:  130,
    [CS_COL_LAST_ORDER_DATE]:130
  };
  Object.entries(colWidths).forEach(([c, w]) => sheet.setColumnWidth(Number(c), w));

  sheet.getRange(2, CS_COL_TOTAL_INVOICED, 999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_TOTAL_PAID,     999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_TOTAL_DUE,      999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_CREDIT_LIMIT,   999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_USED_CREDIT,    999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_AVAIL_CREDIT,   999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, CS_COL_LAST_ORDER_DATE,999, 1).setNumberFormat('dd/mm/yyyy');

  // Conditional formatting: credit status colours
  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=I2<=0`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),   // over limit
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(I2>0,I2/G2<0.2)`)
      .setBackground('#FFF9C4').setRanges([cfRange]).build(),   // < 20 % left
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=I2/G2>=0.2`)
      .setBackground('#E8F5E9').setRanges([cfRange]).build(),   // healthy
  ]);

  // Admin-only protection — this sheet is script-populated
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  const prot = sheet.getRange(2, 1, 1000, headers.length)
    .protect().setDescription('Admin only — Client Summary is script-populated');
  prot.removeEditors(prot.getEditors());
  prot.addEditors(adminEditors);
  prot.setWarningOnly(false);

  sheet.getRange(1, CS_COL_CLIENT_NAME)
    .setNote('Client Summary — one row per client. Auto-refreshed by sync scripts. Do not edit manually.');

  Logger.log('setupHubClientSummarySheet_: "' + HUB_CLIENT_SUMMARY_SHEET + '" ready.');
}

function setupHubDuesSheet_(ss, adminEditors) {
  let sheet = ss.getSheetByName(HUB_DUES_SHEET);
  if (!sheet) sheet = ss.insertSheet(HUB_DUES_SHEET);

  const headers = [
    'Client Name',       // A  DUE_COL_CLIENT_NAME
    'CRM Owner',         // B  DUE_COL_CRM_OWNER
    'PI No',             // C  DUE_COL_PI_NO
    'Inv No',            // D  DUE_COL_INV_NO
    'Final Invoice',     // E  DUE_COL_FINAL_INVOICE
    'Payment Received',  // F  DUE_COL_PAY_RECEIVED
    'Due Balance',       // G  DUE_COL_DUE_BALANCE
    'Due Date',          // H  DUE_COL_DUE_DATE
    'Payment Status'     // I  DUE_COL_PAY_STATUS
  ];

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight('bold').setBackground('#B71C1C').setFontColor('#FFFFFF')
          .setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);

  const colWidths = {
    [DUE_COL_CLIENT_NAME]:   190, [DUE_COL_CRM_OWNER]:    180,
    [DUE_COL_PI_NO]:         100, [DUE_COL_INV_NO]:         80,
    [DUE_COL_FINAL_INVOICE]: 130, [DUE_COL_PAY_RECEIVED]:  130,
    [DUE_COL_DUE_BALANCE]:   130, [DUE_COL_DUE_DATE]:      110,
    [DUE_COL_PAY_STATUS]:    120
  };
  Object.entries(colWidths).forEach(([c, w]) => sheet.setColumnWidth(Number(c), w));

  sheet.getRange(2, DUE_COL_FINAL_INVOICE, 999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DUE_COL_PAY_RECEIVED,  999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DUE_COL_DUE_BALANCE,   999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DUE_COL_DUE_DATE,      999, 1).setNumberFormat('dd/mm/yyyy');

  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(H2<TODAY(),H2<>"")`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),   // overdue
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=I2="Partial"`)
      .setBackground('#FFF9C4').setRanges([cfRange]).build(),   // partial payment
  ]);

  // Fully protected — no manual entry permitted
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  const prot = sheet.getRange(2, 1, 1000, headers.length)
    .protect().setDescription('Admin only — Client Dues is auto-populated. Do not edit manually.');
  prot.removeEditors(prot.getEditors());
  prot.addEditors(adminEditors);
  prot.setWarningOnly(false);

  sheet.getRange(1, DUE_COL_CLIENT_NAME)
    .setNote('Client Dues — auto-populated from Order Ledger (unpaid/partial orders only, sorted by Due Date). Do not edit manually.');

  Logger.log('setupHubDuesSheet_: "' + HUB_DUES_SHEET + '" ready.');
}

function setupHubDashboardSheet_(ss, adminEditors) {
  let sheet = ss.getSheetByName(HUB_DASHBOARD_SHEET);
  if (!sheet) sheet = ss.insertSheet(HUB_DASHBOARD_SHEET);

  const headers = [
    'Client Name',       // A  DB_COL_CLIENT_NAME
    'CRM Owner',         // B  DB_COL_CRM_OWNER
    'Total Due Balance', // C  DB_COL_TOTAL_DUE
    'Total Invoiced',    // D  DB_COL_TOTAL_INVOICED
    'Total Paid',        // E  DB_COL_TOTAL_PAID
    'Unpaid Orders',     // F  DB_COL_OVERDUE_COUNT  (counts Pending + Partial orders)
    'Credit Status',     // G  DB_COL_CREDIT_STATUS
    'Last Order Date'    // H  DB_COL_LAST_ORDER_DATE
  ];

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF')
          .setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);

  const colWidths = {
    [DB_COL_CLIENT_NAME]:    200, [DB_COL_CRM_OWNER]:      180,
    [DB_COL_TOTAL_DUE]:      140, [DB_COL_TOTAL_INVOICED]: 140,
    [DB_COL_TOTAL_PAID]:     130, [DB_COL_OVERDUE_COUNT]:  130,
    [DB_COL_CREDIT_STATUS]:  130, [DB_COL_LAST_ORDER_DATE]:140
  };
  Object.entries(colWidths).forEach(([c, w]) => sheet.setColumnWidth(Number(c), w));

  sheet.getRange(2, DB_COL_TOTAL_DUE,      999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DB_COL_TOTAL_INVOICED, 999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DB_COL_TOTAL_PAID,     999, 1).setNumberFormat('#,##0.00');
  sheet.getRange(2, DB_COL_LAST_ORDER_DATE,999, 1).setNumberFormat('dd/mm/yyyy');

  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=C2>0`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),   // has outstanding dues
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=G2="Over Limit"`)
      .setBackground('#FF8A65').setRanges([cfRange]).build(),   // over credit limit
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=C2=0`)
      .setBackground('#E8F5E9').setRanges([cfRange]).build(),   // fully clear
  ]);

  // Dashboard is script-populated — admin only
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  const prot = sheet.getRange(2, 1, 1000, headers.length)
    .protect().setDescription('Admin only — Dashboard is auto-populated, sorted by Due Balance');
  prot.removeEditors(prot.getEditors());
  prot.addEditors(adminEditors);
  prot.setWarningOnly(false);

  sheet.getRange(1, DB_COL_CLIENT_NAME)
    .setNote('Dashboard — all clients sorted by Total Due Balance (highest first). Auto-refreshed by sync scripts.');

  Logger.log('setupHubDashboardSheet_: "' + HUB_DASHBOARD_SHEET + '" ready.');
}

function setupCreditSheet() {
  try {
    const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
    let creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) creditSheet = hubSS.insertSheet(CREDIT_SHEET_NAME);

    // Visible headers A–G (col H = UUID, hidden)
    const headers = [
      'Client Name',          // A  CC_COL_CLIENT_NAME  = 1
      'CRM Owner (Gmail)',     // B  CC_COL_CRM_OWNER    = 2
      'Credit Limit ($)',      // C  CC_COL_CREDIT_LIMIT = 3
      'Used Credit ($)',       // D  CC_COL_USED_CREDIT  = 4  — script-maintained
      'Available Credit ($)',  // E  CC_COL_AVAIL_CREDIT = 5  — script-maintained
      'Client Email',          // F  CC_COL_CLIENT_EMAIL = 6  NEW v14.0
      'Last Updated',          // G  CC_COL_LAST_UPDATED = 7  (was col J)
      'UUID (Internal)'        // H  CC_COL_UUID         = 8  hidden scratch col
    ];

    const hdrRange = creditSheet.getRange(1, 1, 1, headers.length);
    hdrRange.setValues([headers]);
    hdrRange.setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF').setFontSize(11);
    creditSheet.setFrozenRows(1);

    const widths = [200, 220, 140, 140, 160, 200, 160, 80];
    widths.forEach((w, i) => creditSheet.setColumnWidth(i + 1, w));

    // Hide UUID col (H = col 8)
    creditSheet.hideColumns(CC_COL_UUID);

    // Clear D2 and E2 to remove any old ARRAYFORMULAs (see v8.0 bug fix)
    creditSheet.getRange(2, CC_COL_USED_CREDIT).clearContent();
    creditSheet.getRange(2, CC_COL_AVAIL_CREDIT).clearContent();
    SpreadsheetApp.flush();

    // Data validation: CRM Owner col B — dropdown of CRM-role members only
    // FIX v14.1: use MC_COL_* constants (1-based) with [col-1] for array access
    // MC_COL_GMAIL=1, MC_COL_FILE_ID=2, MC_COL_ROLE=3, MC_COL_NAME=4, MC_COL_ACTIVE=5
    try {
      const membersSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
      if (membersSheet) {
        const mData = membersSheet.getDataRange().getValues();
        const crmOwnerOptions = [];
        for (let m = 1; m < mData.length; m++) {
          const mEmail  = (mData[m][MC_COL_GMAIL  - 1] || '').toString().trim();
          const mRole   = (mData[m][MC_COL_ROLE   - 1] || '').toString().trim();
          const mName   = (mData[m][MC_COL_NAME   - 1] || '').toString().trim();
          const mActive = (mData[m][MC_COL_ACTIVE - 1] || '').toString().toUpperCase().trim();
          if (mEmail && mActive === 'TRUE' && mRole.toUpperCase() === 'CRM') {
            // Format: "Name (email)" — matches the label used everywhere else in the system
            const label = mName ? mName + ' (' + mEmail + ')' : mEmail;
            crmOwnerOptions.push(label);
          }
        }
        if (crmOwnerOptions.length > 0) {
          creditSheet.getRange(2, CC_COL_CRM_OWNER, 999, 1)
            .setDataValidation(
              SpreadsheetApp.newDataValidation()
                .requireValueInList(crmOwnerOptions, true)
                .setAllowInvalid(true)
                .setHelpText('Select the CRM owner for this client (CRM roles only)')
                .build()
            );
        }
      }
    } catch (dvErr) {
      Logger.log('setupCreditSheet: CRM Owner dropdown error: ' + dvErr.message);
    }

    // Data validation: Client Email col F
    creditSheet.getRange(2, CC_COL_CLIENT_EMAIL, 999, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireTextIsEmail()
          .setAllowInvalid(true)   // allow blank
          .setHelpText('Enter the client\'s email address (optional)')
          .build()
      );

    // ── Protections ──────────────────────────────────────────
    const ownerEmail  = Session.getEffectiveUser().getEmail();
    const config      = getCrmConfig_();
    // v15.1 FIX: adminEmails was missing crmSpokeEmails. refreshCreditTotals_()
    // writes to cols D–E under the CRM exec's OWN authorization (it's called
    // synchronously from their sub-tab edit, cross-spreadsheet into the Hub).
    // Without their email on this protection's editor list, that write throws
    // "You are trying to edit a protected cell or object" every single time —
    // which is exactly the error this codebase's own fixHubWriteAccess() exists
    // to patch, but setupCreditSheet() rebuilds protections from scratch on
    // every run (line below: .forEach(p => p.remove())), silently undoing
    // that patch. Fixed at the source so it doesn't need re-patching.
    const crmSpokeEmailsForCredit = Object.keys(config.crmFileIds);
    const adminEmails = [...new Set([ownerEmail, ...config.adminEmails, ...crmSpokeEmailsForCredit])];

    // v14.0: fullEditors = Sales (merged role) + Admin
    const fullEditors = [...new Set([
      ownerEmail,
      ...config.adminEmails,
      ...config.salesEmails
    ])];

    creditSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // Cols A–C: Client Name, CRM Owner, Credit Limit — Sales + Admin
    const abcProt = creditSheet.getRange(2, CC_COL_CLIENT_NAME, 999, 3)
      .protect().setDescription('Sales/Admin — Client Name, CRM Owner, Credit Limit (A–C)');
    abcProt.removeEditors(abcProt.getEditors());
    abcProt.addEditors(fullEditors);
    abcProt.setWarningOnly(false);

    // Cols D–E: Used Credit, Available Credit — Admin only (script-maintained)
    const deProt = creditSheet.getRange(2, CC_COL_USED_CREDIT, 999, 2)
      .protect().setDescription('Admin only — Used Credit & Available Credit (D–E) — script-maintained');
    deProt.removeEditors(deProt.getEditors());
    deProt.addEditors(adminEmails);
    deProt.setWarningOnly(false);

    // Col F: Client Email — Sales + Admin
    const fProt = creditSheet.getRange(2, CC_COL_CLIENT_EMAIL, 999, 1)
      .protect().setDescription('Sales/Admin — Client Email (F)');
    fProt.removeEditors(fProt.getEditors());
    fProt.addEditors(fullEditors);
    fProt.setWarningOnly(false);

    // Col G: Last Updated — script-maintained timestamp, same as D-E.
    // v15.1 FIX: this was using fullEditors (Sales+Admin, no CRM) instead of
    // adminEmails (now CRM-inclusive). refreshCreditTotals_ writes this column
    // as the LAST step of its per-row loop -- when it threw here under the
    // CRM exec's authorization, the function aborted before it ever reached
    // updateClientSummary_(), which is why Client Summary and Dashboard stayed
    // empty even after the D-E fix started working.
    const gProt = creditSheet.getRange(2, CC_COL_LAST_UPDATED, 999, 1)
      .protect().setDescription('Auto-calculated -- Last Updated (G), script only');
    gProt.removeEditors(gProt.getEditors());
    gProt.addEditors(adminEmails);
    gProt.setWarningOnly(false);

    // Conditional formatting
    const cfRange = creditSheet.getRange(2, 1, 999, CC_TOTAL_COLS);
    creditSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(A2<>"",E2<=0)`)
        .setBackground('#FFCDD2').setRanges([cfRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(A2<>"",C2>0,E2/C2<0.2,E2>0)`)
        .setBackground('#FFF9C4').setRanges([cfRange]).build(),
    ]);

    // Populate Used Credit + Available Credit from live Hub data
    refreshCreditTotals_(Session.getActiveUser().getEmail());

    logActivity_(hubSS, 'CREDIT SHEET SETUP v14.0', 'setupCreditSheet()', Session.getActiveUser().getEmail());
    alertIfUi_(
      '✅ Client Credit sheet ready (v14.0).\n\n' +
      'Column changes:\n' +
      '• Removed: Request New Limit, Request Reason, Request Status, Approved By\n' +
      '• Added: Client Email (col F)\n' +
      '• Used Credit and Available Credit refreshed from live Hub data.'
    );
  } catch (err) {
    Logger.log('ERROR in setupCreditSheet: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function setupCRMSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Part 1: Order Ledger CRM (formerly 'Main') ────────────
    // Check for existing sheet under either old or new name
    let mainSheet = ss.getSheetByName(CRM_ORDER_LEDGER_SHEET);
    if (!mainSheet) mainSheet = getCRMMainSheet_(ss);
    if (!mainSheet) mainSheet = ss.insertSheet(CRM_ORDER_LEDGER_SHEET);

    // Rename to 'Order Ledger CRM' if it still has the old name
    if (mainSheet.getName() !== CRM_ORDER_LEDGER_SHEET) {
      mainSheet.setName(CRM_ORDER_LEDGER_SHEET);
      Logger.log('setupCRMSheet: renamed "Main" → "Order Ledger CRM"');
    }

    // Header row — v14.0: CRM Owner (col C) header updated to note it is hidden
    const headers = [
      'UUID (hidden)',             // A  CRM_COL_UUID
      'Customer Name',             // B  CRM_COL_CUSTOMER
      'CRM Owner (hidden)',        // C  CRM_COL_CRM_OWNER — hidden in spoke; owner is implicit
      'PI No',                     // D  CRM_COL_PI_NO
      'Inv No',                    // E  CRM_COL_INV_NO
      'Inv Date',                  // F  CRM_COL_DATE
      'Inv Value',                 // G  CRM_COL_INV_VALUE
      'Payment Status',            // H  CRM_COL_PAY_STATUS
      'Due Date',                  // I  CRM_COL_DUE_DATE
      'Payment Received',          // J  CRM_COL_PAY_RECEIVED
      'Payment Date',              // K  CRM_COL_PAY_DATE
      'Dispatch Status',           // L  CRM_COL_DISPATCH_STATUS
      'Due Balance',               // M  CRM_COL_DUE_BALANCE
      'Approval Required',         // N  CRM_COL_APPROVAL_REQ
      'Approved By',               // O  CRM_COL_APPROVED_BY
      'Remarks'                    // P  CRM_COL_REMARKS
    ];
    mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    mainSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#0D47A1').setFontColor('#FFFFFF').setFontSize(11);
    mainSheet.setFrozenRows(1);

    // Hide UUID (col A) and CRM Owner (col C) — not needed by the CRM user
    mainSheet.hideColumns(CRM_COL_UUID);
    mainSheet.hideColumns(CRM_COL_CRM_OWNER);

    const widths = {
      [CRM_COL_CUSTOMER]: 180,        [CRM_COL_PI_NO]: 100,
      [CRM_COL_INV_NO]: 80,           [CRM_COL_INV_VALUE]: 110,
      [CRM_COL_PAY_STATUS]: 120,      [CRM_COL_DISPATCH_STATUS]: 130,
      [CRM_COL_DUE_BALANCE]: 110,     [CRM_COL_DUE_DATE]: 110,
      [CRM_COL_PAY_RECEIVED]: 130,    [CRM_COL_PAY_DATE]: 110
    };
    Object.entries(widths).forEach(([c, w]) => mainSheet.setColumnWidth(Number(c), w));

    const ownerEmail = Session.getActiveUser().getEmail();
    // Stamp owner email in hidden col C, row 2 (used by sync scripts)
    try { ss.setNamedRange('CRM_OWNER_EMAIL', mainSheet.getRange('C2')); } catch (_) {}
    mainSheet.getRange('C2').setValue(ownerEmail);

    mainSheet.getRange(2, CRM_COL_APPROVAL_REQ, 1000, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Yes', 'No', 'Approved'], true).build()
      );
    mainSheet.getRange(2, CRM_COL_PAY_STATUS, 1000, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Pending', 'Partial', 'Received'], true).build()
      );
    mainSheet.getRange(2, CRM_COL_PAY_RECEIVED, 1000, 3)
      .protect().setDescription('Filled by hub script — do not edit manually').setWarningOnly(true);

    const cfRange = mainSheet.getRange(2, 1, 1000, headers.length);
    mainSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=C2="${ownerEmail}"`)
        .setBackground('#FFF9C4').setRanges([cfRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=N2="Yes"`)
        .setBackground('#FFE0B2').setRanges([cfRange]).build(),
    ]);

    // Add a note on the tab label change
    mainSheet.getRange(1, CRM_COL_CUSTOMER)
      .setNote('This is the Order Ledger CRM tab (v14.0). One row per order. CRM Owner column is hidden — it is pre-filled with your email automatically.');

    // ── Part 1b: Repair formulas on ALL existing client sub-tabs ─
    // Runs every time setupCRMSheet() is called so any tab created
    // before v14.0 gets the correct formula columns automatically.
    fixSubTabFormulas_(ss);

    // ── Part 2: Client Dues sheet (CRM-filtered) ──────────────
    setupCRMDuesSheet_(ss);

    logActivity_(ss, 'CRM SETUP v14.0', 'setupCRMSheet()', ownerEmail);
    alertIfUi_(
      '✅ CRM spoke setup complete (v14.0).\n\n' +
      'Sheets created/updated:\n' +
      '• Order Ledger CRM (Main tab — CRM Owner column now hidden)\n' +
      '• Client Dues (your orders with outstanding balances)\n\n' +
      'Client sub-tabs will now include a "Due Date" column (col P).'
    );
  } catch (err) {
    Logger.log('ERROR in setupCRMSheet: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

// ============================================================
// fixSubTabFormulas_()
// Repairs FOB / CIF / Final Invoice ARRAYFORMULA on every
// client sub-tab in this CRM Spoke. Safe to run multiple times.
// Called automatically by setupCRMSheet() and can also be run
// manually from the script editor at any time.
// ============================================================
function fixSubTabFormulas_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const systemSheets = new Set([
    CRM_ORDER_LEDGER_SHEET, 'Main', CRM_DUES_SHEET,
    LOG_SHEET_NAME, MEMBERS_SHEET_NAME, CREDIT_SHEET_NAME
  ]);
  let fixed = 0;
  ss.getSheets().forEach(sheet => {
    if (systemSheets.has(sheet.getName())) return;
    try {
      // Col I = FOB Value = Item Value (G) + Dev Charges (H)
      sheet.getRange(2, SUB_COL_FOB_VALUE).setFormula(
        '=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(G2:G),0)+IFERROR(VALUE(H2:H),0)))'
      );
      // Col K = CIF Value = FOB (I) + Freight (J)
      sheet.getRange(2, SUB_COL_CIF_VALUE).setFormula(
        '=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(I2:I),0)+IFERROR(VALUE(J2:J),0)))'
      );
      // Col M = Final Invoice = CIF (K) - Discount (L)
      sheet.getRange(2, SUB_COL_FINAL_INV).setFormula(
        '=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(K2:K),0)-IFERROR(VALUE(L2:L),0)))'
      );
      fixed++;
      Logger.log('fixSubTabFormulas_: fixed "' + sheet.getName() + '"');
    } catch (err) {
      Logger.log('fixSubTabFormulas_: error on "' + sheet.getName() + '": ' + err.message);
    }
  });
  Logger.log('fixSubTabFormulas_: ' + fixed + ' tab(s) repaired.');
}

// Public wrapper so it appears in the script editor dropdown
function fixSubTabFormulas() { fixSubTabFormulas_(); }

function setupCRMDuesSheet_(ss) {
  let sheet = ss.getSheetByName(CRM_DUES_SHEET);
  if (!sheet) sheet = ss.insertSheet(CRM_DUES_SHEET);

  const headers = [
    'Client Name',       // A
    'PI No',             // B
    'Inv No',            // C
    'Final Invoice',     // D
    'Payment Received',  // E
    'Due Balance',       // F
    'Due Date',          // G
    'Payment Status'     // H
  ];

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setValues([headers]);
  hdrRange.setFontWeight('bold').setBackground('#B71C1C').setFontColor('#FFFFFF')
          .setFontSize(11).setWrap(true);
  sheet.setFrozenRows(1);

  const colWidths = [190, 100, 80, 130, 130, 130, 110, 120];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.getRange(2, 4, 999, 1).setNumberFormat('#,##0.00');  // Final Invoice
  sheet.getRange(2, 5, 999, 1).setNumberFormat('#,##0.00');  // Payment Received
  sheet.getRange(2, 6, 999, 1).setNumberFormat('#,##0.00');  // Due Balance
  sheet.getRange(2, 7, 999, 1).setNumberFormat('dd/mm/yyyy'); // Due Date

  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(G2<TODAY(),G2<>"")`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),   // overdue
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=H2="Partial"`)
      .setBackground('#FFF9C4').setRanges([cfRange]).build(),   // partial
  ]);

  // Read-only for the CRM user — populated by Hub sync
  sheet.getRange(2, 1, 1000, headers.length)
    .protect()
    .setDescription('Auto-populated by Hub sync — do not edit manually')
    .setWarningOnly(true);

  sheet.getRange(1, 1)
    .setNote('Client Dues — your outstanding balances, auto-refreshed from Hub. Do not edit manually.');

  Logger.log('setupCRMDuesSheet_: "' + CRM_DUES_SHEET + '" ready in CRM spoke.');
}

function createClientTab(sheetFile, customerName, mainSheetRow, uuid) {
  try {
    const tabName = customerName.toString().trim()
      .replace(/[\\\/\*\[\]\:]/g, '')
      .substring(0, 100);

    let clientSheet = sheetFile.getSheetByName(tabName);
    if (!clientSheet) {
      clientSheet = sheetFile.insertSheet(tabName);

      // v14.0: Added 'Due Date' (col P) to client sub-tab headers
      const headers = [
        'PI No',                        // A  SUB_COL_PI_NO           = 1
        'Inv No',                       // B  SUB_COL_INV_NO          = 2
        'Inv Date',                     // C  SUB_COL_INV_DATE        = 3
        'Order Type',                   // D  SUB_COL_ORDER_TYPE      = 4
        'AWB No',                       // E  SUB_COL_AWB_NO          = 5
        'Dispatched Date',              // F  SUB_COL_DISP_DATE       = 6
        'Item Value USD',               // G  SUB_COL_ITEM_VALUE_USD  = 7
        'Development Charges USD',      // H  SUB_COL_DEV_CHARGES_USD = 8
        'FOB Value (G+H)',              // I  SUB_COL_FOB_VALUE       = 9
        'Freight',                      // J  SUB_COL_FREIGHT         = 10
        'CIF Value (I+J)',              // K  SUB_COL_CIF_VALUE       = 11
        'Discount',                     // L  SUB_COL_DISCOUNT        = 12
        'Final Invoice Value (K-L)',    // M  SUB_COL_FINAL_INV       = 13
        'Payment Received',             // N  SUB_COL_PAY_RECEIVED    = 14
        'Payment Date',                 // O  SUB_COL_PAY_DATE        = 15
        'Due Date'                      // P  SUB_COL_DUE_DATE        = 16  NEW v14.0
      ];

      const hdrRange = clientSheet.getRange(1, 1, 1, headers.length);
      hdrRange.setValues([headers]);
      hdrRange.setFontWeight('bold').setBackground('#1565C0').setFontColor('#FFFFFF').setWrap(true);
      clientSheet.setFrozenRows(1);

      // Number formats
      clientSheet.getRange(2, SUB_COL_INV_DATE,  999, 1).setNumberFormat('dd/mm/yyyy');
      clientSheet.getRange(2, SUB_COL_DISP_DATE, 999, 1).setNumberFormat('dd/mm/yyyy');
      clientSheet.getRange(2, SUB_COL_PAY_DATE,  999, 1).setNumberFormat('dd/mm/yyyy');
      clientSheet.getRange(2, SUB_COL_DUE_DATE,  999, 1).setNumberFormat('dd/mm/yyyy'); // NEW v14.0

      const orderTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Production', 'Sample', 'Showroom'], true)
        .setAllowInvalid(false).build();
      clientSheet.getRange(2, SUB_COL_ORDER_TYPE, 999, 1).setDataValidation(orderTypeRule);

      // Formulas (columns G–M unchanged)
      // FOB  (col I) = Item Value USD (G) + Dev Charges USD (H)
      // CIF  (col K) = FOB (I) + Freight (J)
      // Final Invoice (col M) = CIF (K) - Discount (L)
      clientSheet.getRange(2, SUB_COL_FOB_VALUE).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(G2:G),0)+IFERROR(VALUE(H2:H),0)))`
      );
      clientSheet.getRange(2, SUB_COL_CIF_VALUE).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(I2:I),0)+IFERROR(VALUE(J2:J),0)))`
      );
      clientSheet.getRange(2, SUB_COL_FINAL_INV).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(K2:K),0)-IFERROR(VALUE(L2:L),0)))`
      );

      clientSheet.autoResizeColumns(1, headers.length);
      Logger.log('createClientTab: created sub-tab "' + tabName + '" with Due Date col (v14.0)');
    }

    // ── Write HYPERLINK back to CRM Main ──────────────────────
    const mainSheet = getCRMMainSheet_(sheetFile);
    if (!mainSheet) return clientSheet;

    let targetRow = mainSheetRow || null;

    if (!targetRow) {
      const crmData    = mainSheet.getDataRange().getValues();
      const crmDisplay = mainSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < crmDisplay.length; i++) {
        // Match by display name OR by UUID (catches rows where HYPERLINK not yet written)
        const cellVal  = (crmDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim();
        const cellUuid = (crmData[i][CRM_COL_UUID - 1] || '').toString().trim();
        if (cellVal.toLowerCase() === tabName.toLowerCase() ||
            (uuid && cellUuid === uuid)) {
          targetRow = i + 1;
          break;
        }
      }
    }

    if (!targetRow) {
      targetRow = mainSheet.getLastRow() + 1;
      Logger.log('createClientTab: appending new CRM Main row ' + targetRow + ' for "' + tabName + '"');
    }

    // Write UUID before HYPERLINK so onCRMEdit sees it immediately
    if (uuid) {
      const existingUuid = mainSheet.getRange(targetRow, CRM_COL_UUID).getValue();
      if (!existingUuid) {
        mainSheet.getRange(targetRow, CRM_COL_UUID).setValue(uuid);
        SpreadsheetApp.flush();
        Logger.log('createClientTab: UUID written to Main row ' + targetRow + ': ' + uuid);
      }
    }

    // Write HYPERLINK formula to col B
    const sheetId  = clientSheet.getSheetId();
    const fileId   = sheetFile.getId();
    const tabUrl   = `https://docs.google.com/spreadsheets/d/${fileId}/edit#gid=${sheetId}`;
    const formula  = `=HYPERLINK("${tabUrl}","${tabName}")`;
    mainSheet.getRange(targetRow, CRM_COL_CUSTOMER).setFormula(formula);
    SpreadsheetApp.flush();

    Logger.log('createClientTab: HYPERLINK written to Main row ' + targetRow + ' col B for "' + tabName + '"');
    return clientSheet;

  } catch (err) {
    Logger.log('ERROR in createClientTab: ' + err.message);
    return null;
  }
}

function autoAddClientToCredit_(hubSS, clientName, crmOwner) {
  try {
    if (!clientName) return;
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) return;

    const data = creditSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][CC_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase() ===
          clientName.toString().trim().toLowerCase()) {
        return; // already exists
      }
    }

    // New client row — CC_TOTAL_COLS = 7 visible cols + UUID in col 8
    const newRow = Array(CC_TOTAL_COLS + 1).fill('');  // +1 for UUID col H
    newRow[CC_COL_CLIENT_NAME  - 1] = clientName;
    newRow[CC_COL_CRM_OWNER    - 1] = crmOwner || '';
    newRow[CC_COL_CREDIT_LIMIT - 1] = 0;
    newRow[CC_COL_LAST_UPDATED - 1] = new Date();
    // CC_COL_CLIENT_EMAIL (col F) = blank — to be filled in later
    // CC_COL_UUID (col H = index 7) = blank — stamped when needed

    const appendAt = creditSheet.getLastRow() + 1;
    creditSheet.getRange(appendAt, 1, 1, newRow.length).setValues([newRow]);
    creditSheet.getRange(appendAt, CC_COL_LAST_UPDATED).setNumberFormat('dd/mm/yyyy hh:mm');

    refreshCreditTotals_(Session.getActiveUser().getEmail() || 'auto');
    Logger.log('autoAddClientToCredit_: added "' + clientName + '" to Client Credit (v14.0)');
  } catch (err) {
    Logger.log('ERROR in autoAddClientToCredit_: ' + err.message);
  }
}

function setCreditLimit_(clientName, crmOwner, limit) {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) {
      alertIfUi_('Client Credit sheet not found. Run Setup -> Setup Credit sheet first.');
      return;
    }

    const data = creditSheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if ((data[i][CC_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase() ===
          clientName.toLowerCase()) {
        existingRow = i + 1;
        break;
      }
    }

    const userEmail = Session.getActiveUser().getEmail();
    const now       = new Date();

    if (existingRow !== -1) {
      creditSheet.getRange(existingRow, CC_COL_CLIENT_NAME ).setValue(clientName);
      creditSheet.getRange(existingRow, CC_COL_CRM_OWNER   ).setValue(crmOwner);
      creditSheet.getRange(existingRow, CC_COL_CREDIT_LIMIT).setValue(limit);
      creditSheet.getRange(existingRow, CC_COL_LAST_UPDATED).setValue(now);
      logActivity_(hubSS, 'CREDIT UPDATED', `Client:${clientName} | Limit:$${limit}`, userEmail);
      alertIfUi_(`Credit limit updated for ${clientName}: $${limit.toLocaleString('en-US')}`);
    } else {
      // New row — visible cols A–G, UUID in col H
      const newRow = Array(CC_TOTAL_COLS + 1).fill('');
      newRow[CC_COL_CLIENT_NAME  - 1] = clientName;
      newRow[CC_COL_CRM_OWNER    - 1] = crmOwner;
      newRow[CC_COL_CREDIT_LIMIT - 1] = limit;
      newRow[CC_COL_LAST_UPDATED - 1] = now;
      const appendAt = creditSheet.getLastRow() + 1;
      creditSheet.getRange(appendAt, 1, 1, newRow.length).setValues([newRow]);
      logActivity_(hubSS, 'CREDIT ADDED', `Client:${clientName} | Limit:$${limit}`, userEmail);
      alertIfUi_(`Credit limit set for ${clientName}: $${limit.toLocaleString('en-US')}`);
    }
  } catch (err) {
    Logger.log('ERROR in setCreditLimit_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}


// ============================================================
// SECTION 2 — CORE FUNCTIONS (from v13 base, unchanged)
// ============================================================


// ============================================================
// alertIfUi_()
// Safe wrapper for alertIfUi_().
// getUi() throws when called from a non-UI context (e.g. when
// running directly from the Apps Script editor, time-based
// triggers, or internal calls). This wrapper silently falls
// back to Logger.log() so the rest of the function still runs.
// ============================================================


// ============================================================
// NEW v3.0 — getCrmConfig_()
// ─────────────────────────────────────────────────────────────
// Central function that ALL other functions call instead of
// reading hardcoded CRM_FILE_IDS or role email constants.
//
// Returns an object:
// {
//   crmFileIds:       { 'gmail': 'fileId', ... },   // CRM role only
//   salesManagerEmails: ['email', ...],
//   adminEmails:        ['email', ...],
//   logisticsEmails:    ['email', ...],
//   accountsEmails:     ['email', ...],
//   allMembers:         [ { gmail, fileId, role, name, active }, ... ]
// }
//
// Results are cached in CacheService for 6 minutes to avoid
// reading the hub sheet on every single onEdit() call.
// ============================================================


// ============================================================
// NEW v3.0 — invalidateConfigCache_()
// Call this after any change to the CRM Members sheet so
// all subsequent calls get fresh data immediately.
// ============================================================


// ============================================================
// NEW v3.0 — setupMembersSheet()
// Creates the "CRM Members" config tab in Master Hub.
// Run ONCE from the Master Hub Apps Script editor.
// After this, add/edit/remove members directly in this sheet
// — no code changes ever needed again.
// ============================================================


// ============================================================
// NEW v3.0 — onMembersSheetEdit(e)
// Fires when admin edits the CRM Members config sheet.
// Immediately clears the cache so next script run gets
// fresh config without waiting for the 6-minute TTL.
// Install this trigger in setupAllTriggers().
// ============================================================
// ============================================================
// onMembersSheetEdit(e)  — v3.1 AUTO-SHARE + NOTIFY
// Fires when the CRM Members sheet is edited.
// • Always clears the config cache
// • When a row has Gmail + Role + Active=TRUE -> auto-shares
//   the correct file and sends a welcome email to the member
// ============================================================

// ============================================================
// notifyAndShareMember_()
// Called by onMembersSheetEdit when a complete member row is saved.
// 1. Shares the correct Google Sheet with the member (as Editor)
// 2. Sends them a role-specific welcome email
// ============================================================


// ============================================================
// 1. setupAllTriggers()
//    v3.0: Also installs onMembersSheetEdit trigger on hub.
//    Idempotent — safe to run multiple times.
// ============================================================


// ============================================================
// 2. onCRMEdit(e)
//    v3.0: Uses getCrmConfig_() — no hardcoded emails.
//    v4.0: Lock-first architecture — lock acquired before UUID
//          generation and tab creation to prevent concurrent
//          trigger executions from producing duplicate rows,
//          mismatched UUIDs, and missing client tabs.
// ============================================================


// ============================================================
// 3. onMasterHubEdit(e)
//    v3.0: Uses getCrmConfig_() to look up CRM file IDs
//    dynamically — no hardcoded CRM_FILE_IDS object.
// ============================================================


// ============================================================
// extractEmail_()
// Parses a raw email address from a dropdown value that may be
// formatted as "Display Name (email@gmail.com)" or may already
// be a bare email. Returns the lowercased email, or the original
// lowercased string if no parenthesised email is found.
// ============================================================


// ============================================================
// 4. onCreditRequestEdit(e)
//    v3.0: Admin email comes from getCrmConfig_().
//    BUG FIX: extractEmail_() strips "Name (email)" dropdown
//    format before config.crmFileIds lookup and email sends.
// ============================================================


// ============================================================
// sendNewClientNotificationToCRM_()
// ─────────────────────────────────────────────────────────────
// Fires when Sales Manager / Sales / Admin adds a new client
// row in the Client Credit sheet (fills Client Name + CRM Owner).
// Sends a professional email to the assigned CRM owner with:
//   • Client name, UUID, who added them and when
//   • Direct link to their CRM spoke sheet → Main tab
//   • Instructions on what to fill next (PI No, Inv No, etc.)
// ============================================================


// ============================================================
//    v3.0: Sales Manager emails come from getCrmConfig_().
// ============================================================


// ============================================================
// setupCreditSheet()  — v3.0: admin email from config
// ============================================================


// ============================================================
// createClientTab()  — unchanged logic
// ============================================================
// createClientTab(sheetFile, customerName, mainSheetRow, uuid)
// ─────────────────────────────────────────────────────────────
// Creates the client sub-tab if it doesn't exist, then writes
// the HYPERLINK formula to the CRM Main sheet row.
//
// mainSheetRow — existing row number in CRM Main, OR omit/null
//               to have this function append a new placeholder row.
// uuid         — when provided, written to CRM_COL_UUID (col A)
//               BEFORE the HYPERLINK formula, so onCRMEdit sees
//               the UUID immediately and never appends a dup Hub row.


// ============================================================
// exportClientPDF()  — unchanged logic
// ============================================================


// ============================================================
// exportClientCSV()  — unchanged logic
// ============================================================


// ============================================================
// addCustomMenu()
//    v3.0: Added "👥 Manage Members" sub-menu.
// ============================================================


// ── NEW v3.0 Members menu handlers ────────────────────────────


// ============================================================
// menuRepairCreditProtection_()
// Run this from Admin after adding new Sales users to CRM Members.
// Re-applies the Credit Limit column protection with the latest
// Sales + Admin emails so new Sales members can edit col C.
// ============================================================


// ============================================================
// menuRepairHubProtection_()
// BUG 1 FIX — Run from Admin after adding new Logistics or
// Accounts users to CRM Members, or if J/K columns are not
// editable by Accounts or Q/R not editable by Logistics.
// Re-applies Hub column protections with the latest role emails.
// ============================================================
function menuRepairHubProtection_() {
  try {
    const hubSS      = SpreadsheetApp.getActiveSpreadsheet();
    const hubSheet   = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { alertIfUi_('Order Ledger sheet not found. Run Setup first.'); return; }
    const ownerEmail      = Session.getEffectiveUser().getEmail();
    const config          = getCrmConfig_();
    const crmSpokeEmails  = Object.keys(config.crmFileIds);
    const adminEditors     = [...new Set([ownerEmail, ...config.adminEmails,     ...crmSpokeEmails])];
    const logisticsEditors = [...new Set([ownerEmail, ...config.logisticsEmails, ...crmSpokeEmails])];
    const accountsEditors  = [...new Set([ownerEmail, ...config.accountsEmails,  ...crmSpokeEmails])];
    const approvalEditors  = [...new Set([ownerEmail, ...config.adminEmails,
      ...config.salesEmails, ...crmSpokeEmails])];

    hubSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    function applyProt_(range, desc, extraEditors) {
      const p = range.protect().setDescription(desc);
      p.removeEditors(p.getEditors());
      p.addEditors(extraEditors || adminEditors);
      p.setWarningOnly(false);
    }

    // Cols B–I: Admin only
    applyProt_(hubSheet.getRange(2, 2, 1000, 8),  'Admin only — core order columns B–I');
    // Cols J–K: Accounts + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_PAY_RECEIVED, 1000, 2),
               'Accounts only — Payment Received & Payment Date (J–K)', accountsEditors);
    // Cols L–M: Admin only (formula columns)
    applyProt_(hubSheet.getRange(2, 12, 1000, 2), 'Admin only — formula columns L–M');
    // Cols N–P: Sales + Sales Manager + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_APPROVAL_REQ, 1000, 3),
               'Sales+SalesManager — Approval Required, Approved By, Remarks (N–P)', approvalEditors);
    // Cols Q–R: Logistics + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_AWB_NO, 1000, 2),
               'Logistics only — AWB No & Dispatched Date (Q–R)', logisticsEditors);

    logActivity_(hubSS, 'HUB PROTECTION REPAIRED',
      `Admin: ${adminEditors.join(', ')} | Accounts (J-K): ${accountsEditors.join(', ')} | Logistics (Q-R): ${logisticsEditors.join(', ')}`, ownerEmail);

    // Also repair formulas on all existing Hub rows (H, L, M)
    const lastRow = hubSheet.getLastRow();
    if (lastRow >= 2) {
      for (let r = 2; r <= lastRow; r++) setHubFormulas_(hubSheet, r);
      SpreadsheetApp.flush();
    }

    alertIfUi_(
      'Master Hub protection + formulas fixed!\n\n' +
      'Cols B-I (core data):    Admin only\n' +
      'Cols J-K (payment):      Accounts + Admin\n' +
      'Cols L-M (formulas):     Admin only\n' +
      'Cols N-P (approval):     Sales + Sales Manager + Admin\n' +
      'Cols Q-R (dispatch):     Logistics + Admin\n\n' +
      'Due Balance (M), Payment Status (H), Dispatch Status (L)\n' +
      'formulas repaired on all existing rows.\n\n' +
      'Accounts editors:\n' + accountsEditors.join('\n') + '\n\n' +
      'Logistics editors:\n' + logisticsEditors.join('\n') + '\n\n' +
      'Approval editors (N-P):\n' + approvalEditors.join('\n')
    );
  } catch (err) {
    Logger.log('ERROR in menuRepairHubProtection_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

// ============================================================
// fixHubWriteAccess()
// Permanently adds all CRM spoke emails to every non-warning
// range protection in the Master Hub spreadsheet.
// Run this from Master Hub after any protection repair wipes
// CRM access. Also called automatically by setupEverything().
// ============================================================
function fixHubWriteAccess() {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const config = getCrmConfig_();
    const crmEmails = Object.keys(config.crmFileIds);

    if (!crmEmails.length) {
      alertIfUi_('No CRM spoke emails found in CRM Members sheet.\nAdd CRM members first.');
      return;
    }

    let count = 0;
    ss.getSheets().forEach(sheet => {
      sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => {
        if (p.isWarningOnly()) return;
        try {
          p.addEditors(crmEmails);
          count++;
          Logger.log('fixHubWriteAccess: added CRM emails to: ' +
            sheet.getName() + ' > ' + p.getDescription());
        } catch (e) {
          Logger.log('fixHubWriteAccess: skip "' + p.getDescription() +
            '" — ' + e.message);
        }
      });
      // v15.1 FIX: also patch SHEET-level protections (e.g. Activity Log,
      // which uses sheet.protect() rather than a range protection). The loop
      // above only ever checked ProtectionType.RANGE, so a sheet-level
      // protection's "protected cell" error could never actually be fixed by
      // running this function, despite the hardcoded error hints across this
      // codebase telling people to run it for exactly that error.
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
        if (p.isWarningOnly()) return;
        try {
          p.addEditors(crmEmails);
          count++;
          Logger.log('fixHubWriteAccess: added CRM emails to SHEET protection: ' +
            sheet.getName() + ' > ' + p.getDescription());
        } catch (e) {
          Logger.log('fixHubWriteAccess: skip sheet-protection "' + p.getDescription() +
            '" — ' + e.message);
        }
      });
    });

    SpreadsheetApp.flush();
    const msg = '✅ fixHubWriteAccess complete!\n\n' +
      'Added ' + crmEmails.length + ' CRM email(s) to ' + count + ' protection(s).\n\n' +
      'CRM emails:\n' + crmEmails.join('\n');
    Logger.log(msg);
    alertIfUi_(msg);

  } catch (err) {
    Logger.log('ERROR in fixHubWriteAccess: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}


// ============================================================
// BUG 1 FIX — menuSalesAddMemberAndClient_()
// Sales-specific guided workflow:
//   1. Add new CRM member to CRM Members sheet
//   2. Add their client to Client Credit sheet
//   3. Set the client's credit limit — all in one flow.
// Only Sales and Admin roles should use this.
// ============================================================


// ============================================================
// setCreditLimit_()
// ============================================================


// ============================================================
// requestCreditIncrease_()
// ============================================================


// ============================================================
// setupHubSheet_legacy_()  — v3.0 retained for reference only
// This is the old v3 Hub setup — kept so menuRepairHubProtection_
// can reuse its applyProt_ pattern. Not called directly.
// ============================================================
function setupHubSheet_legacy_() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = Session.getEffectiveUser().getEmail();
    const config     = getCrmConfig_();
    const crmSpokeEmails  = Object.keys(config.crmFileIds);
    const adminEditors     = [...new Set([ownerEmail, ...config.adminEmails,     ...crmSpokeEmails])];
    const logisticsEditors = [...new Set([ownerEmail, ...config.logisticsEmails, ...crmSpokeEmails])];
    const accountsEditors  = [...new Set([ownerEmail, ...config.accountsEmails,  ...crmSpokeEmails])];
    const approvalEditors  = [...new Set([ownerEmail, ...config.adminEmails,
      ...config.salesEmails, ...crmSpokeEmails])];

    // Master Hub removed v15.0 — Order Ledger is primary

    const headers = [
      'UUID','Customer Name','CRM Owner (Gmail)','PI No','Inv No','Inv Date',
      'Order Type','Inv Value','Due Date','Payment Status','Payment Received',
      'Payment Date','Dispatch Status','Due Balance','Approval Required',
      'Approved By','Remarks','AWB No','Dispatch Date'
    ];
    mainSheet.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF');

    function applyProt_(range, desc, extraEditors) {
      const p = range.protect().setDescription(desc);
      p.removeEditors(p.getEditors());
      p.addEditors(extraEditors || adminEditors);
      p.setWarningOnly(false);
    }
    mainSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // Cols B–I (2–9): Admin only — core order data, no role editing
    applyProt_(mainSheet.getRange(2, 2, 1000, 8),  'Admin only — core order columns B–I');
    // Cols J–K (10–11): Accounts only (+ Admin)
    applyProt_(mainSheet.getRange(2, HUB_COL_PAY_RECEIVED, 1000, 2),
               'Accounts only — Payment Received & Payment Date (J–K)', accountsEditors);
    // Cols L–M (12–13): Admin only — formula columns (Dispatch Status, Due Balance)
    applyProt_(mainSheet.getRange(2, 12, 1000, 2), 'Admin only — formula columns L–M');
    // Cols N–P (14–16): Sales + Sales Manager + Admin — Approval Required, Approved By, Remarks
    applyProt_(mainSheet.getRange(2, HUB_COL_APPROVAL_REQ, 1000, 3),
               'Sales+SalesManager — Approval Required, Approved By, Remarks (N–P)', approvalEditors);
    // Cols Q–R (17–18): Logistics only (+ Admin)
    applyProt_(mainSheet.getRange(2, HUB_COL_AWB_NO, 1000, 2),
               'Logistics only — AWB No & Dispatched Date (Q–R)', logisticsEditors);

    const cfRange = mainSheet.getRange(2, 1, 1000, headers.length);
    mainSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(I2<TODAY(),H2<>"Received")`)
        .setBackground('#FFCDD2').setRanges([cfRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=N2="Yes"`)
        .setBackground('#FFE0B2').setRanges([cfRange]).build(),
    ]);

    setupCreditSheet();
    logActivity_(ss, 'HUB SETUP', 'setupHubSheet()', Session.getActiveUser().getEmail());
    alertIfUi_('Master Hub setup complete.');
  } catch (err) {
    Logger.log('ERROR in setupHubSheet: ' + err.message);
  }
}


// ============================================================
// setupCRMSheet()  — unchanged logic
// ============================================================


// ============================================================
// syncAllCRMToHub()  — unchanged logic
// ============================================================

// ============================================================
// onCRMSubTabEdit(e) — v11.0 COMPLETE FIX
//
// COLUMN SYNC MAP (sub-tab -> CRM Main -> Hub):
//   Col 1  PI No         -> CRM Main col D  -> Hub col D
//   Col 2  Inv No        -> CRM Main col E  -> Hub col E
//   Col 3  Inv Date      -> CRM Main col F  -> Hub col F
//   Col 5  AWB No        ->                    Hub col Q
//   Col 6  Disp Date     ->                    Hub col R
//   Col 7  Item Value    -> triggers Final Inv push
//   Col 8  Dev Charges   -> triggers Final Inv push
//   Col 9  FOB (formula) -> triggers Final Inv push
//   Col 10 Freight       -> triggers Final Inv push
//   Col 11 CIF (formula) -> triggers Final Inv push
//   Col 12 Discount      -> triggers Final Inv push
//   Col 13 Final Inv     -> CRM Main col G  -> Hub col G
//   Col 14 Pay Received  -> CRM Main col J  -> Hub col J
//   Col 15 Pay Date      -> CRM Main col K  -> Hub col K
// ============================================================


// ============================================================
// addHubFormulasToAllRows()
// ============================================================


// ============================================================
// repairClientTab(customerName)
// ─────────────────────────────────────────────────────────────
// Repairs an existing client sub-tab that was created before
// the formula fix (v8). Rewrites:
//   • FOB, CIF, Final Invoice ARRAYFORMULAs (fixed to handle
//     blank Dev Charges / Freight / Discount gracefully)
//   • Order Type dropdown (col C)
//   • AWB No, Dispatched Date, Payment Received, Payment Date
//     column widths and header formatting
//
// Safe to run on a tab that already has data — it only touches
// the formula seed cells (H2, J2, L2) and validation on col C.
// ============================================================


// ── Repair ALL client tabs at once ────────────────────────────
// Iterates every sheet that is not a system sheet and repairs it.


// ============================================================
// deleteClientTab()
// ============================================================


// ============================================================
// sendPaymentReminder()
// ============================================================


// ============================================================
// setupTimedReminder()
// ============================================================


// ============================================================
// PRIVATE HELPERS
// ============================================================


// ============================================================
// pushFinalInvoiceToMainAndHub_()  — v10.0 FIXED
//
// WHAT CHANGED:
//   • New signature: receives pre-looked-up mainRow, mainData,
//     clientName, piNo, finalInv, uuid, crmOwner — avoids
//     redundant sheet reads that the caller already did.
//   • Sets PROP_SKIP_NEXT_EDIT before writing to CRM Main col G
//     so the resulting onCRMEdit trigger call is suppressed.
//   • Always calls refreshCreditTotals_() at the end.
//   • Calls autoAddClientToCredit_() to ensure credit row exists.
//   • Improved logging to trace execution in Apps Script Logs.
// ============================================================

// ============================================================
// refreshCreditTotals_(userEmail)
// ─────────────────────────────────────────────────────────────
// Recalculates Used Credit (D) and Available Credit (E) for
// every row in the Client Credit sheet by reading Hub col G
// (Inv Value) directly — no cross-file ARRAYFORMULA needed.
//
// WHY: The ARRAYFORMULA in setupCreditSheet references
// 'Master Hub'!G:G across files. Google Sheets treats cross-
// file formula recalculation as a background task that can lag
// minutes. This function writes plain computed values instantly.
//
// Called from pushFinalInvoiceToMainAndHub_() after Hub is
// updated, so credit totals refresh on every sub-tab edit that
// changes Invoice Value, Freight, Discount, Item Value, etc.
// Also callable manually via CRM Tools -> Credit Limit menu.
// ============================================================


// ============================================================
// reEvaluateAllCreditFlags_()
// ─────────────────────────────────────────────────────────────
// Reads Hub and Credit sheet fresh each time — never relies on
// passed-in stale arrays. For each Hub row:
//   • If client's total used credit (InvValue−PayReceived) > limit
//     AND this row is not already Approved -> set N = "Yes"
//   • If within limit and N = "Yes" -> clear to N = "No"
//   • If N = "Approved" -> leave untouched (Sales Manager approved)
// ============================================================


// Menu wrapper — lets Admin manually force a refresh any time
// ============================================================
// menuRepairAllSheetSecurity_()
// ─────────────────────────────────────────────────────────────
// One-click repair for ALL sheet security in Master Hub:
//   • Activity Log  → Admin only, hidden tab
//   • CRM Members   → Admin/Sales/SalesMgr only, hidden tab
//   • Client Credit → role-based column protections
//   • Master Hub    → role-based column protections
// Run this from Master Hub after adding new members or if
// any sheet security seems wrong.
// ============================================================


// ============================================================
// diagnoseTriggers()
// ─────────────────────────────────────────────────────────────
// Run this from the CRM Spoke script editor (or via the menu)
// to see exactly which installable triggers are active for this
// spreadsheet and which user installed them.
//
// Common issue: setupAllTriggers() was run from the Hub while
// the CRM Spoke editor was open — this installs Hub triggers
// against the Spoke file ID instead of Spoke triggers.
// Or: triggers exist but are bound to a DIFFERENT spreadsheet.
//
// HOW TO USE:
//   1. Open the CRM Spoke spreadsheet.
//   2. Extensions > Apps Script > Run > diagnoseTriggers
//   3. Check the popup and Execution Log.
//   Expected: onCRMEdit + onCRMSubTabEdit bound to THIS file.
// ============================================================


// ============================================================
// menuForceSubTabSync_()
// ─────────────────────────────────────────────────────────────
// One-click repair: reads ALL client sub-tabs in this CRM Spoke
// and pushes every row's data to CRM Main + Master Hub.
// Use this after pasting/updating code, or if onCRMSubTabEdit
// wasn't firing and Hub data is stale.
//
// Pushes: PI No, Inv No, Inv Date, Final Invoice, AWB No,
//         Dispatch Date, Payment Received, Payment Date.
// Also calls autoAddClientToCredit_ + refreshCreditTotals_.
// ============================================================


// ============================================================
// pushNToCRM_()
// Pushes Approval Required (N) value to the matching CRM Main
// row identified by UUID. Uses a cache to avoid re-opening the
// same CRM spreadsheet multiple times in one bulk pass.
// ============================================================


// ============================================================
// sendHoldNotificationEmails_()
// Sends dispatch-hold notification to Sales Manager, Sales,
// and Admin when a Hub row is flagged N = "Yes".
// Called by reEvaluateAllCreditFlags_ and checkCreditAndFlag_.
// ============================================================


// ============================================================
// autoAddClientToCredit_()  — Bug 1 fix
// Called from onCRMEdit after hub row is written.
// Adds the client to the Client Credit sheet automatically
// if they don't already have a row there.
// ============================================================


// ============================================================
// applyActivityLogSecurity_()
// Applies protection + hides the Activity Log sheet.
// Called on first creation (logActivity_) and from
// menuRepairAllSheetSecurity_ for repair runs.
// ============================================================


// ============================================================
// nukeClientCreditProtections()
// ─────────────────────────────────────────────────────────────
// EMERGENCY FIX — Run this directly from the Apps Script editor
// (Master Hub) if Sales / Sales Manager still cannot edit the
// Client Credit sheet after running setupCreditSheet.
//
// What it does:
//   1. Removes EVERY range protection on the Client Credit sheet
//      (no description filtering — unconditional full wipe)
//   2. Re-applies ONLY the Admin-only lock on cols H–I
//   3. Leaves col C (Credit Limit) completely unprotected
//
// Run once. Sales and Sales Manager can edit immediately after.
// ============================================================

// ============================================================
// syncSubTabsToMain()
// One-time manual sync: reads every client sub-tab and pushes
// PI No, Inv No, Inv Date, Final Invoice to CRM Main + Hub.
// Run from CRM Spoke script editor when data is out of sync.
// ============================================================


// ─────────────────────────────────────────
// LICENSE CHECKER
// ─────────────────────────────────────────
const LICENSE_SHEET_NAME = "Licenses";


function onOpen() {
  // ── Real CRM Tools menu ──
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('CRM Tools');

  menu.addSubMenu(ui.createMenu('⚙️ Setup')
    .addItem('Setup Hub Sheet',         'setupMasterSheet')
    .addItem('Setup Members Sheet',     'setupMembersSheet')
    .addItem('Setup Credit Sheet',      'setupCreditSheet')
    .addItem('Setup Spoke Sheet',       'setupSpokeSheet')
    .addItem('Set Up Triggers',         'setupTriggers')
    .addItem('Schedule Weekly Reminder','scheduleWeeklyReminder'));

  menu.addSubMenu(ui.createMenu('👥 Manage Members')
    .addItem('View Members',            'viewMembers')
    .addItem('Reload Config Cache',     'reloadConfigCache')
    .addItem('Repair ALL sheet security','repairAllSheetSecurity'));

  menu.addSubMenu(ui.createMenu('🔄 Sync & Repair')
    .addItem('Bulk sync CRM → Hub',     'syncSubTabsToMainUnprotected')
    .addItem('Repair Hub formulas',     'repairHubFormulas')
    .addItem('Repair Client Tab formulas','repairClientTabFormulas'));

  menu.addSubMenu(ui.createMenu('💰 Credit Limit')
    .addItem('Set Credit Limit',        'setCreditLimit')
    .addItem('Request Increase',        'requestCreditIncrease')
    .addItem('Refresh credit totals now','refreshCreditTotals'));

  menu.addSubMenu(ui.createMenu('🧑‍💼 Sales Actions')
    .addItem('Add New CRM Member + Client + Credit Limit', 'salesAddMemberClientCredit'));

  menu.addSubMenu(ui.createMenu('📄 Export')
    .addItem('Export Client Tab as PDF','exportClientTabPdf')
    .addItem('Export Client Tab as CSV','exportClientTabCsv'));

  menu.addSubMenu(ui.createMenu('🗂️ Client Tabs')
    .addItem('🗑️ Delete client (all data)',      'menuDeleteClientFull_')
    .addItem('🗑️ Delete order by PI No / Inv No', 'menuDeleteOrderByRef_'));

  menu.addToUi();
}

function alertIfUi_(msg) {
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (_) {
    Logger.log('ALERT (no UI context): ' + msg);
  }
}

function weeklyReminder() {
  // Don't call checkLicense() here — no UI context
  // Just run the reminder logic directly
}

function onMembersSheetEdit(e) {
  const _licEmail = (e && e.user) ? e.user.email : '';
  if (!checkLicense(_licEmail)) return;
  // Guard: this function requires a real sheet edit event.
  // Running it manually from the editor passes no event object.
  if (!e || !e.range) {
    Logger.log('onMembersSheetEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== MEMBERS_SHEET_NAME) return;

    // Always invalidate cache
    invalidateConfigCache_();

    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row <= 1) return;

    // ONLY fire when the Active checkbox (col E) is checked TRUE
    if (col !== MC_COL_ACTIVE) return;
    if (e.value !== 'TRUE') return; // unchecking = no email

    // Read the full row
    const rowData = sheet.getRange(row, 1, 1, MC_TOTAL_COLS).getValues()[0];
    const gmail  = (rowData[MC_COL_GMAIL  - 1] || '').toString().trim().toLowerCase();
    const fileId = (rowData[MC_COL_FILE_ID - 1] || '').toString().trim();
    const role   = (rowData[MC_COL_ROLE   - 1] || '').toString().trim();
    const name   = (rowData[MC_COL_NAME   - 1] || '').toString().trim();

    if (!gmail || !role) {
      Logger.log('onMembersSheetEdit: Gmail or Role missing — no email sent.');
      return;
    }

    notifyAndShareMember_(gmail, role, name, fileId);

    const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
    logActivity_(hubSS, 'MEMBER NOTIFIED', role + ' | ' + gmail + ' | ' + (name || '—'), Session.getActiveUser().getEmail());

  } catch (err) {
    Logger.log('ERROR in onMembersSheetEdit: ' + err.message);
  }
}

function notifyAndShareMember_(gmail, role, name, fileId) {
  const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
  const hubUrl   = 'https://docs.google.com/spreadsheets/d/' + MASTER_HUB_ID + '/edit';
  const roleLc   = role.toLowerCase();
  const greeting = name ? ('Dear ' + name) : 'Hello';

  let sharedFileUrl  = hubUrl;
  let sharedFileName = 'UV Overseas Master Hub';

  try {
    if (roleLc === 'crm' && fileId) {
      const spokeSS  = SpreadsheetApp.openById(fileId);
      const spokeFile = DriveApp.getFileById(fileId);
      spokeFile.addEditorWithoutNotification(gmail);                        // ← changed
      sharedFileUrl  = 'https://docs.google.com/spreadsheets/d/' + fileId + '/edit';
      sharedFileName = spokeSS.getName();

      DriveApp.getFileById(MASTER_HUB_ID).addViewerWithoutNotification(gmail); // ← changed
      Logger.log('notifyAndShareMember_: CRM spoke shared with ' + gmail);

    } else if (roleLc === 'sales') {
      DriveApp.getFileById(MASTER_HUB_ID).addEditorWithoutNotification(gmail); // ← Sales gets editor on Hub
      Logger.log('notifyAndShareMember_: Hub shared as Editor (Sales) with ' + gmail);

    } else if (roleLc === 'admin') {
      DriveApp.getFileById(MASTER_HUB_ID).addEditorWithoutNotification(gmail); // ← changed
      Logger.log('notifyAndShareMember_: Hub shared as Editor (Admin) with ' + gmail);

    } else {
      DriveApp.getFileById(MASTER_HUB_ID).addEditorWithoutNotification(gmail); // ← changed
      Logger.log('notifyAndShareMember_: Hub shared as Editor (' + role + ') with ' + gmail);
    }
  } catch (shareErr) {
    Logger.log('notifyAndShareMember_: Share failed for ' + gmail + ' — ' + shareErr.message);
  }

  const roleInstructions = {
    'crm': `Your personal CRM sheet has been shared with you. You can:\n` +
           `  • Add new customer orders in the "Main" tab\n` +
           `  • Track payment status, dispatch, and due balances\n` +
           `  • View client sub-tabs for order details\n` +
           `  • Request credit limit increases from the Client Credit tab\n\n` +
           `You also have read-only access to the Master Hub to monitor overall order flow.`,

    'logistics': `You have been added as Logistics on the Master Hub. You can:\n` +
                 `  • Update AWB No (col Q) and Dispatch Date (col R) for any order\n` +
                 `  • All other columns are view-only for your role\n\n` +
                 `Please update dispatch details promptly once shipments go out.`,

    'accounts': `You have been added as Accounts on the Master Hub. You can:\n` +
                `  • Update Payment Received (col J) and Payment Date (col K)\n` +
                `  • All other columns are view-only for your role\n` +
                `  • Approve or reject credit limit increase requests\n\n` +
                `Please keep payment records up to date for accurate balance tracking.`,

    'sales': `You have been added as Sales on the Master Hub. You can:\n` +
             `  • Add new CRM Members in the CRM Members tab\n` +
             `  • Enter Client Name and assign CRM in the Client Credit tab\n` +
             `  • Edit all columns in the Client Credit tab\n` +
             `  • View overall order flow in the Master Hub\n\n` +
             `Note: Your entries in Client Credit will appear in the CRM Assigned Main sheet automatically.`,

    'sales manager': `You have been added as Sales Manager on the Master Hub. You can:\n` +
                     `  • Approve orders in the Approved By (col O) column\n` +
                     `  • Review and approve credit limit increases in the Client Credit tab\n` +
                     `  • Add new clients directly in the Client Credit tab (cols A, B, C, F, G, J)\n` +
                     `  • Edit Credit Limit (col C) and all other editable Client Credit columns\n` +
                     `  • Monitor all CRM activity and order statuses\n\n` +
                     `Orders requiring your approval will appear in col N (Approval Req).`,

    'admin': `You have been added as Admin on the Master Hub. You have full access to:\n` +
             `  • All sheets: Master Hub, CRM Members, Client Credit, Activity Log\n` +
             `  • Add or deactivate team members in the CRM Members tab\n` +
             `  • Run all menu functions (Export, Reminders, Setup)\n\n` +
             `Please do not share this file with anyone not listed in CRM Members.`,
  };

  const instructions = roleInstructions[roleLc] || `You have been added as ${role} to the UV Overseas CRM system.`;

  const subject = `[UV Overseas CRM] You've been added as ${role}`;
  const htmlBody = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;">
      <div style="background:#1A237E;padding:18px 24px;border-radius:6px 6px 0 0;">
        <span style="color:#fff;font-size:18px;font-weight:bold;">UV Overseas CRM System</span>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:24px;border-radius:0 0 6px 6px;">
        <p style="margin:0 0 12px;">${greeting},</p>
        <p>You have been added to the <strong>UV Overseas CRM system</strong> with the role of 
           <strong style="color:#1A237E;">${role}</strong>.</p>

        <div style="background:#F3F4F6;border-left:4px solid #1A237E;padding:12px 16px;margin:16px 0;border-radius:0 4px 4px 0;white-space:pre-line;font-size:13px;">
${instructions}
        </div>

        <p style="margin:16px 0 8px;"><strong>Your access link:</strong></p>
        <a href="${sharedFileUrl}" 
           style="display:inline-block;background:#1A237E;color:#fff;padding:10px 20px;border-radius:4px;text-decoration:none;font-weight:bold;">
          Open ${sharedFileName} ->
        </a>

        <p style="margin-top:20px;font-size:12px;color:#888;">
          This is an automated message from the UV Overseas CRM system.<br>
          If you believe you received this in error, please contact your Admin.
        </p>
      </div>
    </div>`;

  try {
    GmailApp.sendEmail(
      gmail,
      subject,
      `${greeting},\n\nYou have been added to UV Overseas CRM as ${role}.\n\n${instructions}\n\nAccess: ${sharedFileUrl}\n\n— UV Overseas CRM (Auto-generated)`,
      { htmlBody: htmlBody, name: 'UV Overseas CRM System' }
    );
    Logger.log('notifyAndShareMember_: Welcome email sent to ' + gmail);
  } catch (mailErr) {
    Logger.log('notifyAndShareMember_: Email failed for ' + gmail + ' — ' + mailErr.message);
  }
}



// ══════════════════════════════════════════════════════════════════
// CLEANUP HUB SHEETS FROM CRM SPOKE
// Removes any Hub-only sheets accidentally created in a CRM spoke
// by a previous run of setupEverything or setupMembersSheet.
//
// Hub-only sheets that should NOT exist in a CRM spoke:
//   - Order Ledger      (Hub has this; CRM has "Order Ledger CRM")
//   - Client Summary    (Hub only)
//   - Dashboard         (Hub only)
//   - Team Members      (Hub only)
//   - _WipeQueue        (Hub only)
//   - Credit List       (Hub only — CRM has its own credit handling)
//
// Safe to run multiple times. Will not delete if running in Hub.
// ══════════════════════════════════════════════════════════════════
function cleanupHubSheetsFromCRM() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const isHub = ss.getId() === MASTER_HUB_ID;

  if (isHub) {
    try {
      SpreadsheetApp.getUi().alert(
        'Not needed',
        'This function only applies to CRM Spoke files.\n' +
        'You are running it from the Admin Hub — nothing to clean up.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (_) {}
    return;
  }

  // Sheets that belong in Hub only — should not exist in CRM spoke
  const HUB_ONLY_SHEETS = [
    'Order Ledger',    // Hub has this; CRM has "Order Ledger CRM"
    'Client Summary',  // Hub only
    'Dashboard',       // Hub only
    'Team Members',    // Hub only — setupMembersSheet() creates this in Hub, not CRM
    '_WipeQueue',      // Hub only
    'Credit List',     // Hub only
  ];

  const removed = [];
  const skipped = [];
  const notFound = [];

  HUB_ONLY_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      notFound.push(sheetName);
      return;
    }
    // Safety check — don't delete if it has data (ask user)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      skipped.push(sheetName + ' (' + (lastRow - 1) + ' data rows — skipped for safety)');
      return;
    }
    try {
      ss.deleteSheet(sheet);
      removed.push(sheetName);
      Logger.log('cleanupHubSheetsFromCRM: deleted "' + sheetName + '"');
    } catch (e) {
      skipped.push(sheetName + ' (error: ' + e.message + ')');
    }
  });

  const summary =
    'File: ' + ss.getName() + '\n\n' +
    (removed.length  ? '✅ Removed:\n' + removed.map(s => '  - ' + s).join('\n') + '\n\n' : '') +
    (skipped.length  ? '⚠️ Skipped (has data — delete manually if needed):\n' + skipped.map(s => '  - ' + s).join('\n') + '\n\n' : '') +
    (notFound.length ? 'ℹ️ Not found (already clean):\n' + notFound.map(s => '  - ' + s).join('\n') : '');

  Logger.log('cleanupHubSheetsFromCRM:\n' + summary);

  try {
    SpreadsheetApp.getUi().alert(
      removed.length > 0 ? '✅ Cleanup Complete' : 'ℹ️ Nothing to clean',
      summary,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
}

// ══════════════════════════════════════════════════════════════════
// SETUP EVERYTHING — Single function to set up any new file
// Works for both Admin Hub and CRM Spoke automatically.
// Run this ONCE after pasting the code into any new Google Sheet.
//
// What it does:
//   Detects whether this is Hub or CRM spoke (by MASTER_HUB_ID)
//   Then runs ALL required setup steps for that file type.
// ══════════════════════════════════════════════════════════════════
function setupEverything() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const isHub   = ss.getId() === MASTER_HUB_ID;
  const steps   = [];
  const errors  = [];

  // Safe UI helper — works from both editor and menu click
  const showAlert = (title, msg) => {
    try { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
    catch (_) { Logger.log(title + ': ' + msg); }
  };

  const log = (msg) => {
    steps.push(msg);
    Logger.log('setupEverything: ' + msg);
  };

  try {
    showAlert(
      'Setup Everything',
      'Starting setup for: ' + ss.getName() + '\n' +
      'Detected as: ' + (isHub ? '✅ Admin Hub' : '✅ CRM Spoke') + '\n\n' +
      'This will take about 30-60 seconds. Click OK to begin.'
    );

    if (isHub) {
      // ════════════════════════════════
      // ADMIN HUB SETUP STEPS
      // ════════════════════════════════

      // STEP 1: Hub sheet structure
      try {
        setupHubSheet();
        log('✅ Hub sheet structure created');
      } catch (e) {
        log('ℹ️  Hub sheets: ' + e.message + ' (may already exist)');
      }

      // STEP 2: Team Members sheet (Hub only — opens MASTER_HUB_ID)
      try {
        setupMembersSheet();
        log('✅ Team Members sheet setup');
      } catch (e) {
        log('ℹ️  Team Members: ' + e.message + ' (may already exist)');
      }

      // STEP 3: All Hub triggers
      try {
        setupAllTriggers();
        log('✅ Hub triggers installed');
      } catch (e) {
        errors.push('Hub triggers: ' + e.message);
      }

      // STEP 4: Wipe queue for remote delete
      try {
        setupWipeQueue();
        log('✅ Wipe queue created (_WipeQueue)');
      } catch (e) {
        log('ℹ️  Wipe queue: ' + e.message + ' (may already exist)');
      }

      // STEP 5: License sync
      try {
        refreshLicenseStatus_();
        log('✅ License status synced');
      } catch (e) {
        errors.push('License sync: ' + e.message);
      }

      // STEP 6: Initial dashboard refresh
      try {
        menuRefreshDashboard_();
        log('✅ Dashboard refreshed');
      } catch (e) {
        log('ℹ️  Dashboard refresh: ' + e.message);
      }

    } else {
      // ════════════════════════════════════════════════════════
      // CRM SPOKE SETUP STEPS
      // STRICTLY ISOLATED — only creates sheets in THIS file.
      // Never opens MASTER_HUB_ID for sheet creation.
      // Hub sheets (Order Ledger, Client Summary, Dashboard,
      // Team Members, Credit List) are NOT created here.
      // ════════════════════════════════════════════════════════

      // STEP 1: CRM-only sheet structure
      // Creates: Order Ledger CRM, Client Dues (CRM)
      // Does NOT create: Order Ledger, Client Summary, Dashboard,
      //                  Team Members, Credit List (all Hub-only)
      try {
        setupCRMSheet();
        log('✅ CRM sheets created: Order Ledger CRM, Client Dues');
      } catch (e) {
        log('ℹ️  CRM sheets: ' + e.message + ' (may already exist)');
      }

      // STEP 2: CRM triggers only
      // setupAllTriggers() already detects CRM vs Hub via MASTER_HUB_ID
      try {
        setupAllTriggers();
        log('✅ CRM triggers installed: onCRMEdit, onCRMSubTabEdit, onOpen');
      } catch (e) {
        errors.push('CRM triggers: ' + e.message);
      }

      // STEP 3: Async Hub sync queue (_SyncQueue in THIS file only)
      try {
        setupSyncQueue_();
        log('✅ Sync queue created: _SyncQueue sheet + 1-min trigger');
      } catch (e) {
        log('ℹ️  Sync queue: ' + e.message + ' (may already exist)');
      }

      // STEP 4: Annual archive trigger (April 1st, THIS file only)
      try {
        setupArchiveTrigger_();
        log('✅ Archive trigger set: runs automatically every April 1st');
      } catch (e) {
        errors.push('Archive trigger: ' + e.message);
      }

      // STEP 5: License sync (reads License sheet, writes to DocumentProperties)
      try {
        refreshLicenseStatus_();
        log('✅ License status synced from License sheet');
      } catch (e) {
        errors.push('License sync: ' + e.message);
      }

      // STEP 6: Auto-cleanup — remove any Hub sheets accidentally created before
      try {
        const HUB_ONLY = [
          'Order Ledger', 'Client Summary', 'Dashboard',
          'Team Members', '_WipeQueue', 'Credit List'
        ];
        const ss2 = SpreadsheetApp.getActiveSpreadsheet();
        const cleaned = [];
        HUB_ONLY.forEach(name => {
          const s = ss2.getSheetByName(name);
          if (s && s.getLastRow() <= 1) { // only remove if empty/header only
            try { ss2.deleteSheet(s); cleaned.push(name); } catch (_) {}
          }
        });
        if (cleaned.length > 0) {
          log('✅ Cleaned up Hub-only sheets: ' + cleaned.join(', '));
        } else {
          log('ℹ️  No Hub sheets to clean up');
        }
      } catch (e) {
        log('ℹ️  Cleanup check: ' + e.message);
      }
    }

    // ── Done — show summary ────────────────────────────────────────
    const statusIcon = errors.length === 0 ? '✅' : '⚠️';
    showAlert(
      statusIcon + ' Setup Complete — ' + ss.getName(),
      'File type: ' + (isHub ? 'Admin Hub' : 'CRM Spoke') + '\n\n' +
      'Steps completed:\n' + steps.map(s => '  ' + s).join('\n') +
      (errors.length
        ? '\n\nErrors (manual fix needed):\n' + errors.map(e => '  ❌ ' + e).join('\n')
        : '\n\n🎉 All steps completed successfully!') +
      '\n\n' +
      (isHub
        ? 'Next: Add CRM spoke File IDs to the Team Members sheet.'
        : 'Next: Make sure your email is active in the Hub Team Members sheet.'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (fatal) {
    try {
      SpreadsheetApp.getUi().alert(
        '❌ Setup Failed',
        'Fatal error: ' + fatal.message + '\n\n' +
        'Steps completed before error:\n' +
        steps.map(s => '  ' + s).join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (_) {
      Logger.log('setupEverything FATAL: ' + fatal.message);
    }
  }
}

function setupAllTriggers() {
  const currentSS   = SpreadsheetApp.getActive();
  const currentSSId = currentSS.getId();
  const isHub       = currentSSId === MASTER_HUB_ID;

  // Record which account is installing these triggers
  const installingUser = Session.getEffectiveUser().getEmail();
  Logger.log('setupAllTriggers: installing as ' + installingUser + ' | isHub=' + isHub);

  // ── SCOPED DELETE: only remove triggers bound to THIS spreadsheet ─────
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getTriggerSourceId() === currentSSId) {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (isHub) {
    ScriptApp.newTrigger('onMasterHubEdit').forSpreadsheet(currentSS).onEdit().create();
    ScriptApp.newTrigger('onCreditRequestEdit').forSpreadsheet(currentSS).onEdit().create();
    ScriptApp.newTrigger('onMembersSheetEdit').forSpreadsheet(currentSS).onEdit().create();
    ScriptApp.newTrigger('addCustomMenu').forSpreadsheet(currentSS).onOpen().create();
    setupLicenseRefreshTrigger_(); // runs every 5 min to sync license status locally
    refreshLicenseStatus_();       // sync immediately on setup
    Logger.log('setupAllTriggers: Hub triggers installed.');
    alertIfUi_(
      '✅ Hub triggers installed:\n' +
      '• onMasterHubEdit\n• onCreditRequestEdit\n• onMembersSheetEdit\n• addCustomMenu (onOpen)\n\n' +
      '⚠️  Installed by: ' + installingUser + '\n' +
      'Make sure this is the account with Hub Editor access.'
    );

  } else {
    ScriptApp.newTrigger('onCRMEdit').forSpreadsheet(currentSS).onEdit().create();
    ScriptApp.newTrigger('onCRMSubTabEdit').forSpreadsheet(currentSS).onEdit().create();
    ScriptApp.newTrigger('addCustomMenu').forSpreadsheet(currentSS).onOpen().create();
    setupLicenseRefreshTrigger_(); // runs every 5 min to sync license status locally
    refreshLicenseStatus_();       // sync immediately on setup
    setupSyncQueue_();             // FIX 4: async Hub sync queue + 1-min trigger
    Logger.log('setupAllTriggers: CRM spoke triggers installed for ' + currentSS.getName());
    alertIfUi_(
      '✅ CRM Spoke triggers installed:\n' +
      '• onCRMEdit\n• onCRMSubTabEdit\n• addCustomMenu (onOpen)\n\n' +
      'Spreadsheet: ' + currentSS.getName() + '\n\n' +
      '⚠️  Installed by: ' + installingUser + '\n' +
      'This account MUST have Editor access to the Master Hub (' + MASTER_HUB_ID + ').\n\n' +
      'Run CRM Tools > Sync & Repair > Diagnose triggers to verify.'
    );
  }
}

function onCRMEdit(e) {
  invalidateAllSheetCaches_(); // FIX 5: fresh read per trigger
  const _licEmail = (e && e.user) ? e.user.email : '';
  if (!checkLicense(_licEmail)) return;
  if (!e || !e.range) {
    Logger.log('onCRMEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  try {
    const sheet = e.range.getSheet();
    const row   = e.range.getRow();
    const col   = e.range.getColumn();

    if (row <= 1) return;
    // BUG FIX v15: sheet was renamed from 'Main' to 'Order Ledger CRM' in v14.
    // Check both names so the guard works regardless of which name is used.
    const isMainSheet = sheet.getName() === MAIN_SHEET_NAME ||
                        sheet.getName() === CRM_ORDER_LEDGER_SHEET;
    if (!isMainSheet) return;

    // ── v15.1: Lockout guard — same as onCRMSubTabEdit, but for direct
    // edits to the Order Ledger CRM sheet itself (in case CRM edits here
    // instead of a client sub-tab). See onCRMSubTabEdit for full rationale.
    // v15.3: Due Date removed — CRM can edit it again (rate-limited to once
    // per 7 days, enforced in the sub-tab's PATH 2, not blocked outright).
    const LOCKED_FOR_CRM_MAIN_COLS = new Set([
      CRM_COL_INV_NO, CRM_COL_DATE, CRM_COL_AWB_NO, CRM_COL_DISP_DATE,
      CRM_COL_PAY_RECEIVED, CRM_COL_PAY_DATE,
    ]);
    if (LOCKED_FOR_CRM_MAIN_COLS.has(col)) {
      const guardEmail = (e.user && e.user.email) || Session.getActiveUser().getEmail() || 'unknown';
      try { e.range.clearContent(); SpreadsheetApp.flush(); } catch (_) {}
      try {
        logActivity_(SpreadsheetApp.getActiveSpreadsheet(), 'CRM EDIT BLOCKED',
          'sheet:' + sheet.getName() + ' row:' + row + ' col:' + col +
          ' — managed by Logistics/Accounts via Hub', guardEmail);
      } catch (_) {}
      alertIfUi_(
        'This field is entered by Logistics/Accounts on the Hub and updates here automatically.\n' +
        'Please contact Logistics or Accounts directly to change it.'
      );
      return;
    }

    // ── v10.0 FIX: Loop-guard via PropertiesService ────────
    // pushFinalInvoiceToMainAndHub_ sets this flag before writing
    // to CRM Main col G. We read and immediately clear it so that
    // only ONE onCRMEdit call per programmatic write is suppressed.
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_SKIP_NEXT_EDIT) === 'true') {
      props.deleteProperty(PROP_SKIP_NEXT_EDIT);
      Logger.log('onCRMEdit: skipping programmatic write to col ' + col + ' (loop-guard)');
      return;
    }

    // Guard: skip columns written by internal scripts to prevent loops.
    // CRM_COL_UUID (col 1): UUID writes must never re-trigger sync.
    // NOTE: CRM_COL_INV_VALUE intentionally REMOVED from this list in v10.0.
    // It is now guarded by PROP_SKIP_NEXT_EDIT above.
    const SUB_DRIVEN_COLS = [
      CRM_COL_UUID,           // UUID writes must never re-trigger sync
    ];
    if (SUB_DRIVEN_COLS.includes(col)) {
      Logger.log('onCRMEdit: skipping col ' + col + ' — internal/driven column');
      return;

    // ── GLOBAL DUPLICATE CHECK: PI No and Inv No on CRM Main sheet ──────
    // Catches users who type PI No or Inv No directly into the CRM Order
    // Ledger main sheet (col D or col E) instead of via the client sub-tab.
    // Checks across all client sub-tabs in this CRM + Hub Order Ledger
    // (which covers all other CRM spokes too).
    } else if ((col === CRM_COL_PI_NO || col === CRM_COL_INV_NO) && e.range.getValue()) {
      const cellStr = e.range.getValue().toString().trim();
      if (cellStr) {
        const isPI = (col === CRM_COL_PI_NO);
        const subCol = isPI ? SUB_COL_PI_NO : SUB_COL_INV_NO;
        const mainSh = sheet; // this IS the main sheet (checked above)
        let dupClient = '';

        // 1. Check all other rows in CRM Order Ledger main sheet
        const allMainData = mainSh.getDataRange().getValues();
        for (let i = 1; i < allMainData.length; i++) {
          if (i + 1 === row) continue; // skip self
          const chkCol = isPI ? CRM_COL_PI_NO : CRM_COL_INV_NO;
          if ((allMainData[i][chkCol - 1] || '').toString().trim() === cellStr) {
            dupClient = (allMainData[i][CRM_COL_CUSTOMER - 1] || 'another client').toString().trim();
            break;
          }
        }

        // 2. Check all client sub-tabs in this CRM spoke
        const sysSetMain = new Set([
          ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
          CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
          CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME,
        ]);
        if (!dupClient) {
          const activeSheetName = sheet.getName();
          SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(s => {
            if (dupClient) return;
            if (sysSetMain.has(s.getName())) return;
            const sData = s.getDataRange().getValues();
            for (let i = 1; i < sData.length; i++) {
              if ((sData[i][subCol - 1] || '').toString().trim() === cellStr) {
                dupClient = s.getName();
                break;
              }
            }
          });
        }

        // 3. Check Hub Order Ledger (covers all other CRM spokes)
        if (!dupClient) {
          try {
            const hubSS2 = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubOL2 = hubSS2.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubOL2) {
              const hd2 = hubOL2.getDataRange().getValues();
              for (let i = 1; i < hd2.length; i++) {
                const hVal  = (hd2[i][(isPI ? OL_COL_PI_NO : OL_COL_INV_NO) - 1] || '').toString().trim();
                const hName = (hd2[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim();
                if (hVal === cellStr) {
                  dupClient = hName || 'another client (Hub)';
                  break;
                }
              }
            }
          } catch (he) {
            Logger.log('onCRMEdit dup check Hub error: ' + he.message);
          }
        }

        if (dupClient) {
          e.range.clearContent();
          SpreadsheetApp.flush();
          const label = isPI ? 'PI No' : 'Inv No';
          alertIfUi_(
            'DUPLICATE ' + label + ' DETECTED\n\n' +
            '"' + cellStr + '" is already assigned to client:\n  ' + dupClient + '\n\n' +
            label + ' must be unique across ALL clients and CRM files.\n' +
            'The entry has been cleared.'
          );
          const userEmail2 = Session.getActiveUser().getEmail();
          logActivity_(sheet.getParent(), 'DUP ' + label + ' BLOCKED (MAIN)',
            label + ':' + cellStr + ' | conflict with:' + dupClient + ' | row:' + row, userEmail2);
          return;
        }
      }
    }

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getCRMMainSheet_(ss);
    const userEmail = Session.getActiveUser().getEmail();

    // ── LOCK FIRST ────────────────────────────────────────────
    // Acquire the script lock BEFORE reading UUID, generating a
    // new one, or creating the client tab. This prevents two
    // concurrent trigger executions (e.g. the user typing the
    // name AND the HYPERLINK write from createClientTab both
    // firing onCRMEdit simultaneously) from each generating a
    // different UUID, both appending to Hub, and producing two
    // CRM Main rows with mismatched UUIDs.
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000); // FIX v14.3: increased from 15s
    } catch (_) {
      Logger.log('onCRMEdit: could not acquire lock — aborting row ' + row);
      return;
    }

    try {
      // ── Re-read sheet state AFTER acquiring lock ───────────
      // The state may have changed between the edit event and now
      // (another execution may have already written UUID/HYPERLINK).
      const customerFormula = mainSheet.getRange(row, CRM_COL_CUSTOMER).getFormula();
      const alreadyLinked   = customerFormula.toUpperCase().startsWith('=HYPERLINK');

      // Use getDisplayValue when HYPERLINK is already set (getValue
      // returns raw formula text like =HYPERLINK("url","name")).
      const customerName = alreadyLinked
        ? mainSheet.getRange(row, CRM_COL_CUSTOMER).getDisplayValue()
        : (mainSheet.getRange(row, CRM_COL_CUSTOMER).getValue() || '').toString().trim();

      // ── Generate or read UUID (under lock — safe) ──────────
      // CROSS-VM PROPAGATION RULE:
      // When alreadyLinked=TRUE this execution was triggered by the
      // HYPERLINK write from createClientTab() (Hub VM, separate process).
      // That execution already wrote UUID to col A before the HYPERLINK.
      // We MUST NOT generate a new UUID here — the flush() in that Hub
      // execution is NOT guaranteed to be visible to this Spoke VM yet.
      // Strategy: sleep 1 s and retry once, then fall through to the
      // name-based Hub match which will adopt + stamp the correct UUID.
      // Never call generateUUID_() when alreadyLinked=TRUE.
      let uuid = mainSheet.getRange(row, CRM_COL_UUID).getValue();
      if (!uuid && alreadyLinked) {
        Utilities.sleep(1000);
        uuid = mainSheet.getRange(row, CRM_COL_UUID).getValue();
        Logger.log('onCRMEdit: alreadyLinked=TRUE uuid after retry: "' + uuid + '" row=' + row);
      } else if (!uuid) {
        uuid = generateUUID_();
        mainSheet.getRange(row, CRM_COL_UUID).setValue(uuid);
        mainSheet.hideColumns(CRM_COL_UUID);
        SpreadsheetApp.flush();
      }

      // ── Open Hub UNDER LOCK to check for existing row ──────
      const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
      const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
      if (!hubSheet) { Logger.log('onCRMEdit: Hub sheet not found'); return; }
      const hubData  = hubSheet.getDataRange().getValues();

      // ── UUID match in Hub ───────────────────────────────────
      let matchRow = -1;
      for (let i = 1; i < hubData.length; i++) {
        if (hubData[i][HUB_COL_UUID - 1] === uuid) { matchRow = i + 1; break; }
      }

      // ── Name-based match: adopt Hub UUID if client pre-exists ─
      // When Sales/onCreditRequestEdit already created a Hub row for
      // this client, we must adopt the Hub's UUID so we update in
      // place rather than appending a duplicate row.
      // BUG FIX v14.1: only adopt a Hub row by name when EITHER:
      //   (a) The CRM row has no PI No (registration-only edit), OR
      //   (b) The Hub row has the SAME PI No as the CRM row (order match).
      // This prevents the order row from being merged into the registration
      // placeholder row — each order must have its own Hub row.
      // Read PI No directly from sheet — rowData not declared yet at this point
      const crmPiNoForMatch = mainSheet.getRange(row, CRM_COL_PI_NO).getValue().toString().trim();
      if (matchRow === -1 && customerName) {
        for (let i = 1; i < hubData.length; i++) {
          if ((hubData[i][HUB_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase() ===
              customerName.toLowerCase()) {
            const existingHubUuid = (hubData[i][HUB_COL_UUID - 1] || '').toString().trim();
            if (!existingHubUuid) continue; // Hub row has no UUID — skip
            // PI No guard: if CRM row has a PI No, only adopt a Hub row with
            // the same PI No. Never adopt the registration-placeholder row.
            const hubPiNo = (hubData[i][HUB_COL_PI_NO - 1] || '').toString().trim();
            if (crmPiNoForMatch && hubPiNo !== crmPiNoForMatch) {
              Logger.log('onCRMEdit: skipping name-match for Hub row ' + (i+1) +
                ' — CRM PI No "' + crmPiNoForMatch + '" != Hub PI No "' + hubPiNo + '"');
              continue; // do NOT adopt registration placeholder as this order row
            }
            if (existingHubUuid) {
              uuid = existingHubUuid;
              matchRow = i + 1;
              Logger.log('onCRMEdit: adopted Hub UUID for ' + customerName + ' PI:' + crmPiNoForMatch + ' -> ' + uuid);

              // ── Deduplicate CRM Main rows ───────────────────
              // If another CRM Main row already holds this Hub UUID
              // (written by onCreditRequestEdit), the current row is
              // a duplicate typed by the CRM owner. Merge and delete.
              // Cross-VM propagation: the Hub-written UUID may not be
              // visible yet — retry up to 2× with sleep before giving up.
              let primaryCrmRow = -1;
              for (let attempt = 0; attempt < 3 && primaryCrmRow === -1; attempt++) {
                if (attempt > 0) Utilities.sleep(1500); // wait for cross-VM flush
                const crmAllData2    = mainSheet.getDataRange().getValues();
                const crmAllDisplay2 = mainSheet.getDataRange().getDisplayValues();
                for (let j = 1; j < crmAllData2.length; j++) {
                  if (j + 1 === row) continue;
                  const jUuid = (crmAllData2[j][CRM_COL_UUID - 1] || '').toString().trim();
                  const jName = (crmAllDisplay2[j][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
                  if (jUuid === existingHubUuid || jName === customerName.toLowerCase()) {
                    primaryCrmRow = j + 1;
                    break;
                  }
                }
              }

              if (primaryCrmRow !== -1) {
                // Merge data from duplicate row into primary, then delete duplicate
                mainSheet.getRange(primaryCrmRow, CRM_COL_UUID).setValue(uuid);
                const dupData = mainSheet.getRange(row, 1, 1, 16).getValues()[0];
                const priData = mainSheet.getRange(primaryCrmRow, 1, 1, 16).getValues()[0];
                for (let c = CRM_COL_PI_NO; c <= 16; c++) {
                  if (!priData[c - 1] && dupData[c - 1]) {
                    mainSheet.getRange(primaryCrmRow, c).setValue(dupData[c - 1]);
                  }
                }
                mainSheet.deleteRow(row);
                logActivity_(ss, 'DUPLICATE ROW REMOVED',
                  'Merged dup row ' + row + ' into primary row ' + primaryCrmRow +
                  ' for client: ' + customerName, userEmail);
                return; // primary row has correct UUID & data; done
              } else {
                // This IS the primary row — stamp it with the adopted Hub UUID
                mainSheet.getRange(row, CRM_COL_UUID).setValue(uuid);
                SpreadsheetApp.flush();
              }
            }
            break;
          }
        }
      }

      // ── Create client sub-tab (under lock, safe) ────────────
      // Only when user typed a customer name for the first time
      // (col B edit AND no HYPERLINK yet). The HYPERLINK write
      // inside createClientTab will fire another onCRMEdit (col=2,
      // alreadyLinked=TRUE) — that execution will wait for THIS lock
      // to release, then find the UUID and Hub row already present
      // and do a simple UPDATE. No duplicate row or tab is created.
      if (col === CRM_COL_CUSTOMER && customerName && !alreadyLinked) {
        createClientTab(ss, customerName, row, uuid);
        SpreadsheetApp.flush();
        // Re-read uuid in case createClientTab found a different one
        const uuidAfter = mainSheet.getRange(row, CRM_COL_UUID).getValue();
        if (uuidAfter && uuidAfter !== uuid) {
          uuid = uuidAfter;
          // Re-search Hub for updated UUID
          matchRow = -1;
          const hubData2 = hubSheet.getDataRange().getValues();
          for (let i = 1; i < hubData2.length; i++) {
            if (hubData2[i][HUB_COL_UUID - 1] === uuid) { matchRow = i + 1; break; }
          }
        }
      }

      // ── Auto-fill PI No, Inv No & Inv Date into Customer Tab ──────────
      const piNo  = mainSheet.getRange(row, CRM_COL_PI_NO).getValue();
      const invNo = mainSheet.getRange(row, CRM_COL_INV_NO).getValue();
      const invDate = mainSheet.getRange(row, CRM_COL_DATE).getValue();
      const cName = mainSheet.getRange(row, CRM_COL_CUSTOMER).getDisplayValue();
      if (piNo && cName) {
        const clientSheet = ss.getSheetByName(cName);
        if (clientSheet) {
          const clientData = clientSheet.getDataRange().getValues();
          let subRow = -1;
          for (let i = 1; i < clientData.length; i++) {
            if (clientData[i][SUB_COL_PI_NO - 1] === piNo) { subRow = i + 1; break; }
          }
          if (subRow === -1) {
            for (let i = 1; i < clientData.length; i++) {
              if (!clientData[i][SUB_COL_PI_NO - 1]) { subRow = i + 1; break; }
            }
            if (subRow === -1) subRow = clientSheet.getLastRow() + 1;
          }
          clientSheet.getRange(subRow, SUB_COL_PI_NO).setValue(piNo);
          if (invNo) clientSheet.getRange(subRow, SUB_COL_INV_NO).setValue(invNo);
          if (invDate) {
            const invDateCell = clientSheet.getRange(subRow, SUB_COL_INV_DATE);
            invDateCell.setValue(invDate);
            invDateCell.setNumberFormat('dd/mm/yyyy');
          }
          if (col === CRM_COL_INV_VALUE) {
            clientSheet.getRange(subRow, SUB_COL_ITEM_VALUE_USD).setValue(e.range.getValue());
          }
          logActivity_(ss, 'MAIN->SUBTAB', 'PI:' + piNo + ' row:' + subRow, userEmail);
        }
      }

      // ── Set CRM row formulas ────────────────────────────────
      setCRMFormulas_(mainSheet, row);
      SpreadsheetApp.flush();

      // ── Read full CRM row for Hub sync ──────────────────────
      const rowData = mainSheet.getRange(row, 1, 1, 16).getValues()[0];
      // Use display value for col B — getValue returns raw HYPERLINK formula text
      const displayName = mainSheet.getRange(row, CRM_COL_CUSTOMER).getDisplayValue();
      const rowDataForHub = rowData.slice();
      rowDataForHub[CRM_COL_CUSTOMER - 1] = displayName;

      // ── Write to Hub ────────────────────────────────────────
      // BUG FIX v15: Skip Hub write for ANY col=CRM_COL_CUSTOMER + alreadyLinked trigger.
      // When createClientTab writes =HYPERLINK(...) to col B, it fires onCRMEdit with
      // col=CRM_COL_CUSTOMER and alreadyLinked=TRUE. This is purely a registration
      // write — no order data should flow to Hub Order Ledger from this trigger.
      // Previously the guard also required !crmPiNoForMatch, which failed when an
      // Order Ledger CRM row already had a PI No from a prior edit.
      const isRegistrationOnlyRow = (col === CRM_COL_CUSTOMER && alreadyLinked);
      if (isRegistrationOnlyRow) {
        Logger.log('onCRMEdit: HYPERLINK registration trigger — skipping Hub Order Ledger write. row=' + row);
        autoAddClientToCredit_(hubSS, displayName, rowData[CRM_COL_CRM_OWNER - 1]);
        SpreadsheetApp.flush();
        return;
      }

      // Re-read Hub data after all the above writes to get latest state
      const hubDataFresh = hubSheet.getDataRange().getValues();
      if (matchRow === -1) {
        // Final safety check before appending — prevent duplicate if another exec ran.
        // BUG FIX v14.1: name-match must also respect PI No so we don't collapse
        // an order row onto the registration-placeholder row.
        // Priority: (1) UUID match, (2) PI No + name match, (3) name-only match
        // for registration rows (no PI No on either side).
        for (let i = 1; i < hubDataFresh.length; i++) {
          const hUuid  = (hubDataFresh[i][HUB_COL_UUID     - 1] || '').toString().trim();
          const hName  = (hubDataFresh[i][HUB_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
          const hPiNo  = (hubDataFresh[i][HUB_COL_PI_NO   - 1] || '').toString().trim();
          // (1) UUID exact match — always safe
          if (hUuid && hUuid === uuid) {
            matchRow = i + 1;
            mainSheet.getRange(row, CRM_COL_UUID).setValue(uuid);
            break;
          }
          // (2) Name + PI No match — same order, adopt
          if (displayName && hName === displayName.toLowerCase() &&
              crmPiNoForMatch && hPiNo === crmPiNoForMatch) {
            matchRow = i + 1;
            uuid = (hUuid || uuid).toString().trim();
            mainSheet.getRange(row, CRM_COL_UUID).setValue(uuid);
            break;
          }
          // (3) Name-only match — only safe when CRM row has NO PI No (registration edit)
          if (!crmPiNoForMatch && displayName && hName === displayName.toLowerCase()) {
            matchRow = i + 1;
            uuid = (hUuid || uuid).toString().trim();
            mainSheet.getRange(row, CRM_COL_UUID).setValue(uuid);
            break;
          }
        }
      }

      if (matchRow === -1) {
        // GUARD: if alreadyLinked=TRUE and uuid is still empty, the Hub row was
        // created by onCreditRequestEdit but UUID hasn't propagated to this Spoke VM.
        // Abort — do not append a phantom Hub row with a blank UUID.
        // The next edit on this CRM row will retry and find the Hub row by name.
        if (!uuid && alreadyLinked) {
          Logger.log('onCRMEdit: alreadyLinked=TRUE uuid empty after name-match — aborting to prevent phantom Hub row. row=' + row);
          return;
        }

        const newRow   = buildHubRowFromCRM_(uuid, rowDataForHub);
        const appendAt = hubSheet.getLastRow() + 1;
        hubSheet.getRange(appendAt, 1, 1, newRow.length).setValues([newRow]);
        setHubFormulas_(hubSheet, appendAt);
        matchRow = appendAt;
        logActivity_(hubSS, 'APPEND',   'UUID:' + uuid + ' | ' + displayName, userEmail);
        logActivity_(ss,    'SYNC->HUB', 'UUID:' + uuid + ' (new row)',          userEmail);
      } else {
        hubSheet.getRange(matchRow, HUB_COL_UUID).setValue(uuid);
        hubSheet.getRange(matchRow, HUB_COL_CUSTOMER).setValue(displayName);
        hubSheet.getRange(matchRow, HUB_COL_CRM_OWNER).setValue(rowData[CRM_COL_CRM_OWNER - 1]);
        hubSheet.getRange(matchRow, HUB_COL_PI_NO).setValue(rowData[CRM_COL_PI_NO - 1]);
        hubSheet.getRange(matchRow, HUB_COL_INV_NO).setValue(rowData[CRM_COL_INV_NO - 1]);
        hubSheet.getRange(matchRow, HUB_COL_DATE).setValue(rowData[CRM_COL_DATE - 1]);
        hubSheet.getRange(matchRow, HUB_COL_INV_VALUE).setValue(rowData[CRM_COL_INV_VALUE - 1]);
        hubSheet.getRange(matchRow, HUB_COL_DUE_DATE).setValue(rowData[CRM_COL_DUE_DATE - 1]);
        // ── Payment columns: sync CRM Main -> Hub ────────────────
        // These are written here when the CRM user types them directly
        // in CRM Main, or when onCRMSubTabEdit pushes them to Main first.
        // onMasterHubEdit will NOT loop back: it only fires when an
        // Accounts user edits Hub J/K in the spreadsheet UI, not when
        // this script writes them programmatically.
        // BUG FIX: use != null check instead of !== '' so numeric 0
        // values (valid payment amounts) are not accidentally skipped.
        const payReceived = rowData[CRM_COL_PAY_RECEIVED - 1];
        const payDate     = rowData[CRM_COL_PAY_DATE     - 1];
        if (payReceived !== '' && payReceived != null) {
          hubSheet.getRange(matchRow, HUB_COL_PAY_RECEIVED).setValue(payReceived);
        }
        if (payDate !== '' && payDate != null) {
          hubSheet.getRange(matchRow, HUB_COL_PAY_DATE).setValue(payDate);
        }
        hubSheet.getRange(matchRow, HUB_COL_APPROVAL_REQ).setValue(rowData[CRM_COL_APPROVAL_REQ - 1]);
        hubSheet.getRange(matchRow, HUB_COL_APPROVED_BY).setValue(rowData[CRM_COL_APPROVED_BY - 1]);
        if (rowData[CRM_COL_REMARKS - 1] !== '') {
          hubSheet.getRange(matchRow, HUB_COL_REMARKS).setValue(rowData[CRM_COL_REMARKS - 1]);
        }
        logActivity_(hubSS, 'UPDATE',   'UUID:' + uuid + ' | col:' + col, userEmail);
        logActivity_(ss,    'SYNC->HUB', 'UUID:' + uuid + ' (update)',     userEmail);
      }

      // ── Credit check ────────────────────────────────────────
      if (col === CRM_COL_INV_VALUE || col === CRM_COL_CUSTOMER) {
        const invValue = rowData[CRM_COL_INV_VALUE - 1];
        const crmOwner = rowData[CRM_COL_CRM_OWNER - 1];
        if (invValue && displayName) {
          checkCreditAndFlag_(hubSS, hubSheet, matchRow, uuid, invValue, displayName, crmOwner, mainSheet, row, userEmail);
        }
      }

      // FIX: When CRM updates Payment Received directly in Main, refresh
      // credit totals so Client Credit and col N stay in sync.
      if (col === CRM_COL_PAY_RECEIVED || col === CRM_COL_PAY_DATE) {
        refreshCreditTotals_(userEmail);
      }

      // ── Auto-add to Credit sheet ────────────────────────────
      // Use display value (not getValue) to avoid HYPERLINK formula text
      // being written as the client name, which would create a duplicate
      // Credit row and trigger a second onCreditRequestEdit cycle.
      autoAddClientToCredit_(hubSS, displayName, rowData[CRM_COL_CRM_OWNER-1]);

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    Logger.log('ERROR in onCRMEdit: ' + err.message);
  }
}

function extractEmail_(value) {
  if (!value) return '';
  const match = value.toString().match(/\(([^)]+@[^)]+)\)\s*$/);
  return match ? match[1].trim().toLowerCase() : value.toString().trim().toLowerCase();
}

function onCreditRequestEdit(e) {
  if (!e || !e.range) {
    Logger.log('onCreditRequestEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  try {
    const sheet = e.range.getSheet();
    const row   = e.range.getRow();
    const col   = e.range.getColumn();

    if (row <= 1) return;
    if (sheet.getName() !== CREDIT_SHEET_NAME) return;

    // ── NEW: When Sales Manager, Sales, or Admin enters Client Name (col A) or CRM Owner (col B) ──
    // Sync the row to the Master Hub sheet AND create the client sub-tab
    // in the correct CRM spoke file with a hyperlink back in CRM Main.
    if (col === CC_COL_CLIENT_NAME || col === CC_COL_CRM_OWNER) {
      // ── Acquire lock for the entire Hub + Spoke write ──────────
      // Without a lock, two rapid edits (Sales types Name then immediately
      // assigns CRM Owner) each fire this handler concurrently. Both read
      // an empty col K UUID, both generate a new UUID, both append Hub rows
      // and Spoke rows -> 2 or 3 duplicate rows with different UUIDs.
      const lock = LockService.getScriptLock();
      try { lock.waitLock(30000); } catch (_) { // FIX v14.3: increased from 20s
        Logger.log('onCreditRequestEdit: could not acquire lock — aborting row ' + row);
        return;
      }

      try {
        // Re-read the row AFTER acquiring the lock — another execution may
        // have already written the UUID and CRM owner while we waited.
        const rowData     = sheet.getRange(row, 1, 1, CC_TOTAL_COLS + 1).getValues()[0];
        const clientName  = (rowData[CC_COL_CLIENT_NAME - 1] || '').toString().trim();
        const crmOwnerRaw = (rowData[CC_COL_CRM_OWNER   - 1] || '').toString().trim();
        if (!clientName) return;

        // FIX: extract bare email from "Name (email@gmail.com)" dropdown format.
        // crmOwnerRaw   — display string kept for human-readable Hub writes and logs.
        // crmOwnerEmail — bare email used for all config / file-ID lookups.
        const crmOwnerEmail = extractEmail_(crmOwnerRaw);
        const crmOwner      = crmOwnerRaw;

        const hubSS    = e.source;
        const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
        if (!hubSheet) return;

        const userEmail = Session.getActiveUser().getEmail();

        // CC_COL_UUID = col H (8) — hidden scratch column (v15.0: was col K=11, now col H=8)
        const CC_COL_UUID = 8; // CC_COL_UUID constant
        let existingUuid = (rowData[CC_COL_UUID - 1] || '').toString().trim();

        const hubData = hubSheet.getDataRange().getValues();
        let hubRow = -1;

        // 1. Match by UUID (most reliable — survives name changes)
        if (existingUuid) {
          for (let i = 1; i < hubData.length; i++) {
            if ((hubData[i][HUB_COL_UUID - 1] || '').toString().trim() === existingUuid) {
              hubRow = i + 1;
              break;
            }
          }
        }

        // 2. Fall back to name match (first time, before UUID is stamped)
        if (hubRow === -1) {
          for (let i = 1; i < hubData.length; i++) {
            if ((hubData[i][HUB_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase() ===
                clientName.toLowerCase()) {
              hubRow = i + 1;
              existingUuid = (hubData[i][HUB_COL_UUID - 1] || '').toString().trim();
              break;
            }
          }
        }

        // ── FIX v14.2: Do NOT write a blank placeholder row to Hub Order Ledger ──
        // Order Ledger rows must only be created when a real order (PI No) is entered
        // in the CRM sub-tab. At client-assignment time, only a UUID needs to be
        // stored in Credit List col H so the CRM Spoke can inherit it on first order.
        // If a real Hub Order Ledger row already exists (hubRow !== -1),
        // just update the CRM owner — never create a new blank row.
        if (hubRow === -1) {
          // No real order row yet — generate UUID for tracking but do NOT append to Order Ledger
          const uuid = generateUUID_();
          existingUuid = uuid;
          // hubRow stays -1 intentionally — no blank row written to Order Ledger
          logActivity_(hubSS, 'CREDIT->UUID GENERATED (no Hub row yet)',
            'Client:' + clientName + ' | CRM:' + crmOwner + ' | UUID:' + uuid, userEmail);
        } else {
          // Real Hub Order Ledger row exists — update CRM owner if we have one
          if (crmOwner) hubSheet.getRange(hubRow, HUB_COL_CRM_OWNER).setValue(crmOwner);
          logActivity_(hubSS, 'CREDIT->HUB UPDATE (Sales/SalesMgr/Admin)',
            'Client:' + clientName + ' | CRM:' + crmOwner + ' | UUID:' + existingUuid, userEmail);
        }

        // Always stamp UUID back to Credit sheet col K so next edit finds it
        sheet.getRange(row, CC_COL_UUID).setValue(existingUuid);

        // ── Resolve CRM owner email for Spoke lookup ───────────
        // If the dropdown was blank when col A was saved, pull the owner
        // from the Hub row we just upserted.
        const resolvedOwnerRaw = crmOwnerRaw ||
          (hubData[hubRow - 1] ? (hubData[hubRow - 1][HUB_COL_CRM_OWNER - 1] || '').toString().trim() : '');
        const resolvedOwnerEmail = extractEmail_(resolvedOwnerRaw);

        // ── Write to CRM Spoke (only when we have owner + UUID) ─
        if (!resolvedOwnerEmail || !existingUuid) {
          // No CRM owner assigned yet — Hub row is ready; Spoke write will
          // happen when Sales later assigns the CRM Owner (col B edit).
          Logger.log('onCreditRequestEdit: no CRM owner yet for "' + clientName
            + '" — Hub row written, Spoke write deferred. UUID=' + existingUuid);
          return;
        }

        const config    = getCrmConfig_();
        const crmFileId = config.crmFileIds[resolvedOwnerEmail];
        if (!crmFileId) {
          Logger.log('onCreditRequestEdit: no CRM file ID for email: ' + resolvedOwnerEmail
            + ' (raw: "' + resolvedOwnerRaw + '")');
          return;
        }

        try {
          const crmSS   = SpreadsheetApp.openById(crmFileId);
          const crmMain = getCRMMainSheet_(crmSS);
          if (!crmMain) return;

          // ── DEDUPLICATE Spoke rows ────────────────────────────
          // Collect ALL rows that match by UUID or client name, keep the
          // first one as the authoritative write slot, delete the rest
          // bottom-up so row indices stay valid.
          const crmValues  = crmMain.getDataRange().getValues();
          const crmDisplay = crmMain.getDataRange().getDisplayValues();
          const matchedRows = [];
          for (let i = 1; i < crmValues.length; i++) {
            const rUuid = (crmValues[i][CRM_COL_UUID     - 1] || '').toString().trim();
            const rName = (crmDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
            if (rUuid === existingUuid || rName === clientName.toLowerCase()) {
              matchedRows.push(i + 1); // 1-based sheet row
            }
          }

          let crmMainRow;
          if (matchedRows.length === 0) {
            crmMainRow = crmMain.getLastRow() + 1; // append fresh row
          } else {
            crmMainRow = matchedRows[0]; // authoritative slot
            // Delete duplicates bottom-up (indices > 0 only, never the primary)
            if (matchedRows.length > 1) {
              for (let d = matchedRows.length - 1; d >= 1; d--) {
                crmMain.deleteRow(matchedRows[d]);
                Logger.log('onCreditRequestEdit: deleted dup Spoke row '
                  + matchedRows[d] + ' for: ' + clientName);
              }
              SpreadsheetApp.flush();
              logActivity_(hubSS, 'CREDIT->CRM DEDUP',
                'Removed ' + (matchedRows.length - 1) + ' dup row(s) for: ' + clientName, userEmail);
            }
          }

          // ── FIX v14.2: Write UUID + Name + Owner to Order Ledger CRM ──
          // Set row background to white so it does not trigger pink conditional
          // formatting (which fires on Payment Status = Pending formula).
          // No setCRMFormulas_() here — this is a registration-only row, not an order.
          // The Payment Status formula (and pink color) will be applied by
          // onCRMSubTabEdit when the first real order PI No is entered.
          const batchRow = Array(CRM_COL_CRM_OWNER).fill('');
          batchRow[CRM_COL_UUID      - 1] = existingUuid;
          batchRow[CRM_COL_CUSTOMER  - 1] = clientName;
          batchRow[CRM_COL_CRM_OWNER - 1] = resolvedOwnerEmail;
          crmMain.getRange(crmMainRow, 1, 1, CRM_COL_CRM_OWNER).setValues([batchRow]);
          // Force white background on registration row to suppress conditional formatting
          try {
            const lastCol = Math.max(crmMain.getLastColumn(), 18);
            crmMain.getRange(crmMainRow, 1, 1, lastCol).setBackground('#FFFFFF');
          } catch (_) {}
          SpreadsheetApp.flush();

          // ── Create / update client sub-tab + HYPERLINK ────────
          createClientTab(crmSS, clientName, crmMainRow, existingUuid);
          SpreadsheetApp.flush();

          logActivity_(hubSS, 'CREDIT->CRM MAIN+TAB (Sales/SalesMgr/Admin)',
            'Client:' + clientName + ' | Row:' + crmMainRow + ' | UUID:' + existingUuid, userEmail);

          // ── NOTIFY CRM OWNER: New client assigned ─────────────
          // DEDUP RULE: send ONE email per client assignment.
          // onCreditRequestEdit fires separately for col A (Client Name)
          // AND col B (CRM Owner). Without a guard both executions reach
          // here with resolvedOwnerEmail filled → CRM owner gets 2 emails.
          //
          // Fix: only send when col B (CRM Owner) is the column being edited.
          // That is the definitive "assignment complete" moment — the user
          // has just picked the owner, so both name and owner are final.
          //
          // Edge case — col A edited while col B already has an owner:
          // This is a client NAME change after assignment, not a new
          // assignment, so we do NOT send. The original col-B email already
          // covered this client.
          if (col === CC_COL_CRM_OWNER && resolvedOwnerEmail && clientName) {
            sendNewClientNotificationToCRM_(
              resolvedOwnerEmail,
              clientName,
              existingUuid,
              userEmail,       // who added it (Sales / Sales Manager / Admin)
              hubSS
            );
          }

        } catch (crmErr) {
          Logger.log('onCreditRequestEdit: CRM write error — ' + crmErr.message);
        }

      } finally {
        lock.releaseLock();
      }

      return;  // Done — no further processing for these columns
    }

    // ── Bug 3 fix: Notify CRM owner when Admin sets Approved/Rejected ──
    if (col === CC_COL_REQ_STATUS) {
      const status      = e.range.getValue();
      if (status !== 'Approved' && status !== 'Rejected') return;

      const data        = sheet.getDataRange().getValues();
      const rowData     = data[row - 1];
      const clientName  = rowData[CC_COL_CLIENT_NAME - 1];
      const crmOwnerRaw = (rowData[CC_COL_CRM_OWNER  - 1] || '').toString().trim();
      const reqAmount   = rowData[CC_COL_REQ_INCREASE - 1];
      const approvedBy  = rowData[CC_COL_APPROVED_BY_CC  - 1];

      if (!crmOwnerRaw) return;

      // FIX (Bug 3): extract bare email before using as GmailApp recipient
      const crmOwnerEmail = extractEmail_(crmOwnerRaw);
      if (!crmOwnerEmail) return;

      const color    = status === 'Approved' ? '#2E7D32' : '#C62828';
      const emoji = '';

      try {
        GmailApp.sendEmail(
          crmOwnerEmail, // FIX: bare email, not "Name (email)" string
          `Credit Limit Request ${status} - ${clientName}`,
          '',
          { htmlBody: `
            <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
              <p>Hi,</p>
              <p>Your credit limit increase request for <strong>${clientName}</strong> has been
                 <strong style="color:${color};">${status}</strong>.</p>
              <table style="border-collapse:collapse;font-size:13px;margin:12px 0;">
                <tr style="background:#1A237E;color:#fff;">
                  <th style="padding:8px 12px;" colspan="2">Credit Request Update</th>
                </tr>
                <tr><td style="padding:6px 12px;border:1px solid #ddd;">Client</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;"><strong>${clientName}</strong></td></tr>
                <tr><td style="padding:6px 12px;border:1px solid #ddd;">Requested Amount</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">$${reqAmount}</td></tr>
                <tr><td style="padding:6px 12px;border:1px solid #ddd;">Status</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;color:${color};font-weight:bold;">${status}</td></tr>
                <tr><td style="padding:6px 12px;border:1px solid #ddd;">Actioned By</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">${approvedBy || 'Admin'}</td></tr>
              </table>
              <p style="color:#999;font-size:12px;">— CRM System (Auto-generated)</p>
            </div>`
          }
        );
      } catch (mailErr) {
        Logger.log('Credit approval notify error: ' + mailErr.message);
      }
      return;
    }

    if (col !== CC_COL_REQ_INCREASE) return;

    const reqAmount = e.range.getValue();
    if (!reqAmount || isNaN(Number(reqAmount))) return;

    const data        = sheet.getDataRange().getValues();
    const rowData     = data[row - 1];
    const clientName  = rowData[CC_COL_CLIENT_NAME  - 1];
    const crmOwnerRaw = (rowData[CC_COL_CRM_OWNER   - 1] || '').toString().trim();
    const creditLimit = rowData[CC_COL_CREDIT_LIMIT - 1];
    const usedCredit  = rowData[CC_COL_USED_CREDIT  - 1];
    const availCredit = rowData[CC_COL_AVAIL_CREDIT - 1];
    const reqReason   = rowData[CC_COL_REQ_REASON   - 1];
    const userEmail   = Session.getActiveUser().getEmail();

    // Display value for the email body (keep human-readable if available)
    const crmOwnerDisplay = crmOwnerRaw;

    // Update request status
    sheet.getRange(row, CC_COL_REQ_STATUS).setValue('Pending');
    sheet.getRange(row, CC_COL_LAST_UPDATED).setValue(new Date());

    // ── v3.0: Get admin emails from config sheet ──────────────
    const config      = getCrmConfig_();
    const adminEmails = config.adminEmails;
    if (!adminEmails.length) {
      Logger.log('onCreditRequestEdit: No Admin emails in CRM Members sheet.');
      return;
    }

    try {
      adminEmails.forEach(adminEmail => {
        GmailApp.sendEmail(
          adminEmail,
          `Credit Increase Request - ${clientName} - Action Required`,
          '',
          {
            htmlBody: `
              <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
                <p>Dear Admin,</p>
                <p>A <strong>credit limit increase</strong> has been requested.</p>
                <table style="border-collapse:collapse;font-size:13px;margin:12px 0;width:100%;">
                  <tr style="background:#1A237E;color:#fff;">
                    <th style="padding:8px 12px;text-align:left;" colspan="2">Credit Increase Request</th>
                  </tr>
                  <tr style="background:#E3F2FD;">
                    <td style="padding:6px 12px;border:1px solid #ddd;">Client Name</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;"><strong>${clientName}</strong></td>
                  </tr>
                  <tr>
                    <td style="padding:6px 12px;border:1px solid #ddd;">Requested By (CRM Owner)</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">${crmOwnerDisplay}</td>
                  </tr>
                  <tr style="background:#E3F2FD;">
                    <td style="padding:6px 12px;border:1px solid #ddd;">Current Credit Limit</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">$${creditLimit}</td>
                  </tr>
                  <tr>
                    <td style="padding:6px 12px;border:1px solid #ddd;">Used Credit</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">$${usedCredit}</td>
                  </tr>
                  <tr style="background:#E3F2FD;">
                    <td style="padding:6px 12px;border:1px solid #ddd;">Available Credit</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">$${availCredit}</td>
                  </tr>
                  <tr>
                    <td style="padding:6px 12px;border:1px solid #ddd;"><strong>Requested New Limit</strong></td>
                    <td style="padding:6px 12px;border:1px solid #ddd;color:#C62828;"><strong>$${reqAmount}</strong></td>
                  </tr>
                  <tr style="background:#E3F2FD;">
                    <td style="padding:6px 12px;border:1px solid #ddd;">Reason</td>
                    <td style="padding:6px 12px;border:1px solid #ddd;">${reqReason || '(no reason given)'}</td>
                  </tr>
                </table>
                <p>Please open the <strong>Client Credit</strong> sheet in Master Hub to approve or reject.</p>
                <p style="color:#999;font-size:12px;">— CRM System (Auto-generated)</p>
              </div>`
          }
        );
      });
    } catch (mailErr) {
      Logger.log('Credit request email error: ' + mailErr.message);
    }

    const hubSS = e.source;
    logActivity_(hubSS, 'CREDIT REQUEST', `Client:${clientName} | Req:$${reqAmount} | By:${crmOwnerDisplay}`, userEmail);
    Logger.log('Credit increase request emailed to admin(s) for: ' + clientName);
  } catch (err) {
    Logger.log('ERROR in onCreditRequestEdit: ' + err.message);
  }
}

function sendNewClientNotificationToCRM_(crmOwnerEmail, clientName, uuid, addedByEmail, hubSS) {
  try {
    if (!crmOwnerEmail || !clientName) return;

    const config     = getCrmConfig_();
    const crmFileId  = config.crmFileIds[crmOwnerEmail.toLowerCase()];
    const today      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');

    // Resolve display name for the CRM owner from CRM Members
    const ownerMember  = config.allMembers.find(m => m.gmail === crmOwnerEmail.toLowerCase());
    const ownerName    = (ownerMember && ownerMember.name) ? ownerMember.name : crmOwnerEmail;

    // Resolve full details for whoever added this client from CRM Members sheet.
    // Shows: "Name (email) — Role" so the CRM owner knows exactly who assigned them.
    const adderMember   = config.allMembers.find(m => m.gmail === (addedByEmail || '').toLowerCase());
    const adderName     = adderMember && adderMember.name  ? adderMember.name  : '';
    const adderRole     = adderMember && adderMember.role  ? adderMember.role  : '';
    const adderEmail    = addedByEmail || '';
    // Build display string: "Full Name (email@domain.com) — Role"
    // Falls back gracefully when name/role are missing.
    let adderDisplay;
    if (adderName && adderRole) {
      adderDisplay = adderName + ' (' + adderEmail + ') — ' + adderRole;
    } else if (adderName) {
      adderDisplay = adderName + ' (' + adderEmail + ')';
    } else if (adderEmail) {
      adderDisplay = adderEmail + (adderRole ? ' — ' + adderRole : '');
    } else {
      adderDisplay = 'Unknown';
    }

    // Build the CRM spoke URL — deep-link to the client sub-tab if we can find its gid
    let spokeUrl = crmFileId
      ? 'https://docs.google.com/spreadsheets/d/' + crmFileId + '/edit'
      : 'https://docs.google.com/spreadsheets/d/' + MASTER_HUB_ID + '/edit';

    let clientTabUrl = spokeUrl; // fallback — will update below
    if (crmFileId) {
      try {
        const crmSS2      = SpreadsheetApp.openById(crmFileId);
        const clientSheet = crmSS2.getSheetByName(clientName.toString().trim());
        if (clientSheet) {
          clientTabUrl = 'https://docs.google.com/spreadsheets/d/' + crmFileId +
                         '/edit#gid=' + clientSheet.getSheetId();
        }
      } catch (_) {}
    }

    const subject = '[UV Overseas CRM] New Client Assigned — ' + clientName;

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:640px;">
        <div style="background:#1A237E;padding:18px 26px;border-radius:6px 6px 0 0;">
          <span style="color:#fff;font-size:18px;font-weight:bold;">New Client Assigned to You</span>
        </div>
        <div style="border:1px solid #ddd;border-top:none;padding:24px 26px;border-radius:0 0 6px 6px;">
          <p style="margin:0 0 14px;">Dear ${ownerName},</p>
          <p>A new client has been assigned to you in the <strong>UV Overseas CRM</strong> system.
             Please open your CRM sheet and complete the order details as soon as possible.</p>

          <table style="border-collapse:collapse;font-size:13px;width:100%;margin:16px 0;">
            <tr style="background:#1A237E;color:#fff;">
              <th style="padding:9px 14px;text-align:left;" colspan="2">New Client Details</th>
            </tr>
            <tr style="background:#E8EAF6;">
              <td style="padding:8px 14px;border:1px solid #ddd;width:40%;">Client Name</td>
              <td style="padding:8px 14px;border:1px solid #ddd;"><strong>${clientName}</strong></td>
            </tr>
            <tr>
              <td style="padding:8px 14px;border:1px solid #ddd;">Assigned CRM</td>
              <td style="padding:8px 14px;border:1px solid #ddd;">${ownerName} (${crmOwnerEmail})</td>
            </tr>
            <tr style="background:#E8EAF6;">
              <td style="padding:8px 14px;border:1px solid #ddd;">Assigned By</td>
              <td style="padding:8px 14px;border:1px solid #ddd;"><strong>${adderName || adderEmail}</strong><br>
                <span style="color:#555;font-size:12px;">${adderEmail}</span><br>
                <span style="display:inline-block;margin-top:3px;background:#E8EAF6;color:#1A237E;
                      font-size:11px;font-weight:bold;padding:2px 8px;border-radius:10px;">${adderRole || 'Unknown Role'}</span>
              </td>
            </tr>
            <tr>
              <td style="padding:8px 14px;border:1px solid #ddd;">Date Added</td>
              <td style="padding:8px 14px;border:1px solid #ddd;">${today}</td>
            </tr>
            <tr style="background:#E8EAF6;">
              <td style="padding:8px 14px;border:1px solid #ddd;">Order Reference (UUID)</td>
              <td style="padding:8px 14px;border:1px solid #ddd;font-family:monospace;font-size:12px;">${uuid || '—'}</td>
            </tr>
          </table>

          <p style="margin:16px 0 8px;font-weight:bold;">What to do next:</p>
          <ol style="font-size:13px;line-height:1.9;margin:0 0 16px;padding-left:20px;">
            <li>Open your CRM sheet using the button below.</li>
            <li>Find the row for <strong>${clientName}</strong> in the <strong>Main</strong> tab
                (a client sub-tab has been created automatically).</li>
            <li>Enter the <strong>PI No</strong>, <strong>Inv No</strong>, <strong>Inv Date</strong>,
                and <strong>Invoice Value</strong>.</li>
            <li>Fill in the <strong>Due Date</strong> and <strong>CRM Owner</strong> email if not already set.</li>
            <li>Open the client sub-tab to enter order line details
                (Item Value, Dev Charges, Freight, Discount, etc.).</li>
          </ol>

          <p style="margin:0 0 16px;">
            <a href="${clientTabUrl}"
               style="display:inline-block;background:#1A237E;color:#fff;padding:11px 22px;
                      border-radius:4px;text-decoration:none;font-weight:bold;font-size:14px;">
              Open ${clientName} Tab in My CRM →
            </a>
            &nbsp;
            <a href="${spokeUrl}"
               style="display:inline-block;background:#455A64;color:#fff;padding:11px 22px;
                      border-radius:4px;text-decoration:none;font-weight:bold;font-size:14px;">
              Open CRM Main Sheet →
            </a>
          </p>

          <div style="background:#FFF8E1;border-left:4px solid #F9A825;padding:10px 14px;
                      border-radius:0 4px 4px 0;font-size:13px;margin-bottom:16px;">
            <strong>Note:</strong> A credit limit may already be set for this client.
            Check the <strong>Client Credit</strong> tab in the Master Hub before processing large orders.
            Orders exceeding the credit limit will be flagged for approval.
          </div>

          <p style="color:#999;font-size:12px;margin-top:20px;">
            This is an automated message from the UV Overseas CRM system.<br>
            If you believe this was assigned in error, please contact your Sales Manager or Admin.
          </p>
        </div>
      </div>`;

    const plainText =
      'Dear ' + ownerName + ',\n\n' +
      'A new client has been assigned to you: ' + clientName + '\n' +
      'Assigned By: ' + adderDisplay + '\n' +
      'Date: ' + today + '\n' +
      'UUID: ' + (uuid || '—') + '\n\n' +
      'Please open your CRM sheet and fill in the order details:\n' +
      '  1. PI No, Inv No, Inv Date, Invoice Value in the Main tab\n' +
      '  2. Order line details in the ' + clientName + ' sub-tab\n\n' +
      'CRM Sheet: ' + spokeUrl + '\n\n' +
      '— UV Overseas CRM System (Auto-generated)';

    GmailApp.sendEmail(
      crmOwnerEmail,
      subject,
      plainText,
      { htmlBody: htmlBody, name: 'UV Overseas CRM System' }
    );

    logActivity_(hubSS, 'NEW CLIENT EMAIL SENT',
      'Client:' + clientName + ' | To CRM:' + crmOwnerEmail + ' | By:' + addedByEmail, addedByEmail || 'system');

    Logger.log('sendNewClientNotificationToCRM_: email sent to ' + crmOwnerEmail + ' for client: ' + clientName);
  } catch (err) {
    Logger.log('ERROR in sendNewClientNotificationToCRM_: ' + err.message);
  }
}

function checkCreditAndFlag_(hubSS, hubSheet, hubRow, uuid, invValue, customerName, crmOwner, crmMainSheet, crmRow, userEmail) {
  try {
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) {
      Logger.log('checkCreditAndFlag_: Client Credit sheet not found — skipping check');
      return;
    }

    const creditData = creditSheet.getDataRange().getValues();
    let creditRow = -1;
    for (let i = 1; i < creditData.length; i++) {
      if (creditData[i][CC_COL_CLIENT_NAME - 1].toString().trim().toLowerCase() ===
          customerName.toString().trim().toLowerCase()) {
        creditRow = i;
        break;
      }
    }

    if (creditRow === -1) {
      Logger.log('checkCreditAndFlag_: No credit record for: ' + customerName);
      return;
    }

    const creditLimit = Number(creditData[creditRow][CC_COL_CREDIT_LIMIT - 1]) || 0;

    // v15.1 FIX: this function was using HUB_COL_* constants (the OLD/legacy
    // "Master Hub" sheet layout, 18 cols) to read from hubSheet, which is
    // actually the Order Ledger sheet (NEW v14.0 layout, 20 cols, different
    // column meanings at the same index!). HUB_COL_INV_VALUE(7) on Order
    // Ledger is actually "Order Type", and HUB_COL_PAY_RECEIVED(10) is
    // actually "Freight" — so usedCredit was being computed from completely
    // wrong data. Fixed to use OL_COL_* throughout.
    const hubData = hubSheet.getDataRange().getValues();
    let totalInv = 0, totalPay = 0;
    for (let i = 1; i < hubData.length; i++) {
      const rowName = (hubData[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
      if (rowName === customerName.toString().trim().toLowerCase()) {
        totalInv += Number(hubData[i][OL_COL_FINAL_INVOICE - 1]) || 0;
        totalPay += Number(hubData[i][OL_COL_PAY_RECEIVED  - 1]) || 0;
      }
    }
    const usedCredit  = Math.max(0, totalInv - totalPay);
    const availCredit = creditLimit - usedCredit;
    const numInvValue = Number(invValue) || 0;

    if (creditLimit === 0) {
      Logger.log('checkCreditAndFlag_: Credit limit is 0 for: ' + customerName + ' — skipping');
      return;
    }

    if (numInvValue > availCredit) {
      const overage = numInvValue - availCredit;

      // v15.1 FIX: removed the write to hubSheet via HUB_COL_APPROVAL_REQ(14) —
      // on the Order Ledger sheet, column 14 is "Dispatch Date", not "Approval
      // Required" (that constant is only valid on the old Master Hub layout).
      // Writing here was silently overwriting Logistics' Dispatch Date with
      // the text "Yes"/"No". Order Ledger has no Approval Required column at
      // all today — Order Ledger CRM (crmMainSheet) is the only sheet where
      // this status correctly exists, so that's the only write kept below.
      crmMainSheet.getRange(crmRow, CRM_COL_APPROVAL_REQ).setValue('Yes');

      // Send unified hold notification to Sales Manager, Sales, Admin
      sendHoldNotificationEmails_(hubSS, uuid, customerName, crmOwner,
        null, numInvValue, usedCredit, creditLimit, userEmail);

      logActivity_(hubSS, 'CREDIT EXCEEDED',
        `Client:${customerName} | Inv:$${numInvValue} | Avail:$${availCredit} | Over:$${overage}`, userEmail);

    } else {
      // v15.1 FIX: read the "currently flagged" state from Order Ledger CRM
      // (the only sheet that actually has this column) instead of hubSheet.
      const currentFlag = crmMainSheet.getRange(crmRow, CRM_COL_APPROVAL_REQ).getValue();
      if (currentFlag === 'Yes') {
        crmMainSheet.getRange(crmRow, CRM_COL_APPROVAL_REQ).setValue('No');
        Logger.log('checkCreditAndFlag_: Cleared stale credit flag for UUID: ' + uuid);
      }
    }
  } catch (err) {
    Logger.log('ERROR in checkCreditAndFlag_: ' + err.message);
  }
}

function exportClientPDF(customerName, recipientEmail) {
  try {
    if (!customerName || !recipientEmail) { Browser.msgBox('Customer name and email required.'); return; }
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const tabName     = customerName.toString().trim();
    const clientSheet = ss.getSheetByName(tabName);
    if (!clientSheet) { Browser.msgBox('Sheet not found: ' + tabName); return; }

    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${ss.getId()}/export` +
      `?format=pdf&gid=${clientSheet.getSheetId()}&size=A4&portrait=true` +
      `&fitw=true&fith=false&gridlines=false&printtitle=false&sheetnames=false&pagenumbers=false&attachment=true`;

    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) { Browser.msgBox('PDF export failed. Check logs.'); return; }

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    GmailApp.sendEmail(recipientEmail, `Invoice Summary — ${tabName} — ${today}`,
      `Dear Client,\n\nPlease find the invoice summary for ${tabName} attached.\n\nRegards`,
      { attachments: [response.getBlob().setName(tabName + '.pdf')], name: 'CRM System' }
    );
    logActivity_(ss, 'PDF EXPORTED', 'Customer:' + tabName + ' | To:' + recipientEmail, Session.getActiveUser().getEmail());
  } catch (err) {
    Logger.log('ERROR in exportClientPDF: ' + err.message);
    Browser.msgBox('Error: ' + err.message);
  }
}

function exportClientCSV(customerName, recipientEmail) {
  try {
    if (!customerName || !recipientEmail) { Browser.msgBox('Customer name and email required.'); return; }
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const tabName     = customerName.toString().trim();
    const clientSheet = ss.getSheetByName(tabName);
    if (!clientSheet) { Browser.msgBox('Sheet not found: ' + tabName); return; }

    const data = clientSheet.getDataRange().getValues();
    const csvRows = data.map(row =>
      row.map(cell => {
        let val = (cell === null || cell === undefined) ? '' : cell.toString();
        if (cell instanceof Date) val = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        val = val.replace(/"/g, '""');
        if (val.search(/(,|"|\n|\r)/) >= 0) val = `"${val}"`;
        return val;
      }).join(',')
    );
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    GmailApp.sendEmail(recipientEmail, `Invoice Summary — ${tabName} — ${today}`,
      `Dear Client,\n\nPlease find the CSV data for ${tabName} attached.\n\nRegards`,
      { attachments: [Utilities.newBlob(csvRows.join('\r\n'), 'text/csv', tabName + '.csv')], name: 'CRM System' }
    );
    logActivity_(ss, 'CSV EXPORTED', 'Customer:' + tabName + ' | To:' + recipientEmail, Session.getActiveUser().getEmail());
  } catch (err) {
    Logger.log('ERROR in exportClientCSV: ' + err.message);
    Browser.msgBox('Error: ' + err.message);
  }
}

function addCustomMenu() {
  // Sync license on every file open so status is current
  try { refreshLicenseStatus_(); } catch (_) {}
  // Auto-scroll Order Ledger CRM to last row on open
  try { scrollToLastRow_(); } catch (_) {}
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CRM Tools')
    .addSubMenu(ui.createMenu('📄 Export')
      .addItem('Export PDF to client', 'menuExportPDF_')
      .addItem('Export CSV to client', 'menuExportCSV_'))
    .addSubMenu(ui.createMenu('⚙️ Setup')
      .addItem('🚀 Setup Everything (run first)', 'setupEverything')
      .addItem('🧹 Cleanup Hub sheets from CRM',  'cleanupHubSheetsFromCRM')
      .addSeparator()
      .addItem('Setup Master Hub sheet',     'setupHubSheet')
      .addItem('Setup CRM Spoke sheet',      'setupCRMSheet')
      .addItem('Setup Credit sheet',         'setupCreditSheet')
      .addItem('Setup Members sheet',        'setupMembersSheet')
      .addItem('Setup triggers',             'setupAllTriggers')
      .addItem('Setup weekly reminder',      'setupTimedReminder'))
    .addSubMenu(ui.createMenu('👥 Manage Members')                  // NEW v3.0
      .addItem('View / edit members list',   'menuViewMembers_')
      .addItem('Reload member config now',   'menuReloadConfig_')
      .addItem('Repair Credit sheet access', 'menuRepairCreditProtection_')
      .addItem('Repair Hub sheet access',    'menuRepairHubProtection_')
      .addItem('🔑 Fix CRM write access',    'fixHubWriteAccess')
      .addItem('🔒 Repair ALL sheet security', 'menuRepairAllSheetSecurity_'))
    .addSubMenu(ui.createMenu('🔄 Sync & Repair')
      .addItem('Bulk sync this CRM -> Hub',        'syncAllCRMToHub')
      .addItem('Repair Hub formulas',             'addHubFormulasToAllRows')
      .addItem('🔧 Repair client tab formulas',   'menuRepairClientTab_'))
    .addSubMenu(ui.createMenu('💰 Credit Limit')
      .addItem('Add / update client credit limit', 'menuSetCreditLimit_')
      .addItem('View credit status (all clients)',  'menuViewCreditStatus_')
      .addItem('Request credit increase',           'menuRequestCreditIncrease_')
      .addItem('🔄 Refresh credit totals now',      'menuRefreshCreditTotals_'))
    .addSubMenu(ui.createMenu('🧑‍💼 Sales Actions')
      .addItem('➕ Add New CRM Member + Client + Credit Limit', 'menuSalesAddMemberAndClient_'))
    .addSubMenu(ui.createMenu('🗂️ Client Tabs')
      .addItem('🗑️ Delete client (all data)',      'menuDeleteClientFull_')
      .addItem('🗑️ Delete order by PI No / Inv No', 'menuDeleteOrderByRef_')
      .addItem('Delete client tab only',            'menuDeleteClientTab_'))
    .addItem('💳 Send payment reminder',     'menuSendPaymentReminder_')
    .addSubMenu(ui.createMenu('🗄️ Archive')
      .addItem('Archive data now (manual)',       'menuArchiveNow_')
      .addItem('Setup auto archive (Apr 1st)',    'setupArchiveTrigger_'))
    .addToUi();
}

function menuViewMembers_() {
  try {
    const hubSS      = SpreadsheetApp.openById(MASTER_HUB_ID);
    const membSheet  = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
    if (!membSheet) {
      alertIfUi_('CRM Members sheet not found.\nRun Setup -> Setup Members sheet first.');
      return;
    }
    const url = `https://docs.google.com/spreadsheets/d/${MASTER_HUB_ID}/edit#gid=${membSheet.getSheetId()}`;
    alertIfUi_(
      '👥 CRM Members Sheet:\n\n' + url +
      '\n\n(Copy and open in browser to add/edit/remove members)\n\n' +
      'Tip: To add a new CRM member:\n' +
      '1. Add their Gmail in col A\n' +
      '2. Add their CRM spoke File ID in col B\n' +
      '3. Set Role = CRM\n' +
      '4. Set Active = '
    );
  } catch (err) {
    alertIfUi_('Error: ' + err.message);
  }
}

function menuReloadConfig_() {
  invalidateConfigCache_();
  const config = getCrmConfig_();
  const crmCount = Object.keys(config.crmFileIds).length;
  alertIfUi_(
    '✅ Config reloaded successfully!\n\n' +
    `CRM members: ${crmCount}\n` +
    `Sales Managers: ${config.salesManagerEmails.length}\n` +
    `Sales: ${config.salesEmails.length}\n` +
    `Accounts: ${config.accountsEmails.length}\n` +
    `Logistics: ${config.logisticsEmails.length}\n` +
    `Admins: ${config.adminEmails.length}`
  );
}

function menuRepairCreditProtection_() {
  try {
    invalidateConfigCache_();
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) {
      alertIfUi_('Credit List sheet not found. Run Setup -> Setup Credit sheet first.');
      return;
    }

    const ownerEmail  = Session.getEffectiveUser().getEmail();
    const config      = getCrmConfig_();
    // v15.1 FIX: same bug as setupCreditSheet — adminEmails was missing
    // crmSpokeEmails, which blocks refreshCreditTotals_ when it runs under a
    // CRM exec's authorization.
    const crmSpokeEmailsForCredit = Object.keys(config.crmFileIds);
    const adminEmails = [...new Set([ownerEmail, ...config.adminEmails, ...crmSpokeEmailsForCredit])];

    // Full editors = Sales + Admin (v15.0: Sales Manager merged into Sales)
    const fullEditors = [...new Set([
      ownerEmail,
      ...config.adminEmails,
      ...config.salesEmails
    ])];

    // Remove ALL existing protections and start clean
    creditSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    // Cols A–C: Client Name, CRM Owner, Credit Limit — Sales + Admin
    const abcProt = creditSheet.getRange(2, CC_COL_CLIENT_NAME, 999, 3)
      .protect().setDescription('Sales/Admin — Client Name, CRM Owner, Credit Limit (A–C)');
    abcProt.removeEditors(abcProt.getEditors());
    abcProt.addEditors(fullEditors);
    abcProt.setWarningOnly(false);

    // Cols D–E: Used Credit, Available Credit — Admin only (script-maintained)
    const deProt = creditSheet.getRange(2, CC_COL_USED_CREDIT, 999, 2)
      .protect().setDescription('Admin only — Used Credit & Available Credit (D–E) — script-maintained');
    deProt.removeEditors(deProt.getEditors());
    deProt.addEditors(adminEmails);
    deProt.setWarningOnly(false);

    // Col F: Client Email — Sales + Admin
    const fProt = creditSheet.getRange(2, CC_COL_CLIENT_EMAIL, 999, 1)
      .protect().setDescription('Sales/Admin — Client Email (F)');
    fProt.removeEditors(fProt.getEditors());
    fProt.addEditors(fullEditors);
    fProt.setWarningOnly(false);

    // Col G: Last Updated — script-maintained, same as D–E.
    // v15.1 FIX: was fullEditors (no CRM) — this single throw used to abort
    // refreshCreditTotals_ before it reached updateClientSummary_.
    const gProt = creditSheet.getRange(2, CC_COL_LAST_UPDATED, 999, 1)
      .protect().setDescription('Auto-calculated — Last Updated (G), script only');
    gProt.removeEditors(gProt.getEditors());
    gProt.addEditors(adminEmails);
    gProt.setWarningOnly(false);

    logActivity_(hubSS, 'CREDIT PROTECTION REPAIRED',
      `A-C, F, G editable by Sales/Admin. D-E locked to Admin: ${adminEmails.join(', ')}`, ownerEmail);

    alertIfUi_(
      '✅ Credit List sheet protection fixed!\n\n' +
      'Cols A-C (Client Name, CRM Owner, Credit Limit): Sales + Admin.\n' +
      'Cols D-E (Used/Available Credit): Admin only (script-maintained).\n' +
      'Col F (Client Email): Sales + Admin.\n' +
      'Col G (Last Updated): Sales + Admin.\n\n' +
      'Full editors:\n' + fullEditors.join('\n')
    );
  } catch (err) {
    Logger.log('ERROR in menuRepairCreditProtection_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function menuRepairHubProtection_() {
  try {
    invalidateConfigCache_();
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) {
      alertIfUi_('Master Hub sheet not found. Run Setup -> Setup Master Hub sheet first.');
      return;
    }

    const ownerEmail       = Session.getEffectiveUser().getEmail();
    const config           = getCrmConfig_();
    const adminEditors     = [ownerEmail, ...config.adminEmails].filter((v, i, a) => a.indexOf(v) === i);
    const logisticsEditors = [ownerEmail, ...config.logisticsEmails].filter((v, i, a) => a.indexOf(v) === i);
    const accountsEditors  = [ownerEmail, ...config.accountsEmails ].filter((v, i, a) => a.indexOf(v) === i);
    const approvalEditors  = [ownerEmail,
      ...config.adminEmails, ...config.salesManagerEmails, ...config.salesEmails
    ].filter((v, i, a) => a.indexOf(v) === i);

    // Wipe ALL existing range protections unconditionally
    hubSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    function applyProt_(range, desc, extraEditors) {
      const p = range.protect().setDescription(desc);
      p.removeEditors(p.getEditors());
      const editors = [...new Set([...adminEditors, ...(extraEditors || [])])];
      p.addEditors(editors);
      p.setWarningOnly(false);
    }

    // Cols B–I: Admin only
    applyProt_(hubSheet.getRange(2, 2, 1000, 8),  'Admin only — core order columns B–I');
    // Cols J–K: Accounts + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_PAY_RECEIVED, 1000, 2),
               'Accounts only — Payment Received & Payment Date (J–K)', accountsEditors);
    // Cols L–M: Admin only (formula columns)
    applyProt_(hubSheet.getRange(2, 12, 1000, 2), 'Admin only — formula columns L–M');
    // Cols N–P: Sales + Sales Manager + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_APPROVAL_REQ, 1000, 3),
               'Sales+SalesManager — Approval Required, Approved By, Remarks (N–P)', approvalEditors);
    // Cols Q–R: Logistics + Admin
    applyProt_(hubSheet.getRange(2, HUB_COL_AWB_NO, 1000, 2),
               'Logistics only — AWB No & Dispatched Date (Q–R)', logisticsEditors);

    logActivity_(hubSS, 'HUB PROTECTION REPAIRED',
      `Admin: ${adminEditors.join(', ')} | Accounts (J-K): ${accountsEditors.join(', ')} | Logistics (Q-R): ${logisticsEditors.join(', ')}`, ownerEmail);

    // Also repair formulas on all existing Hub rows (H, L, M)
    const lastRow = hubSheet.getLastRow();
    if (lastRow >= 2) {
      for (let r = 2; r <= lastRow; r++) setHubFormulas_(hubSheet, r);
      SpreadsheetApp.flush();
    }

    alertIfUi_(
      'Master Hub protection + formulas fixed!\n\n' +
      'Cols B-I (core data):    Admin only\n' +
      'Cols J-K (payment):      Accounts + Admin\n' +
      'Cols L-M (formulas):     Admin only\n' +
      'Cols N-P (approval):     Sales + Sales Manager + Admin\n' +
      'Cols Q-R (dispatch):     Logistics + Admin\n\n' +
      'Due Balance (M), Payment Status (H), Dispatch Status (L)\n' +
      'formulas repaired on all existing rows.\n\n' +
      'Accounts editors:\n' + accountsEditors.join('\n') + '\n\n' +
      'Logistics editors:\n' + logisticsEditors.join('\n') + '\n\n' +
      'Approval editors (N-P):\n' + approvalEditors.join('\n')
    );
  } catch (err) {
    Logger.log('ERROR in menuRepairHubProtection_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function menuExportPDF_() {
  const name  = Browser.inputBox('Export PDF', 'Enter Customer Name (tab name):', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel' || !name) return;
  const email = Browser.inputBox('Export PDF', 'Enter recipient email:', Browser.Buttons.OK_CANCEL);
  if (email === 'cancel' || !email) return;
  exportClientPDF(name.trim(), email.trim());
}

function menuExportCSV_() {
  const name  = Browser.inputBox('Export CSV', 'Enter Customer Name (tab name):', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel' || !name) return;
  const email = Browser.inputBox('Export CSV', 'Enter recipient email:', Browser.Buttons.OK_CANCEL);
  if (email === 'cancel' || !email) return;
  exportClientCSV(name.trim(), email.trim());
}

// ============================================================
// DELETE CLIENT — FULL WIPE (v14.1)
// Removes a client completely from:
//   CRM  : Order Ledger CRM rows + client sub-tab sheet
//   Hub  : Order Ledger rows + Client Dues rows + Client Summary row
//          + Credit List row + Dashboard row
// Then refreshes all Hub aggregates so totals stay accurate.
// Runs from: CRM Tools -> Client Tabs -> Delete client (all data)
// ============================================================
function menuDeleteClientFull_() {
  try {
    const ui  = SpreadsheetApp.getUi();
    const ss  = SpreadsheetApp.getActiveSpreadsheet();

    // Detect whether running from Hub or CRM spoke
    const isHub = (ss.getId() === MASTER_HUB_ID);
    let hubSS, crmSS, mainSheet;

    if (isHub) {
      // ── Running from Admin Hub ─────────────────────────────────────────
      // Uses config.crmFileIds to open every CRM spoke and delete there too.
      hubSS = ss;
      const config = getCrmConfig_();
      const spokeFileIds = Object.values(config.crmFileIds); // all CRM spoke file IDs

      // Build client list from Hub Order Ledger
      const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
      if (!hubOL) { ui.alert('Error', 'Hub Order Ledger sheet not found.', ui.ButtonSet.OK); return; }
      const hubClientSet = new Set();
      const hubAllData = hubOL.getDataRange().getValues();
      for (let i = 1; i < hubAllData.length; i++) {
        const n = (hubAllData[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim();
        if (n) hubClientSet.add(n);
      }
      if (!hubClientSet.size) { ui.alert('No clients found in Hub Order Ledger.', '', ui.ButtonSet.OK); return; }

      // Step 1: Ask which client
      const clientList = [...hubClientSet].sort().join('\n  - ');
      const nameRaw = Browser.inputBox(
        'Delete Client - Step 1 of 2',
        'Clients in Hub Order Ledger:\n\n  - ' + clientList +
        '\n\nType the client name exactly as shown above:',
        Browser.Buttons.OK_CANCEL
      );
      if (nameRaw === 'cancel' || !nameRaw.trim()) return;
      const clientName = nameRaw.trim();
      if (![...hubClientSet].some(n => n.toLowerCase() === clientName.toLowerCase())) {
        ui.alert('Client "' + clientName + '" not found.\nCheck spelling and try again.', '', ui.ButtonSet.OK);
        return;
      }
      const nameLower  = clientName.toLowerCase();
      const userEmail  = Session.getActiveUser().getEmail();

      // Step 2: Confirm — now covering ALL files
      const spokeCount = spokeFileIds.length;
      const confirm = ui.alert(
        'CONFIRM FULL DELETE (Hub + All CRM Spokes)',
        'Permanently delete ALL data for:\n\n  Client: ' + clientName + '\n\n' +
        'This will remove from Hub:\n' +
        '  - Hub Order Ledger rows\n' +
        '  - Hub Credit List row\n' +
        '  - Hub Client Summary row\n' +
        '  - Hub Client Dues rows\n' +
        '  - Hub Dashboard row\n\n' +
        'AND from every CRM spoke (' + spokeCount + ' file(s)):\n' +
        '  - CRM Order Ledger rows\n' +
        '  - CRM Client Dues rows\n' +
        '  - Client sub-tab sheet\n\n' +
        'WARNING: Cannot be undone. Proceed?',
        ui.ButtonSet.YES_NO
      );
      if (confirm !== ui.Button.YES) return;

      const deletedLog = [];

      // ── Hub sheets ─────────────────────────────────────────────────────
      const deleteMatchingRows_ = (sheet, colIdx) => {
        if (!sheet) return 0;
        const d = sheet.getDataRange().getValues();
        const rows = [];
        for (let i = 1; i < d.length; i++) {
          if ((d[i][colIdx] || '').toString().trim().toLowerCase() === nameLower) rows.push(i + 1);
        }
        for (let r = rows.length - 1; r >= 0; r--) sheet.deleteRow(rows[r]);
        return rows.length;
      };

      const hubOL2 = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
      deletedLog.push('Hub Order Ledger: ' + deleteMatchingRows_(hubOL2, OL_COL_CLIENT_NAME - 1) + ' row(s) deleted');

      const crSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
      deletedLog.push('Hub Credit List: ' + deleteMatchingRows_(crSheet, CC_COL_CLIENT_NAME - 1) + ' row(s) deleted');

      deletedLog.push('Hub Client Summary: ' + deleteMatchingRows_(hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME), 0) + ' row(s) deleted');
      deletedLog.push('Hub Client Dues: '    + deleteMatchingRows_(hubSS.getSheetByName(CLIENT_DUES_SHEET_NAME),    0) + ' row(s) deleted');
      deletedLog.push('Hub Dashboard: '      + deleteMatchingRows_(hubSS.getSheetByName(DASHBOARD_SHEET_NAME),      0) + ' row(s) deleted');

      SpreadsheetApp.flush();

      // ── Every CRM spoke ────────────────────────────────────────────────
      let spokeErrors = [];
      for (const fileId of spokeFileIds) {
        if (!fileId) continue;
        try {
          const spokeSS = SpreadsheetApp.openById(fileId);
          const spokeName = spokeSS.getName();

          // CRM Order Ledger rows
          const spokeMain = getCRMMainSheet_(spokeSS);
          if (spokeMain) {
            const sd = spokeMain.getDataRange().getValues();
            const sRows = [];
            for (let i = 1; i < sd.length; i++) {
              const n = (sd[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
              if (n === nameLower) sRows.push(i + 1);
            }
            for (let r = sRows.length - 1; r >= 0; r--) spokeMain.deleteRow(sRows[r]);
            if (sRows.length) deletedLog.push(spokeName + ' CRM Order Ledger: ' + sRows.length + ' row(s) deleted');
          }

          // Client sub-tab sheet
          let subSheet = spokeSS.getSheetByName(clientName);
          if (!subSheet) spokeSS.getSheets().forEach(sh => {
            if (sh.getName().toLowerCase() === nameLower) subSheet = sh;
          });
          if (subSheet) {
            spokeSS.deleteSheet(subSheet);
            deletedLog.push(spokeName + ' sub-tab "' + clientName + '" deleted');
          }

          // CRM Client Dues spoke sheet
          const crmDues = spokeSS.getSheetByName(CRM_DUES_SHEET);
          if (crmDues) {
            const dd = crmDues.getDataRange().getValues();
            const dRows = [];
            for (let i = 1; i < dd.length; i++) {
              if ((dd[i][0] || '').toString().trim().toLowerCase() === nameLower) dRows.push(i + 1);
            }
            for (let r = dRows.length - 1; r >= 0; r--) crmDues.deleteRow(dRows[r]);
            if (dRows.length) deletedLog.push(spokeName + ' CRM Client Dues: ' + dRows.length + ' row(s) deleted');
          }

          SpreadsheetApp.flush();

        } catch (spokeErr) {
          const msg = 'Spoke ' + fileId + ': ' + spokeErr.message;
          Logger.log('menuDeleteClientFull_ Hub spoke error: ' + msg);
          spokeErrors.push(msg);
        }
      }

      // ── Refresh Hub aggregates ─────────────────────────────────────────
      try { refreshCreditTotals_(userEmail); } catch(e) {}
      try { refreshClientDues_(userEmail);   } catch(e) {}
      try { refreshDashboard_(userEmail);    } catch(e) {}
      deletedLog.push('Hub aggregates refreshed');

      if (spokeErrors.length) deletedLog.push('Spoke errors: ' + spokeErrors.join('; '));

      logActivity_(hubSS, 'CLIENT DELETED (HUB+SPOKES)', clientName + ' | ' + deletedLog.join(' | '), userEmail);

      ui.alert(
        'Client Fully Deleted',
        '"' + clientName + '" removed from Hub and all CRM spokes.\n\n' +
        deletedLog.map(l => '  - ' + l).join('\n'),
        ui.ButtonSet.OK
      );
      return;
    }

    // ── Running from CRM spoke ────────────────────────────────────────────
    crmSS = ss;
    mainSheet = getCRMMainSheet_(crmSS);

    // Step 1: Build list of unique client names from CRM Order Ledger
    if (!mainSheet) {
      ui.alert('Error', 'Order Ledger CRM sheet not found.', ui.ButtonSet.OK);
      return;
    }
    const crmData = mainSheet.getDataRange().getDisplayValues();
    const clientSet = new Set();
    for (let i = 1; i < crmData.length; i++) {
      const name = (crmData[i][CRM_COL_CUSTOMER - 1] || '').toString().trim();
      if (name) clientSet.add(name);
    }
    if (!clientSet.size) {
      ui.alert('No clients found in Order Ledger CRM.', '', ui.ButtonSet.OK);
      return;
    }

    // Step 2: Ask which client to delete
    const clientList = [...clientSet].sort().join('\n  - ');
    const nameRaw = Browser.inputBox(
      'Delete Client - Step 1 of 2',
      'Current clients:\n\n  - ' + clientList +
      '\n\nType the client name exactly as shown above:',
      Browser.Buttons.OK_CANCEL
    );
    if (nameRaw === 'cancel' || !nameRaw.trim()) return;
    const clientName = nameRaw.trim();

    // Validate the name exists
    if (![...clientSet].some(n => n.toLowerCase() === clientName.toLowerCase())) {
      ui.alert('Client "' + clientName + '" not found in Order Ledger CRM.\nCheck spelling and try again.', '', ui.ButtonSet.OK);
      return;
    }

    // Step 3: Confirmation with full scope warning
    const confirm = ui.alert(
      'CONFIRM FULL DELETE',
      'You are about to permanently delete ALL data for:\n\n' +
      '  Client: ' + clientName + '\n\n' +
      'This will remove:\n' +
      '  - All rows in CRM Order Ledger (this file)\n' +
      '  - Client sub-tab sheet "' + clientName + '" (this file)\n' +
      '  - All matching rows in Hub Order Ledger\n' +
      '  - Client row in Hub Credit List\n' +
      '  - Client row in Hub Client Summary\n' +
      '  - Client rows in Hub Client Dues\n' +
      '  - Client row in Hub Dashboard\n\n' +
      'WARNING: This CANNOT be undone. Proceed?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    const userEmail  = Session.getActiveUser().getEmail();
    const nameLower  = clientName.toLowerCase();
    const deletedLog = [];

    // Step 4: Delete from CRM Order Ledger (bottom-up to preserve row indices)
    {
      const data = mainSheet.getDataRange().getDisplayValues();
      const rowsToDelete = [];
      for (let i = 1; i < data.length; i++) {
        const name = (data[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
        if (name === nameLower) rowsToDelete.push(i + 1);
      }
      for (let r = rowsToDelete.length - 1; r >= 0; r--) {
        mainSheet.deleteRow(rowsToDelete[r]);
      }
      deletedLog.push('CRM Order Ledger: ' + rowsToDelete.length + ' row(s) deleted');
    }

    // Step 5: Delete client sub-tab from CRM
    {
      let tabSheet = ss.getSheetByName(clientName);
      if (!tabSheet) {
        ss.getSheets().forEach(sh => {
          if (sh.getName().toLowerCase() === nameLower) tabSheet = sh;
        });
      }
      if (tabSheet) {
        ss.deleteSheet(tabSheet);
        deletedLog.push('CRM sub-tab "' + clientName + '" deleted');
      } else {
        deletedLog.push('CRM sub-tab not found (may not have been created)');
      }
    }

    // Step 6: Hub operations
    try {
      const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);

      // 6a. Hub Order Ledger
      const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
      if (hubOL) {
        const hubData = hubOL.getDataRange().getValues();
        const hubRows = [];
        for (let i = 1; i < hubData.length; i++) {
          const name = (hubData[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
          if (name === nameLower) hubRows.push(i + 1);
        }
        for (let r = hubRows.length - 1; r >= 0; r--) hubOL.deleteRow(hubRows[r]);
        deletedLog.push('Hub Order Ledger: ' + hubRows.length + ' row(s) deleted');
      }

      // 6b. Hub Credit List
      const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
      if (creditSheet) {
        const creditData = creditSheet.getDataRange().getValues();
        const creditRows = [];
        for (let i = 1; i < creditData.length; i++) {
          const name = (creditData[i][CC_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
          if (name === nameLower) creditRows.push(i + 1);
        }
        for (let r = creditRows.length - 1; r >= 0; r--) creditSheet.deleteRow(creditRows[r]);
        deletedLog.push('Hub Credit List: ' + creditRows.length + ' row(s) deleted');
      }

      // 6c. Hub Client Summary
      const summarySheet = hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME);
      if (summarySheet) {
        const sumData = summarySheet.getDataRange().getValues();
        const sumRows = [];
        for (let i = 1; i < sumData.length; i++) {
          const name = (sumData[i][0] || '').toString().trim().toLowerCase();
          if (name === nameLower) sumRows.push(i + 1);
        }
        for (let r = sumRows.length - 1; r >= 0; r--) summarySheet.deleteRow(sumRows[r]);
        deletedLog.push('Hub Client Summary: ' + sumRows.length + ' row(s) deleted');
      }

      // 6d. Hub Client Dues
      const hubDues = hubSS.getSheetByName(CLIENT_DUES_SHEET_NAME);
      if (hubDues) {
        const duesData = hubDues.getDataRange().getValues();
        const duesRows = [];
        for (let i = 1; i < duesData.length; i++) {
          const name = (duesData[i][0] || '').toString().trim().toLowerCase();
          if (name === nameLower) duesRows.push(i + 1);
        }
        for (let r = duesRows.length - 1; r >= 0; r--) hubDues.deleteRow(duesRows[r]);
        deletedLog.push('Hub Client Dues: ' + duesRows.length + ' row(s) deleted');
      }

      // 6e. Hub Dashboard
      const dashSheet = hubSS.getSheetByName(DASHBOARD_SHEET_NAME);
      if (dashSheet) {
        const dashData = dashSheet.getDataRange().getValues();
        const dashRows = [];
        for (let i = 1; i < dashData.length; i++) {
          const name = (dashData[i][0] || '').toString().trim().toLowerCase();
          if (name === nameLower) dashRows.push(i + 1);
        }
        for (let r = dashRows.length - 1; r >= 0; r--) dashSheet.deleteRow(dashRows[r]);
        deletedLog.push('Hub Dashboard: ' + dashRows.length + ' row(s) deleted');
      }

      // 6f. CRM Client Dues spoke sheet
      const crmDues = ss.getSheetByName(CRM_DUES_SHEET);
      if (crmDues) {
        const crmDuesData = crmDues.getDataRange().getValues();
        const crmDuesRows = [];
        for (let i = 1; i < crmDuesData.length; i++) {
          const name = (crmDuesData[i][0] || '').toString().trim().toLowerCase();
          if (name === nameLower) crmDuesRows.push(i + 1);
        }
        for (let r = crmDuesRows.length - 1; r >= 0; r--) crmDues.deleteRow(crmDuesRows[r]);
        deletedLog.push('CRM Client Dues: ' + crmDuesRows.length + ' row(s) deleted');
      }

      SpreadsheetApp.flush();

      // Step 7: Refresh all Hub aggregates
      try { refreshCreditTotals_(userEmail); } catch (e) {
        Logger.log('menuDeleteClientFull_: refreshCreditTotals_ error: ' + e.message);
      }
      try { refreshClientDues_(userEmail); } catch (e) {
        Logger.log('menuDeleteClientFull_: refreshClientDues_ error: ' + e.message);
      }
      try { refreshDashboard_(userEmail); } catch (e) {
        Logger.log('menuDeleteClientFull_: refreshDashboard_ error: ' + e.message);
      }
      deletedLog.push('Hub aggregates refreshed (Credit totals + Client Dues + Dashboard)');

    } catch (hubErr) {
      Logger.log('menuDeleteClientFull_: Hub error: ' + hubErr.message);
      deletedLog.push('Hub error: ' + hubErr.message);
    }

    // Step 8: Activity log + summary popup
    logActivity_(ss, 'CLIENT DELETED', clientName + ' | ' + deletedLog.join(' | '), userEmail);

    ui.alert(
      'Client Deleted',
      '"' + clientName + '" fully removed.\n\n' +
      deletedLog.map(l => '  - ' + l).join('\n'),
      ui.ButtonSet.OK
    );

  } catch (err) {
    Logger.log('ERROR in menuDeleteClientFull_: ' + err.message);
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
  }
}

// ============================================================
// DELETE SINGLE ORDER BY PI No or Inv No (v14.1)
// Removes one order row from every sheet it lives in:
//   CRM  : Order Ledger CRM row
//   CRM  : Matching row in client sub-tab (by PI No)
//   Hub  : Order Ledger row
//   Hub  : Client Dues row (if present)
//   CRM  : CRM Client Dues row (if present)
// Then refreshes Hub aggregates so totals stay accurate.
// Runs from: CRM Tools -> Client Tabs -> Delete order by PI No / Inv No
// ============================================================
function menuDeleteOrderByRef_() {
  try {
    const ui  = SpreadsheetApp.getUi();
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const isHub = (ss.getId() === MASTER_HUB_ID);

    // Step 1: Ask for PI No or Inv No
    const refRaw = Browser.inputBox(
      'Delete Order - Step 1 of 2',
      'Enter the PI No OR Inv No of the order to delete:\n' +
      '(e.g. #UV101  or  INV-2025-001)',
      Browser.Buttons.OK_CANCEL
    );
    if (refRaw === 'cancel' || !refRaw.trim()) return;
    const ref = refRaw.trim().toLowerCase();

    // Step 2: Find matching row(s) — search Hub Order Ledger if on Hub, CRM if on CRM
    const matches = [];  // { sheetRow, piNo, invNo, clientName, sheet }

    if (isHub) {
      // Search Hub Order Ledger (OL_COL_* columns)
      const hubOL = ss.getSheetByName(ORDER_LEDGER_SHEET_NAME);
      if (!hubOL) { ui.alert('Error', 'Hub Order Ledger not found.', ui.ButtonSet.OK); return; }
      const hubData = hubOL.getDataRange().getValues();
      for (let i = 1; i < hubData.length; i++) {
        const piNo  = (hubData[i][OL_COL_PI_NO  - 1] || '').toString().trim();
        const invNo = (hubData[i][OL_COL_INV_NO - 1] || '').toString().trim();
        if (piNo.toLowerCase() === ref || invNo.toLowerCase() === ref) {
          const clientName = (hubData[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim();
          matches.push({ sheetRow: i + 1, piNo, invNo, clientName, sheet: hubOL });
        }
      }
    } else {  // CRM spoke
      // Search CRM Order Ledger (CRM_COL_* columns)
      const mainSheet = getCRMMainSheet_(ss);
      if (!mainSheet) { ui.alert('Error', 'Order Ledger CRM sheet not found.', ui.ButtonSet.OK); return; }
      const crmData = mainSheet.getDataRange().getValues();
      for (let i = 1; i < crmData.length; i++) {
        const piNo  = (crmData[i][CRM_COL_PI_NO  - 1] || '').toString().trim();
        const invNo = (crmData[i][CRM_COL_INV_NO - 1] || '').toString().trim();
        if (piNo.toLowerCase() === ref || invNo.toLowerCase() === ref) {
          const clientName = mainSheet.getRange(i + 1, CRM_COL_CUSTOMER).getDisplayValue().trim();
          matches.push({ sheetRow: i + 1, piNo, invNo, clientName, sheet: mainSheet });
        }
      }
    }

    if (!matches.length) {
      ui.alert('Not Found', 'No order matching "' + refRaw.trim() + '" found in CRM Order Ledger.', ui.ButtonSet.OK);
      return;
    }

    // If multiple matches (edge case), list them and ask user to be more specific
    if (matches.length > 1) {
      const list = matches.map(m =>
        '  PI No: ' + (m.piNo || '-') + '  |  Inv No: ' + (m.invNo || '-') + '  |  Client: ' + m.clientName
      ).join('\n');
      ui.alert(
        'Multiple Matches Found',
        matches.length + ' rows match "' + refRaw.trim() + '":\n\n' + list +
        '\n\nPlease use a more specific reference (PI No is always unique).',
        ui.ButtonSet.OK
      );
      return;
    }

    const m = matches[0];

    // Step 3: Confirm
    const confirm = ui.alert(
      'CONFIRM DELETE ORDER',
      'Permanently delete this order?\n\n' +
      '  Client  : ' + m.clientName + '\n' +
      '  PI No   : ' + (m.piNo  || '(none)') + '\n' +
      '  Inv No  : ' + (m.invNo || '(none)') + '\n\n' +
      'This will remove the row from:\n' +
      '  - CRM Order Ledger\n' +
      '  - Client sub-tab "' + m.clientName + '"\n' +
      '  - Hub Order Ledger\n' +
      '  - Hub Client Dues (if present)\n' +
      '  - CRM Client Dues (if present)\n\n' +
      'WARNING: Cannot be undone. Proceed?',
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;

    const userEmail = Session.getActiveUser().getEmail();
    const deletedLog = [];
    const piNoLower  = m.piNo.toLowerCase();
    const invNoLower = m.invNo.toLowerCase();

    // Helper: does a value match either ref?
    const isMatch = (pi, inv) => {
      const p = (pi  || '').toString().trim().toLowerCase();
      const v = (inv || '').toString().trim().toLowerCase();
      return (piNoLower  && p === piNoLower)  ||
             (invNoLower && v === invNoLower);
    };

    // Step 4: Delete from the source sheet (CRM Order Ledger if on CRM, Hub OL if on Hub)
    m.sheet.deleteRow(m.sheetRow);
    deletedLog.push((isHub ? 'Hub' : 'CRM') + ' Order Ledger: row ' + m.sheetRow + ' deleted');

    // Step 5: Delete matching row from client sub-tab (CRM only — Hub has no sub-tabs)
    if (!isHub && m.clientName) {
      let subSheet = ss.getSheetByName(m.clientName);
      if (!subSheet) {
        const nl = m.clientName.toLowerCase();
        ss.getSheets().forEach(sh => { if (sh.getName().toLowerCase() === nl) subSheet = sh; });
      }
      if (subSheet) {
        const subData = subSheet.getDataRange().getValues();
        const subRows = [];
        for (let i = 1; i < subData.length; i++) {
          const pi  = (subData[i][SUB_COL_PI_NO  - 1] || '').toString().trim();
          const inv = (subData[i][SUB_COL_INV_NO - 1] || '').toString().trim();
          if (isMatch(pi, inv)) subRows.push(i + 1);
        }
        for (let r = subRows.length - 1; r >= 0; r--) subSheet.deleteRow(subRows[r]);
        deletedLog.push('Client sub-tab "' + m.clientName + '": ' + subRows.length + ' row(s) deleted');
      } else {
        deletedLog.push('Client sub-tab not found — skipped');
      }
    }

    // Step 6: Hub operations
    try {
      const hubSS = isHub ? ss : SpreadsheetApp.openById(MASTER_HUB_ID);

      // 6a. Hub Order Ledger — skip if we already deleted from it in Step 4 (isHub)
      if (!isHub) {
        const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
        if (hubOL) {
          const hubData = hubOL.getDataRange().getValues();
          const hubRows = [];
          for (let i = 1; i < hubData.length; i++) {
            const pi  = hubData[i][OL_COL_PI_NO  - 1];
            const inv = hubData[i][OL_COL_INV_NO - 1];
            if (isMatch(pi, inv)) hubRows.push(i + 1);
          }
          for (let r = hubRows.length - 1; r >= 0; r--) hubOL.deleteRow(hubRows[r]);
          deletedLog.push('Hub Order Ledger: ' + hubRows.length + ' row(s) deleted');
        }
      }

      // 6a-ii. If running from Hub, also delete matching rows from every CRM spoke
      if (isHub) {
        const config = getCrmConfig_();
        const spokeFileIds = Object.values(config.crmFileIds);
        let spokeErrors = [];
        for (const fileId of spokeFileIds) {
          if (!fileId) continue;
          try {
            const spokeSS   = SpreadsheetApp.openById(fileId);
            const spokeName = spokeSS.getName();

            // CRM Order Ledger rows
            const spokeMain = getCRMMainSheet_(spokeSS);
            if (spokeMain) {
              const sd = spokeMain.getDataRange().getValues();
              const sRows = [];
              for (let i = 1; i < sd.length; i++) {
                const pi  = (sd[i][CRM_COL_PI_NO  - 1] || '').toString().trim();
                const inv = (sd[i][CRM_COL_INV_NO - 1] || '').toString().trim();
                if (isMatch(pi, inv)) sRows.push(i + 1);
              }
              for (let r = sRows.length - 1; r >= 0; r--) spokeMain.deleteRow(sRows[r]);
              if (sRows.length) deletedLog.push(spokeName + ' CRM Order Ledger: ' + sRows.length + ' row(s) deleted');
            }

            // Client sub-tab row
            if (m.clientName) {
              let subSheet = spokeSS.getSheetByName(m.clientName);
              if (!subSheet) spokeSS.getSheets().forEach(sh => {
                if (sh.getName().toLowerCase() === m.clientName.toLowerCase()) subSheet = sh;
              });
              if (subSheet) {
                const subData = subSheet.getDataRange().getValues();
                const subRows = [];
                for (let i = 1; i < subData.length; i++) {
                  const pi  = (subData[i][SUB_COL_PI_NO  - 1] || '').toString().trim();
                  const inv = (subData[i][SUB_COL_INV_NO - 1] || '').toString().trim();
                  if (isMatch(pi, inv)) subRows.push(i + 1);
                }
                for (let r = subRows.length - 1; r >= 0; r--) subSheet.deleteRow(subRows[r]);
                if (subRows.length) deletedLog.push(spokeName + ' sub-tab "' + m.clientName + '": ' + subRows.length + ' row(s) deleted');
              }
            }

            SpreadsheetApp.flush();
          } catch (spokeErr) {
            const msg = 'Spoke ' + fileId + ': ' + spokeErr.message;
            Logger.log('menuDeleteOrderByRef_ Hub spoke error: ' + msg);
            spokeErrors.push(msg);
          }
        }
        if (spokeErrors.length) deletedLog.push('Spoke errors: ' + spokeErrors.join('; '));
      }

      // 6b. Hub Client Dues (col 3 = PI No, col 4 = Inv No)
      const hubDues = hubSS.getSheetByName(CLIENT_DUES_SHEET_NAME);
      if (hubDues) {
        const duesData = hubDues.getDataRange().getValues();
        const duesRows = [];
        for (let i = 1; i < duesData.length; i++) {
          const pi  = duesData[i][2];  // col C = PI No (0-based index 2)
          const inv = duesData[i][3];  // col D = Inv No (0-based index 3)
          if (isMatch(pi, inv)) duesRows.push(i + 1);
        }
        for (let r = duesRows.length - 1; r >= 0; r--) hubDues.deleteRow(duesRows[r]);
        deletedLog.push('Hub Client Dues: ' + duesRows.length + ' row(s) deleted');
      }

      // 6c. CRM Client Dues spoke sheet (same column layout)
      const crmDues = ss.getSheetByName(CRM_DUES_SHEET);
      if (crmDues) {
        const crmDuesData = crmDues.getDataRange().getValues();
        const crmDuesRows = [];
        for (let i = 1; i < crmDuesData.length; i++) {
          const pi  = crmDuesData[i][2];
          const inv = crmDuesData[i][3];
          if (isMatch(pi, inv)) crmDuesRows.push(i + 1);
        }
        for (let r = crmDuesRows.length - 1; r >= 0; r--) crmDues.deleteRow(crmDuesRows[r]);
        deletedLog.push('CRM Client Dues: ' + crmDuesRows.length + ' row(s) deleted');
      }

      SpreadsheetApp.flush();

      // Step 7: Refresh Hub aggregates
      try { refreshCreditTotals_(userEmail); } catch (e) {
        Logger.log('menuDeleteOrderByRef_: refreshCreditTotals_ error: ' + e.message);
      }
      try { refreshClientDues_(userEmail); } catch (e) {
        Logger.log('menuDeleteOrderByRef_: refreshClientDues_ error: ' + e.message);
      }
      try { refreshDashboard_(userEmail); } catch (e) {
        Logger.log('menuDeleteOrderByRef_: refreshDashboard_ error: ' + e.message);
      }
      deletedLog.push('Hub aggregates refreshed');

    } catch (hubErr) {
      Logger.log('menuDeleteOrderByRef_: Hub error: ' + hubErr.message);
      deletedLog.push('Hub error: ' + hubErr.message);
    }

    // Step 8: Log + summary
    const logSS = isHub ? ss : ss;
    logActivity_(logSS, isHub ? 'ORDER DELETED (HUB+SPOKES)' : 'ORDER DELETED',
      'PI:' + m.piNo + ' | Inv:' + m.invNo + ' | Client:' + m.clientName +
      ' | ' + deletedLog.join(' | '),
      userEmail
    );

    ui.alert(
      'Order Deleted',
      'Order removed from ' + (isHub ? 'Hub and all CRM spokes' : 'CRM and Hub') + '.\n\n' +
      '  Client : ' + m.clientName + '\n' +
      '  PI No  : ' + (m.piNo  || '-') + '\n' +
      '  Inv No : ' + (m.invNo || '-') + '\n\n' +
      deletedLog.map(l => '  - ' + l).join('\n'),
      ui.ButtonSet.OK
    );

  } catch (err) {
    Logger.log('ERROR in menuDeleteOrderByRef_: ' + err.message);
    SpreadsheetApp.getUi().alert('Error: ' + err.message);
  }
}

function menuDeleteClientTab_() {
  const name = Browser.inputBox('Delete Tab Only', 'Enter Customer Name (deletes sub-tab only, NOT row data):', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel' || !name) return;
  deleteClientTab(name.trim());
}

function menuSendPaymentReminder_() {
  const email = Browser.inputBox('Payment Reminder', 'Send reminder to (email):', Browser.Buttons.OK_CANCEL);
  if (email === 'cancel' || !email) return;
  const name  = Browser.inputBox('Payment Reminder', 'Filter by Customer Name (blank = ALL overdue):', Browser.Buttons.OK_CANCEL);
  if (name === 'cancel') return;
  sendPaymentReminder(email.trim(), name.trim() || null);
}

function menuSetCreditLimit_() {
  const client = Browser.inputBox('Set Credit Limit', 'Enter Client Name (must match CRM exactly):', Browser.Buttons.OK_CANCEL);
  if (client === 'cancel' || !client) return;
  const owner  = Browser.inputBox('Set Credit Limit', 'Enter CRM Owner Gmail for this client:', Browser.Buttons.OK_CANCEL);
  if (owner === 'cancel' || !owner) return;
  const limit  = Browser.inputBox('Set Credit Limit', 'Enter Credit Limit ($) — numbers only:', Browser.Buttons.OK_CANCEL);
  if (limit === 'cancel' || !limit || isNaN(Number(limit))) { Browser.msgBox('Invalid amount.'); return; }
  setCreditLimit_(client.trim(), owner.trim(), Number(limit));
}

function menuViewCreditStatus_() {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) { alertIfUi_('Client Credit sheet not found. Run Setup -> Setup Credit sheet first.'); return; }
    const url = `https://docs.google.com/spreadsheets/d/${MASTER_HUB_ID}/edit#gid=${creditSheet.getSheetId()}`;
    alertIfUi_('Credit Status Sheet:\n\n' + url + '\n\n(Copy and open in browser)');
  } catch (err) {
    alertIfUi_('Error: ' + err.message);
  }
}

function menuRequestCreditIncrease_() {
  const client = Browser.inputBox('Request Credit Increase', 'Enter Client Name:', Browser.Buttons.OK_CANCEL);
  if (client === 'cancel' || !client) return;
  const amount = Browser.inputBox('Request Credit Increase', 'Enter requested new credit limit ($):', Browser.Buttons.OK_CANCEL);
  if (amount === 'cancel' || !amount || isNaN(Number(amount))) { Browser.msgBox('Invalid amount.'); return; }
  const reason = Browser.inputBox('Request Credit Increase', 'Enter reason for request:', Browser.Buttons.OK_CANCEL);
  if (reason === 'cancel') return;
  requestCreditIncrease_(client.trim(), Number(amount), reason.trim());
}

function menuSalesAddMemberAndClient_() {
  try {
    const ui = SpreadsheetApp.getUi();

    // Step 1: CRM Member details
    const memberGmail = Browser.inputBox(
      '[Sales] Add CRM Member — Step 1/5',
      'Enter new CRM member\'s Gmail (Google Account email):',
      Browser.Buttons.OK_CANCEL
    );
    if (memberGmail === 'cancel' || !memberGmail.trim()) return;

    const memberFileId = Browser.inputBox(
      '[Sales] Add CRM Member — Step 2/5',
      'Enter their CRM Spoke Spreadsheet File ID\n(from the URL: docs.google.com/spreadsheets/d/[FILE_ID]/edit):',
      Browser.Buttons.OK_CANCEL
    );
    if (memberFileId === 'cancel') return;

    const memberName = Browser.inputBox(
      '[Sales] Add CRM Member — Step 3/5',
      'Enter display name for this CRM member:',
      Browser.Buttons.OK_CANCEL
    );
    if (memberName === 'cancel' || !memberName.trim()) return;

    // Step 2: Client details
    const clientName = Browser.inputBox(
      '[Sales] Add Client — Step 4/5',
      'Enter client name to add to this CRM member (must match CRM exactly):',
      Browser.Buttons.OK_CANCEL
    );
    if (clientName === 'cancel' || !clientName.trim()) return;

    // Step 3: Credit Limit
    const creditLimitStr = Browser.inputBox(
      '[Sales] Set Credit Limit — Step 5/5',
      'Enter credit limit ($) for this client (numbers only, e.g. 500000).\nEnter 0 to set later:',
      Browser.Buttons.OK_CANCEL
    );
    if (creditLimitStr === 'cancel' || isNaN(Number(creditLimitStr))) {
      Browser.msgBox('Invalid credit limit. Aborting.');
      return;
    }

    const gmailClean      = memberGmail.trim().toLowerCase();
    const fileIdClean     = memberFileId.trim();
    const nameClean       = memberName.trim();
    const clientClean     = clientName.trim();
    const creditLimit     = Number(creditLimitStr);
    const salesUserEmail  = Session.getActiveUser().getEmail();

    // ── Write to CRM Members sheet ──────────────────────────
    const hubSS      = SpreadsheetApp.openById(MASTER_HUB_ID);
    const membSheet  = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
    if (!membSheet) {
      Browser.msgBox('CRM Members sheet not found. Ask Admin to run Setup -> Setup Members Sheet first.');
      return;
    }

    // Check if member already exists
    const membData = membSheet.getDataRange().getValues();
    let memberExists = false;
    for (let i = 1; i < membData.length; i++) {
      if ((membData[i][MC_COL_GMAIL - 1] || '').toString().trim().toLowerCase() === gmailClean) {
        memberExists = true;
        Browser.msgBox(`Note: ${gmailClean} already exists in CRM Members. Skipping member row — updating client credit only.`);
        break;
      }
    }

    if (!memberExists) {
      const newMembRow = membSheet.getLastRow() + 1;
      membSheet.getRange(newMembRow, MC_COL_GMAIL  ).setValue(gmailClean);
      membSheet.getRange(newMembRow, MC_COL_FILE_ID).setValue(fileIdClean);
      membSheet.getRange(newMembRow, MC_COL_ROLE   ).setValue('CRM');
      membSheet.getRange(newMembRow, MC_COL_NAME   ).setValue(nameClean);
      membSheet.getRange(newMembRow, MC_COL_ACTIVE ).setValue(true);
      invalidateConfigCache_();
      logActivity_(hubSS, 'SALES: MEMBER ADDED', `Gmail:${gmailClean} | Name:${nameClean}`, salesUserEmail);

      // Auto-share and notify (same logic as onMembersSheetEdit)
      try {
        notifyAndShareMember_(gmailClean, 'CRM', nameClean, fileIdClean);
      } catch (shareErr) {
        Logger.log('menuSalesAddMemberAndClient_: share/notify error — ' + shareErr.message);
      }
    }

    // ── Add client to Credit sheet + set limit ───────────────
    setCreditLimit_(clientClean, gmailClean, creditLimit);
    logActivity_(hubSS, 'SALES: CLIENT+CREDIT ADDED',
      `Client:${clientClean} | CRM:${gmailClean} | Limit:$${creditLimit}`, salesUserEmail);

    Browser.msgBox(
      `✅ Done!\n\n` +
      `CRM Member: ${nameClean} (${gmailClean})\n` +
      `Client: ${clientClean}\n` +
      `Credit Limit: $${creditLimit.toLocaleString('en-US')}\n\n` +
      `The new member has been emailed their access link.\n` +
      `Credit limit is live immediately.`
    );
  } catch (err) {
    Logger.log('ERROR in menuSalesAddMemberAndClient_: ' + err.message);
    Browser.msgBox('Error: ' + err.message);
  }
}

function requestCreditIncrease_(clientName, amount, reason) {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) {
      alertIfUi_('Client Credit sheet not found. Ask admin to run Setup -> Setup Credit sheet.');
      return;
    }

    const data = creditSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][CC_COL_CLIENT_NAME - 1].toString().trim().toLowerCase() ===
          clientName.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      alertIfUi_(`Client "${clientName}" not found in Credit sheet.\nAsk admin to add this client first.`);
      return;
    }

    creditSheet.getRange(targetRow, CC_COL_REQ_INCREASE).setValue(amount);
    creditSheet.getRange(targetRow, CC_COL_REQ_REASON).setValue(reason);
    creditSheet.getRange(targetRow, CC_COL_LAST_UPDATED).setValue(new Date());

    logActivity_(hubSS, 'CREDIT REQUEST SUBMITTED',
      `Client:${clientName} | Req:$${amount}`, Session.getActiveUser().getEmail());

    alertIfUi_(
      `Credit increase request submitted for ${clientName}.\n` +
      `Requested limit: $${amount.toLocaleString('en-US')}\n` +
      `Admin has been notified by email.`
    );
  } catch (err) {
    Logger.log('ERROR in requestCreditIncrease_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function syncAllCRMToHub() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getCRMMainSheet_(ss);
    if (!mainSheet) { Logger.log('No Main sheet found.'); return; }

    const crmData  = mainSheet.getDataRange().getValues();
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { Logger.log('syncAllCRMToHub: Hub sheet not found'); return; }
    const hubData   = hubSheet.getDataRange().getValues();
    const userEmail = Session.getActiveUser().getEmail();

    const uuidMap = {};
    for (let i = 1; i < hubData.length; i++) {
      const u = hubData[i][HUB_COL_UUID - 1];
      if (u) uuidMap[u] = i + 1;
    }

    const lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch (_) { alertIfUi_('Server busy — retry.'); return; }

    let appended = 0, updated = 0;
    try {
      for (let i = 1; i < crmData.length; i++) {
        const row = crmData[i];
        if (!row[CRM_COL_PI_NO - 1]) continue;

        let uuid = row[CRM_COL_UUID - 1];
        if (!uuid) {
          uuid = generateUUID_();
          mainSheet.getRange(i + 1, CRM_COL_UUID).setValue(uuid);
          row[CRM_COL_UUID - 1] = uuid;
        }

        const newRow = buildHubRowFromCRM_(uuid, row);
        if (uuidMap[uuid]) {
          const hubRowNum = uuidMap[uuid];
          hubSheet.getRange(hubRowNum, HUB_COL_UUID, 1, 9).setValues([[
            uuid, row[CRM_COL_CUSTOMER-1], row[CRM_COL_CRM_OWNER-1],
            row[CRM_COL_PI_NO-1], row[CRM_COL_INV_NO-1],
            row[CRM_COL_DATE-1], row[CRM_COL_INV_VALUE-1],
            '', row[CRM_COL_DUE_DATE-1],
          ]]);
          hubSheet.getRange(hubRowNum, HUB_COL_APPROVAL_REQ).setValue(row[CRM_COL_APPROVAL_REQ-1]);
          hubSheet.getRange(hubRowNum, HUB_COL_APPROVED_BY).setValue(row[CRM_COL_APPROVED_BY-1]);
          if (row[CRM_COL_REMARKS-1]) hubSheet.getRange(hubRowNum, HUB_COL_REMARKS).setValue(row[CRM_COL_REMARKS-1]);
          setHubFormulas_(hubSheet, hubRowNum);
          updated++;
        } else {
          const appendAt = hubSheet.getLastRow() + 1;
          hubSheet.getRange(appendAt, 1, 1, newRow.length).setValues([newRow]);
          setHubFormulas_(hubSheet, appendAt);
          uuidMap[uuid] = appendAt;
          appended++;
        }
        if (i % 20 === 0) { SpreadsheetApp.flush(); Utilities.sleep(500); }
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    const summary = `Bulk sync done. Appended: ${appended}, Updated: ${updated}`;
    logActivity_(ss,    'BULK SYNC',          summary, userEmail);
    logActivity_(hubSS, 'BULK SYNC RECEIVED', summary, userEmail);
    alertIfUi_(summary);
  } catch (err) {
    Logger.log('ERROR in syncAllCRMToHub: ' + err.message);
    alertIfUi_('Sync failed: ' + err.message);
  }
}

function addHubFormulasToAllRows() {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) return;
    const lastRow = hubSheet.getLastRow();
    if (lastRow < 2) return;
    for (let r = 2; r <= lastRow; r++) setHubFormulas_(hubSheet, r);
    logActivity_(hubSS, 'FORMULAS REPAIRED', `rows 2–${lastRow}`, Session.getActiveUser().getEmail());
    alertIfUi_(`Formulas applied to ${lastRow - 1} rows.`);
  } catch (err) {
    Logger.log('ERROR in addHubFormulasToAllRows: ' + err.message);
  }
}

function repairClientTab(customerName) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = (customerName || '').toString().trim();
    if (!tabName) { alertIfUi_('No customer name provided.'); return; }

    const clientSheet = ss.getSheetByName(tabName);
    if (!clientSheet) {
      alertIfUi_(`Tab "${tabName}" not found.`);
      return;
    }

    // ── Rewrite ARRAYFORMULAs ────────────────────────────────
    // Clear old formula first so setFormula doesn't conflict with
    // an existing ARRAYFORMULA covering the same range.
    clientSheet.getRange(2, SUB_COL_FOB_VALUE).clearContent();
    clientSheet.getRange(2, SUB_COL_CIF_VALUE).clearContent();
    clientSheet.getRange(2, SUB_COL_FINAL_INV).clearContent();
    SpreadsheetApp.flush();

    clientSheet.getRange(2, SUB_COL_FOB_VALUE).setFormula(
      `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(G2:G),0)+IFERROR(VALUE(H2:H),0)))`
    );
    clientSheet.getRange(2, SUB_COL_CIF_VALUE).setFormula(
      `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(I2:I),0)+IFERROR(VALUE(J2:J),0)))`
    );
    clientSheet.getRange(2, SUB_COL_FINAL_INV).setFormula(
      `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(K2:K),0)-IFERROR(VALUE(L2:L),0)))`
    );

    // ── Re-apply Order Type dropdown ─────────────────────────
    const orderTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Production', 'Sample', 'Showroom'], true)
      .setAllowInvalid(false)
      .build();
    clientSheet.getRange(2, SUB_COL_ORDER_TYPE, 999, 1).setDataValidation(orderTypeRule);

    // ── Auto-resize all columns ──────────────────────────────
    clientSheet.autoResizeColumns(1, 15);

    SpreadsheetApp.flush();
    logActivity_(ss, 'CLIENT TAB REPAIRED', tabName, Session.getActiveUser().getEmail());
    alertIfUi_(`✅ Tab "${tabName}" repaired!\n\nFOB, CIF, Final Invoice formulas rewritten.\nOrder Type dropdown restored.`);
  } catch (err) {
    Logger.log('ERROR in repairClientTab: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function repairAllClientTabs() {
  try {
    const ss           = SpreadsheetApp.getActiveSpreadsheet();
    const systemSheets = new Set([
      MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME
    ]);
    const userEmail = Session.getActiveUser().getEmail();
    let count = 0;

    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (systemSheets.has(name)) return;

      // Only repair sheets that look like client tabs
      // (have the expected 15-column header structure)
      const lastCol = sheet.getLastColumn();
      if (lastCol < 13) return;

      sheet.getRange(2, SUB_COL_FOB_VALUE).clearContent();
      sheet.getRange(2, SUB_COL_CIF_VALUE).clearContent();
      sheet.getRange(2, SUB_COL_FINAL_INV).clearContent();
      SpreadsheetApp.flush();

      sheet.getRange(2, SUB_COL_FOB_VALUE).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(G2:G),0)+IFERROR(VALUE(H2:H),0)))`
      );
      sheet.getRange(2, SUB_COL_CIF_VALUE).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(I2:I),0)+IFERROR(VALUE(J2:J),0)))`
      );
      sheet.getRange(2, SUB_COL_FINAL_INV).setFormula(
        `=ARRAYFORMULA(IF(G2:G="","",IFERROR(VALUE(K2:K),0)-IFERROR(VALUE(L2:L),0)))`
      );

      const orderTypeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Production', 'Sample', 'Showroom'], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(2, SUB_COL_ORDER_TYPE, 999, 1).setDataValidation(orderTypeRule);

      sheet.autoResizeColumns(1, 15);
      count++;
      logActivity_(ss, 'CLIENT TAB REPAIRED', name, userEmail);
    });

    SpreadsheetApp.flush();
    alertIfUi_(`✅ Repaired ${count} client tab(s).\n\nFOB, CIF, Final Invoice formulas rewritten.\nOrder Type dropdowns restored on all tabs.`);
  } catch (err) {
    Logger.log('ERROR in repairAllClientTabs: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function menuRepairClientTab_() {
  const name = Browser.inputBox(
    'Repair Client Tab',
    'Enter client name (tab name) to repair, or leave blank to repair ALL client tabs:',
    Browser.Buttons.OK_CANCEL
  );
  if (name === 'cancel') return;
  if (name.trim() === '') {
    repairAllClientTabs();
  } else {
    repairClientTab(name.trim());
  }
}

function deleteClientTab(customerName) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const tabName = (customerName || '').toString().trim();
    if (!tabName) { alertIfUi_('No customer name provided.'); return; }

    const ui      = SpreadsheetApp.getUi();
    const confirm = ui.alert('Confirm Delete', `Delete tab "${tabName}"? This cannot be undone.`, ui.ButtonSet.YES_NO);
    if (confirm !== ui.Button.YES) return;

    const sheet = ss.getSheetByName(tabName);
    if (!sheet) { ui.alert(`Tab "${tabName}" not found.`); return; }

    const mainSheet = getCRMMainSheet_(ss);
    if (mainSheet) {
      const data = mainSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const formula = mainSheet.getRange(i + 1, CRM_COL_CUSTOMER).getFormula();
        if (formula.includes(tabName)) {
          mainSheet.getRange(i + 1, CRM_COL_CUSTOMER).setValue(tabName);
          break;
        }
      }
    }
    ss.deleteSheet(sheet);
    logActivity_(ss, 'TAB DELETED', tabName, Session.getActiveUser().getEmail());
    ui.alert(`Tab "${tabName}" deleted.`);
  } catch (err) {
    Logger.log('ERROR in deleteClientTab: ' + err.message);
  }
}

function sendPaymentReminder(recipientEmail, customerName) {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) return;
    const data = hubSheet.getDataRange().getValues();

    const overdueRows = [];
    for (let i = 1; i < data.length; i++) {
      const row       = data[i];
      const name      = row[HUB_COL_CUSTOMER-1];
      const status    = row[HUB_COL_PAY_STATUS-1];
      const balance   = row[HUB_COL_DUE_BALANCE-1];
      const nameMatch = !customerName || name === customerName;
      if (nameMatch && (status === 'Partial' || status === 'Pending') && balance > 0) {
        const dd = row[HUB_COL_DUE_DATE-1];
        overdueRows.push({
          name, invNo: row[HUB_COL_INV_NO-1], invValue: row[HUB_COL_INV_VALUE-1],
          balance, status,
          dueDate: dd ? Utilities.formatDate(new Date(dd), Session.getScriptTimeZone(), 'dd-MMM-yyyy') : '—'
        });
      }
    }

    if (!overdueRows.length) { alertIfUi_('No outstanding balances found.'); return; }

    const tableRows = overdueRows.map(r => `
      <tr>
        <td style="padding:6px 10px;border:1px solid #ddd;">${r.name}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;">${r.invNo}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;">${r.invValue}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;color:#c62828;font-weight:bold;">${r.balance}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;">${r.status}</td>
        <td style="padding:6px 10px;border:1px solid #ddd;">${r.dueDate}</td>
      </tr>`).join('');

    const today  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');
    const target = recipientEmail || Session.getActiveUser().getEmail();
    GmailApp.sendEmail(target, `Payment Reminder — Outstanding Balance — ${today}`,
      'Please view in HTML-compatible email client.',
      {
        htmlBody: `
          <div style="font-family:Arial,sans-serif;font-size:14px;color:#333;">
            <p>Dear Team,</p><p>Outstanding balances as of <strong>${today}</strong>:</p>
            <table style="border-collapse:collapse;width:100%;font-size:13px;">
              <thead><tr style="background:#1A237E;color:#fff;">
                <th style="padding:8px 10px;text-align:left;">Customer</th>
                <th style="padding:8px 10px;text-align:left;">Inv No</th>
                <th style="padding:8px 10px;text-align:left;">Inv Value</th>
                <th style="padding:8px 10px;text-align:left;">Due Balance</th>
                <th style="padding:8px 10px;text-align:left;">Status</th>
                <th style="padding:8px 10px;text-align:left;">Due Date</th>
              </tr></thead>
              <tbody>${tableRows}</tbody>
            </table>
            <p style="margin-top:16px;">Please follow up at the earliest.</p>
            <p>— CRM System (Auto-generated)</p>
          </div>`
      }
    );
    logActivity_(hubSS, 'REMINDER SENT', `To:${target} | Rows:${overdueRows.length}`, Session.getActiveUser().getEmail());
    alertIfUi_(`Reminder sent to ${target} for ${overdueRows.length} invoice(s).`);
  } catch (err) {
    Logger.log('ERROR in sendPaymentReminder: ' + err.message);
  }
}

function setupTimedReminder() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendPaymentReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendPaymentReminder')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  alertIfUi_('Weekly reminder scheduled for Monday 9 AM.');
}

function generateUUID_() {
  const date  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let s = '';
  for (let i = 0; i < 4; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return `PI-${date}-${s}`;
}

function getMainSheet_(ss) {
  return getCRMMainSheet_(ss) || ss.getSheets()[0];
}

function buildHubRowFromCRM_(uuid, row) {
  return [
    uuid,
    row[CRM_COL_CUSTOMER-1], row[CRM_COL_CRM_OWNER-1],
    row[CRM_COL_PI_NO-1],    row[CRM_COL_INV_NO-1],
    row[CRM_COL_DATE-1],     row[CRM_COL_INV_VALUE-1],
    '',                       row[CRM_COL_DUE_DATE-1],
    '', '',
    '', '',
    row[CRM_COL_APPROVAL_REQ-1],
    row[CRM_COL_APPROVED_BY-1],
    row[CRM_COL_REMARKS-1] || '',
    '', '',
  ];
}

function setHubFormulas_(hubSheet, r) {
  hubSheet.getRange(r, HUB_COL_PAY_STATUS).setFormula(
    `=IF(OR(J${r}="",J${r}=0),"Pending",IF(J${r}<G${r},"Partial","Received"))`
  );
  // BUG 3 FIX: Dispatch logic based on credit approval only:
  //   • No invoice value yet             -> blank
  //   • Approval NOT required (N = No or empty) -> "Ready to Dispatch" (credit within limit)
  //   • Approval required + Approved (N = Approved) -> "Ready to Dispatch" (Sales/Admin approved)
  //   • Approval required but NOT yet approved (N = Yes) -> "Hold – Approval Required"
  hubSheet.getRange(r, HUB_COL_DISPATCH_STATUS).setFormula(
    `=IF(G${r}="","",IF(OR(N${r}="No",N${r}="",N${r}="Approved"),"Ready to Dispatch","Hold – Approval Required"))`
  );
  hubSheet.getRange(r, HUB_COL_DUE_BALANCE).setFormula(
    `=IF(G${r}="","",G${r}-J${r})`
  );
}

function setCRMFormulas_(sheet, r) {
  sheet.getRange(r, CRM_COL_PAY_STATUS).setFormula(
    `=IF(OR(J${r}="",J${r}=0),"Pending",IF(J${r}<G${r},"Partial","Received"))`
  );
  // BUG 3 FIX: Mirror hub logic — dispatch allowed unless approval is pending (N=Yes)
  sheet.getRange(r, CRM_COL_DISPATCH_STATUS).setFormula(
    `=IF(G${r}="","",IF(OR(N${r}="No",N${r}="",N${r}="Approved"),"Ready to Dispatch","Hold – Approval Required"))`
  );
  sheet.getRange(r, CRM_COL_DUE_BALANCE).setFormula(
    `=IF(G${r}="","",G${r}-J${r})`
  );
}

function pushFinalInvoiceToMainAndHub_(
    ss, mainSheet, mainRow, mainData,
    clientName, piNo, finalInv, uuid, crmOwner, userEmail,
    precomputed) {
  // precomputed: optional { itemVal, devChg, freight, discount } from PATH 5
  // When provided, skips re-reading formula columns from sub-tab (avoids stale ARRAYFORMULA reads)

  Logger.log('pushFinalInvoiceToMainAndHub_: START | client=' + clientName +
    ' pi=' + piNo + ' finalInv=' + finalInv + ' uuid=' + uuid);

  if (!finalInv && finalInv !== 0) {
    Logger.log('pushFinalInvoiceToMainAndHub_: finalInv is empty — skipping write');
    return;
  }

  try {
    // ── Write this order's Final Invoice to Order Ledger CRM col G ───
    // Order Ledger CRM has ONE ROW PER ORDER (one row per PI No).
    // Write the single row's finalInv directly — do NOT sum across rows.
    // Each sub-tab row maps to its own Order Ledger CRM row via PI No.
    PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
    mainSheet.getRange(mainRow, CRM_COL_INV_VALUE).setValue(finalInv);
    // Also refresh M (Due Balance) and L (Dispatch Status) formulas on this row
    setCRMFormulas_(mainSheet, mainRow);
    SpreadsheetApp.flush();
    Logger.log('pushFinalInvoiceToMainAndHub_: wrote finalInv=' + finalInv + ' to Order Ledger CRM row ' + mainRow + ' (PI=' + piNo + ')');
  } catch (mainErr) {
    Logger.log('pushFinalInvoiceToMainAndHub_: Order Ledger CRM write error: ' + mainErr.message);
    PropertiesService.getScriptProperties().deleteProperty(PROP_SKIP_NEXT_EDIT);
  }

  if (!uuid) {
    Logger.log('pushFinalInvoiceToMainAndHub_: uuid empty — cannot write to Hub');
    return;
  }

  let hubSS;
  try {
    hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
  } catch (openErr) {
    Logger.log('pushFinalInvoiceToMainAndHub_: Cannot open Hub — ' + openErr.message);
    return;
  }

  // ── Read full sub-tab row for complete data push ────────────
  // The sub-tab sheet name = clientName. We need to find the row
  // in the sub-tab that has this piNo in col A.
  const subSheet = ss.getSheetByName(clientName);
  let subRowData = [];
  if (subSheet) {
    const subData = subSheet.getDataRange().getValues();
    for (let i = 1; i < subData.length; i++) {
      if ((subData[i][SUB_COL_PI_NO - 1] || '').toString().trim() === piNo.toString().trim()) {
        // Re-read this row with formula values resolved
        subRowData = subSheet.getRange(i + 1, 1, 1,
          Math.max(subSheet.getLastColumn(), SUB_COL_DUE_DATE)).getValues()[0];
        break;
      }
    }
  }

  const invNo    = subRowData[SUB_COL_INV_NO   - 1] || '';
  const invDate  = subRowData[SUB_COL_INV_DATE - 1] || '';
  const ordType  = subRowData[SUB_COL_ORDER_TYPE - 1] || '';
  // BUG FIX v14.1: use precomputed values from PATH 5 when available
  // (avoids reading stale ARRAYFORMULA results from FOB/CIF/Final cols)
  const itemVal  = (precomputed && precomputed.itemVal  != null) ? precomputed.itemVal  : (Number(subRowData[SUB_COL_ITEM_VALUE_USD  - 1]) || 0);
  const devChg   = (precomputed && precomputed.devChg   != null) ? precomputed.devChg   : (Number(subRowData[SUB_COL_DEV_CHARGES_USD - 1]) || 0);
  const freight  = (precomputed && precomputed.freight  != null) ? precomputed.freight  : (Number(subRowData[SUB_COL_FREIGHT  - 1]) || 0);
  const discount = (precomputed && precomputed.discount != null) ? precomputed.discount : (Number(subRowData[SUB_COL_DISCOUNT - 1]) || 0);
  const awbNo    = subRowData[SUB_COL_AWB_NO   - 1] || '';
  const dispDate = subRowData[SUB_COL_DISP_DATE - 1] || '';
  // v15.3 FIX: this was `|| 0`, not `|| ''`. Every other field here correctly
  // falls back to '' when the sub-tab cell is blank, which upsertOrderLedgerRow_'s
  // setVal() guard skips (val === '' → don't write, leave Hub's existing value
  // alone). But 0 is a real number — it sailed straight past that guard and
  // OVERWROTE Accounts' real Payment Received on the Hub with 0, every single
  // time CRM touched Item Value/Dev Charges/Freight/Discount (since CRM no
  // longer has Payment Received in their own sub-tab to read a real value from).
  // This is almost certainly the "values wipe out" bug — fixed at the source.
  const payRec   = subRowData[SUB_COL_PAY_RECEIVED - 1] || '';
  const payDate  = subRowData[SUB_COL_PAY_DATE - 1] || '';
  const dueDate  = subRowData[SUB_COL_DUE_DATE - 1] || '';

  // ── Upsert to Hub Order Ledger ──────────────────────────────
  try {
    upsertOrderLedgerRow_(hubSS, {
      uuid, clientName, crmOwner, piNo, invNo, invDate,
      ordType, itemVal, devChg, freight, discount, finalInv,
      awbNo, dispDate, payRec, payDate, dueDate
    }, userEmail);
  } catch (olErr) {
    Logger.log('pushFinalInvoiceToMainAndHub_: Order Ledger upsert error: ' + olErr.message);
  }

  // Master Hub removed v15.0 — no legacy write

  if (clientName && crmOwner) {
    autoAddClientToCredit_(hubSS, clientName, crmOwner);
  }
  refreshCreditTotals_(userEmail);
  Logger.log('pushFinalInvoiceToMainAndHub_: DONE for client=' + clientName);
}

// ============================================================
// upsertOrderLedgerRow_()
// Finds existing Order Ledger row by UUID+PI No, or creates
// a new one. Updates all provided fields.
// ============================================================
function upsertOrderLedgerRow_(hubSS, data, userEmail) {
  const olSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
  if (!olSheet) {
    Logger.log('upsertOrderLedgerRow_: Order Ledger sheet not found');
    return;
  }

  const {uuid, clientName, crmOwner, piNo, invNo, invDate,
         ordType, itemVal, devChg, freight, discount, finalInv,
         awbNo, dispDate, payRec, payDate, dueDate} = data;

  const olData = olSheet.getDataRange().getValues();
  let olRow = -1;

  // Find by UUID + PI No (most specific)
  for (let i = 1; i < olData.length; i++) {
    const rowUUID = (olData[i][OL_COL_UUID  - 1] || '').toString().trim();
    const rowPI   = (olData[i][OL_COL_PI_NO - 1] || '').toString().trim();
    if (rowUUID === uuid && rowPI === piNo.toString().trim()) {
      olRow = i + 1; break;
    }
  }

  // NOTE: UUID-only fallback removed — it caused "last order only" syndrome
  // by overwriting an existing row for a different PI No of the same client.
  // Every order row in Order Ledger is keyed by UUID + PI No together.
  // If no match found, always append a brand new row.
  if (olRow === -1) {
    olRow = olSheet.getLastRow() + 1;
    Logger.log('upsertOrderLedgerRow_: NEW row ' + olRow + ' for UUID=' + uuid + ' PI=' + piNo);
  } else {
    Logger.log('upsertOrderLedgerRow_: UPDATING row ' + olRow + ' for UUID=' + uuid + ' PI=' + piNo);
  }

  // Write all fields
  const currencyCols = new Set([
    OL_COL_ITEM_VALUE, OL_COL_DEV_CHARGES, OL_COL_FREIGHT,
    OL_COL_DISCOUNT,   OL_COL_FINAL_INVOICE, OL_COL_PAY_RECEIVED,
    OL_COL_DUE_BALANCE
  ]);
  const setVal = (col, val) => {
    if (val === '' || val === null || val === undefined) return;
    const cell = olSheet.getRange(olRow, col);
    cell.setValue(val);
    // FIX v14.3: Always enforce number format on currency cols so previous
    // date formatting (from old PATH 2 bug) cannot corrupt display.
    if (currencyCols.has(col)) cell.setNumberFormat('#,##0.00');
  };
  const setDate = (col, val) => {
    if (!val) return;
    const c = olSheet.getRange(olRow, col);
    c.setValue(val);
    c.setNumberFormat('dd/mm/yyyy');
  };

  // v15.3 NEW: show the CRM person's name alongside their email in Column C,
  // matching the same "Name (email)" format Credit List's own dropdown already
  // uses. Only changes what's WRITTEN to the cell — the underlying crmOwner
  // variable (a bare email, used for lookups/protections elsewhere) is
  // untouched, so this can't break anything that depends on an exact email.
  let crmOwnerDisplay = crmOwner;
  try {
    const memberMatch = getCrmConfig_().allMembers.find(
      m => m.gmail === (crmOwner || '').toString().trim().toLowerCase()
    );
    if (memberMatch && memberMatch.name) {
      crmOwnerDisplay = memberMatch.name + ' (' + crmOwner + ')';
    }
  } catch (_) { /* fall back to bare email if lookup fails */ }

  setVal(OL_COL_UUID,          uuid);
  setVal(OL_COL_CLIENT_NAME,   clientName);
  setVal(OL_COL_CRM_OWNER,     crmOwnerDisplay);
  setVal(OL_COL_PI_NO,         piNo);
  setVal(OL_COL_INV_NO,        invNo);
  setDate(OL_COL_INV_DATE,     invDate);
  setVal(OL_COL_ORDER_TYPE,    ordType);
  setVal(OL_COL_ITEM_VALUE,    itemVal);
  setVal(OL_COL_DEV_CHARGES,   devChg);
  setVal(OL_COL_FREIGHT,       freight);
  setVal(OL_COL_DISCOUNT,      discount);
  setVal(OL_COL_FINAL_INVOICE, finalInv);
  setVal(OL_COL_AWB_NO,        awbNo);
  setDate(OL_COL_DISPATCH_DATE, dispDate);
  setVal(OL_COL_PAY_RECEIVED,  payRec);
  setDate(OL_COL_PAY_DATE,     payDate);
  setDate(OL_COL_DUE_DATE,     dueDate);

  // v15.1: Payment Status & Due Balance — script-computed (replaces the old
  // live ARRAYFORMULA, which inflated getLastRow() and misplaced new orders
  // far down the sheet). Computed from this row's own Final Invoice and
  // whatever Payment Received currently sits in the row (existing value on
  // update, or the payRec just written above on a brand-new order).
  const finalInvNum = Number(finalInv) || 0;
  const payRecNum   = (payRec !== '' && payRec != null) ? Number(payRec) || 0
                       : Number(olSheet.getRange(olRow, OL_COL_PAY_RECEIVED).getValue()) || 0;
  const dueBalanceNow = Math.max(0, finalInvNum - payRecNum);
  const payStatusNow  = dueBalanceNow <= 0 ? 'Received' : (payRecNum > 0 ? 'Partial' : 'Pending');
  setVal(OL_COL_DUE_BALANCE, dueBalanceNow);
  setVal(OL_COL_PAY_STATUS,  payStatusNow);

  SpreadsheetApp.flush();
  logActivity_(hubSS, 'UPSERT ORDER LEDGER',
    'UUID:' + uuid + ' | PI:' + piNo + ' | Client:' + clientName + ' | FinalInv:' + finalInv,
    userEmail || 'system');
}

function reEvaluateAllCreditFlags_(hubSS, userEmail) {
  try {
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    const hubSheet    = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!creditSheet || !hubSheet) return;

    // Build credit limit map from Credit sheet (fresh read)
    const creditData  = creditSheet.getDataRange().getValues();
    const limitMap    = {};  // clientName.lower -> creditLimit
    for (let i = 1; i < creditData.length; i++) {
      const name  = (creditData[i][CC_COL_CLIENT_NAME  - 1] || '').toString().trim().toLowerCase();
      const limit = Number(creditData[i][CC_COL_CREDIT_LIMIT - 1]) || 0;
      if (name && limit > 0) limitMap[name] = limit;
    }

    // Compute used credit per client from Hub (fresh read)
    const hubData  = hubSheet.getDataRange().getValues();
    const invMap   = {};  // clientName.lower -> total InvValue
    const payMap   = {};  // clientName.lower -> total PayReceived
    for (let i = 1; i < hubData.length; i++) {
      // Order Ledger uses OL_COL_* layout (not HUB_COL_*)
      const name = (hubData[i][OL_COL_CLIENT_NAME   - 1] || '').toString().trim().toLowerCase();
      const inv  = Number(hubData[i][OL_COL_FINAL_INVOICE - 1]) || 0;
      const pay  = Number(hubData[i][OL_COL_PAY_RECEIVED  - 1]) || 0;
      if (name) {
        invMap[name] = (invMap[name] || 0) + inv;
        payMap[name] = (payMap[name] || 0) + pay;
      }
    }

    // For each Hub data row, set N based on whether client total used > limit
    const config     = getCrmConfig_();
    const crmFileIds = config.crmFileIds;
    const crmSSCache = {};  // avoid re-opening same CRM file repeatedly

    for (let i = 1; i < hubData.length; i++) {
      const row        = hubData[i];
      const hubRowNum  = i + 1;
      // BUG FIX v14.1: Order Ledger uses OL_COL_* layout, NOT HUB_COL_*.
      // HUB_COL_INV_VALUE=7 (col G) but OL_COL_FINAL_INVOICE=12 (col L) — very different.
      const clientName = (row[OL_COL_CLIENT_NAME    - 1] || '').toString().trim();
      const piNo       = (row[OL_COL_PI_NO          - 1] || '').toString().trim();
      const invValue   = Number(row[OL_COL_FINAL_INVOICE - 1]) || 0;
      const uuid       = (row[OL_COL_UUID           - 1] || '').toString().trim();
      // v15.3 FIX: Column C can now contain "Name (email)" display format —
      // extractEmail_ safely returns a bare email either way.
      const crmOwner   = extractEmail_(row[OL_COL_CRM_OWNER - 1]);
      // v15.2 FIX: this used to be hardcoded to '' with a comment saying Order
      // Ledger had no Approval Required column. It now does (col U) — read the
      // REAL current value. This also fixes a duplicate-email bug: with a
      // hardcoded '', currentN could never equal 'Yes', so the AUTO-FLAG branch
      // (and its hold-notification email) re-fired on every single call for
      // any client still over their limit, instead of only once when the
      // state actually changed from No → Yes.
      const currentN   = (row[OL_COL_APPROVAL_REQ   - 1] || '').toString().trim();

      if (!clientName || !piNo || !invValue) continue; // skip registration-only rows

      const key   = clientName.toLowerCase();
      const limit = limitMap[key];
      if (!limit) continue;

      const used       = Math.max(0, (invMap[key] || 0) - (payMap[key] || 0));
      const shouldFlag = used > limit;

      if (shouldFlag && currentN !== 'Approved') {
        if (currentN !== 'Yes') {
          // v15.2: write the real column now that it exists. Clear Approved By
          // (V) too — a fresh breach needs a fresh sign-off, even if this same
          // order was approved once before and later paid back down.
          hubSheet.getRange(hubRowNum, OL_COL_APPROVAL_REQ).setValue('Yes');
          hubSheet.getRange(hubRowNum, OL_COL_APPROVED_BY).setValue('');
          pushNToCRM_(uuid, crmOwner, 'Yes', crmFileIds, crmSSCache);
          logActivity_(hubSS, 'AUTO-FLAG N=Yes',
            `UUID:${uuid} | ${clientName} | Used:${used} > Limit:${limit}`, userEmail || 'system');
          sendHoldNotificationEmails_(hubSS, uuid, clientName, crmOwner,
            row[OL_COL_INV_NO-1], row[OL_COL_FINAL_INVOICE-1], used, limit, userEmail);
        }
      } else if (!shouldFlag && currentN === 'Yes') {
        // v15.2: write the real column now that it exists.
        hubSheet.getRange(hubRowNum, OL_COL_APPROVAL_REQ).setValue('No');
        pushNToCRM_(uuid, crmOwner, 'No', crmFileIds, crmSSCache);
        logActivity_(hubSS, 'AUTO-CLEAR N=No',
          `UUID:${uuid} | ${clientName} | Used:${used} ≤ Limit:${limit}`, userEmail || 'system');
      } else if (!currentN && !shouldFlag) {
        // First time this row is evaluated and it's fine — set the baseline
        // "No" instead of leaving the cell blank forever.
        hubSheet.getRange(hubRowNum, OL_COL_APPROVAL_REQ).setValue('No');
      }
    }
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('ERROR in reEvaluateAllCreditFlags_: ' + err.message);
  }
}

function menuRepairAllSheetSecurity_() {
  try {
    const hubSS      = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = Session.getEffectiveUser().getEmail();
    const config     = getCrmConfig_();

    const adminOnly  = [ownerEmail, ...config.adminEmails].filter((v, i, a) => a.indexOf(v) === i);
    const salesEditors = [...new Set([
      ownerEmail, ...config.adminEmails,
      ...config.salesManagerEmails, ...config.salesEmails
    ])];

    let report = '✅ Sheet Security Repaired:\n\n';

    // ── 1. Activity Log ──────────────────────────────────────
    const logSheet = hubSS.getSheetByName(LOG_SHEET_NAME);
    if (logSheet) {
      applyActivityLogSecurity_(hubSS, logSheet);
      report += '• Activity Log: Admin only, tab hidden\n';
    }

    // ── 2. CRM Members ───────────────────────────────────────
    const membSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
    if (membSheet) {
      membSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
      membSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
      const membProt = membSheet.getRange(2, 1, 999, MC_TOTAL_COLS)
        .protect().setDescription('Admin/Sales/SalesMgr only — CRM Members data');
      membProt.removeEditors(membProt.getEditors());
      membProt.addEditors(salesEditors);
      membProt.setWarningOnly(false);
      membSheet.hideSheet();
      report += '• CRM Members: Admin/Sales/SalesMgr only, tab hidden\n';
    }

    // ── 3. Client Credit col K (UUID) ────────────────────────
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (creditSheet) {
      // Ensure col K is hidden and protected
      try { creditSheet.hideColumns(11); } catch(_) {}
      const kProt = creditSheet.getRange(1, 11, 999, 1)
        .protect().setDescription('Admin only — UUID internal column (K)');
      kProt.removeEditors(kProt.getEditors());
      kProt.addEditors(adminOnly);
      kProt.setWarningOnly(false);
      report += '• Client Credit col K (UUID): hidden, Admin only\n';
    }

    // ── 4. Hub + Credit full protection repair ────────────────
    menuRepairHubProtection_();
    menuRepairCreditProtection_();
    report += '• Master Hub columns: role-based protections refreshed\n';
    report += '• Client Credit columns: role-based protections refreshed\n';

    logActivity_(hubSS, 'SECURITY REPAIR', 'All sheet security repaired', ownerEmail);
    alertIfUi_(report);

  } catch (err) {
    Logger.log('ERROR in menuRepairAllSheetSecurity_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function menuRefreshCreditTotals_() {
  const userEmail = Session.getActiveUser().getEmail();
  refreshCreditTotals_(userEmail);
  alertIfUi_('Credit totals + col N refreshed!\nUsed Credit, Available Credit, and Approval Required (N) updated from Hub data.');
}

function diagnoseTriggers() {
  try {
    const currentSS   = SpreadsheetApp.getActive();
    const currentSSId = currentSS.getId();
    const isHub       = currentSSId === MASTER_HUB_ID;
    const allTriggers = ScriptApp.getProjectTriggers();

    const thisFileTriggers  = [];
    const otherFileTriggers = [];

    allTriggers.forEach(t => {
      const entry = {
        fn:     t.getHandlerFunction(),
        type:   t.getEventType().toString(),
        fileId: t.getTriggerSourceId(),
      };
      if (t.getTriggerSourceId() === currentSSId) {
        thisFileTriggers.push(entry);
      } else {
        otherFileTriggers.push(entry);
      }
    });

    const required = isHub
      ? ['onMasterHubEdit', 'onCreditRequestEdit', 'onMembersSheetEdit']
      : ['onCRMEdit', 'onCRMSubTabEdit'];

    const found    = thisFileTriggers.map(t => t.fn);
    const missing  = required.filter(fn => !found.includes(fn));
    const extra    = found.filter(fn => !required.includes(fn) && fn !== 'addCustomMenu');

    let msg = '=== Trigger Diagnosis ===\n';
    msg += 'File: ' + currentSS.getName() + '\n';
    msg += 'File ID: ' + currentSSId + '\n';
    msg += 'Role detected: ' + (isHub ? 'MASTER HUB' : 'CRM SPOKE') + '\n\n';

    msg += '--- Triggers on THIS file (' + thisFileTriggers.length + ') ---\n';
    if (thisFileTriggers.length === 0) {
      msg += '  ⚠️ NONE — triggers are missing! Run Setup Triggers.\n';
    } else {
      thisFileTriggers.forEach(t => { msg += '  ✅ ' + t.fn + ' (' + t.type + ')\n'; });
    }

    if (missing.length > 0) {
      msg += '\n⛔ MISSING required triggers:\n';
      missing.forEach(fn => { msg += '  • ' + fn + '\n'; });
      msg += '\n→ Run: CRM Tools > Setup > Setup triggers\n';
    }

    if (extra.length > 0) {
      msg += '\n⚠️ Extra/unexpected triggers on this file:\n';
      extra.forEach(fn => { msg += '  • ' + fn + '\n'; });
    }

    if (otherFileTriggers.length > 0) {
      msg += '\n--- Triggers on OTHER files (' + otherFileTriggers.length + ') ---\n';
      otherFileTriggers.forEach(t => {
        msg += '  • ' + t.fn + ' → fileId: ' + t.fileId + '\n';
      });
    }

    if (missing.length === 0 && thisFileTriggers.length > 0) {
      msg += '\n✅ All required triggers are correctly installed!';
    }

    // ── v10.0: Also verify Hub reachability ─────────────────
    msg += '\n\n--- Hub Reachability Check ---\n';
    try {
      const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
      msg += '✅ Hub reachable: ' + hubSS.getName() + '\n';
      msg += '   Current user: ' + Session.getEffectiveUser().getEmail();
    } catch (reachErr) {
      msg += '❌ Hub NOT reachable! Error: ' + reachErr.message + '\n';
      msg += '   This means triggers will fail silently on every sub-tab edit.\n';
      msg += '   Fix: run setupAllTriggers() as the account that has Hub Editor access.';
    }

    Logger.log(msg);
    alertIfUi_(msg);

  } catch (err) {
    Logger.log('ERROR in diagnoseTriggers: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function menuForceSubTabSync_() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = getCRMMainSheet_(ss);
    if (!mainSheet) { alertIfUi_('Main sheet not found.'); return; }

    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { alertIfUi_('Master Hub sheet not found.'); return; }

    const userEmail  = Session.getActiveUser().getEmail();
    const mainData   = mainSheet.getDataRange().getValues();
    const mainDisp   = mainSheet.getDataRange().getDisplayValues();
    const hubData    = hubSheet.getDataRange().getValues();

    const systemSheets = new Set([
      MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME
    ]);

    let rowsUpdated = 0;

    ss.getSheets().forEach(sheet => {
      if (systemSheets.has(sheet.getName())) return;
      const sheetName = sheet.getName();
      const subData   = sheet.getDataRange().getValues();

      for (let s = 1; s < subData.length; s++) {
        const piNo       = subData[s][SUB_COL_PI_NO   - 1];
        const invNo      = subData[s][SUB_COL_INV_NO  - 1];
        const invDate    = subData[s][SUB_COL_INV_DATE - 1];
        const awbNo      = subData[s][SUB_COL_AWB_NO  - 1];
        const dispDate   = subData[s][SUB_COL_DISP_DATE - 1];
        const payRec     = subData[s][SUB_COL_PAY_RECEIVED - 1];
        const payDate    = subData[s][SUB_COL_PAY_DATE - 1];
        const finalInv   = sheet.getRange(s + 1, SUB_COL_FINAL_INV).getValue();
        if (!piNo) continue;

        // ── Find Main row ───────────────────────────────────
        let mainRow = -1, uuid = null, crmOwner = null;
        for (let m = 1; m < mainData.length; m++) {
          if (mainData[m][CRM_COL_PI_NO - 1] === piNo) {
            mainRow  = m + 1;
            uuid     = mainData[m][CRM_COL_UUID      - 1];
            crmOwner = mainData[m][CRM_COL_CRM_OWNER - 1];
            break;
          }
        }
        // Fallback: name match
        if (mainRow === -1) {
          for (let m = 1; m < mainData.length; m++) {
            const mName = (mainDisp[m][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
            if (mName === sheetName.toLowerCase()) {
              mainRow  = m + 1;
              uuid     = mainData[m][CRM_COL_UUID      - 1];
              crmOwner = mainData[m][CRM_COL_CRM_OWNER - 1];
              if (!mainData[m][CRM_COL_PI_NO - 1]) mainSheet.getRange(mainRow, CRM_COL_PI_NO).setValue(piNo);
              break;
            }
          }
        }
        if (mainRow === -1) continue;

        // ── Write to CRM Main ────────────────────────────────
        if (invNo)    mainSheet.getRange(mainRow, CRM_COL_INV_NO).setValue(invNo);
        if (invDate) {
          const dc = mainSheet.getRange(mainRow, CRM_COL_DATE);
          dc.setValue(invDate); dc.setNumberFormat('dd/mm/yyyy');
        }
        if (finalInv) mainSheet.getRange(mainRow, CRM_COL_INV_VALUE).setValue(finalInv);
        if (payRec !== '' && payRec != null) mainSheet.getRange(mainRow, CRM_COL_PAY_RECEIVED).setValue(payRec);
        if (payDate !== '' && payDate != null) {
          const pdc = mainSheet.getRange(mainRow, CRM_COL_PAY_DATE);
          pdc.setValue(payDate); pdc.setNumberFormat('dd/mm/yyyy');
        }

        // ── Recover UUID from Hub if missing ─────────────────
        if (!uuid) {
          for (let h = 1; h < hubData.length; h++) {
            const hName = (hubData[h][HUB_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
            if (hName === sheetName.toLowerCase()) {
              uuid = (hubData[h][HUB_COL_UUID - 1] || '').toString().trim();
              if (uuid) { mainSheet.getRange(mainRow, CRM_COL_UUID).setValue(uuid); }
              if (!crmOwner) crmOwner = (hubData[h][HUB_COL_CRM_OWNER - 1] || '').toString().trim();
              break;
            }
          }
        }
        if (!uuid) continue;

        // ── Write to Hub ─────────────────────────────────────
        let hubRow = -1;
        for (let h = 1; h < hubData.length; h++) {
          if (hubData[h][HUB_COL_UUID - 1] === uuid) { hubRow = h + 1; break; }
        }
        if (hubRow === -1) continue;

        if (invNo)   hubSheet.getRange(hubRow, HUB_COL_INV_NO).setValue(invNo);
        if (invDate) {
          const hdc = hubSheet.getRange(hubRow, HUB_COL_DATE);
          hdc.setValue(invDate); hdc.setNumberFormat('dd/mm/yyyy');
        }
        if (finalInv) hubSheet.getRange(hubRow, HUB_COL_INV_VALUE).setValue(finalInv);
        if (awbNo)    hubSheet.getRange(hubRow, HUB_COL_AWB_NO).setValue(awbNo);
        if (dispDate) {
          const ddc = hubSheet.getRange(hubRow, HUB_COL_DISP_DATE);
          ddc.setValue(dispDate); ddc.setNumberFormat('dd/mm/yyyy');
        }
        if (payRec !== '' && payRec != null) hubSheet.getRange(hubRow, HUB_COL_PAY_RECEIVED).setValue(payRec);
        if (payDate !== '' && payDate != null) {
          const phdc = hubSheet.getRange(hubRow, HUB_COL_PAY_DATE);
          phdc.setValue(payDate); phdc.setNumberFormat('dd/mm/yyyy');
        }

        // ── Ensure Credit sheet has this client ──────────────
        const clientDisplayName = mainSheet.getRange(mainRow, CRM_COL_CUSTOMER).getDisplayValue() || sheetName;
        autoAddClientToCredit_(hubSS, clientDisplayName, crmOwner || '');

        rowsUpdated++;
        logActivity_(hubSS, 'FORCE SUBTAB SYNC', 'PI:' + piNo + ' | ' + sheetName, userEmail);
      }
    });

    SpreadsheetApp.flush();

    // Refresh credit totals after all writes
    try { refreshCreditTotals_(userEmail); } catch (_) {}

    alertIfUi_(
      '✅ Force Sub-Tab Sync complete!\n\n' +
      rowsUpdated + ' row(s) synced to CRM Main + Master Hub.\n\n' +
      'Check Master Hub now — PI No, Inv No, Final Invoice,\n' +
      'AWB No, Dispatch Date, and Payment columns should be updated.'
    );

  } catch (err) {
    Logger.log('ERROR in menuForceSubTabSync_: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function pushNToCRM_(uuid, crmOwner, nValue, crmFileIds, crmSSCache) {
  try {
    if (!uuid || !crmOwner) return;
    const fileId = crmFileIds[crmOwner.toLowerCase()];
    if (!fileId) return;
    if (!crmSSCache[fileId]) {
      crmSSCache[fileId] = SpreadsheetApp.openById(fileId);
    }
    const crmSS    = crmSSCache[fileId];
    const crmSheet = getCRMMainSheet_(crmSS) || crmSS.getSheets()[0];
    const crmData  = crmSheet.getDataRange().getValues();
    for (let i = 1; i < crmData.length; i++) {
      if ((crmData[i][CRM_COL_UUID - 1] || '').toString() === uuid) {
        crmSheet.getRange(i + 1, CRM_COL_APPROVAL_REQ).setValue(nValue);
        break;
      }
    }
  } catch (err) {
    Logger.log('pushNToCRM_ error for UUID:' + uuid + ' — ' + err.message);
  }
}

function sendHoldNotificationEmails_(hubSS, uuid, customerName, crmOwner, invNo, invValue, usedCredit, creditLimit, userEmail) {
  try {
    const config             = getCrmConfig_();
    const salesManagerEmails = config.salesManagerEmails;
    const salesEmails        = config.salesEmails;
    const adminEmails        = config.adminEmails;
    const recipients         = [...new Set([...salesManagerEmails, ...salesEmails, ...adminEmails])];
    if (!recipients.length) return;

    const overage = Math.max(0, usedCredit - creditLimit);
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy');

    const htmlBody = `
      <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:620px;">
        <div style="background:#B71C1C;padding:16px 24px;border-radius:6px 6px 0 0;">
          <span style="color:#fff;font-size:17px;font-weight:bold;">Dispatch On Hold - Credit Limit Exceeded</span>
        </div>
        <div style="border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 6px 6px;">
          <p style="margin:0 0 12px;">Dear Team,</p>
          <p>An order for <strong>${customerName}</strong> has been placed <strong style="color:#B71C1C;">on hold</strong>
             because the client's credit limit has been exceeded. Logistics <strong>cannot dispatch</strong> until this is resolved.</p>
          <table style="border-collapse:collapse;font-size:13px;width:100%;margin:14px 0;">
            <tr style="background:#B71C1C;color:#fff;">
              <th style="padding:8px 12px;text-align:left;" colspan="2">Hold Details — ${today}</th>
            </tr>
            <tr style="background:#FFEBEE;">
              <td style="padding:7px 12px;border:1px solid #ddd;">Client</td>
              <td style="padding:7px 12px;border:1px solid #ddd;"><strong>${customerName}</strong></td>
            </tr>
            <tr>
              <td style="padding:7px 12px;border:1px solid #ddd;">Invoice No</td>
              <td style="padding:7px 12px;border:1px solid #ddd;">${invNo || '—'}</td>
            </tr>
            <tr style="background:#FFEBEE;">
              <td style="padding:7px 12px;border:1px solid #ddd;">Invoice Value</td>
              <td style="padding:7px 12px;border:1px solid #ddd;">$${Number(invValue || 0).toLocaleString('en-US')}</td>
            </tr>
            <tr>
              <td style="padding:7px 12px;border:1px solid #ddd;">CRM Owner</td>
              <td style="padding:7px 12px;border:1px solid #ddd;">${crmOwner || '—'}</td>
            </tr>
            <tr style="background:#FFEBEE;">
              <td style="padding:7px 12px;border:1px solid #ddd;">Credit Limit</td>
              <td style="padding:7px 12px;border:1px solid #ddd;">$${Number(creditLimit).toLocaleString('en-US')}</td>
            </tr>
            <tr>
              <td style="padding:7px 12px;border:1px solid #ddd;">Total Used Credit</td>
              <td style="padding:7px 12px;border:1px solid #ddd;color:#C62828;font-weight:bold;">$${Number(usedCredit).toLocaleString('en-US')}</td>
            </tr>
            <tr style="background:#FFEBEE;">
              <td style="padding:7px 12px;border:1px solid #ddd;"><strong>Amount Over Limit</strong></td>
              <td style="padding:7px 12px;border:1px solid #ddd;color:#B71C1C;font-weight:bold;">$${Number(overage).toLocaleString('en-US')}</td>
            </tr>
          </table>
          <p><strong>To resolve — choose one:</strong></p>
          <ul style="font-size:13px;line-height:1.8;">
            <li><strong>Sales Manager:</strong> Approve the order — open Master Hub -> col O (Approved By) -> type your name. Logistics will be unblocked immediately.</li>
            <li><strong>Admin / Sales:</strong> Increase the credit limit — open Client Credit sheet -> update col C (Credit Limit) for <strong>${customerName}</strong>.</li>
          </ul>
          <p style="color:#999;font-size:12px;margin-top:16px;">— UV Overseas CRM System (Auto-generated) | UUID: ${uuid}</p>
        </div>
      </div>`;

    recipients.forEach(email => {
      try {
        GmailApp.sendEmail(
          email,
          `Dispatch On Hold - ${customerName} | Credit Limit Exceeded - Action Required`,
          `Order for ${customerName} is on hold. Credit used $${usedCredit} exceeds limit $${creditLimit}. Approve in col O or increase credit limit.`,
          { htmlBody, name: 'UV Overseas CRM System' }
        );
      } catch (mailErr) {
        Logger.log('sendHoldNotificationEmails_: failed for ' + email + ' — ' + mailErr.message);
      }
    });

    logActivity_(hubSS, 'HOLD EMAIL SENT',
      `Client:${customerName} | Used:$${usedCredit} > Limit:$${creditLimit} | To:${recipients.join(', ')}`, userEmail || 'system');
  } catch (err) {
    Logger.log('ERROR in sendHoldNotificationEmails_: ' + err.message);
  }
}

function logActivity_(ss, action, detail, user) {
  try {
    let log = ss.getSheetByName(LOG_SHEET_NAME);
    if (!log) {
      log = ss.insertSheet(LOG_SHEET_NAME);
      const hdr = ['Timestamp', 'User (Gmail)', 'Action', 'Detail / Reference'];
      log.getRange(1, 1, 1, hdr.length).setValues([hdr]);
      log.getRange(1, 1, 1, hdr.length)
        .setFontWeight('bold').setBackground('#F57F17').setFontColor('#FFFFFF');
      log.setFrozenRows(1);
      log.setColumnWidth(1, 160);
      log.setColumnWidth(2, 200);
      log.setColumnWidth(3, 180);
      log.setColumnWidth(4, 320);
      // BUG FIX v13: Protect Activity Log on first creation.
      // No role should be able to edit or delete log entries.
      // Hide the sheet tab so Accounts/Logistics don't see it.
      applyActivityLogSecurity_(ss, log);
    }
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    log.appendRow([ts, user || 'unknown', action, detail || '']);
  } catch (err) {
    // Protection errors are expected if CRM email not yet added to Hub protection list.
    // Run fixHubWriteAccess() from Hub to grant CRM emails write access.
    Logger.log('ERROR in logActivity_: ' + err.message +
      (err.message.includes('protected') ? ' — run fixHubWriteAccess() from Hub to fix.' : ''));
  }
}

function applyActivityLogSecurity_(ss, logSheet) {
  try {
    if (!logSheet) logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) return;

    const ownerEmail = Session.getEffectiveUser().getEmail();
    const config     = getCrmConfig_();
    // v15.1 FIX: adminOnly was missing every operational role's emails.
    // logActivity_() is called from dozens of places across this codebase,
    // executing under whichever role's authorization triggered that code path
    // (CRM, Logistics, Accounts, Sales — not just Admin). Excluding them here
    // caused "You are trying to edit a protected cell or object" on nearly
    // every logActivity_ call made outside an Admin-run context. The sheet is
    // still hidden and still blocks casual manual editing — this only adds
    // the identities the SCRIPT itself needs to append rows on their behalf.
    const crmSpokeEmailsForLog = Object.keys(config.crmFileIds);
    const adminOnly  = [...new Set([
      ownerEmail, ...config.adminEmails, ...config.salesEmails,
      ...config.logisticsEmails, ...config.accountsEmails, ...crmSpokeEmailsForLog
    ])];

    // Remove any existing protections
    logSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    logSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());

    // Protect entire sheet — Admin only, no one else can edit
    const prot = logSheet.protect().setDescription('Activity Log — Admin read-only, no edits');
    prot.removeEditors(prot.getEditors());
    prot.addEditors(adminOnly);
    prot.setWarningOnly(false);

    // Hide sheet tab — Accounts and Logistics should not see this
    logSheet.hideSheet();

    Logger.log('applyActivityLogSecurity_: Activity Log protected and hidden.');
  } catch (err) {
    Logger.log('ERROR in applyActivityLogSecurity_: ' + err.message);
  }
}

function cleanupHubTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    ScriptApp.deleteTrigger(t);
    Logger.log('Deleted: ' + t.getHandlerFunction());
  });
  alertIfUi_('All triggers cleaned up from Hub.');
}

function nukeClientCreditProtections() {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) {
      Logger.log('nukeClientCreditProtections: Client Credit sheet not found.');
      alertIfUi_('Client Credit sheet not found.');
      return;
    }

    // Step 1: Remove every single range protection — no filtering
    const all = creditSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    Logger.log('nukeClientCreditProtections: removing ' + all.length + ' protection(s)...');
    all.forEach((p, i) => {
      Logger.log('  Removing: "' + p.getDescription() + '" on range ' + p.getRange().getA1Notation());
      p.remove();
    });

    // Step 2: Reapply protections per new role model:
    //   • Cols A-C, F-G, J: Sales Manager + Sales + Admin (add clients, edit all details)
    //   • Cols D-E: Admin only (formula columns — do not overwrite)
    //   • Cols H-I: Admin only (Request Status, Approved By)
    const ownerEmail  = Session.getEffectiveUser().getEmail();
    const config      = getCrmConfig_();
    const adminEmails = [ownerEmail, ...config.adminEmails].filter((v, i, a) => a.indexOf(v) === i);
    const fullEditors = [...new Set([
      ownerEmail,
      ...config.adminEmails,
      ...config.salesManagerEmails,
      ...config.salesEmails
    ])];

    // Cols A–C: Sales Manager + Sales + Admin
    const abcProt = creditSheet.getRange(2, 1, 999, 3)
      .protect().setDescription('Sales/SalesMgr/Admin — Client Name, CRM Owner, Credit Limit (A–C)');
    abcProt.removeEditors(abcProt.getEditors());
    abcProt.addEditors(fullEditors);
    abcProt.setWarningOnly(false);

    // Cols D–E: Admin only (formula columns)
    const deProt = creditSheet.getRange(2, CC_COL_USED_CREDIT, 999, 2)
      .protect().setDescription('Admin only — Used Credit & Available Credit formula cols (D–E)');
    deProt.removeEditors(deProt.getEditors());
    deProt.addEditors(adminEmails);
    deProt.setWarningOnly(false);

    // Cols F–G: Sales Manager + Sales + Admin
    const fgProt = creditSheet.getRange(2, CC_COL_REQ_INCREASE, 999, 2)
      .protect().setDescription('Sales/SalesMgr/Admin — Request Limit & Reason (F–G)');
    fgProt.removeEditors(fgProt.getEditors());
    fgProt.addEditors(fullEditors);
    fgProt.setWarningOnly(false);

    // Cols H–I: Admin only
    const sProt = creditSheet.getRange(2, CC_COL_REQ_STATUS, 999, 2)
      .protect().setDescription('Admin only — Request Status & Approved By (H-I)');
    sProt.removeEditors(sProt.getEditors());
    sProt.addEditors(adminEmails);
    sProt.setWarningOnly(false);

    // Col J: Sales Manager + Sales + Admin
    const jProt = creditSheet.getRange(2, CC_COL_LAST_UPDATED, 999, 1)
      .protect().setDescription('Sales/SalesMgr/Admin — Last Updated (J)');
    jProt.removeEditors(jProt.getEditors());
    jProt.addEditors(fullEditors);
    jProt.setWarningOnly(false);

    logActivity_(hubSS, 'NUKE CREDIT PROTECTIONS',
      `Removed ${all.length} protection(s). New model: A-C,F-G,J=Sales/SalesMgr/Admin; D-E,H-I=Admin only.`, ownerEmail);

    Logger.log('nukeClientCreditProtections: done. New permission model applied.');
    alertIfUi_(
      '✅ Done!\n\n' +
      `Removed ${all.length} old protection(s) from Client Credit sheet.\n\n` +
      'New permissions:\n' +
      'Cols A-C (Client Name, CRM Owner, Credit Limit): Sales Manager + Sales + Admin\n' +
      'Cols D-E (Used/Available Credit): Admin only (formula cols)\n' +
      'Cols F-G (Request New Limit, Reason): Sales Manager + Sales + Admin\n' +
      'Cols H-I (Request Status, Approved By): Admin only\n' +
      'Col J (Last Updated): Sales Manager + Sales + Admin\n\n' +
      'Sales Manager, Sales, and Admin can now add new clients and edit all columns.'
    );
  } catch (err) {
    Logger.log('ERROR in nukeClientCreditProtections: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}

function fixClientTabValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const systemSheets = new Set([
    MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
    CREDIT_SHEET_NAME
  ]);

  ss.getSheets().forEach(sheet => {
    if (systemSheets.has(sheet.getName())) return;
    if (sheet.getLastColumn() < 13) return;

    // Clear ALL data validations by setting null over entire sheet range
    sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getLastColumn())
      .setDataValidation(null);

    // Re-apply ONLY the correct Order Type dropdown on col D (SUB_COL_ORDER_TYPE = 4)
    const orderTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Production', 'Sample', 'Showroom'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, SUB_COL_ORDER_TYPE, 999, 1).setDataValidation(orderTypeRule);

    Logger.log('Fixed validations on: ' + sheet.getName());
  });

  SpreadsheetApp.flush();
  alertIfUi_('✅ Client tab validations fixed. Order Type dropdown now on col D only.');
}

function syncSubTabsToMainUnprotected() {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { alertIfUi_('Hub sheet not found.'); return; }

    const hubData    = hubSheet.getDataRange().getValues();
    const userEmail  = Session.getEffectiveUser().getEmail();
    const config     = getCrmConfig_();
    let count = 0;

    // Loop every CRM spoke file
    for (const [gmail, fileId] of Object.entries(config.crmFileIds)) {
      if (!fileId) continue;
      try {
        const crmSS    = SpreadsheetApp.openById(fileId);
        const mainSheet = getCRMMainSheet_(crmSS);
        if (!mainSheet) continue;

        const systemSheets = new Set([
          MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
          CREDIT_SHEET_NAME
        ]);

        crmSS.getSheets().forEach(sheet => {
          if (systemSheets.has(sheet.getName())) return;
          if (sheet.getLastColumn() < 13) return;

          const sheetName = sheet.getName();
          const subData   = sheet.getDataRange().getValues();
          const mainData  = mainSheet.getDataRange().getValues();
          const mainDisp  = mainSheet.getDataRange().getDisplayValues();

          for (let s = 1; s < subData.length; s++) {
            const piNo    = subData[s][SUB_COL_PI_NO   - 1];
            const invNo   = subData[s][SUB_COL_INV_NO  - 1];
            const invDate = subData[s][SUB_COL_INV_DATE - 1];
            const finalInv = sheet.getRange(s + 1, SUB_COL_FINAL_INV).getValue();
            if (!piNo) continue;

            // Find Main row
            let mainRow = -1, uuid = null;
            for (let m = 1; m < mainData.length; m++) {
              if (mainData[m][CRM_COL_PI_NO - 1] === piNo) {
                mainRow = m + 1; uuid = mainData[m][CRM_COL_UUID - 1]; break;
              }
            }
            if (mainRow === -1) {
              for (let m = 1; m < mainData.length; m++) {
                if ((mainDisp[m][CRM_COL_CUSTOMER-1]||'').toString().trim().toLowerCase()
                    === sheetName.toLowerCase()) {
                  mainRow = m + 1; uuid = mainData[m][CRM_COL_UUID - 1]; break;
                }
              }
            }
            if (mainRow === -1) continue;

            // Write to CRM Main (Hub script runs as owner — bypasses protections)
            if (invNo) mainSheet.getRange(mainRow, CRM_COL_INV_NO).setValue(invNo);
            if (invDate) {
              const dc = mainSheet.getRange(mainRow, CRM_COL_DATE);
              dc.setValue(invDate);
              dc.setNumberFormat('dd/mm/yyyy');
            }
            if (finalInv) mainSheet.getRange(mainRow, CRM_COL_INV_VALUE).setValue(finalInv);

            // Recover UUID from Hub by name if missing
            if (!uuid) {
              for (let h = 1; h < hubData.length; h++) {
                if ((hubData[h][HUB_COL_CUSTOMER-1]||'').toString().trim().toLowerCase()
                    === sheetName.toLowerCase()) {
                  uuid = (hubData[h][HUB_COL_UUID-1]||'').toString().trim();
                  if (uuid) mainSheet.getRange(mainRow, CRM_COL_UUID).setValue(uuid);
                  break;
                }
              }
            }
            if (!uuid) continue;

            // Write to Hub
            let hubRow = -1;
            for (let h = 1; h < hubData.length; h++) {
              if (hubData[h][HUB_COL_UUID-1] === uuid) { hubRow = h + 1; break; }
            }
            if (hubRow === -1) continue;

            if (invNo)   hubSheet.getRange(hubRow, HUB_COL_INV_NO).setValue(invNo);
            if (invDate) {
              const hdc = hubSheet.getRange(hubRow, HUB_COL_DATE);
              hdc.setValue(invDate);
              hdc.setNumberFormat('dd/mm/yyyy');
            }
            if (finalInv) hubSheet.getRange(hubRow, HUB_COL_INV_VALUE).setValue(finalInv);

            count++;
            logActivity_(hubSS, 'SUBTAB SYNC', 'PI:' + piNo + ' | ' + sheetName, userEmail);
          }
        });

        SpreadsheetApp.flush();
      } catch(e) {
        Logger.log('syncSubTabsToMainUnprotected: error for ' + gmail + ' — ' + e.message);
      }
    }

    alertIfUi_('✅ Sync done. ' + count + ' row(s) updated in CRM Main + Hub.');
  } catch(err) {
    Logger.log('ERROR in syncSubTabsToMainUnprotected: ' + err.message);
    alertIfUi_('Error: ' + err.message);
  }
}


// ══════════════════════════════════════════════════════════════════
// LICENSE GATE — Simple onEdit (always has e.user, no network needed)
// ══════════════════════════════════════════════════════════════════
// Architecture:
//   1. A time-based trigger runs refreshLicenseStatus_() every 5 min
//      → fetches License sheet + Members sheet → stores result in
//      DocumentProperties keyed by email: "lic_EMAIL" = "1" or "0"
//   2. Simple onEdit reads the local property — instant, no network.
//      If status is "0" → clear cell + show alert.
//   3. When you change a license in your sheet, it takes effect within
//      5 minutes automatically, or instantly via syncLicenseNow().

function onEdit(e) {
  if (!e || !e.user || !e.user.email) return;
  const userEmail = e.user.email.toLowerCase().trim();
  const props     = PropertiesService.getDocumentProperties();
  const status    = props.getProperty('lic_' + userEmail);

  // "0" = explicitly blocked. Anything else (null = not yet synced) = allow.
  // This means first edit before first sync always passes — sync runs on open too.
  if (status === '0') {
    try { e.range.clearContent(); } catch (_) {}
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'LICENSE ISSUE: Your license has expired or is not active. Contact: work.mannnish@gmail.com | WhatsApp: +91-9511565884',
      'License Issue', 10
    );
  }
}

// ── Sync license status from License sheet into DocumentProperties ──
// Run by time-based trigger every 5 min AND on file open.
// Stores "lic_EMAIL" = "1" (valid) or "0" (blocked) per email.
function refreshLicenseStatus_() {
  try {
    const url = PropertiesService.getScriptProperties().getProperty('LICENSE_URL');
    if (!url) return;

    const props   = PropertiesService.getDocumentProperties();
    const today   = new Date();
    const hubId   = MASTER_HUB_ID;

    // Build a set of all emails that are ACTIVE + not expired in License sheet
    const validEmails = new Set();

    try {
      const licSS    = SpreadsheetApp.openByUrl(url);
      const licSheet = licSS.getSheetByName(LICENSE_SHEET_NAME);
      if (licSheet) {
        const data = licSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const email  = (data[i][0] || '').toString().toLowerCase().trim();
          const status = (data[i][1] || '').toString().toUpperCase().trim();
          let   expiry = data[i][3];
          if (!email) continue;
          if (status !== 'ACTIVE') continue;
          if (expiry) {
            if (!(expiry instanceof Date)) expiry = new Date(expiry);
            const expiryEnd = new Date(expiry);
            expiryEnd.setDate(expiryEnd.getDate() + 1);
            if (today > expiryEnd) continue; // expired
          }
          validEmails.add(email);
        }
      }
    } catch (licErr) {
      Logger.log('refreshLicenseStatus_: license sheet error: ' + licErr.message);
    }

    // Tier 2: add active Team Members from Hub Members sheet
    // BUT ONLY if the Hub owner's license is ACTIVE in the License sheet.
    // If the Hub license is expired, ALL members are blocked too.
    // Hub owner = the email that has a row in the License sheet (not just Members sheet).
    const hubLicenseActive = validEmails.size > 0; // at least one valid email from License sheet
    if (hubLicenseActive) {
      try {
        const hubSS   = SpreadsheetApp.openById(hubId);
        const membSh  = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
        if (membSh) {
          const mData = membSh.getDataRange().getValues();
          for (let m = 1; m < mData.length; m++) {
            const mEmail  = (mData[m][MC_COL_GMAIL  - 1] || '').toString().toLowerCase().trim();
            const mActive =  mData[m][MC_COL_ACTIVE - 1];
            const isActive = (mActive === true || mActive === 'TRUE' || mActive === 'true');
            if (mEmail && isActive) validEmails.add(mEmail);
          }
        }
      } catch (membErr) {
        Logger.log('refreshLicenseStatus_: members sheet error: ' + membErr.message);
      }
    } else {
      Logger.log('refreshLicenseStatus_: Hub license expired — Members fallback skipped, all users blocked.');
    }

    // Write results to DocumentProperties
    // Only update emails we know about — don't overwrite unknown emails
    validEmails.forEach(email => props.setProperty('lic_' + email, '1'));

    // Mark ALL known emails as "0" first, then overwrite valid ones with "1"
    // This ensures Members-only users get blocked when Hub license expires.
    // Step 1: collect all known emails (License sheet + Members sheet)
    const allKnownEmails = new Set();
    try {
      const licSS2    = SpreadsheetApp.openByUrl(url);
      const licSheet2 = licSS2.getSheetByName(LICENSE_SHEET_NAME);
      if (licSheet2) {
        const data2 = licSheet2.getDataRange().getValues();
        for (let i = 1; i < data2.length; i++) {
          const email = (data2[i][0] || '').toString().toLowerCase().trim();
          if (email) allKnownEmails.add(email);
        }
      }
    } catch (_) {}
    try {
      const hubSS4  = SpreadsheetApp.openById(hubId);
      const membSh4 = hubSS4.getSheetByName(MEMBERS_SHEET_NAME);
      if (membSh4) {
        const mData4 = membSh4.getDataRange().getValues();
        for (let m = 1; m < mData4.length; m++) {
          const mEmail = (mData4[m][MC_COL_GMAIL - 1] || '').toString().toLowerCase().trim();
          if (mEmail) allKnownEmails.add(mEmail);
        }
      }
    } catch (_) {}

    // Step 2: mark all known emails as blocked by default
    allKnownEmails.forEach(email => props.setProperty('lic_' + email, '0'));

    // Step 3: overwrite valid emails with "1"
    validEmails.forEach(email => props.setProperty('lic_' + email, '1'));

    Logger.log('refreshLicenseStatus_: synced. Valid: ' + validEmails.size + ' | Blocked: ' + (allKnownEmails.size - validEmails.size));
  } catch (e) {
    Logger.log('refreshLicenseStatus_ error: ' + e.message);
  }
}

// ── Manual sync — run this to apply license changes immediately ──
function syncLicenseNow() {
  refreshLicenseStatus_();
  try {
    SpreadsheetApp.getUi().alert(
      '✅ License status synced!\n\n' +
      'Changes from your License sheet are now active immediately.'
    );
  } catch (_) {
    // Running from editor or trigger — no UI context, that's fine
    Logger.log('syncLicenseNow: completed successfully (no UI context)');
  }
}

// ── Setup time-based trigger for license refresh (run once) ──
function setupLicenseRefreshTrigger_() {
  // Remove existing refresh triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshLicenseStatus_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('refreshLicenseStatus_')
    .timeBased().everyMinutes(30).create(); // FIX v14.3: 30 min saves trigger quota (was 5 min)
  Logger.log('License refresh trigger created — runs every 30 minutes.');
}

// ── ONE-TIME CACHE CLEAR (run once from Apps Script editor after update) ──
// Run this manually from: Apps Script editor → Select function → Run
// Clears old stale license cache entries so the fresh check runs immediately.
function clearLicenseCache() {
  try {
    CacheService.getScriptCache().remove('license_ok');   // old shared key
    CacheService.getScriptCache().remove('license_ok_v2');
    CacheService.getUserCache().remove('license_ok');
    CacheService.getUserCache().remove('license_ok_v2');
    Logger.log('✅ License cache cleared. Next edit will do a fresh license check.');
    SpreadsheetApp.getUi().alert('✅ License cache cleared!\n\nThe next edit will perform a fresh license check.');
  } catch (e) {
    Logger.log('clearLicenseCache error: ' + e.message);
  }
}

function checkLicense(knownEmail) {
  try {
    // ── Cache: only check once per 6 hours per execution ────
    // Prevents double-popup when both onCRMEdit + onCRMSubTabEdit
    // fire on the same cell edit.
    // FIX v14.3: per-user cache (getUserCache) so one user's failure
    // doesn't block others. Failures never cached — corrected license
    // takes effect on next edit without waiting for cache to expire.
    // FIX v14.3: Cache fully removed — every edit does a live license check
    // so EXPIRED/ACTIVE changes take effect instantly with no wait time.
    // (Cache was causing up to 6hr delay before blocking took effect)

    const url = PropertiesService.getScriptProperties()
                  .getProperty("LICENSE_URL");

    if (!url) {
      console.log("LICENSE_URL not set in Script Properties");
      showLicenseAlert("Cannot reach the license server. Contact: work.mannnish@gmail.com");
      // (failure not cached — corrected license takes effect immediately)
      return false;
    }

    // FIX v14.3: accept email from trigger event (Session returns empty in installable triggers)
    let userEmail = knownEmail || '';
    if (!userEmail) userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) userEmail = Session.getEffectiveUser().getEmail();

    console.log("License check for: " + userEmail);

    if (!userEmail || userEmail === "") {
      // Cannot determine user — BLOCK by default (fail closed, not open)
      console.log("WARNING: Could not determine user email — BLOCKING access");
      showLicenseAlert("Unable to verify your license. Please contact: work.mannnish@gmail.com");
      return false;
    }

    const ss    = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName(LICENSE_SHEET_NAME);

    if (!sheet) {
      console.log("Licenses sheet not found");
      showLicenseAlert("License configuration error. Contact: work.mannnish@gmail.com");
      // (failure not cached — corrected license takes effect immediately)
      return false;
    }

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const licEmail  = data[i][0].toString().toLowerCase().trim();
      const licStatus = data[i][1].toString().toUpperCase().trim();
      // FIX v14.3: robust expiry parsing — handles Date objects, strings, and serial numbers
      let licExpiry = null;
      const rawExpiry = data[i][3];
      if (rawExpiry) {
        if (rawExpiry instanceof Date) {
          licExpiry = rawExpiry;
        } else {
          const parsed = new Date(rawExpiry);
          if (!isNaN(parsed.getTime())) licExpiry = parsed;
        }
      }

      if (licEmail === userEmail.toLowerCase().trim()) {
        if (licStatus !== "ACTIVE") {
          showLicenseAlert("Your license is not active. Contact: work.mannnish@gmail.com | WhatsApp: +91-9511565884");
          // (failure not cached — corrected license takes effect immediately)
          return false;
        }
        // License expiring on today's date is valid all day.
        // Compare using end-of-day in IST (UTC+5:30) so users in India
        // don't get falsely expired during the day.
        if (licExpiry) {
          const expiryEndOfDay = new Date(licExpiry);
          expiryEndOfDay.setDate(expiryEndOfDay.getDate() + 1); // start of NEXT day = end of expiry day
          expiryEndOfDay.setHours(0, 0, 0, 0);
          // Add IST offset (5h 30m = 19800s) so expiry is truly end of day in India
          const expiryWithIST = new Date(expiryEndOfDay.getTime() - (5.5 * 60 * 60 * 1000));
          if (new Date() > expiryWithIST) {
            showLicenseAlert("Your license has expired. Contact: work.mannnish@gmail.com | WhatsApp: +91-9511565884");
            // (failure not cached — corrected license takes effect immediately)
            return false;
          }
        }
        console.log("✅ License valid for: " + userEmail);
        // cache removed — live check every edit for instant license control
        return true;
      }
    }

    // Fallback: check if user is an active Team Member in this CRM
    try {
      const activeSS = SpreadsheetApp.openById(MASTER_HUB_ID);
      const membersSheet = activeSS.getSheetByName(MEMBERS_SHEET_NAME);
      if (membersSheet) {
        const mData = membersSheet.getDataRange().getValues();
        for (let m = 1; m < mData.length; m++) {
          // FIX v14.3: Members sheet cols are A=Gmail, B=FileID, C=Role, D=Name, E=Active
          // Previous code read col B (index 1) as email and col C (index 2) as Active — both wrong.
          const mEmail  = (mData[m][MC_COL_GMAIL  - 1] || '').toString().toLowerCase().trim(); // col A
          const mActive =  mData[m][MC_COL_ACTIVE - 1];  // col E — checkbox = true/false boolean
          const isActive = (mActive === true || mActive === 'TRUE' || mActive === 'true');
          if (mEmail === userEmail.toLowerCase() && isActive) {
            // cache removed — live check every edit for instant license control
            return true;
          }
        }
      }
    } catch (_) {}
    showLicenseAlert("License not found for your account. Contact: work.mannnish@gmail.com");
    // (failure not cached — corrected license takes effect immediately)
    return false;

  } catch (e) {
    console.log("License check error: " + e.message);
    // On error — allow access rather than blocking (fail open)
    // to prevent legitimate users being locked out by network issues
    return true;
  }
}

function showLicenseAlert(msg) {
  try {
    SpreadsheetApp.getUi().alert(
      "🔒 License Issue",
      msg,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    console.log("License alert (no UI): " + msg);
  }
}

function fixLicenseTest() {
  const url = PropertiesService.getScriptProperties()
                .getProperty("LICENSE_URL");
  Logger.log("URL = " + url);
  
  try {
    const ss = SpreadsheetApp.openByUrl(url);
    Logger.log("SUCCESS: " + ss.getName());
  } catch(e) {
    Logger.log("FAILED: " + e.message);
  }
}


// ============================================================
// SECTION 3 — TRIGGER & SYNC FUNCTIONS (v14.0)
// ============================================================

// ============================================================
// === SESSION 2 OUTPUT ===
// CRM MULTI-FILE SYSTEM — Trigger & Sync Functions
// VERSION 14.0 — Order Ledger + Dues + Dashboard
// ============================================================
//
// PREREQUISITES (assumed set by Session 1):
//   • ORDER_LEDGER_SHEET_NAME   = 'Order Ledger'    (Hub — replaces 'Master Hub')
//   • ORDER_LEDGER_CRM_NAME     = 'Order Ledger CRM' (CRM spoke — replaces 'Main')
//   • CLIENT_DUES_SHEET_NAME    = 'Client Dues'
//   • CLIENT_SUMMARY_SHEET_NAME = 'Client Summary'
//   • DASHBOARD_SHEET_NAME      = 'Dashboard'
//   • SUB_COL_DUE_DATE          = 16   (new col P on sub-tab)
//   • HUB_COL_DUE_DATE          = 9    (unchanged — col I on Order Ledger)
//   • CRM_COL_DUE_DATE          = 9    (unchanged — col I on Order Ledger CRM)
//
// NOTE: Every reference to the old 'Master Hub' sheet name is now
//   ORDER_LEDGER_SHEET_NAME, and every reference to the old 'Main'
//   sheet is now ORDER_LEDGER_CRM_NAME. The old constants
//   HUB_MAIN_SHEET_NAME and MAIN_SHEET_NAME should be aliased or
//   replaced in Session 1 constants block. This file uses the NEW
//   names throughout. If Session 1 kept the old const names as
//   aliases (e.g. const HUB_MAIN_SHEET_NAME = ORDER_LEDGER_SHEET_NAME)
//   this code will still compile correctly.
//
// NEW SHEET NAME CONSTANTS (paste into constants block if not already):
//   const ORDER_LEDGER_SHEET_NAME   = 'Order Ledger';
//   const ORDER_LEDGER_CRM_NAME     = 'Order Ledger CRM';
//   const CLIENT_DUES_SHEET_NAME    = 'Client Dues';
//   const CLIENT_SUMMARY_SHEET_NAME = 'Client Summary';
//   const DASHBOARD_SHEET_NAME      = 'Dashboard';
//   const SUB_COL_DUE_DATE          = 16;   // col P on client sub-tab
//
// NEW SYSTEM SHEET GUARD (add to SYSTEM_SHEETS Sets in onCRMSubTabEdit etc.):
//   ORDER_LEDGER_CRM_NAME, CLIENT_DUES_SHEET_NAME,
//   CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME
//
// UPDATED HUB_ALLOWED_PUSH_COLS — now also includes:
//   HUB_COL_INV_NO, HUB_COL_DATE, HUB_COL_DUE_DATE
//
// UPDATED HUB_TO_CRM_MAIN — now also maps:
//   HUB_COL_INV_NO  -> CRM_COL_INV_NO
//   HUB_COL_DATE    -> CRM_COL_DATE
//   HUB_COL_DUE_DATE -> CRM_COL_DUE_DATE
// ============================================================


// ============================================================
// STATUS COLOUR HELPER
// ─────────────────────────────────────────────────────────────
// Returns a background hex colour for a given order status
// string.  Used by onCRMSubTabEdit to paint sub-tab rows and
// by refreshClientDues_ / refreshDashboard_ for dues/dashboard.
// ============================================================


// ============================================================
// deriveOrderStatus_()
// ─────────────────────────────────────────────────────────────
// Derives a human-readable order status from Hub/CRM row data.
// Called by row-colouring code in onCRMSubTabEdit and by the
// dues / dashboard refresh functions.
//
// Parameters (all from the row values array, 0-indexed):
//   payStatus   — col H display value  ('Pending'|'Partial'|'Received')
//   dueDate     — col I value          (Date or '')
//   invValue    — col G value          (number or '')
//   awbNo       — col Q value          (string)  [Hub only; pass '' for CRM]
//   invNo       — col E value          (string)
// ============================================================


// ============================================================
// 1. onCRMSubTabEdit(e) — v14.0
// ─────────────────────────────────────────────────────────────
// Changes from v11.0:
//   • Syncs Inv No + Inv Date to Hub Order Ledger (two-way)
//   • Syncs Due Date (col P / SUB_COL_DUE_DATE) to:
//       – CRM Order Ledger CRM col I
//       – Hub Order Ledger col I
//   • Pushes all written data to Order Ledger CRM (renamed Main)
//   • Duplicate PI No check on entry — alerts if PI already exists
//     in another row of this client sub-tab
//   • Duplicate Inv No check on entry — alerts if Inv No already
//     exists anywhere in this CRM spoke's sub-tabs
//   • Due Date lock: if Due Date was set < 7 days ago, blocks
//     changes with a warning and reverts the cell
//   • Applies row background colour based on derived order status
//     (Created/Invoiced/Dispatched/Payment Pending/Overdue/
//      Partial/Paid) after every sync
// ============================================================


// ============================================================
// applySubTabRowColour_(sheet, row, uuid, mainRowData)
// ─────────────────────────────────────────────────────────────
// Derives the current order status for the sub-tab row and
// paints the row background accordingly.
// mainRowData is the 0-indexed values array for the matching
// Order Ledger CRM row (used to get dueDate, invValue etc.).
// ============================================================


// ============================================================
// 2. onMasterHubEdit(e) — v14.0
// ─────────────────────────────────────────────────────────────
// Changes from v13:
//   • Now listens to 'Order Ledger' sheet (not 'Master Hub')
//   • HUB_ALLOWED_PUSH_COLS expanded to include:
//       HUB_COL_INV_NO, HUB_COL_DATE, HUB_COL_DUE_DATE
//   • Logistics enters Inv No / Inv Date → push to CRM spoke
//       sub-tab (Inv No col B, Inv Date col C) AND Order Ledger CRM
//   • Sales enters Due Date (col I) → push to:
//       CRM sub-tab col P, Order Ledger CRM col I,
//       Hub Client Dues, CRM Client Dues
//   • Admin/Sales edits any CRM-mapped field → push to CRM sub-tab
//   • After any payment edit → refreshClientDues_ + refreshDashboard_
//   • After any edit → refreshDashboard_
// ============================================================


// ============================================================
// 3. refreshCreditTotals_(userEmail) — v14.0
// ─────────────────────────────────────────────────────────────
// Changes from v8:
//   • Reads from 'Order Ledger' (not old 'Master Hub' sheet)
//   • Updates Client Summary sheet totals (new sheet)
//   • Otherwise identical calculation: UsedCredit = InvValue − PayReceived
// ============================================================


// ============================================================
// updateClientSummary_(hubSS, invTotals, payTotals, userEmail)
// ─────────────────────────────────────────────────────────────
// Writes aggregated per-client totals to the Client Summary
// sheet (columns: Client Name | CRM Owner | Total Invoiced |
// Total Paid | Total Due | Credit Limit | Available Credit).
// Called by refreshCreditTotals_ after computing totals.
// ============================================================


// ============================================================
// 4. refreshClientDues_(userEmail) — NEW in v14.0
// ─────────────────────────────────────────────────────────────
// Reads all Order Ledger rows.
// Filters where Due Balance > 0 (outstanding debt).
// Derives order status per row.
// Sorts: Overdue (red) rows first, Partial (yellow) below,
//        others after.
// Writes result to Hub 'Client Dues' sheet.
// Then iterates every active CRM spoke and writes a filtered
// version (only that spoke's clients) to the CRM's own
// 'Client Dues' sheet.
//
// Column layout written:
//   A: Customer | B: CRM Owner | C: PI No | D: Inv No |
//   E: Inv Date | F: Inv Value | G: Due Date | H: Pay Status |
//   I: Due Balance | J: Status | K: UUID
// ============================================================


// ── writeDuesSheet_ helper ─────────────────────────────────────


// ============================================================
// 5. refreshDashboard_(userEmail) — NEW in v14.0
// ─────────────────────────────────────────────────────────────
// Reads Client Summary sheet totals.
// Writes all clients sorted by Total Due (high → low) to the
// 'Dashboard' sheet in Master Hub.
// Applies row colours:
//   Red    = Total Due > 0 AND any order is Overdue
//   Yellow = Total Due > 0, no Overdue, some Partial
//   Green  = Total Due = 0 (fully paid / no outstanding)
//
// Dashboard columns:
//   A: Rank | B: Client Name | C: CRM Owner |
//   D: Total Invoiced | E: Total Paid | F: Total Due |
//   G: Credit Limit | H: Available Credit | I: Status
// ============================================================


// ============================================================
// Menu wrappers — expose new refresh functions to Admin menu
// Add these items to the CRM Tools menu in onOpen():
//   .addItem('Refresh Client Dues',  'menuRefreshClientDues_')
//   .addItem('Refresh Dashboard',    'menuRefreshDashboard_')
// ============================================================


function statusColour_(status) {
  const s = (status || '').toString().trim().toLowerCase();
  // Priority map — most actionable conditions first
  if (s === 'overdue')          return '#FFCDD2'; // red-100
  if (s === 'payment pending')  return '#FFE0B2'; // orange-100
  if (s === 'partial')          return '#FFF9C4'; // yellow-100
  if (s === 'invoiced')         return '#E3F2FD'; // blue-100
  if (s === 'dispatched')       return '#E8F5E9'; // green-100
  if (s === 'paid')             return '#C8E6C9'; // green-200
  if (s === 'created')          return '#F3E5F5'; // purple-100
  return '#FFFFFF';                               // default
}

function deriveOrderStatus_(payStatus, dueDate, invValue, awbNo, invNo) {
  const now = new Date();
  const pay = (payStatus || '').toString().trim().toLowerCase();

  if (pay === 'received') return 'Paid';

  // Overdue: due date is in the past and not fully paid
  if (dueDate && dueDate instanceof Date && dueDate < now && pay !== 'received') {
    return 'Overdue';
  }

  if (pay === 'partial') return 'Partial';

  // Dispatched: AWB No is filled
  if (awbNo && awbNo.toString().trim()) return 'Dispatched';

  // Invoiced: Invoice No is filled
  if (invNo && invNo.toString().trim()) return 'Invoiced';

  // Invoice Value entered but not yet invoiced
  if (invValue && Number(invValue) > 0) return 'Payment Pending';

  return 'Created';
}

function onCRMSubTabEdit(e) {
  if (!e || !e.range) {
    Logger.log('onCRMSubTabEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  // FIX 5: invalidate sheet cache at start of each execution
  invalidateAllSheetCaches_();
  try {
    // ── LICENSE CHECK — FIX v14.3: was missing from onCRMSubTabEdit ──
    // Clear the edited cell and return if license is invalid,
    // so the edit is visually reverted (same pattern as duplicate check).
    const _licEmail = (e && e.user) ? e.user.email : '';
    if (!checkLicense(_licEmail)) {
      try { e.range.clearContent(); } catch (_) {}
      return;
    }

    const sheet     = e.range.getSheet();
    const sheetName = sheet.getName();
    const row       = e.range.getRow();
    const col       = e.range.getColumn();

    if (row <= 1) return; // header row

    // ── Guard: skip system sheets ────────────────────────────
    const SYSTEM_SHEETS = new Set([
      // Legacy names (in case old constants are still present)
      'Main', 'Master Hub',
      // New names
      ORDER_LEDGER_CRM_NAME,
      ORDER_LEDGER_SHEET_NAME,
      // Unchanged system sheets
      LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME,
      CLIENT_DUES_SHEET_NAME,
      CLIENT_SUMMARY_SHEET_NAME,
      DASHBOARD_SHEET_NAME,
    ]);
    if (SYSTEM_SHEETS.has(sheetName)) return;

    // ── v15.1: Lockout guard — Logistics/Accounts-exclusive columns ──
    // Inv No, Inv Date, AWB No, Dispatched Date, Payment Received, Payment Date
    // are entered ONLY on the Hub (by Logistics/Accounts) and flow DOWN into
    // this sub-tab via onMasterHubEdit. CRM can no longer enter them here.
    // v15.3: Due Date removed from this set — CRM can edit it again here, just
    // rate-limited to once every 7 days (see PATH 2 below), not blocked
    // outright. NOTE: sheet protection alone cannot enforce the remaining
    // locks, because each CRM exec is the Drive OWNER of their own spoke file,
    // and Sheets protection can never block a file's owner. This guard is the
    // real enforcement: any human edit to one of these cells is immediately
    // reverted. Script-driven writes from onMasterHubEdit do NOT fire onEdit
    // at all (Apps Script never fires onEdit for programmatic SpreadsheetApp
    // writes), so legitimate Hub→Spoke pushes are never touched by this guard.
    const LOCKED_FOR_CRM_SUB_COLS = new Set([
      SUB_COL_INV_NO, SUB_COL_INV_DATE, SUB_COL_AWB_NO, SUB_COL_DISP_DATE,
      SUB_COL_PAY_RECEIVED, SUB_COL_PAY_DATE,
    ]);
    if (LOCKED_FOR_CRM_SUB_COLS.has(col)) {
      const guardEmail = (e.user && e.user.email) || Session.getActiveUser().getEmail() || 'unknown';
      try { e.range.clearContent(); SpreadsheetApp.flush(); } catch (_) {}
      try {
        logActivity_(SpreadsheetApp.getActiveSpreadsheet(), 'CRM EDIT BLOCKED',
          'sheet:' + sheetName + ' row:' + row + ' col:' + col +
          ' — managed by Logistics/Accounts via Hub', guardEmail);
      } catch (_) {}
      alertIfUi_(
        'This field is entered by Logistics/Accounts on the Hub and updates here automatically.\n' +
        'Please contact Logistics or Accounts directly to change it.'
      );
      return;
    }

    // ── Guard: must have at least 2 columns ──────────────────
    if (sheet.getLastColumn() < 2) return;

    // ── Editable sub-tab columns ──────────────────────────────
    // v15.1: Inv No, Inv Date, AWB No, Dispatched Date, Payment Received,
    // Payment Date removed — now Logistics/Accounts-exclusive via the Hub
    // (see LOCKED_FOR_CRM_SUB_COLS guard above).
    // v15.3: Due Date is back — CRM can edit it, rate-limited (PATH 2 below).
    const EDITABLE_SUB_COLS = [
      SUB_COL_PI_NO,           // 1   → Order Ledger CRM col D + Hub col D
      SUB_COL_ITEM_VALUE_USD,  // 7   → triggers Final Inv push
      SUB_COL_DEV_CHARGES_USD, // 8   → triggers Final Inv push
      SUB_COL_FOB_VALUE,       // 9   → triggers Final Inv push
      SUB_COL_FREIGHT,         // 10  → triggers Final Inv push
      SUB_COL_CIF_VALUE,       // 11  → triggers Final Inv push
      SUB_COL_DISCOUNT,        // 12  → triggers Final Inv push
      SUB_COL_FINAL_INV,       // 13  → Order Ledger CRM col G + Hub col G
      SUB_COL_DUE_DATE,        // 16  → v15.3: CRM can set, 7-day rate limit
    ];
    if (!EDITABLE_SUB_COLS.includes(col)) {
      Logger.log('onCRMSubTabEdit: col ' + col + ' not in EDITABLE_SUB_COLS — skipping');
      return;
    }

    Logger.log('onCRMSubTabEdit: FIRING | sheet="' + sheetName + '" row=' + row + ' col=' + col);

    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName(ORDER_LEDGER_CRM_NAME);
    const userEmail = Session.getActiveUser().getEmail() || 'trigger-auto';

    if (!mainSheet) {
      Logger.log('onCRMSubTabEdit: Order Ledger CRM sheet not found in ' + ss.getName());
      return;
    }

    const lock = LockService.getScriptLock();
    // FIX v14.3: Increased lock timeouts to handle more concurrent users
    // and larger datasets. Previously 8s+10s (18s max), now 15s+20s (35s max).
    const gotLock = lock.tryLock(15000);
    if (!gotLock) {
      // Another trigger is syncing this row right now.
      // BUG FIX v14.1: Do NOT return — the other trigger only syncs the
      // column IT was fired for. If user edited G, H, J, L fast, we have
      // 4 simultaneous triggers but only 1 gets the lock. The rest must
      // wait and then do a FULL-ROW sync so no data is lost.
      Logger.log('onCRMSubTabEdit: lock busy — waiting 15s then doing full-row sync | col=' + col);
      Utilities.sleep(15000);  // wait for the first trigger to finish
      // Try once more after waiting
      const gotLock2 = lock.tryLock(20000);
      if (!gotLock2) {
        Logger.log('onCRMSubTabEdit: lock still busy after 30s — giving up | col=' + col);
        return;
      }
      // We now have the lock — fall through and do a full-row sync below.
      // The col variable below will be ignored; we always read the full row.
    }

    try {
      // ── Find matching Order Ledger CRM row by client name + PI No ───
      // Order Ledger CRM has ONE ROW PER ORDER (not per client).
      // We must match on both customer name AND PI No so that each
      // sub-tab row maps to its own Order Ledger CRM row.
      const mainData    = mainSheet.getDataRange().getValues();
      const mainDisplay = mainSheet.getDataRange().getDisplayValues();

      // Read the PI No from the sub-tab row being edited
      const subRowForPI = sheet.getRange(row, 1, 1,
        Math.max(sheet.getLastColumn(), SUB_COL_PI_NO)).getValues()[0];
      const editedPiNo  = (subRowForPI[SUB_COL_PI_NO - 1] || '').toString().trim();

      let mainRow = -1;

      // Priority 1: match client name + PI No (exact per-order row)
      if (editedPiNo) {
        for (let i = 1; i < mainDisplay.length; i++) {
          const rowClient = (mainDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
          const rowPiNo   = (mainDisplay[i][CRM_COL_PI_NO   - 1] || '').toString().trim();
          if (rowClient === sheetName.toLowerCase() && rowPiNo === editedPiNo) {
            mainRow = i + 1;
            break;
          }
        }
      }

      // ── NEW ORDER ROW: PI No entered but no matching row found ────
      if (mainRow === -1 && col === SUB_COL_PI_NO && editedPiNo) {

        // ── FIX v14.3: DUPLICATE PI No CHECK — must run BEFORE writing any row ──
        // PATH 1 previously returned early (before the duplicate check block below),
        // so a brand-new PI No was never validated. This caused duplicates across clients.
        {
          const piStrNew = editedPiNo;
          let dupPiClientNew = '';

          // 1. Check all client sub-tabs in this CRM spoke
          const sysSetNew = new Set([
            ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
            CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
            CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME,
          ]);
          ss.getSheets().forEach(s => {
            if (dupPiClientNew) return;
            if (sysSetNew.has(s.getName())) return;
            const sData = s.getDataRange().getValues();
            for (let i = 1; i < sData.length; i++) {
              if (s.getName() === sheetName && i + 1 === row) continue; // skip self
              if ((sData[i][SUB_COL_PI_NO - 1] || '').toString().trim() === piStrNew) {
                dupPiClientNew = s.getName();
                break;
              }
            }
          });

          // 2. Check CRM Order Ledger main sheet
          if (!dupPiClientNew) {
            for (let i = 1; i < mainData.length; i++) {
              if ((mainData[i][CRM_COL_PI_NO - 1] || '').toString().trim() === piStrNew) {
                dupPiClientNew = (mainData[i][CRM_COL_CUSTOMER - 1] || 'another client').toString().trim();
                break;
              }
            }
          }

          // 3. Check Hub Order Ledger — catches ALL other CRM spokes
          if (!dupPiClientNew) {
            try {
              const hubSSNew = SpreadsheetApp.openById(MASTER_HUB_ID);
              const hubOLNew = hubSSNew.getSheetByName(ORDER_LEDGER_SHEET_NAME);
              if (hubOLNew) {
                const hdNew = hubOLNew.getDataRange().getValues();
                for (let i = 1; i < hdNew.length; i++) {
                  const hPi   = (hdNew[i][OL_COL_PI_NO       - 1] || '').toString().trim();
                  const hName = (hdNew[i][OL_COL_CLIENT_NAME  - 1] || '').toString().trim();
                  if (hPi === piStrNew) {
                    dupPiClientNew = hName || 'another client (Hub)';
                    break;
                  }
                }
              }
            } catch (hubErrNew) {
              Logger.log('DUP PI NO (PATH1) Hub check error: ' + hubErrNew.message);
            }
          }

          if (dupPiClientNew) {
            e.range.clearContent();
            SpreadsheetApp.flush();
            alertIfUi_(
              'DUPLICATE PI No DETECTED\n\n' +
              'PI No "' + piStrNew + '" is already assigned to client:\n  ' + dupPiClientNew + '\n\n' +
              'PI Numbers must be unique across ALL clients and CRM files.\n' +
              'The entry has been cleared.'
            );
            logActivity_(ss, 'DUP PI NO BLOCKED (PATH1)',
              'Client:' + sheetName + ' | PI:' + piStrNew + ' | conflict with:' + dupPiClientNew + ' | row:' + row,
              userEmail);
            return;
          }
        }
        // ── END DUPLICATE CHECK ───────────────────────────────────────────────────

        mainRow = mainSheet.getLastRow() + 1;

        // Inherit UUID and CRM Owner from existing row for this client
        let inheritUuid     = '';
        let inheritCrmOwner = '';

        // Priority 1: find from Order Ledger CRM (same file)
        for (let i = 1; i < mainData.length; i++) {
          const rowClient = (mainDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
          if (rowClient === sheetName.toLowerCase()) {
            inheritUuid     = (mainData[i][CRM_COL_UUID      - 1] || '').toString().trim();
            inheritCrmOwner = (mainData[i][CRM_COL_CRM_OWNER - 1] || '').toString().trim();
            if (inheritUuid) break; // stop once we find a non-empty UUID
          }
        }

        // Priority 2: if UUID still empty, look in Hub Order Ledger by client name
        if (!inheritUuid) {
          try {
            const hubSS2 = SpreadsheetApp.openById(MASTER_HUB_ID);
            const olSheet2 = hubSS2.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (olSheet2) {
              const olData2 = olSheet2.getDataRange().getValues();
              for (let i = 1; i < olData2.length; i++) {
                const hName = (olData2[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
                const hUUID = (olData2[i][OL_COL_UUID        - 1] || '').toString().trim();
                // v15.3 FIX: Column C can now contain "Name (email)" display
                // format — extractEmail_ safely returns a bare email either way.
                const hOwner= extractEmail_(olData2[i][OL_COL_CRM_OWNER - 1]);
                if (hName === sheetName.toLowerCase() && hUUID) {
                  inheritUuid     = hUUID;
                  inheritCrmOwner = inheritCrmOwner || hOwner;
                  break;
                }
              }
            }
          } catch (_) {}
        }

        Logger.log('onCRMSubTabEdit: NEW ORDER ROW | client=' + sheetName +
          ' PI=' + editedPiNo + ' inheritUuid=' + inheritUuid);

        // Write the new row skeleton: UUID + Customer Name + CRM Owner + PI No
        const sheetId = ss.getSheetByName(sheetName) ? ss.getSheetByName(sheetName).getSheetId() : null;
        const fileId  = ss.getId();
        const tabUrl  = sheetId
          ? `https://docs.google.com/spreadsheets/d/${fileId}/edit#gid=${sheetId}`
          : null;
        const nameFormula = tabUrl
          ? `=HYPERLINK("${tabUrl}","${sheetName}")`
          : sheetName;

        PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
        if (inheritUuid)     mainSheet.getRange(mainRow, CRM_COL_UUID).setValue(inheritUuid);
        if (inheritCrmOwner) mainSheet.getRange(mainRow, CRM_COL_CRM_OWNER).setValue(inheritCrmOwner);
        mainSheet.getRange(mainRow, CRM_COL_CUSTOMER).setFormula(nameFormula);
        mainSheet.getRange(mainRow, CRM_COL_PI_NO).setValue(editedPiNo);
        setCRMFormulas_(mainSheet, mainRow);
        SpreadsheetApp.flush();

        // ── FIX v14.2: Delete the blank placeholder row (no PI No) for this client ──
        // onCreditRequestEdit creates a placeholder row when a client is first assigned.
        // Now that a real order row (with PI No) exists, remove the placeholder to prevent
        // blank rows appearing in Order Ledger CRM and Hub Order Ledger.
        try {
          // 1. Delete blank placeholder in Order Ledger CRM (CRM Spoke — same file)
          const freshMainData    = mainSheet.getDataRange().getValues();
          const freshMainDisplay = mainSheet.getDataRange().getDisplayValues();
          for (let i = freshMainData.length - 1; i >= 1; i--) { // bottom-up to keep row indices valid
            const rClient = (freshMainDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
            const rPiNo   = (freshMainData[i][CRM_COL_PI_NO - 1] || '').toString().trim();
            const rUuid   = (freshMainData[i][CRM_COL_UUID  - 1] || '').toString().trim();
            // Match: same client name, no PI No, has a UUID (= placeholder), not the row we just wrote
            if (rClient === sheetName.toLowerCase() && !rPiNo && rUuid && (i + 1) !== mainRow) {
              mainSheet.deleteRow(i + 1);
              Logger.log('onCRMSubTabEdit FIX: deleted blank placeholder row ' + (i + 1) +
                ' from Order Ledger CRM for client: ' + sheetName);
            }
          }
          SpreadsheetApp.flush();

          // 2. Delete blank placeholder in Hub Order Ledger
          const hubSS3   = SpreadsheetApp.openById(MASTER_HUB_ID);
          const olSheet3 = hubSS3.getSheetByName(ORDER_LEDGER_SHEET_NAME);
          if (olSheet3) {
            const olData3 = olSheet3.getDataRange().getValues();
            for (let i = olData3.length - 1; i >= 1; i--) {
              const hClient = (olData3[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
              const hPiNo   = (olData3[i][OL_COL_PI_NO       - 1] || '').toString().trim();
              const hUuid   = (olData3[i][OL_COL_UUID         - 1] || '').toString().trim();
              if (hClient === sheetName.toLowerCase() && !hPiNo && hUuid) {
                olSheet3.deleteRow(i + 1);
                Logger.log('onCRMSubTabEdit FIX: deleted blank placeholder row ' + (i + 1) +
                  ' from Hub Order Ledger for client: ' + sheetName);
              }
            }
            SpreadsheetApp.flush();
          }
        } catch (cleanupErr) {
          Logger.log('onCRMSubTabEdit: placeholder cleanup error: ' + cleanupErr.message);
        }

        // BUG FIX v15: Sort Order Ledger CRM by Customer Name (col B) then PI No (col D)
        // so rows always appear in a predictable order as orders accumulate.
        try {
          const lastDataRow = mainSheet.getLastRow();
          // FIX v14.3: Only sort when 3+ rows exist (avoid sort on brand-new rows
          // which can cause race condition — next trigger can't find the row).
          // Also flush AFTER sort so changes are committed before lock releases.
          if (lastDataRow > 3) {
            mainSheet.getRange(2, 1, lastDataRow - 1, mainSheet.getLastColumn())
              .sort([
                { column: CRM_COL_CUSTOMER, ascending: true },
                { column: CRM_COL_PI_NO,    ascending: true },
              ]);
            SpreadsheetApp.flush(); // ensure sort is committed before lock releases
          }
        } catch (sortErr) {
          Logger.log('onCRMSubTabEdit: sort error: ' + sortErr.message);
        }

        Logger.log('onCRMSubTabEdit: NEW ORDER ROW appended at Order Ledger CRM row ' + mainRow +
          ' for client="' + sheetName + '" PI="' + editedPiNo + '"');
        logActivity_(ss, 'NEW ORDER ROW IN ORDER LEDGER CRM',
          'Client:' + sheetName + ' | PI:' + editedPiNo + ' | row:' + mainRow, userEmail);

        // PI No already written above — nothing more to do for PATH 1 in this case
        applySubTabRowColour_(sheet, row, inheritUuid, []);
        return;
      }

      // Fallback for non-PI-No column edits: match by client name only
      // (handles Inv No, Inv Date, value columns etc. when PI No already written)
      if (mainRow === -1) {
        // Match client name + same sub-tab row number as a proxy
        // (use PI No from sub-tab to narrow down if possible)
        for (let i = 1; i < mainDisplay.length; i++) {
          const rowClient = (mainDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase();
          const rowPiNo   = (mainDisplay[i][CRM_COL_PI_NO   - 1] || '').toString().trim();
          if (rowClient === sheetName.toLowerCase()) {
            // Prefer a row where PI No is blank (not yet assigned) or matches
            if (!rowPiNo || rowPiNo === editedPiNo) {
              mainRow = i + 1;
              break;
            }
          }
        }
      }

      if (mainRow === -1) {
        Logger.log('onCRMSubTabEdit: No matching Order Ledger CRM row for "' + sheetName + '" piNo="' + editedPiNo + '"');
        return;
      }

      const uuid     = (mainData[mainRow - 1][CRM_COL_UUID      - 1] || '').toString().trim();
      const crmOwner = (mainData[mainRow - 1][CRM_COL_CRM_OWNER - 1] || '').toString().trim();
      const cellVal  = e.range.getValue();

      // ── DUPLICATE PI No CHECK ─────────────────────────────────
      // BUG FIX v14.1: PI No must be globally unique across ALL clients
      // in this CRM spoke AND all other CRM spokes (via Hub Order Ledger).
      // Previous check only looked within the same client sub-tab.
      if (col === SUB_COL_PI_NO && cellVal) {
        const piStr = cellVal.toString().trim();
        let dupPiClient = '';  // name of client that already has this PI No

        // 1. Check all client sub-tabs in this CRM spoke
        const sysSetPi = new Set([
          ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
          CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
          CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME,
        ]);
        ss.getSheets().forEach(s => {
          if (dupPiClient) return;
          if (sysSetPi.has(s.getName())) return;
          const sData = s.getDataRange().getValues();
          for (let i = 1; i < sData.length; i++) {
            if (s.getName() === sheetName && i + 1 === row) continue; // skip self
            if ((sData[i][SUB_COL_PI_NO - 1] || '').toString().trim() === piStr) {
              dupPiClient = s.getName();
              break;
            }
          }
        });

        // 2. Check CRM Order Ledger (catches same CRM Main sheet too)
        if (!dupPiClient) {
          const mainSh = getCRMMainSheet_(ss);
          if (mainSh) {
            const md = mainSh.getDataRange().getValues();
            for (let i = 1; i < md.length; i++) {
              if ((md[i][CRM_COL_PI_NO - 1] || '').toString().trim() === piStr) {
                dupPiClient = (md[i][CRM_COL_CUSTOMER - 1] || 'another client').toString().trim();
                break;
              }
            }
          }
        }

        // 3. Check Hub Order Ledger — catches ALL CRM spokes (including same client different order)
        // FIX v14.3: Removed wrong `hName !== sheetName` exclusion which allowed the same
        // PI No to be reused by the same client. Now skips only the exact same order row
        // (matched by UUID) so editing an existing PI No doesn't self-flag as duplicate.
        if (!dupPiClient) {
          try {
            const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubOL) {
              const hd = hubOL.getDataRange().getValues();
              for (let i = 1; i < hd.length; i++) {
                const hPi   = (hd[i][OL_COL_PI_NO  - 1] || '').toString().trim();
                const hUuid = (hd[i][OL_COL_UUID    - 1] || '').toString().trim();
                const hName = (hd[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim();
                // Skip only the exact same order row (by UUID) to avoid self-flagging on re-edit
                if (hPi === piStr && !(uuid && hUuid && hUuid === uuid)) {
                  dupPiClient = hName || 'another client (Hub)';
                  break;
                }
              }
            }
          } catch (hubErr) {
            Logger.log('DUP PI NO: Hub check error: ' + hubErr.message);
            // Non-fatal — continue without Hub check if Hub unreachable
          }
        }

        if (dupPiClient) {
          e.range.clearContent();
          SpreadsheetApp.flush();
          alertIfUi_(
            'DUPLICATE PI No DETECTED\n\n' +
            'PI No "' + piStr + '" is already assigned to client:\n  ' + dupPiClient + '\n\n' +
            'PI Numbers must be unique across ALL clients and CRM files.\n' +
            'The entry has been cleared.'
          );
          logActivity_(ss, 'DUP PI NO BLOCKED',
            'Client:' + sheetName + ' | PI:' + piStr + ' | conflict with:' + dupPiClient + ' | row:' + row,
            userEmail);
          return;
        }
      }

      // ── DUPLICATE Inv No CHECK ─────────────────────────────────
      // BUG FIX v14.1: Inv No must be globally unique across ALL clients
      // in this CRM spoke AND all other CRM spokes (via Hub Order Ledger).
      // Previous check only looked within this CRM spoke.
      if (col === SUB_COL_INV_NO && cellVal) {
        const invNoStr = cellVal.toString().trim();
        let dupInvClient = '';

        // 1. Check all client sub-tabs in this CRM spoke
        const sysSetInv = new Set([
          ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
          CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
          CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME,
        ]);
        ss.getSheets().forEach(s => {
          if (dupInvClient) return;
          if (sysSetInv.has(s.getName())) return;
          const sData = s.getDataRange().getValues();
          for (let i = 1; i < sData.length; i++) {
            if (s.getName() === sheetName && i + 1 === row) continue; // skip self
            if ((sData[i][SUB_COL_INV_NO - 1] || '').toString().trim() === invNoStr) {
              dupInvClient = s.getName();
              break;
            }
          }
        });

        // 2. Check CRM Order Ledger
        if (!dupInvClient) {
          const mainSh = getCRMMainSheet_(ss);
          if (mainSh) {
            const md = mainSh.getDataRange().getValues();
            for (let i = 1; i < md.length; i++) {
              if ((md[i][CRM_COL_INV_NO - 1] || '').toString().trim() === invNoStr) {
                dupInvClient = (md[i][CRM_COL_CUSTOMER - 1] || 'another client').toString().trim();
                break;
              }
            }
          }
        }

        // 3. Check Hub Order Ledger — catches ALL CRM spokes (including same client different order)
        // FIX v14.3: Removed wrong `hName !== sheetName` exclusion. Skip only the exact same
        // order row by UUID so editing an existing Inv No doesn't self-flag as duplicate.
        if (!dupInvClient) {
          try {
            const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubOL) {
              const hd = hubOL.getDataRange().getValues();
              for (let i = 1; i < hd.length; i++) {
                const hInv  = (hd[i][OL_COL_INV_NO  - 1] || '').toString().trim();
                const hUuid = (hd[i][OL_COL_UUID     - 1] || '').toString().trim();
                const hName = (hd[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim();
                // Skip only the exact same order row (by UUID) to avoid self-flagging on re-edit
                if (hInv === invNoStr && !(uuid && hUuid && hUuid === uuid)) {
                  dupInvClient = hName || 'another client (Hub)';
                  break;
                }
              }
            }
          } catch (hubErr) {
            Logger.log('DUP INV NO: Hub check error: ' + hubErr.message);
          }
        }

        if (dupInvClient) {
          e.range.clearContent();
          SpreadsheetApp.flush();
          alertIfUi_(
            'DUPLICATE Invoice No DETECTED\n\n' +
            'Inv No "' + invNoStr + '" is already assigned to client:\n  ' + dupInvClient + '\n\n' +
            'Invoice numbers must be unique across ALL clients and CRM files.\n' +
            'The entry has been cleared.'
          );
          logActivity_(ss, 'DUP INV NO BLOCKED',
            'Client:' + sheetName + ' | InvNo:' + invNoStr + ' | conflict with:' + dupInvClient + ' | row:' + row,
            userEmail);
          return;
        }
      }

      // ── DUE DATE LOCK: block edits if Due Date set < 7 days ago ──
      // We store the Due Date set-timestamp in script properties
      // keyed by UUID+row so we can detect premature changes.
      if (col === SUB_COL_DUE_DATE && cellVal) {
        const lockKey = 'duedate_set_' + uuid + '_r' + row;
        const props   = PropertiesService.getScriptProperties();
        const setTs   = props.getProperty(lockKey);
        if (setTs) {
          const setDate  = new Date(Number(setTs));
          const daysSince = (Date.now() - setDate.getTime()) / 86400000;
          if (daysSince < 7) {
            // Revert: restore old value from mainData
            const oldDueDate = mainData[mainRow - 1][CRM_COL_DUE_DATE - 1];
            e.range.setValue(oldDueDate || '');
            SpreadsheetApp.flush();
            alertIfUi_(
              '🔒 Due Date Locked!\n\n' +
              'The Due Date for this order was set ' + Math.floor(daysSince) + ' day(s) ago.\n' +
              'Due Dates cannot be changed within 7 days of being set.\n\n' +
              'Contact your Sales Manager or Admin to override.'
            );
            logActivity_(ss, 'DUE DATE LOCK',
              'Client:' + sheetName + ' | UUID:' + uuid + ' | row:' + row +
              ' | DaysSinceSet:' + Math.floor(daysSince), userEmail);
            return;
          }
        }
        // First-time set or past lock period — record the timestamp
        props.setProperty(lockKey, Date.now().toString());
      }

      // ─────────────────────────────────────────────────────────
      // PATH 1: PI No / Inv No / Inv Date
      // Writes to: Order Ledger CRM + Hub Order Ledger + Master Hub
      // v15.1 NOTE: the Inv No / Inv Date branches below are now unreachable —
      // the LOCKED_FOR_CRM_SUB_COLS guard earlier in this function returns
      // before col can ever equal SUB_COL_INV_NO or SUB_COL_INV_DATE here.
      // Left in place rather than deleted to minimize surgery on working code;
      // safe to remove in a future cleanup pass.
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_PI_NO || col === SUB_COL_INV_NO || col === SUB_COL_INV_DATE) {
        const mainColMap = {
          [SUB_COL_PI_NO]:    CRM_COL_PI_NO,
          [SUB_COL_INV_NO]:   CRM_COL_INV_NO,
          [SUB_COL_INV_DATE]: CRM_COL_DATE,
        };
        const olColMap = {
          [SUB_COL_PI_NO]:    OL_COL_PI_NO,
          [SUB_COL_INV_NO]:   OL_COL_INV_NO,
          [SUB_COL_INV_DATE]: OL_COL_INV_DATE,
        };
        const hubColMap = {
          [SUB_COL_PI_NO]:    HUB_COL_PI_NO,
          [SUB_COL_INV_NO]:   HUB_COL_INV_NO,
          [SUB_COL_INV_DATE]: HUB_COL_DATE,
        };
        const mainTargetCol = mainColMap[col];
        const olTargetCol   = olColMap[col];
        const hubTargetCol  = hubColMap[col];
        const isDate = col === SUB_COL_INV_DATE;

        // Write to Order Ledger CRM
        PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
        const mainCell = mainSheet.getRange(mainRow, mainTargetCol);
        mainCell.setValue(cellVal);
        if (isDate) mainCell.setNumberFormat('dd/mm/yyyy');
        SpreadsheetApp.flush();

        if (uuid) {
          try {
            const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);

            // Write to Hub Order Ledger (new primary sheet)
            const olSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (olSheet) {
              const olData = olSheet.getDataRange().getValues();
              let olRow = -1;
              // Match by UUID + PI No for multi-order clients
              const subPiNo = sheet.getRange(row, SUB_COL_PI_NO).getValue().toString().trim();
              for (let i = 1; i < olData.length; i++) {
                const rUUID = (olData[i][OL_COL_UUID  - 1] || '').toString().trim();
                const rPI   = (olData[i][OL_COL_PI_NO - 1] || '').toString().trim();
                if (rUUID === uuid && rPI === subPiNo) { olRow = i + 1; break; }
              }
              // Fallback: UUID-only if PI No is blank (new row not yet assigned PI)
              if (olRow === -1) {
                for (let i = 1; i < olData.length; i++) {
                  if ((olData[i][OL_COL_UUID - 1] || '').toString().trim() === uuid
                      && !(olData[i][OL_COL_PI_NO - 1] || '').toString().trim()) {
                    olRow = i + 1; break;
                  }
                }
              }
              if (olRow !== -1) {
                const olCell = olSheet.getRange(olRow, olTargetCol);
                olCell.setValue(cellVal);
                if (isDate) olCell.setNumberFormat('dd/mm/yyyy');
              }
            }

            // Master Hub removed v15.0

            SpreadsheetApp.flush();
            logActivity_(hubSS, 'SUBTAB PI/INV/DATE->OL+HUB',
              'UUID:' + uuid + ' | col:' + col + '=' + cellVal, userEmail);
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH1 Hub write error: ' + hubErr.message);
          }
        }

        logActivity_(ss, 'SUBTAB PI/INV/DATE->ORDER LEDGER CRM',
          'Client:' + sheetName + ' | col:' + col + '=' + cellVal, userEmail);

        // FIX v14.3: Refresh Dashboard when Inv Date is entered
        // so Last Order Date (col H) updates immediately.
        if (col === SUB_COL_INV_DATE && cellVal) {
          try { refreshDashboard_(userEmail); } catch (e) {
            Logger.log('PATH1 refreshDashboard_ error: ' + e.message);
          }
        }

        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 2: Due Date (col P)
      // Push to: Order Ledger CRM col I, Hub Order Ledger col S,
      //          Hub Client Dues, CRM Client Dues (via refreshClientDues_)
      // v15.3: re-activated (was locked out entirely in v15.1/v15.2). Per
      // clarification, CRM CAN set/change Due Date in their own sub-tab — just
      // rate-limited to once every 7 days. (Admin/Sales/Accounts editing Due
      // Date directly on the Hub are NOT subject to this — see onMasterHubEdit.)
      // Lock state is tracked in Script Properties keyed by UUID rather than a
      // sheet column, since adding a hidden column to every dynamically-created
      // client sub-tab would mean changing createClientTab_ too.
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_DUE_DATE) {
        if (uuid) {
          const lockKey = 'dueDateLock_' + uuid;
          const props   = PropertiesService.getScriptProperties();
          const oldRaw  = (e.oldValue || '').toString().trim();
          if (oldRaw) {
            // This is a CHANGE to an already-set Due Date — check the lock.
            const lockSetOnStr = props.getProperty(lockKey);
            const lockSetOn    = lockSetOnStr ? new Date(lockSetOnStr) : null;
            const daysSince    = lockSetOn ? (Date.now() - lockSetOn.getTime()) / 86400000 : 999;
            if (daysSince < 7) {
              e.range.setValue(e.oldValue);
              SpreadsheetApp.flush();
              const daysLeft = Math.ceil(7 - daysSince);
              alertIfUi_(
                'Due Date is locked.\n\n' +
                'You already changed it within the last 7 days.\n' +
                'You can change it again in ' + daysLeft + ' day(s).'
              );
              logActivity_(ss, 'CRM DUE DATE CHANGE BLOCKED',
                'UUID:' + uuid + ' | row:' + row + ' | ' + daysLeft + ' day(s) remaining — reverted', userEmail);
              return;
            }
          }
          // Allowed (first-time set, or 7+ days since last change) — stamp now.
          props.setProperty(lockKey, new Date().toISOString());
        }

        // Write to Order Ledger CRM col I
        PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
        const crmDueCell = mainSheet.getRange(mainRow, CRM_COL_DUE_DATE);
        crmDueCell.setValue(cellVal);
        if (cellVal) crmDueCell.setNumberFormat('dd/mm/yyyy');
        SpreadsheetApp.flush();

        // Write to Hub Order Ledger col I
        if (uuid) {
          try {
            const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubSheet) {
              const hubData = hubSheet.getDataRange().getValues();
              let hubRow = -1;
              const subPiNoP2 = sheet.getRange(row, SUB_COL_PI_NO).getValue().toString().trim();
              for (let i = 1; i < hubData.length; i++) {
                const rUUID = (hubData[i][OL_COL_UUID  - 1] || '').toString().trim();
                const rPI   = (hubData[i][OL_COL_PI_NO - 1] || '').toString().trim();
                if (rUUID === uuid && rPI === subPiNoP2) { hubRow = i + 1; break; }
              }
              if (hubRow === -1) {
                for (let i = 1; i < hubData.length; i++) {
                  if ((hubData[i][OL_COL_UUID - 1] || '').toString().trim() === uuid
                      && !(hubData[i][OL_COL_PI_NO - 1] || '').toString().trim()) {
                    hubRow = i + 1; break;
                  }
                }
              }
              if (hubRow !== -1) {
                // FIX v14.3: Use OL_COL_DUE_DATE (col S=19) not HUB_COL_DUE_DATE (col I=9)
                // HUB_COL_DUE_DATE points to Master Hub legacy sheet col I (Freight),
                // but the Order Ledger sheet has Due Date at OL_COL_DUE_DATE = col S (19).
                const hubDueCell = hubSheet.getRange(hubRow, OL_COL_DUE_DATE);
                hubDueCell.setValue(cellVal);
                if (cellVal) hubDueCell.setNumberFormat('dd/mm/yyyy');
                SpreadsheetApp.flush();
                logActivity_(hubSS, 'SUBTAB DUE DATE->ORDER LEDGER',
                  'UUID:' + uuid + ' | DueDate:' + cellVal + ' | col:' + OL_COL_DUE_DATE, userEmail);
              }
            }
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH2 Hub Due Date write error: ' + hubErr.message);
          }
        }

        // Refresh both Client Dues sheets (Hub + this CRM spoke)
        refreshClientDues_(userEmail);

        logActivity_(ss, 'SUBTAB DUE DATE->MAIN+HUB+DUES',
          'Client:' + sheetName + ' | row:' + row + ' | DueDate:' + cellVal, userEmail);

        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 3: AWB No / Dispatch Date
      // Writes to: Hub Order Ledger + Master Hub legacy cols
      // v15.1 NOTE: unreachable — AWB No / Dispatched Date are now
      // Logistics-exclusive via the Hub (see LOCKED_FOR_CRM_SUB_COLS guard
      // above). Left in place rather than deleted; safe to remove later.
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_AWB_NO || col === SUB_COL_DISP_DATE) {
        const olColTarget  = SUB_TO_ORDER_LEDGER[col];
        const hubColTarget = SUB_TO_HUB[col];
        const isDate = col === SUB_COL_DISP_DATE;
        if (uuid) {
          try {
            const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);

            // Write to Hub Order Ledger (PATH3 AWB/Dispatch)
            const olSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (olSheet && olColTarget) {
              const olData = olSheet.getDataRange().getValues();
              let olRow = -1;
              const subPiNoP3 = sheet.getRange(row, SUB_COL_PI_NO).getValue().toString().trim();
              for (let i = 1; i < olData.length; i++) {
                const rUUID = (olData[i][OL_COL_UUID  - 1] || '').toString().trim();
                const rPI   = (olData[i][OL_COL_PI_NO - 1] || '').toString().trim();
                if (rUUID === uuid && rPI === subPiNoP3) { olRow = i + 1; break; }
              }
              if (olRow === -1) {
                for (let i = 1; i < olData.length; i++) {
                  if ((olData[i][OL_COL_UUID - 1] || '').toString().trim() === uuid
                      && !(olData[i][OL_COL_PI_NO - 1] || '').toString().trim()) {
                    olRow = i + 1; break;
                  }
                }
              }
              if (olRow !== -1) {
                const olCell = olSheet.getRange(olRow, olColTarget);
                // AWB No — always write as string to avoid type errors
                const writeVal = col === SUB_COL_AWB_NO ? cellVal.toString() : cellVal;
                olCell.setValue(writeVal);
                if (isDate && cellVal) olCell.setNumberFormat('dd/mm/yyyy');
              }
            }

            // Master Hub removed v15.0

            SpreadsheetApp.flush();
            logActivity_(hubSS, 'SUBTAB AWB/DISP->OL+HUB',
              'UUID:' + uuid + ' col:' + col + '=' + cellVal, userEmail);
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH3 AWB/Disp Hub write error: ' + hubErr.message);
          }
        }
        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 4: Payment Received / Payment Date
      // v15.1 NOTE: unreachable — Payment Received/Date are now
      // Accounts-exclusive via the Hub (see LOCKED_FOR_CRM_SUB_COLS guard
      // above). Also note CRM_COL_PAY_RECEIVED/CRM_COL_PAY_DATE were already
      // commented "filled by Hub script" — this path contradicted that intent
      // even before today's change. Left in place rather than deleted.
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_PAY_RECEIVED || col === SUB_COL_PAY_DATE) {
        const mainColTarget = col === SUB_COL_PAY_RECEIVED ? CRM_COL_PAY_RECEIVED : CRM_COL_PAY_DATE;
        PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
        mainSheet.getRange(mainRow, mainColTarget).setValue(cellVal);
        SpreadsheetApp.flush();

        if (uuid) {
          try {
            const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubSheet) {
              const hubData = hubSheet.getDataRange().getValues();
              let hubRow = -1;
              const subPiNoP4 = sheet.getRange(row, SUB_COL_PI_NO).getValue().toString().trim();
              for (let i = 1; i < hubData.length; i++) {
                const rUUID = (hubData[i][OL_COL_UUID  - 1] || '').toString().trim();
                const rPI   = (hubData[i][OL_COL_PI_NO - 1] || '').toString().trim();
                if (rUUID === uuid && rPI === subPiNoP4) { hubRow = i + 1; break; }
              }
              if (hubRow === -1) {
                for (let i = 1; i < hubData.length; i++) {
                  if ((hubData[i][OL_COL_UUID - 1] || '').toString().trim() === uuid
                      && !(hubData[i][OL_COL_PI_NO - 1] || '').toString().trim()) {
                    hubRow = i + 1; break;
                  }
                }
              }
              if (hubRow !== -1) {
                const hubColTarget = col === SUB_COL_PAY_RECEIVED ? HUB_COL_PAY_RECEIVED : HUB_COL_PAY_DATE;
                const hCell = hubSheet.getRange(hubRow, hubColTarget);
                hCell.setValue(cellVal);
                if (col === SUB_COL_PAY_DATE && cellVal) hCell.setNumberFormat('dd/mm/yyyy');
              }
            }
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH4 Payment Hub write error: ' + hubErr.message);
          }
        }

        // After payment edit: refresh credit totals, client dues, and dashboard
        refreshCreditTotals_(userEmail);
        refreshClientDues_(userEmail);
        refreshDashboard_(userEmail);

        logActivity_(ss, 'SUBTAB PAY->ORDER LEDGER CRM+HUB+DUES',
          'Client:' + sheetName + ' col:' + col, userEmail);

        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 5: Value columns → Full-row sync
      // BUG FIX v14.1 (final): Always read the ENTIRE sub-tab row and push
      // all values in one shot. This way it doesn't matter which trigger
      // fires first or which column was edited — every column is synced.
      //
      // Why: when user edits G, H, J, L fast, 4 triggers fire simultaneously.
      // Only 1 gets the lock; the others wait 12s and then get the lock too
      // (from the new lock strategy above). Each one does a full-row sync,
      // so the last one to run will always have the final correct values.
      //
      // Final Invoice: compute from raw input columns (never read ARRAYFORMULA
      // col M — it may not have recalculated yet). The arithmetic is identical
      // to the formula: FOB = G+H, CIF = FOB+J, Final = CIF-L.
      // ─────────────────────────────────────────────────────────────────────
      SpreadsheetApp.flush();
      // Read the FULL sub-tab row in one call
      const lastSubCol = Math.max(sheet.getLastColumn(), SUB_COL_DUE_DATE);
      const subRowData = sheet.getRange(row, 1, 1, lastSubCol).getValues()[0];
      const piNo       = (subRowData[SUB_COL_PI_NO - 1] || '').toString().trim();

      if (!piNo) {
        Logger.log('onCRMSubTabEdit: PATH 5 — no PI No on row ' + row + ', skipping sync');
      } else {
        // Read raw input values — never formula columns (G, H, J, L are raw inputs)
        const adjItemVal  = Number(subRowData[SUB_COL_ITEM_VALUE_USD  - 1]) || 0;
        const adjDevChg   = Number(subRowData[SUB_COL_DEV_CHARGES_USD - 1]) || 0;
        const adjFreight  = Number(subRowData[SUB_COL_FREIGHT          - 1]) || 0;
        const adjDiscount = Number(subRowData[SUB_COL_DISCOUNT         - 1]) || 0;

        // Override with the just-edited cell value (most current — sheet read may lag)
        const editedVal = Number(e.range.getValue()) || 0;
        const adjItem   = (col === SUB_COL_ITEM_VALUE_USD)  ? editedVal : adjItemVal;
        const adjDev    = (col === SUB_COL_DEV_CHARGES_USD) ? editedVal : adjDevChg;
        const adjFr     = (col === SUB_COL_FREIGHT)         ? editedVal : adjFreight;
        const adjDisc   = (col === SUB_COL_DISCOUNT)        ? editedVal : adjDiscount;

        // Compute Final Invoice — same formula as ARRAYFORMULA in col M
        const fob      = adjItem + adjDev;
        const cif      = fob + adjFr;
        const finalInv = cif - adjDisc;

        Logger.log('onCRMSubTabEdit: PATH 5 FULL-ROW | piNo=' + piNo +
          ' G=' + adjItem + ' H=' + adjDev + ' J=' + adjFr + ' L=' + adjDisc +
          ' FOB=' + fob + ' CIF=' + cif + ' Final=' + finalInv);

        pushFinalInvoiceToMainAndHub_(
          ss, mainSheet, mainRow, mainData,
          sheetName, piNo, finalInv, uuid, crmOwner, userEmail,
          { itemVal: adjItem, devChg: adjDev, freight: adjFr, discount: adjDisc }
        );

        refreshCreditTotals_(userEmail);
        refreshClientDues_(userEmail);
        refreshDashboard_(userEmail);

        logActivity_(ss, 'SUBTAB FULL-ROW SYNC->OL+HUB',
          'Client:' + sheetName + ' row:' + row + ' col:' + col +
          ' G=' + adjItem + ' H=' + adjDev + ' J=' + adjFr + ' L=' + adjDisc +
          ' Final=' + finalInv, userEmail);

        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
      }

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log('ERROR in onCRMSubTabEdit: ' + err.message + '\n' + err.stack);
  }
}

function applySubTabRowColour_(sheet, row, uuid, mainRowData) {
  try {
    if (!sheet || !row || row <= 1) return;

    // Read live values from the sub-tab row for dispatchable columns
    const lastCol = Math.max(sheet.getLastColumn(), SUB_COL_FINAL_INV);
    const subRow  = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

    const payReceived = Number(subRow[SUB_COL_PAY_RECEIVED - 1]) || 0;
    const finalInv    = Number(subRow[SUB_COL_FINAL_INV    - 1]) || 0;
    const awbNo       = (subRow[SUB_COL_AWB_NO - 1]  || '').toString().trim();
    const invNo       = (subRow[SUB_COL_INV_NO - 1]  || '').toString().trim();
    const dueDateRaw  = subRow[SUB_COL_DUE_DATE - 1];
    const dueDate     = dueDateRaw instanceof Date ? dueDateRaw : (dueDateRaw ? new Date(dueDateRaw) : null);

    // Derive payment status string to mirror hub formula logic
    let payStatus = 'Pending';
    if (payReceived > 0 && finalInv > 0) {
      payStatus = payReceived >= finalInv ? 'Received' : 'Partial';
    }

    const status = deriveOrderStatus_(payStatus, dueDate, finalInv, awbNo, invNo);
    const colour = statusColour_(status);

    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(colour);
  } catch (err) {
    Logger.log('applySubTabRowColour_ error: ' + err.message);
  }
}

function onMasterHubEdit(e) {
  invalidateAllSheetCaches_(); // FIX 5: fresh read per trigger
  const _licEmail = (e && e.user) ? e.user.email : '';
  if (!checkLicense(_licEmail)) return;
  if (!e || !e.range) {
    Logger.log('onMasterHubEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  try {
    const sheet = e.range.getSheet();
    const row   = e.range.getRow();
    const col   = e.range.getColumn();

    if (row <= 1) return;
    // Now watches Order Ledger (not 'Master Hub')
    if (sheet.getName() !== ORDER_LEDGER_SHEET_NAME) return;

    // Expanded allowed columns: original + Inv No, Inv Date, Due Date
    const ALLOWED_COLS_V14 = [
      ...HUB_ALLOWED_PUSH_COLS,  // existing: pay, logistics, approved_by, approval_req, dispatch_status
      OL_COL_INV_NO,             // E  — Logistics enters Inv No
      OL_COL_INV_DATE,           // F  — Logistics enters Inv Date
      OL_COL_DUE_DATE,           // S  — Accounts enters Due Date (7-day lock applies)
    ];
    if (!ALLOWED_COLS_V14.includes(col)) return;

    const hubSS    = e.source;
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { Logger.log('onMasterHubEdit: Order Ledger sheet not found'); return; }
    const rowData  = hubSheet.getRange(row, 1, 1, HUB_TOTAL_COLS).getValues()[0];

    const uuid      = rowData[OL_COL_UUID      - 1];
    // v15.3 FIX: Column C can now contain "Name (email)" display format —
    // extractEmail_ safely returns a bare email either way. This is critical
    // here specifically: crmOwner is used directly as the crmFileIds lookup
    // key for every Hub→CRM push below.
    const crmOwner  = extractEmail_(rowData[OL_COL_CRM_OWNER - 1]);
    const newValue  = e.range.getValue();
    const userEmail = Session.getActiveUser().getEmail();

    if (!uuid || !crmOwner) return;

    // ── v15.2 FIX: LOGISTICS BLOCK — prevent AWB/Dispatch until approved ──
    // Previously used HUB_COL_AWB_NO(17)/HUB_COL_DISP_DATE(18)/HUB_COL_APPROVAL_REQ(14)
    // — all legacy "Master Hub" column numbers that don't match Order Ledger's
    // actual layout (AWB No=13, Dispatch Date=14, Approval Required=21). This
    // block could never fire correctly before; fixed to use the real columns,
    // and to the requested rule: blocked only while Approval Required="Yes"
    // AND Approved By is still blank. Once Admin/Sales fills Approved By,
    // Logistics may proceed even though Approval Required stays "Yes" as an
    // audit record of the original breach.
    if (col === OL_COL_AWB_NO || col === OL_COL_DISPATCH_DATE) {
      const approvalStatus = (rowData[OL_COL_APPROVAL_REQ - 1] || '').toString().trim();
      const approvedBy     = (rowData[OL_COL_APPROVED_BY  - 1] || '').toString().trim();
      if (approvalStatus === 'Yes' && !approvedBy) {
        e.range.clearContent();
        SpreadsheetApp.flush();
        alertIfUi_(
          'Dispatch Blocked!\n\n' +
          'This order is on Hold — credit limit has been exceeded.\n' +
          'Col U (Approval Required) = "Yes"\n\n' +
          'Dispatch cannot proceed until Admin or Sales fills in\n' +
          'col V (Approved By) for this row.\n\n' +
          'Contact your Sales Manager or Admin.'
        );
        logActivity_(hubSS, 'LOGISTICS BLOCKED',
          'UUID:' + uuid + ' | row:' + row + ' | col:' + col + ' | U=Yes, V blank — reverted', userEmail);
        return;
      }
    }

    // v15.3: Due Date 7-day lock REMOVED from here. Per clarification, this
    // restriction is meant for CRM editing Due Date in their OWN spoke sheet
    // only — Admin/Sales/Accounts editing Due Date directly on the Hub should
    // be unrestricted, multiple times. The lock now lives in onCRMSubTabEdit
    // instead (see LOCKED_FOR_CRM_SUB_COLS / PATH 2 there). Column X
    // (Due Date Set On) is no longer used on the Hub side as a result — left
    // in place rather than removed, since deleting a live column is riskier
    // than leaving an unused one.

    const config    = getCrmConfig_();
    const crmFileId = config.crmFileIds[(crmOwner || '').toString().toLowerCase()];
    if (!crmFileId) {
      Logger.log('onMasterHubEdit: No CRM file ID for: ' + crmOwner);
      return;
    }

    const lock = LockService.getScriptLock();
    // FIX v14.3: increased from 10s to 25s for large datasets
    try { lock.waitLock(25000); } catch (_) { Logger.log('onMasterHubEdit: lock timeout'); return; }

    try {
      const crmSS    = SpreadsheetApp.openById(crmFileId);
      const crmSheet = crmSS.getSheetByName(ORDER_LEDGER_CRM_NAME);
      if (!crmSheet) { Logger.log('onMasterHubEdit: Order Ledger CRM not found in spoke'); return; }
      const crmData  = crmSheet.getDataRange().getValues();

      let matchRow = -1;
      for (let i = 1; i < crmData.length; i++) {
        if (crmData[i][CRM_COL_UUID - 1] === uuid) { matchRow = i + 1; break; }
      }
      if (matchRow === -1) { Logger.log('onMasterHubEdit: UUID not found in Order Ledger CRM: ' + uuid); return; }

      // ── Push mapped column to Order Ledger CRM ─────────────
      // BUG FIX v14: use OL_COL_* (Order Ledger cols) not HUB_COL_*
      const crmMainCol = HUB_TO_CRM_MAIN[col];
      if (crmMainCol) {
        const pushValue = (col === OL_COL_DISPATCH_STATUS)
          ? hubSheet.getRange(row, col).getDisplayValue()
          : newValue;
        const crmCell = crmSheet.getRange(matchRow, crmMainCol);
        crmCell.setValue(pushValue);
        const isDateOL = [OL_COL_INV_DATE, OL_COL_DISPATCH_DATE,
                          OL_COL_PAY_DATE, OL_COL_DUE_DATE].includes(col);
        if (isDateOL && pushValue) crmCell.setNumberFormat('dd/mm/yyyy');
        logActivity_(crmSS, 'HUB->ORDER LEDGER CRM col:' + col, 'UUID:' + uuid, userEmail);
        logActivity_(hubSS, 'PUSHED TO ORDER LEDGER CRM',        'UUID:' + uuid, userEmail);
      }

      // ── Push to CRM sub-tab ───────────────────────────
      const customerName = rowData[OL_COL_CLIENT_NAME - 1];
      const piNo         = rowData[OL_COL_PI_NO       - 1];
      const clientSheet  = crmSS.getSheetByName((customerName || '').toString().trim());

      if (clientSheet) {
        const subCol = HUB_TO_SUB_TAB[col];
        if (subCol) {
          const clientData = clientSheet.getDataRange().getValues();
          for (let i = 1; i < clientData.length; i++) {
            if ((clientData[i][SUB_COL_PI_NO - 1] || '').toString().trim()
                === (piNo || '').toString().trim()) {
              const targetCell = clientSheet.getRange(i + 1, subCol);
              targetCell.setValue(newValue);
              const isDateSub = [SUB_COL_DISP_DATE, SUB_COL_PAY_DATE,
                                  SUB_COL_INV_DATE,  SUB_COL_DUE_DATE].includes(subCol);
              if (isDateSub && newValue) targetCell.setNumberFormat('dd/mm/yyyy');
              applySubTabRowColour_(clientSheet, i + 1, uuid, crmData[matchRow - 1]);
              logActivity_(crmSS, 'HUB->SUBTAB col:' + subCol, 'PI:' + piNo, userEmail);
              break;
            }
          }
        }
      }

      // ── Order Approved notification ─────────────────────
      // Approval cols not on Order Ledger sheet — no auto-approve trigger needed here

      // ── After payment edit: refresh Credit + Client Dues ─────
      if (col === OL_COL_PAY_RECEIVED || col === OL_COL_PAY_DATE) {
        // v15.1: recompute Payment Status & Due Balance for this row —
        // script-computed (replaces the old live ARRAYFORMULA). Re-read the
        // row fresh since Payment Received may have just changed.
        const freshRow   = hubSheet.getRange(row, 1, 1, HUB_TOTAL_COLS).getValues()[0];
        const finalInvN  = Number(freshRow[OL_COL_FINAL_INVOICE - 1]) || 0;
        const payRecN    = Number(freshRow[OL_COL_PAY_RECEIVED  - 1]) || 0;
        const dueBalN    = Math.max(0, finalInvN - payRecN);
        const payStatusN = dueBalN <= 0 ? 'Received' : (payRecN > 0 ? 'Partial' : 'Pending');
        hubSheet.getRange(row, OL_COL_DUE_BALANCE).setValue(dueBalN);
        hubSheet.getRange(row, OL_COL_PAY_STATUS).setValue(payStatusN);

        refreshCreditTotals_(userEmail);
        refreshClientDues_(userEmail);
      }

      // ── After Due Date edit: refresh Client Dues ──────────
      if (col === OL_COL_DUE_DATE) {
        refreshClientDues_(userEmail);
      }

      // ── After ANY edit: refresh Dashboard ───────────────
      refreshDashboard_(userEmail);

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log('ERROR in onMasterHubEdit: ' + err.message);
  }
}

function refreshCreditTotals_(userEmail) {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    const hubSheet    = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME); // UPDATED
    if (!creditSheet || !hubSheet) return;

    const hubData    = hubSheet.getDataRange().getValues();
    const creditData = creditSheet.getDataRange().getValues();

    // Build per-client totals from Order Ledger
    // BUG FIX: use OL_COL_* constants (Order Ledger columns)
    // not HUB_COL_* (Master Hub columns) — different sheet layouts.
    // BUG FIX v14.1: skip rows with no PI No — these are client-registration
    // placeholder rows (written by onCreditRequestEdit) that have no order data.
    // Counting them inflates Total Invoiced and Total Orders.
    const invTotals  = {};  // key=name.lower → total Final Invoice
    const payTotals  = {};  // key=name.lower → total Payment Received
    const nameDisplay = {}; // key=name.lower → original display name
    for (let i = 1; i < hubData.length; i++) {
      const name = (hubData[i][OL_COL_CLIENT_NAME  - 1] || '').toString().trim();
      const piNo = (hubData[i][OL_COL_PI_NO        - 1] || '').toString().trim();
      if (!name || !piNo) continue;  // skip registration-only rows (no PI No)
      const inv  = Number(hubData[i][OL_COL_FINAL_INVOICE - 1]) || 0;
      const pay  = Number(hubData[i][OL_COL_PAY_RECEIVED  - 1]) || 0;
      const key = name.toLowerCase();
      invTotals[key]   = (invTotals[key]  || 0) + inv;
      payTotals[key]   = (payTotals[key]  || 0) + pay;
      // Keep the first non-lowercase display name seen
      if (!nameDisplay[key]) nameDisplay[key] = name;
    }

    const now = new Date();
    for (let i = 1; i < creditData.length; i++) {
      const clientName  = (creditData[i][CC_COL_CLIENT_NAME - 1] || '').toString().trim();
      if (!clientName) continue;
      const creditLimit = Number(creditData[i][CC_COL_CREDIT_LIMIT - 1]) || 0;
      const key         = clientName.toLowerCase();
      const usedCredit  = Math.max(0, (invTotals[key] || 0) - (payTotals[key] || 0));
      const availCredit = creditLimit - usedCredit;
      const sheetRow    = i + 1;
      creditSheet.getRange(sheetRow, CC_COL_USED_CREDIT ).setValue(usedCredit);
      creditSheet.getRange(sheetRow, CC_COL_AVAIL_CREDIT).setValue(availCredit);
      creditSheet.getRange(sheetRow, CC_COL_LAST_UPDATED).setValue(now);
    }
    SpreadsheetApp.flush();

    // ── Update Client Summary sheet ───────────────────────────
    updateClientSummary_(hubSS, invTotals, payTotals, nameDisplay, userEmail);

    // Re-evaluate col N credit flags
    reEvaluateAllCreditFlags_(hubSS, userEmail);

    Logger.log('refreshCreditTotals_: refreshed ' + (creditData.length - 1) + ' row(s)');
  } catch (err) {
    Logger.log('ERROR in refreshCreditTotals_: ' + err.message);
  }
}

function updateClientSummary_(hubSS, invTotals, payTotals, nameDisplay, userEmail) {
  try {
    const summarySheet = hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME);
    if (!summarySheet) return;

    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) return;
    const creditData  = creditSheet.getDataRange().getValues();

    // Build credit limit map — preserve original display name from Credit sheet
    const limitMap    = {};
    const crmOwnerMap = {};
    const creditNameDisplay = {}; // lowercase → Credit sheet display name
    for (let i = 1; i < creditData.length; i++) {
      const name  = (creditData[i][CC_COL_CLIENT_NAME  - 1] || '').toString().trim();
      const owner = (creditData[i][CC_COL_CRM_OWNER    - 1] || '').toString().trim();
      const limit = Number(creditData[i][CC_COL_CREDIT_LIMIT - 1]) || 0;
      if (name) {
        const key = name.toLowerCase();
        limitMap[key]          = limit;
        crmOwnerMap[key]       = owner;
        creditNameDisplay[key] = name; // exact case from Credit sheet
      }
    }

    // Build per-client order count and last order date from Order Ledger
    // BUG FIX v14.1: skip rows with no PI No — registration-only placeholder rows
    // must not be counted as orders (they inflate Total Orders count).
    const hubSheet2   = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    const hubData2    = hubSheet2 ? hubSheet2.getDataRange().getValues() : [];
    const orderCount  = {};  // key → number of orders
    const lastOrdDate = {};  // key → most recent inv date
    for (let i = 1; i < hubData2.length; i++) {
      const name    = (hubData2[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
      const piNo    = (hubData2[i][OL_COL_PI_NO       - 1] || '').toString().trim();
      const invDate = hubData2[i][OL_COL_INV_DATE - 1];
      if (!name || !piNo) continue;  // skip registration-only rows (no PI No)
      orderCount[name] = (orderCount[name] || 0) + 1;
      if (invDate) {
        const d = invDate instanceof Date ? invDate : new Date(invDate);
        if (!lastOrdDate[name] || d > lastOrdDate[name]) lastOrdDate[name] = d;
      }
    }

    // Build output rows — columns aligned to CS_COL_* constants:
    // A=Client Name | B=CRM Owner | C=Total Orders | D=Total Invoiced
    // E=Total Paid  | F=Total Due | G=Credit Limit | H=Used Credit
    // I=Available Credit | J=Credit Status | K=Last Order Date
    const outputRows = [];
    const allClients = new Set([
      ...Object.keys(invTotals),
      ...Object.keys(limitMap),
    ]);
    allClients.forEach(key => {
      const totalInv  = invTotals[key] || 0;
      const totalPay  = payTotals[key] || 0;
      const totalDue  = Math.max(0, totalInv - totalPay);
      const limit     = limitMap[key] || 0;
      const usedCred  = Math.max(0, totalInv - totalPay);
      const availCred = limit - usedCred;
      let creditStatus = 'OK';
      if (limit > 0 && usedCred >= limit)       creditStatus = 'Exceeded';
      else if (limit > 0 && usedCred >= limit * 0.9) creditStatus = 'Near Limit';
      const displayName = creditNameDisplay[key]
                       || (nameDisplay && nameDisplay[key])
                       || key;
      outputRows.push([
        displayName,                  // A  CS_COL_CLIENT_NAME
        crmOwnerMap[key] || '',       // B  CS_COL_CRM_OWNER
        orderCount[key]  || 0,        // C  CS_COL_TOTAL_ORDERS
        totalInv,                     // D  CS_COL_TOTAL_INVOICED
        totalPay,                     // E  CS_COL_TOTAL_PAID
        totalDue,                     // F  CS_COL_TOTAL_DUE
        limit,                        // G  CS_COL_CREDIT_LIMIT
        usedCred,                     // H  CS_COL_USED_CREDIT
        availCred,                    // I  CS_COL_AVAIL_CREDIT
        creditStatus,                 // J  CS_COL_CREDIT_STATUS
        lastOrdDate[key] || '',       // K  CS_COL_LAST_ORDER_DATE
      ]);
    });

    if (!outputRows.length) return;

    const headers = [
      'Client Name', 'CRM Owner', 'Total Orders', 'Total Invoiced',
      'Total Paid', 'Total Due', 'Credit Limit', 'Used Credit',
      'Available Credit', 'Credit Status', 'Last Order Date'
    ];
    summarySheet.clearContents();
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF');
    summarySheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    // Format numeric money cols D/E/F/G/H/I
    summarySheet.getRange(2, 4, outputRows.length, 6).setNumberFormat('#,##0.00');
    // Format date col K
    summarySheet.getRange(2, 11, outputRows.length, 1).setNumberFormat('dd/mm/yyyy');
    SpreadsheetApp.flush();

    Logger.log('updateClientSummary_: wrote ' + outputRows.length + ' client row(s)');
  } catch (err) {
    Logger.log('ERROR in updateClientSummary_: ' + err.message);
  }
}

function refreshClientDues_(userEmail) {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) return;

    const hubData = hubSheet.getDataRange().getValues();
    const now     = new Date();

    // Build dues rows from Order Ledger
    // BUG FIX: use OL_COL_* not HUB_COL_* — different sheet layouts
    const duesRows = [];
    for (let i = 1; i < hubData.length; i++) {
      const row        = hubData[i];
      const customer   = (row[OL_COL_CLIENT_NAME   - 1] || '').toString().trim();
      // v15.3 FIX: Column C can now contain "Name (email)" display format —
      // extractEmail_ safely returns a bare email either way. Critical here:
      // this gets matched exactly against crmFileIds keys below to decide
      // which CRM spoke's own Client Dues sheet each row belongs to.
      const crmOwner   = extractEmail_(row[OL_COL_CRM_OWNER - 1]);
      const piNo       = (row[OL_COL_PI_NO          - 1] || '').toString().trim();
      const invNo      = (row[OL_COL_INV_NO         - 1] || '').toString().trim();
      const invDate    = row[OL_COL_INV_DATE        - 1];
      const invValue   = Number(row[OL_COL_FINAL_INVOICE - 1]) || 0;
      const dueDate    = row[OL_COL_DUE_DATE        - 1];
      const payRec     = Number(row[OL_COL_PAY_RECEIVED  - 1]) || 0;
      const payStatus  = (row[OL_COL_PAY_STATUS     - 1] || '').toString().trim();
      const dueBalance = Number(row[OL_COL_DUE_BALANCE   - 1]) || Math.max(0, invValue - payRec);
      const awbNo      = (row[OL_COL_AWB_NO         - 1] || '').toString().trim();
      const uuid       = (row[OL_COL_UUID           - 1] || '').toString().trim();

      if (!customer) continue;
      if (dueBalance <= 0) continue;  // only outstanding rows

      const dueDateObj = dueDate instanceof Date ? dueDate : (dueDate ? new Date(dueDate) : null);
      const status     = deriveOrderStatus_(payStatus, dueDateObj, invValue, awbNo, invNo);

      duesRows.push({
        customer, crmOwner, piNo, invNo, invDate, invValue,
        dueDate: dueDateObj, payStatus, payRec, dueBalance, status, uuid,
        // v15.2 FIX: column H now shows the actual amount received, not the
        // Pending/Partial/Received status text (that's still derived into
        // the separate "Status" column J via deriveOrderStatus_ below).
        row: [customer, crmOwner, piNo, invNo, invDate, invValue,
              dueDateObj, payRec, dueBalance, status, uuid],
      });
    }

    // Sort: Overdue first, Partial second, others after (alphabetical within groups)
    const sortPriority = s => {
      if (s === 'Overdue')  return 0;
      if (s === 'Partial')  return 1;
      return 2;
    };
    duesRows.sort((a, b) => {
      const pDiff = sortPriority(a.status) - sortPriority(b.status);
      if (pDiff !== 0) return pDiff;
      return a.customer.toLowerCase().localeCompare(b.customer.toLowerCase());
    });

    // Hub Client Dues removed v15.0 — writing to Client Dues only

    // ── Write filtered version to each CRM spoke ─────────────
    const config = getCrmConfig_();
    Object.entries(config.crmFileIds).forEach(([gmail, fileId]) => {
      try {
        const crmSS       = SpreadsheetApp.openById(fileId);
        const spokeDues   = duesRows.filter(r =>
          r.crmOwner.toLowerCase() === gmail.toLowerCase()
        );
        writeDuesSheet_(crmSS, spokeDues, CLIENT_DUES_SHEET_NAME);
      } catch (spokeErr) {
        Logger.log('refreshClientDues_: spoke write error for ' + gmail + ': ' + spokeErr.message);
      }
    });

    // ── Also write to Hub's own CLIENT_DUES_SHEET_NAME ───────
    writeDuesSheet_(hubSS, duesRows, CLIENT_DUES_SHEET_NAME);

    Logger.log('refreshClientDues_: wrote ' + duesRows.length + ' due row(s)');
  } catch (err) {
    Logger.log('ERROR in refreshClientDues_: ' + err.message);
  }
}

function writeDuesSheet_(ss, duesRows, sheetName) {
  try {
    let duesSheet = ss.getSheetByName(sheetName);
    if (!duesSheet) {
      duesSheet = ss.insertSheet(sheetName);
    }

    const headers = [
      'Customer', 'CRM Owner', 'PI No', 'Inv No', 'Inv Date',
      'Inv Value', 'Due Date', 'Amount Received', 'Due Balance', 'Status', 'UUID'
    ];

    duesSheet.clearContents();
    duesSheet.clearFormats();

    // Header row
    duesSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#1A237E')
      .setFontColor('#FFFFFF');

    if (!duesRows.length) {
      SpreadsheetApp.flush();
      return;
    }

    // Data rows
    const dataMatrix = duesRows.map(r => r.row);
    duesSheet.getRange(2, 1, dataMatrix.length, headers.length).setValues(dataMatrix);

    // Apply date format to Inv Date (col E=5) and Due Date (col G=7)
    duesSheet.getRange(2, 5, dataMatrix.length, 1).setNumberFormat('dd/mm/yyyy');
    duesSheet.getRange(2, 7, dataMatrix.length, 1).setNumberFormat('dd/mm/yyyy');
    // v15.2 NEW: Amount Received (col H=8) is now a number, not a status string
    duesSheet.getRange(2, 8, dataMatrix.length, 1).setNumberFormat('#,##0.00');

    // Apply row colour based on status (col J = index 9)
    duesRows.forEach((r, idx) => {
      const colour = statusColour_(r.status);
      duesSheet.getRange(idx + 2, 1, 1, headers.length).setBackground(colour);
    });

    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('writeDuesSheet_ error for "' + sheetName + '": ' + err.message);
  }
}

function refreshDashboard_(userEmail) {
  try {
    const hubSS        = SpreadsheetApp.openById(MASTER_HUB_ID);
    const summarySheet = hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME);
    const dashSheet    = hubSS.getSheetByName(DASHBOARD_SHEET_NAME);
    if (!summarySheet || !dashSheet) return;

    const summaryData = summarySheet.getDataRange().getValues();
    if (summaryData.length < 2) return;

    // Also read Order Ledger to know which clients have overdue rows
    const hubSheet  = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    const hubData   = hubSheet ? hubSheet.getDataRange().getValues() : [];

    // Build per-client status flags from Order Ledger
    // BUG FIX v14.1: skip rows with no PI No — registration-only rows must not
    // contribute overdue/partial flags or last-order-date to the dashboard.
    // BUG FIX: use OL_COL_* not HUB_COL_*
    const hasOverdue  = {};
    const hasPartial  = {};
    const now         = new Date();
    for (let i = 1; i < hubData.length; i++) {
      const row      = hubData[i];
      const customer = (row[OL_COL_CLIENT_NAME   - 1] || '').toString().trim().toLowerCase();
      const piNo     = (row[OL_COL_PI_NO         - 1] || '').toString().trim();
      if (!customer || !piNo) continue;  // skip registration-only rows
      const payStatus= (row[OL_COL_PAY_STATUS    - 1] || '').toString().trim();
      const dueDate  = row[OL_COL_DUE_DATE       - 1];
      const invNo    = (row[OL_COL_INV_NO        - 1] || '').toString().trim();
      const invValue = Number(row[OL_COL_FINAL_INVOICE - 1]) || 0;
      const awbNo    = (row[OL_COL_AWB_NO        - 1] || '').toString().trim();
      const dueDateObj = dueDate instanceof Date ? dueDate : (dueDate ? new Date(dueDate) : null);
      const status   = deriveOrderStatus_(payStatus, dueDateObj, invValue, awbNo, invNo);
      if (status === 'Overdue') hasOverdue[customer] = true;
      if (status === 'Partial') hasPartial[customer] = true;
    }

    // Read Client Summary — use CS_COL_* constants (1-based) mapped to 0-based array indices
    // CS_COL_CLIENT_NAME=1(A), CS_COL_CRM_OWNER=2(B), CS_COL_TOTAL_ORDERS=3(C),
    // CS_COL_TOTAL_INVOICED=4(D), CS_COL_TOTAL_PAID=5(E), CS_COL_TOTAL_DUE=6(F),
    // CS_COL_CREDIT_LIMIT=7(G), CS_COL_USED_CREDIT=8(H), CS_COL_AVAIL_CREDIT=9(I),
    // CS_COL_CREDIT_STATUS=10(J), CS_COL_LAST_ORDER_DATE=11(K)
    const clients = [];
    for (let i = 1; i < summaryData.length; i++) {
      const r = summaryData[i];
      const clientName  = (r[CS_COL_CLIENT_NAME    - 1] || '').toString().trim();
      const crmOwner    = (r[CS_COL_CRM_OWNER      - 1] || '').toString().trim();
      const totalInv    = Number(r[CS_COL_TOTAL_INVOICED - 1]) || 0;
      const totalPaid   = Number(r[CS_COL_TOTAL_PAID     - 1]) || 0;
      const totalDue    = Number(r[CS_COL_TOTAL_DUE      - 1]) || 0;
      const creditLimit = Number(r[CS_COL_CREDIT_LIMIT   - 1]) || 0;
      const availCredit = Number(r[CS_COL_AVAIL_CREDIT   - 1]) || 0;
      if (!clientName) continue;
      clients.push({ clientName, crmOwner, totalInv, totalPaid, totalDue, creditLimit, availCredit });
    }

    // Sort: Total Due high → low; ties alphabetically
    clients.sort((a, b) => {
      if (b.totalDue !== a.totalDue) return b.totalDue - a.totalDue;
      return a.clientName.toLowerCase().localeCompare(b.clientName.toLowerCase());
    });

    // Build dashboard rows — columns aligned to DB_COL_* constants:
    // A=Client Name | B=CRM Owner | C=Total Due | D=Total Invoiced
    // E=Total Paid  | F=Overdue Orders | G=Credit Status | H=Last Order Date
    const headers = [
      'Client Name', 'CRM Owner', 'Total Due', 'Total Invoiced',
      'Total Paid', 'Unpaid Orders', 'Credit Status', 'Last Order Date'
    ];

    dashSheet.clearContents();
    dashSheet.clearFormats();

    dashSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#0D47A1')
      .setFontColor('#FFFFFF');

    if (!clients.length) { SpreadsheetApp.flush(); return; }

    // Build per-client last-order-date map from Order Ledger
    // BUG FIX v14.1: skip rows with no PI No (registration-only rows)
    const lastOrderMap = {};
    for (let i = 1; i < hubData.length; i++) {
      const name    = (hubData[i][OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
      const piNo    = (hubData[i][OL_COL_PI_NO       - 1] || '').toString().trim();
      const invDate = hubData[i][OL_COL_INV_DATE - 1];
      if (name && piNo && invDate) {
        const d = invDate instanceof Date ? invDate : new Date(invDate);
        if (!lastOrderMap[name] || d > lastOrderMap[name]) lastOrderMap[name] = d;
      }
    }

    const dataMatrix = clients.map((c) => {
      const key         = c.clientName.toLowerCase();
      // Credit status derived from totalDue vs hasOverdue/hasPartial
      let creditStatus = 'Settled';
      if (c.totalDue > 0) {
        if (hasOverdue[key])      creditStatus = 'Overdue';
        else if (hasPartial[key]) creditStatus = 'Partial';
        else                      creditStatus = 'Outstanding';
      }
      // Unpaid order count = all orders for this client where payment is Pending or Partial
      const overdueCount = hubData.slice(1).filter(r => {
        const rName = (r[OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase();
        if (rName !== key) return false;
        const piNo      = (r[OL_COL_PI_NO      - 1] || '').toString().trim();
        if (!piNo) return false;  // skip registration-only rows
        const payStatus = (r[OL_COL_PAY_STATUS - 1] || '').toString().trim().toLowerCase();
        return payStatus === 'pending' || payStatus === 'partial';
      }).length;

      return [
        c.clientName,           // A  DB_COL_CLIENT_NAME
        c.crmOwner,             // B  DB_COL_CRM_OWNER
        c.totalDue,             // C  DB_COL_TOTAL_DUE
        c.totalInv,             // D  DB_COL_TOTAL_INVOICED
        c.totalPaid,            // E  DB_COL_TOTAL_PAID
        overdueCount,           // F  DB_COL_OVERDUE_COUNT
        creditStatus,           // G  DB_COL_CREDIT_STATUS
        lastOrderMap[key] || '',// H  DB_COL_LAST_ORDER_DATE
      ];
    });

    dashSheet.getRange(2, 1, dataMatrix.length, headers.length).setValues(dataMatrix);

    // Apply row colours based on Credit Status (col G = index 6)
    dataMatrix.forEach((row, idx) => {
      const dashRow    = idx + 2;
      const totalDue   = row[2]; // C = index 2
      const status     = row[6]; // G = index 6

      let colour = '#E8F5E9'; // green — settled
      if (totalDue > 0) {
        if (status === 'Overdue')      colour = '#FFCDD2'; // red
        else if (status === 'Partial') colour = '#FFF9C4'; // yellow
        else                           colour = '#FFE0B2'; // orange — outstanding
      }
      dashSheet.getRange(dashRow, 1, 1, headers.length).setBackground(colour);
    });

    // Format numeric columns — C=TotalDue, D=TotalInvoiced, E=TotalPaid
    dashSheet.getRange(2, 3, dataMatrix.length, 3).setNumberFormat('#,##0.00');
    // Format H=Last Order Date
    if (dataMatrix.length > 0) {
      dashSheet.getRange(2, 8, dataMatrix.length, 1).setNumberFormat('dd/mm/yyyy');
    }

    // Freeze header and auto-resize
    dashSheet.setFrozenRows(1);
    dashSheet.autoResizeColumns(1, headers.length);

    SpreadsheetApp.flush();
    Logger.log('refreshDashboard_: wrote ' + dataMatrix.length + ' client row(s)');
  } catch (err) {
    Logger.log('ERROR in refreshDashboard_: ' + err.message);
  }
}

function menuRefreshClientDues_() {
  const userEmail = Session.getActiveUser().getEmail();
  refreshClientDues_(userEmail);
  alertIfUi_('✅ Client Dues refreshed on Hub + all CRM spokes.');
}

function runRefreshDashboard() { menuRefreshDashboard_(); }

function menuRefreshDashboard_() {
  const userEmail = Session.getActiveUser().getEmail();
  refreshCreditTotals_(userEmail);  // also updates Client Summary
  refreshClientDues_(userEmail);
  refreshDashboard_(userEmail);
  alertIfUi_('✅ Dashboard refreshed (Credit Totals → Client Summary → Client Dues → Dashboard).');
}


// ============================================================
// SECTION 4 — AUTOMATION & ALERTS (Session 3 — new in v14.0)
// ============================================================

// ============================================================


// ── Script Property keys (Session 3) ─────────────────────────
const PROP_REMINDERS_ON       = 'daily_reminders_on';       // 'true' | 'false'
const PROP_REMINDER_TRIGGER   = 'daily_reminder_trigger_id'; // trigger UID
const PROP_SOFT_DUE_PREFIX    = 'soft_due_';                 // + uuid = ISO date string


// ============================================================
// 1. toggleReminders_()
// ─────────────────────────────────────────────────────────────
// Menu function — toggles daily 9AM reminder trigger ON/OFF.
// State is persisted in Script Properties so it survives
// Apps Script restarts.  Menu label is rebuilt after each
// toggle so the user always sees the current state.
// ============================================================


// ============================================================
// buildAutomationMenu_()
// ─────────────────────────────────────────────────────────────
// Called by toggleReminders_() and onOpen() to add the
// Automation sub-menu with the current ON/OFF label.
// Safe to call multiple times — addToUi() appends or updates.
// ============================================================


// ============================================================
// 2. runDailyReminders_()
// ─────────────────────────────────────────────────────────────
// Triggered daily at 9AM (or run manually).
// Reads ALL Order Ledger rows from the Master Hub and emails
// the assigned CRM (and escalation targets) based on due-date
// state.  Every email sent is logged in the Activity Log.
//
// Alert rules (evaluated in priority order per row):
//   a) Due Date is exactly 7 days away     → warning to CRM
//   b) Due Date is today                   → due-today to CRM
//   c) Due Date passed 3+ days ago         → escalation to CRM + Sales + Admin
//   d) No Due Date, Invoice Date > 30 days → alert Sales: "No Due Date set"
// ============================================================


// ── Email helpers for runDailyReminders_ ─────────────────────


// Shared HTML email template for reminder/warning/due-today/escalation


// Safe GmailApp.sendEmail wrapper — logs errors without throwing


// ============================================================
// 3. menuSendOrderSummary_()
// ─────────────────────────────────────────────────────────────
// Manual menu button on the CRM Spoke sheet.
// • Active sheet name = client name
// • Reads client email from Client Credit sheet col F
//   (CC_COL_REQ_INCREASE = 6, but email is expected in col F
//    per spec; mapped to a separate lookup below)
// • Reads all order rows from the client sub-tab
// • Builds a Gmail DRAFT (does NOT send)
// Subject: "Payment Summary — [Client Name] — [Date]"
// Body: professional HTML table with totals + dispatch details
// ============================================================


// ============================================================
// 4. checkDuplicateEntry_(sheet, row, col, value)
// ─────────────────────────────────────────────────────────────
// Helper — checks whether `value` already exists anywhere else
// in the same column on `sheet` (skips the header row and the
// current row itself so a re-save of the same value is safe).
//
// Returns: { isDuplicate: Boolean, existingRow: Number|null }
// Called from onCRMSubTabEdit when col === SUB_COL_PI_NO or
// col === SUB_COL_INV_NO to prevent duplicate PI / Inv numbers.
// ============================================================


// ============================================================
// 5. applySubTabRowColour_(sheet, row, rowData)
// ─────────────────────────────────────────────────────────────
// Helper — applies a background colour to the entire data row
// on a client sub-tab based on the order's current state.
//
// State is derived from rowData (0-based array, one sub-tab row):
//   • Payment Received (col 14) vs Final Invoice (col 13)
//   • Dispatch Date    (col 6)
//   • AWB No           (col 5)
//   • Invoice No       (col 2)
//   • Due Date comes from Hub — not available here, so Overdue
//     is signalled by the caller passing overdue=true optionally
//     via rowData[15] = 'OVERDUE' sentinel (see usage note below).
//
// Colour map:
//   White         (#FFFFFF)  = Created (PI No set, no invoice yet)
//   Light Blue    (#E3F2FD)  = Invoiced
//   Light Orange  (#FFF3E0)  = Dispatched
//   Yellow        (#FFF9C4)  = Payment Pending
//   Red           (#FFCDD2)  = Overdue
//   Light Yellow  (#FFFFF0)  = Partial payment received
//   Light Green   (#E8F5E9)  = Paid in full
//
// Usage note: to signal an Overdue row from onCRMSubTabEdit,
// pass 'OVERDUE' as rowData[15] (a synthetic 16th element).
// ============================================================


// ============================================================
// 6. autoSetDueDate_()
// ─────────────────────────────────────────────────────────────
// Called when an Invoice Date is entered on a client sub-tab.
// Computes Invoice Date + 30 days and stores it as a
// "soft alert threshold" in Script Properties (keyed by UUID).
//
// This does NOT auto-fill the Due Date cell in the sheet.
// The CRM must enter the Due Date manually.  The stored
// threshold is used by runDailyReminders_() via rule (d)
// to detect invoices with no Due Date set and alert Sales.
//
// HOW TO CALL:
//   autoSetDueDate_(uuid, invoiceDateValue)
//   e.g. from onCRMSubTabEdit when col === SUB_COL_INV_DATE:
//     const uuid = mainData[mainRow-1][CRM_COL_UUID - 1];
//     autoSetDueDate_(uuid, e.range.getValue());
//
// Parameters:
//   uuid         — the order's UUID string (from CRM Main col A)
//   invDateValue — the raw Date value from the Invoice Date cell
// ============================================================


// ── Companion reader: retrieve soft-due threshold for a UUID ─
// Returns a Date or null.  Used by runDailyReminders_() rule (d)
// to check "no Due Date, invoice date > 30 days ago".


// ============================================================


function toggleReminders_() {
  try {
    const props     = PropertiesService.getScriptProperties();
    const isOn      = props.getProperty(PROP_REMINDERS_ON) === 'true';
    const ui        = SpreadsheetApp.getUi();

    if (isOn) {
      // ── Turn OFF: delete the time-based trigger ────────────
      const savedId = props.getProperty(PROP_REMINDER_TRIGGER);
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'runDailyReminders_') {
          if (!savedId || t.getUniqueId() === savedId) {
            ScriptApp.deleteTrigger(t);
          }
        }
      });
      props.setProperty(PROP_REMINDERS_ON, 'false');
      props.deleteProperty(PROP_REMINDER_TRIGGER);
      Logger.log('toggleReminders_: Daily reminders turned OFF');
      alertIfUi_('🔕 Daily payment reminders turned OFF.\n\nNo further automated emails will be sent until you turn them back on.');

    } else {
      // ── Turn ON: install daily 9AM trigger ────────────────
      // Remove any stale runDailyReminders_ triggers first
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'runDailyReminders_') {
          ScriptApp.deleteTrigger(t);
        }
      });

      const trigger = ScriptApp.newTrigger('runDailyReminders_')
        .timeBased()
        .everyDays(1)
        .atHour(9)
        .create();

      props.setProperty(PROP_REMINDERS_ON,     'true');
      props.setProperty(PROP_REMINDER_TRIGGER, trigger.getUniqueId());
      Logger.log('toggleReminders_: Daily reminders turned ON, trigger=' + trigger.getUniqueId());
      alertIfUi_('🔔 Daily payment reminders turned ON.\n\nRunDailyReminders_() will fire every day at 9 AM and email CRM owners, Sales, and Admin as needed.');
    }

    // Rebuild menu so label reflects new state
    buildAutomationMenu_();

  } catch (err) {
    Logger.log('ERROR in toggleReminders_: ' + err.message);
    alertIfUi_('Error toggling reminders: ' + err.message);
  }
}

function buildAutomationMenu_() {
  try {
    const props  = PropertiesService.getScriptProperties();
    const isOn   = props.getProperty(PROP_REMINDERS_ON) === 'true';
    const label  = isOn
      ? '🔔 Daily Reminders: ON  (click to turn OFF)'
      : '🔕 Daily Reminders: OFF (click to turn ON)';

    SpreadsheetApp.getUi()
      .createMenu('⚡ Automation')
      .addItem(label,                          'toggleReminders_')
      .addItem('📧 Send Order Summary Draft',  'menuSendOrderSummary_')
      .addItem('▶ Run Daily Reminders Now',    'runDailyReminders_')
      .addToUi();
  } catch (err) {
    // No UI context (e.g. time trigger calling this) — silently skip
    Logger.log('buildAutomationMenu_: no UI context — ' + err.message);
  }
}

function runDailyReminders_() {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) {
      Logger.log('runDailyReminders_: Hub sheet not found');
      return;
    }

    const config   = getCrmConfig_();
    const data     = hubSheet.getDataRange().getValues();
    const tz       = Session.getScriptTimeZone();
    const today    = new Date();
    today.setHours(0, 0, 0, 0);

    const MS_PER_DAY = 24 * 60 * 60 * 1000;

    let emailsSent = 0;

    for (let i = 1; i < data.length; i++) {
      const row        = data[i];
      const customer   = (row[HUB_COL_CUSTOMER    - 1] || '').toString().trim();
      const crmOwner   = (row[HUB_COL_CRM_OWNER   - 1] || '').toString().trim().toLowerCase();
      const piNo       = (row[HUB_COL_PI_NO        - 1] || '').toString().trim();
      const invNo      = (row[HUB_COL_INV_NO       - 1] || '').toString().trim();
      const invValue   = row[HUB_COL_INV_VALUE     - 1];
      const payStatus  = (row[HUB_COL_PAY_STATUS   - 1] || '').toString().trim();
      const dueBalance = row[HUB_COL_DUE_BALANCE   - 1];
      const dueDateRaw = row[HUB_COL_DUE_DATE      - 1];
      const invDateRaw = row[HUB_COL_DATE          - 1];

      if (!customer || !piNo) continue;

      // Only process unpaid / partially-paid orders
      const isPaid = (payStatus === 'Received' || payStatus === 'Paid');
      if (isPaid) continue;
      // Also skip rows with no invoice value set yet
      if (!invValue || Number(invValue) === 0) continue;

      // Resolve CRM owner email
      const crmEmail = crmOwner || null;
      if (!crmEmail) continue;

      // Format helpers
      const fmtDate = d => d instanceof Date
        ? Utilities.formatDate(d, tz, 'dd-MMM-yyyy')
        : '—';

      // ── Parse dates ────────────────────────────────────────
      let dueDate  = dueDateRaw  instanceof Date ? new Date(dueDateRaw)  : null;
      let invDate  = invDateRaw  instanceof Date ? new Date(invDateRaw)  : null;
      if (dueDate)  dueDate.setHours(0, 0, 0, 0);
      if (invDate)  invDate.setHours(0, 0, 0, 0);

      const daysUntilDue  = dueDate ? Math.round((dueDate - today) / MS_PER_DAY) : null;
      const daysSinceInv  = invDate ? Math.round((today   - invDate) / MS_PER_DAY) : null;

      // ── Common email data object ───────────────────────────
      const orderInfo = {
        customer,
        piNo,
        invNo    : invNo       || '—',
        invValue : invValue     || '—',
        balance  : dueBalance  || '—',
        dueDate  : fmtDate(dueDate),
        invDate  : fmtDate(invDate),
        payStatus,
        crmEmail,
      };

      let triggered = false;

      // ── Rule (c): OVERDUE 3+ days ago — highest priority ──
      if (dueDate && daysUntilDue <= -3) {
        const daysOverdue = Math.abs(daysUntilDue);
        sendEscalationEmail_(hubSS, orderInfo, daysOverdue, config);
        logActivity_(hubSS, 'DAILY REMINDER — OVERDUE ESCALATION',
          `PI:${piNo} | ${customer} | ${daysOverdue} days overdue | To:${crmEmail}+Sales+Admin`,
          'system-trigger');
        emailsSent++;
        triggered = true;
      }

      // ── Rule (b): DUE TODAY ────────────────────────────────
      if (!triggered && dueDate && daysUntilDue === 0) {
        sendDueTodayEmail_(hubSS, orderInfo, crmEmail);
        logActivity_(hubSS, 'DAILY REMINDER — DUE TODAY',
          `PI:${piNo} | ${customer} | To:${crmEmail}`, 'system-trigger');
        emailsSent++;
        triggered = true;
      }

      // ── Rule (a): DUE IN 7 DAYS — warning ─────────────────
      if (!triggered && dueDate && daysUntilDue === 7) {
        sendWarningEmail_(hubSS, orderInfo, crmEmail);
        logActivity_(hubSS, 'DAILY REMINDER — 7-DAY WARNING',
          `PI:${piNo} | ${customer} | Due:${orderInfo.dueDate} | To:${crmEmail}`,
          'system-trigger');
        emailsSent++;
        triggered = true;
      }

      // ── Rule (d): NO DUE DATE, Invoice > 30 days old ──────
      if (!triggered && !dueDate && daysSinceInv !== null && daysSinceInv > 30) {
        const salesEmails = config.salesEmails;
        if (salesEmails.length) {
          sendNoDueDateAlert_(hubSS, orderInfo, salesEmails, daysSinceInv);
          logActivity_(hubSS, 'DAILY REMINDER — NO DUE DATE',
            `PI:${piNo} | ${customer} | Inv ${daysSinceInv} days ago | To:Sales`,
            'system-trigger');
          emailsSent++;
        }
      }
    }

    Logger.log('runDailyReminders_: complete — ' + emailsSent + ' email(s) sent');

  } catch (err) {
    Logger.log('ERROR in runDailyReminders_: ' + err.message);
  }
}

function sendWarningEmail_(hubSS, o, toEmail) {
  const subject = `⚠️ Payment Due in 7 Days — ${o.customer} — PI #${o.piNo}`;
  const html    = reminderEmailHtml_({
    headerColor : '#F57F17',
    headerLabel : '⚠️ Payment Due in 7 Days',
    badgeColor  : '#F57F17',
    badgeText   : '7-DAY WARNING',
    message     : `The invoice for <strong>${o.customer}</strong> is due in <strong>7 days</strong>. Please follow up with the client to ensure payment is arranged on time.`,
    orderInfo   : o,
  });
  safeSendEmail_(toEmail, subject, html);
}

function sendDueTodayEmail_(hubSS, o, toEmail) {
  const subject = `🔴 Payment Due TODAY — ${o.customer} — PI #${o.piNo}`;
  const html    = reminderEmailHtml_({
    headerColor : '#E65100',
    headerLabel : '🔴 Payment Due Today',
    badgeColor  : '#E65100',
    badgeText   : 'DUE TODAY',
    message     : `The invoice for <strong>${o.customer}</strong> is due <strong>today</strong>. Please contact the client immediately and confirm payment receipt.`,
    orderInfo   : o,
  });
  safeSendEmail_(toEmail, subject, html);
}

function sendEscalationEmail_(hubSS, o, daysOverdue, config) {
  const subject = `🚨 OVERDUE ${daysOverdue} Days — ${o.customer} — PI #${o.piNo} — Escalation`;
  const html    = reminderEmailHtml_({
    headerColor : '#B71C1C',
    headerLabel : `🚨 Payment Overdue — ${daysOverdue} Day${daysOverdue !== 1 ? 's' : ''}`,
    badgeColor  : '#B71C1C',
    badgeText   : `OVERDUE ${daysOverdue} DAYS`,
    message     : `<strong>Escalation Notice:</strong> The invoice for <strong>${o.customer}</strong> was due <strong>${daysOverdue} day${daysOverdue !== 1 ? 's' : ''} ago</strong> and remains unpaid. Immediate action is required.`,
    orderInfo   : o,
  });

  // Recipients: assigned CRM + all Sales + all Admin
  const recipients = [
    ...new Set([
      o.crmEmail,
      ...config.salesEmails,
      ...config.adminEmails,
    ])
  ].filter(Boolean);

  recipients.forEach(email => safeSendEmail_(email, subject, html));
}

function sendNoDueDateAlert_(hubSS, o, salesEmails, daysSinceInv) {
  const subject = `⚡ No Due Date Set — PI #${o.piNo} — ${o.customer}`;
  const bodyHtml = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:620px;">
      <div style="background:#6A1B9A;padding:16px 24px;border-radius:6px 6px 0 0;">
        <span style="color:#fff;font-size:17px;font-weight:bold;">⚡ No Due Date Set for Invoice</span>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 6px 6px;">
        <p style="margin:0 0 12px;">Dear Sales Team,</p>
        <p>The following proforma invoice has <strong>no Due Date set</strong> in the Master Hub and the invoice date is <strong>${daysSinceInv} days ago</strong>. Please ensure the CRM owner enters a Due Date at the earliest.</p>
        <table style="border-collapse:collapse;font-size:13px;width:100%;margin:14px 0;">
          <tr style="background:#6A1B9A;color:#fff;">
            <th style="padding:8px 12px;text-align:left;" colspan="2">Invoice Without Due Date</th>
          </tr>
          <tr style="background:#F3E5F5;"><td style="padding:7px 12px;border:1px solid #ddd;">PI No</td><td style="padding:7px 12px;border:1px solid #ddd;"><strong>${o.piNo}</strong></td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">Client</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.customer}</td></tr>
          <tr style="background:#F3E5F5;"><td style="padding:7px 12px;border:1px solid #ddd;">Invoice No</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.invNo}</td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">Invoice Date</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.invDate}</td></tr>
          <tr style="background:#F3E5F5;"><td style="padding:7px 12px;border:1px solid #ddd;">Days Since Invoice</td><td style="padding:7px 12px;border:1px solid #ddd;color:#6A1B9A;font-weight:bold;">${daysSinceInv}</td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">Invoice Value</td><td style="padding:7px 12px;border:1px solid #ddd;">$${Number(o.invValue || 0).toLocaleString('en-US')}</td></tr>
          <tr style="background:#F3E5F5;"><td style="padding:7px 12px;border:1px solid #ddd;">CRM Owner</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.crmEmail}</td></tr>
        </table>
        <p>Please ask the assigned CRM to enter the Due Date in Master Hub → col I (Due Date) immediately.</p>
        <p style="color:#999;font-size:12px;margin-top:16px;">— UV Overseas CRM System (Auto-generated daily alert)</p>
      </div>
    </div>`;

  salesEmails.forEach(email => safeSendEmail_(email, subject, bodyHtml));
}

function reminderEmailHtml_(opts) {
  const o = opts.orderInfo;
  return `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:620px;">
      <div style="background:${opts.headerColor};padding:16px 24px;border-radius:6px 6px 0 0;">
        <span style="color:#fff;font-size:17px;font-weight:bold;">${opts.headerLabel}</span>
      </div>
      <div style="border:1px solid #ddd;border-top:none;padding:20px 24px;border-radius:0 0 6px 6px;">
        <p style="margin:0 0 12px;">Dear CRM Team,</p>
        <p>${opts.message}</p>
        <table style="border-collapse:collapse;font-size:13px;width:100%;margin:14px 0;">
          <tr style="background:${opts.headerColor};color:#fff;">
            <th style="padding:8px 12px;text-align:left;" colspan="2">
              Order Details &nbsp;
              <span style="background:#fff;color:${opts.badgeColor};font-size:11px;padding:2px 8px;border-radius:3px;font-weight:bold;">${opts.badgeText}</span>
            </th>
          </tr>
          <tr style="background:#FAFAFA;"><td style="padding:7px 12px;border:1px solid #ddd;">Client</td><td style="padding:7px 12px;border:1px solid #ddd;"><strong>${o.customer}</strong></td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">PI No</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.piNo}</td></tr>
          <tr style="background:#FAFAFA;"><td style="padding:7px 12px;border:1px solid #ddd;">Invoice No</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.invNo}</td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">Invoice Value</td><td style="padding:7px 12px;border:1px solid #ddd;">$${Number(o.invValue || 0).toLocaleString('en-US')}</td></tr>
          <tr style="background:#FAFAFA;"><td style="padding:7px 12px;border:1px solid #ddd;">Due Balance</td><td style="padding:7px 12px;border:1px solid #ddd;color:#C62828;font-weight:bold;">$${Number(o.balance || 0).toLocaleString('en-US')}</td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;">Payment Status</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.payStatus}</td></tr>
          <tr style="background:#FAFAFA;"><td style="padding:7px 12px;border:1px solid #ddd;">Invoice Date</td><td style="padding:7px 12px;border:1px solid #ddd;">${o.invDate}</td></tr>
          <tr><td style="padding:7px 12px;border:1px solid #ddd;"><strong>Due Date</strong></td><td style="padding:7px 12px;border:1px solid #ddd;font-weight:bold;color:${opts.headerColor};">${o.dueDate}</td></tr>
        </table>
        <p>Please update the payment status in your CRM sheet once payment is received.</p>
        <p style="color:#999;font-size:12px;margin-top:16px;">— UV Overseas CRM System (Auto-generated daily alert)</p>
      </div>
    </div>`;
}

function safeSendEmail_(toEmail, subject, htmlBody) {
  try {
    if (!toEmail) return;
    GmailApp.sendEmail(
      toEmail,
      subject,
      'Please view this email in an HTML-compatible client.',
      { htmlBody: htmlBody, name: 'UV Overseas CRM System' }
    );
  } catch (err) {
    Logger.log('safeSendEmail_: Failed to ' + toEmail + ' — ' + err.message);
  }
}

function menuSendOrderSummary_() {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const clientName  = activeSheet.getName();

    // Guard: don't run on system sheets
    const SYSTEM = new Set([
      MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME
    ]);
    if (SYSTEM.has(clientName)) {
      alertIfUi_('⚠️ Please select a CLIENT tab first (not a system sheet).');
      return;
    }

    // ── Get client email from Client Credit sheet col F ──────
    // Col F in Client Credit = CC_COL_REQ_INCREASE (6) per existing constants.
    // Per spec, the client contact email lives in col F of Client Credit.
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    let   clientEmail = '';

    if (creditSheet) {
      const creditData = creditSheet.getDataRange().getValues();
      for (let i = 1; i < creditData.length; i++) {
        const name = (creditData[i][CC_COL_CLIENT_NAME - 1] || '').toString().trim();
        if (name.toLowerCase() === clientName.toLowerCase()) {
          // Col F (index 5) — client email stored here per spec
          clientEmail = (creditData[i][5] || '').toString().trim();
          break;
        }
      }
    }

    // ── Read order rows from the active client sub-tab ───────
    const subData  = activeSheet.getDataRange().getValues();
    const subDisp  = activeSheet.getDataRange().getDisplayValues();
    if (subData.length <= 1) {
      alertIfUi_('No order data found in this client tab.');
      return;
    }

    // ── Also pull Hub data for this client (Due Date, status) ─
    const hubSheet  = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    const hubData   = hubSheet ? hubSheet.getDataRange().getValues()  : [];
    const hubDisp   = hubSheet ? hubSheet.getDataRange().getDisplayValues() : [];

    // Build a PI-No keyed map from Hub for quick lookup
    const hubByPi = {};
    for (let h = 1; h < hubData.length; h++) {
      const pi = (hubData[h][HUB_COL_PI_NO - 1] || '').toString().trim();
      if (pi) hubByPi[pi] = { data: hubData[h], disp: hubDisp[h] };
    }

    const tz      = Session.getScriptTimeZone();
    const today   = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy');
    const subject = `Payment Summary — ${clientName} — ${today}`;

    // ── Build order rows for the HTML table ───────────────────
    let tableRows  = '';
    let totalInv   = 0;
    let totalPaid  = 0;
    let totalBal   = 0;

    for (let r = 1; r < subData.length; r++) {
      const row     = subData[r];
      const piNo    = (row[SUB_COL_PI_NO   - 1] || '').toString().trim();
      const invNo   = (row[SUB_COL_INV_NO  - 1] || '').toString().trim();
      if (!piNo && !invNo) continue;

      const invDateD = row[SUB_COL_INV_DATE - 1];
      const invDate  = invDateD instanceof Date
        ? Utilities.formatDate(invDateD, tz, 'dd-MMM-yyyy') : (subDisp[r][SUB_COL_INV_DATE - 1] || '—');
      const awbNo    = (row[SUB_COL_AWB_NO  - 1] || '').toString().trim()  || '—';
      const dispDateD= row[SUB_COL_DISP_DATE - 1];
      const dispDate = dispDateD instanceof Date
        ? Utilities.formatDate(dispDateD, tz, 'dd-MMM-yyyy') : (subDisp[r][SUB_COL_DISP_DATE - 1] || '—');

      // Final Invoice from sub-tab col M (formula value)
      const finalInv = Number(activeSheet.getRange(r + 1, SUB_COL_FINAL_INV).getValue()) || 0;
      const payRec   = Number(row[SUB_COL_PAY_RECEIVED - 1]) || 0;
      const balance  = finalInv - payRec;

      // Enrich from Hub
      const hubRow  = hubByPi[piNo] || null;
      const dueDate = hubRow
        ? (() => { const d = hubRow.data[HUB_COL_DUE_DATE - 1];
                   return d instanceof Date ? Utilities.formatDate(d, tz, 'dd-MMM-yyyy') : (hubRow.disp[HUB_COL_DUE_DATE - 1] || '—'); })()
        : '—';
      const status  = hubRow ? (hubRow.data[HUB_COL_PAY_STATUS - 1] || '—').toString() : '—';
      const dispatch= hubRow ? (hubRow.data[HUB_COL_DISPATCH_STATUS - 1] || '—').toString() : '—';

      totalInv  += finalInv;
      totalPaid += payRec;
      totalBal  += balance;

      const bgColor = r % 2 === 0 ? '#F9F9F9' : '#FFFFFF';
      const balColor = balance > 0 ? '#C62828' : '#2E7D32';

      tableRows += `
        <tr style="background:${bgColor};">
          <td style="padding:7px 10px;border:1px solid #ddd;">${piNo  || '—'}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;">${invNo || '—'}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;">${invDate}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;text-align:right;">$${finalInv.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;text-align:right;">$${payRec.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;text-align:right;color:${balColor};font-weight:bold;">$${balance.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;">${dueDate}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;">${status}</td>
          <td style="padding:7px 10px;border:1px solid #ddd;font-size:12px;">${awbNo} / ${dispDate}<br><span style="color:#555;">${dispatch}</span></td>
        </tr>`;
    }

    // ── Total row ─────────────────────────────────────────────
    const totalRow = `
      <tr style="background:#1A237E;color:#fff;font-weight:bold;">
        <td style="padding:8px 10px;border:1px solid #0D47A1;" colspan="3">TOTAL</td>
        <td style="padding:8px 10px;border:1px solid #0D47A1;text-align:right;">$${totalInv.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
        <td style="padding:8px 10px;border:1px solid #0D47A1;text-align:right;">$${totalPaid.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
        <td style="padding:8px 10px;border:1px solid #0D47A1;text-align:right;">$${totalBal.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
        <td style="padding:8px 10px;border:1px solid #0D47A1;" colspan="3"></td>
      </tr>`;

    // ── Build full HTML body ───────────────────────────────────
    const htmlBody = `
      <div style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:900px;">
        <div style="background:#1A237E;padding:18px 24px;border-radius:6px 6px 0 0;">
          <span style="color:#fff;font-size:20px;font-weight:bold;">UV Overseas</span>
          <span style="color:#90CAF9;font-size:14px;margin-left:12px;">Payment Summary</span>
        </div>
        <div style="border:1px solid #ddd;border-top:none;padding:24px;border-radius:0 0 6px 6px;">
          <table style="width:100%;margin-bottom:12px;" cellpadding="0" cellspacing="0">
            <tr>
              <td><p style="margin:0;font-size:15px;">Dear <strong>${clientName}</strong>,</p>
                  <p style="margin:8px 0 0;">Please find below a summary of your outstanding and completed orders as of <strong>${today}</strong>.</p></td>
              <td style="text-align:right;vertical-align:top;">
                <div style="background:#E8F5E9;border:1px solid #A5D6A7;border-radius:4px;padding:10px 14px;display:inline-block;text-align:right;">
                  <div style="font-size:11px;color:#555;margin-bottom:2px;">Total Due Balance</div>
                  <div style="font-size:22px;font-weight:bold;color:${totalBal > 0 ? '#C62828' : '#2E7D32'};">$${totalBal.toLocaleString('en-US', {minimumFractionDigits:2})}</div>
                </div>
              </td>
            </tr>
          </table>

          <div style="overflow-x:auto;">
            <table style="border-collapse:collapse;font-size:13px;width:100%;min-width:820px;">
              <thead>
                <tr style="background:#1A237E;color:#fff;">
                  <th style="padding:9px 10px;text-align:left;">PI No</th>
                  <th style="padding:9px 10px;text-align:left;">Inv No</th>
                  <th style="padding:9px 10px;text-align:left;">Inv Date</th>
                  <th style="padding:9px 10px;text-align:right;">Invoice Value</th>
                  <th style="padding:9px 10px;text-align:right;">Paid</th>
                  <th style="padding:9px 10px;text-align:right;">Due Balance</th>
                  <th style="padding:9px 10px;text-align:left;">Due Date</th>
                  <th style="padding:9px 10px;text-align:left;">Status</th>
                  <th style="padding:9px 10px;text-align:left;">AWB / Dispatch</th>
                </tr>
              </thead>
              <tbody>${tableRows}</tbody>
              <tfoot>${totalRow}</tfoot>
            </table>
          </div>

          <p style="margin-top:20px;font-size:13px;color:#555;">
            Kindly ensure payment is made by the respective due dates. For any queries, please contact your account manager.
          </p>

          <div style="background:#F5F5F5;border-top:3px solid #1A237E;padding:14px 16px;margin-top:20px;border-radius:0 0 4px 4px;font-size:12px;color:#555;">
            <strong>UV Overseas</strong><br>
            Email: accounts@uvoverseas.com &nbsp;|&nbsp; WhatsApp: +91-XXXXXXXXXX<br>
            This is an automatically generated payment summary. Please do not reply to this email.
          </div>
        </div>
      </div>`;

    // ── Create Gmail DRAFT ─────────────────────────────────────
    const toAddress = clientEmail || Session.getActiveUser().getEmail();
    GmailApp.createDraft(
      toAddress,
      subject,
      'Please view this email in an HTML-compatible client.',
      { htmlBody: htmlBody, name: 'UV Overseas Accounts' }
    );

    const hubSS2 = SpreadsheetApp.openById(MASTER_HUB_ID);
    logActivity_(hubSS2, 'ORDER SUMMARY DRAFT CREATED',
      `Client:${clientName} | To:${toAddress} | Orders:${tableRows ? 'yes' : 'none'}`,
      Session.getActiveUser().getEmail());

    alertIfUi_(
      `✅ Gmail DRAFT created for ${clientName}!\n\n` +
      `To: ${toAddress}\n` +
      `Subject: ${subject}\n\n` +
      `Open Gmail → Drafts to review and send.`
    );

  } catch (err) {
    Logger.log('ERROR in menuSendOrderSummary_: ' + err.message);
    alertIfUi_('Error creating summary draft: ' + err.message);
  }
}

function checkDuplicateEntry_(sheet, row, col, value) {
  try {
    if (!value || value.toString().trim() === '') {
      return { isDuplicate: false, existingRow: null };
    }

    const trimmed  = value.toString().trim().toLowerCase();
    const lastRow  = sheet.getLastRow();
    if (lastRow < 2) return { isDuplicate: false, existingRow: null };

    const colData  = sheet.getRange(2, col, lastRow - 1, 1).getValues(); // skip header

    for (let i = 0; i < colData.length; i++) {
      const sheetRow = i + 2; // 1-based, offset by header
      if (sheetRow === row) continue; // skip the row being edited
      const cellVal  = (colData[i][0] || '').toString().trim().toLowerCase();
      if (cellVal === trimmed) {
        return { isDuplicate: true, existingRow: sheetRow };
      }
    }

    return { isDuplicate: false, existingRow: null };

  } catch (err) {
    Logger.log('ERROR in checkDuplicateEntry_: ' + err.message);
    return { isDuplicate: false, existingRow: null };
  }
}

function autoSetDueDate_(uuid, invDateValue) {
  try {
    if (!uuid || !invDateValue) return;

    const invDate = invDateValue instanceof Date
      ? new Date(invDateValue)
      : new Date(invDateValue);

    if (isNaN(invDate.getTime())) {
      Logger.log('autoSetDueDate_: invalid date for UUID=' + uuid);
      return;
    }

    // Compute soft threshold: Invoice Date + 30 days
    const softDue = new Date(invDate.getTime() + 30 * 24 * 60 * 60 * 1000);
    const isoStr  = softDue.toISOString();          // e.g. "2025-03-15T00:00:00.000Z"

    const propKey = PROP_SOFT_DUE_PREFIX + uuid;
    PropertiesService.getScriptProperties().setProperty(propKey, isoStr);

    Logger.log(
      'autoSetDueDate_: UUID=' + uuid +
      ' | InvDate=' + Utilities.formatDate(invDate, Session.getScriptTimeZone(), 'dd-MMM-yyyy') +
      ' | SoftDue=' + Utilities.formatDate(softDue, Session.getScriptTimeZone(), 'dd-MMM-yyyy') +
      ' stored (NOT written to sheet)'
    );

  } catch (err) {
    Logger.log('ERROR in autoSetDueDate_: ' + err.message);
  }
}

function getSoftDueDate_(uuid) {
  try {
    if (!uuid) return null;
    const val = PropertiesService.getScriptProperties()
      .getProperty(PROP_SOFT_DUE_PREFIX + uuid);
    if (!val) return null;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  } catch (_) {
    return null;
  }
}

// ============================================================
// setupEverything()
// ─────────────────────────────────────────────────────────────
// ONE-CLICK master setup for the Master Hub.
// Runs ALL setup in the correct order — call this ONCE from
// the Hub Apps Script editor after a fresh install or major
// upgrade.
//
// ORDER OF OPERATIONS:
//   1. Setup all triggers (Hub: onMasterHubEdit, onCreditRequestEdit,
//      onMembersSheetEdit, addCustomMenu)
//   2. Setup all Hub sheets + headers + protections:
//      Master Hub, Order Ledger, Client Summary, Client Dues,
//      Dashboard, Client Credit
//   3. Setup CRM Members sheet
//   4. Migrate / repair existing data:
//      • Repair Hub formulas on all existing rows
//      • Sync sub-tab data to Hub (if running on Hub script)
//   5. Repair ALL sheet protections (Hub + Credit + Members + Log)
//   6. Refresh Dashboard + Client Dues sheets
//   7. Show success summary popup with counts
//
// ⚠️  RUN THIS FROM THE HUB SPREADSHEET ONLY.
//     Run setupAllTriggers() separately from each CRM Spoke.
// ============================================================


// ══════════════════════════════════════════════════════════════════
// REMOTE WIPE SYSTEM
// Architecture:
//   1. License sheet writes a wipe request to a hidden "_WipeQueue"
//      sheet inside the Hub (License sheet has editor access to Hub).
//   2. Hub's remoteWipeCheck_() runs every 5 min via time trigger.
//      It reads pending requests, executes the wipe WITH full Hub
//      permissions (bypasses sheet protection), then marks done.
//   3. License sheet polls "_WipeQueue" for status and reports back.
//
// This keeps sheet protection intact — wipe always runs as Hub owner.
// ══════════════════════════════════════════════════════════════════

const WIPE_QUEUE_SHEET  = '_WipeQueue';   // hidden sheet in Hub
const WIPE_COL_EMAIL    = 1;  // A — target email
const WIPE_COL_TYPE     = 2;  // B — FULL or TARGETED
const WIPE_COL_NAME     = 3;  // C — client display name (targeted only)
const WIPE_COL_FILE_ID  = 4;  // D — CRM file ID (targeted only)
const WIPE_COL_STATUS   = 5;  // E — PENDING / DONE / ERROR
const WIPE_COL_RESULT   = 6;  // F — result summary
const WIPE_COL_TIME     = 7;  // G — timestamp

// ── Setup wipe queue sheet (run once from Hub) ────────────────────
function setupWipeQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let wipeSheet = ss.getSheetByName(WIPE_QUEUE_SHEET);
  if (!wipeSheet) {
    wipeSheet = ss.insertSheet(WIPE_QUEUE_SHEET);
    wipeSheet.hideSheet();
    wipeSheet.getRange(1, 1, 1, 7).setValues([[
      'target_email', 'wipe_type', 'client_name',
      'crm_file_id', 'status', 'result', 'timestamp'
    ]]);
    wipeSheet.setFrozenRows(1);
    Logger.log('setupWipeQueue: _WipeQueue sheet created and hidden.');
  } else {
    Logger.log('setupWipeQueue: _WipeQueue sheet already exists.');
  }
  // Install the polling trigger
  setupRemoteWipeTrigger_();
}

// ── Install time-based trigger for remote wipe polling ────────────
function setupRemoteWipeTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'remoteWipeCheck_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('remoteWipeCheck_')
    .timeBased().everyMinutes(5).create();
  Logger.log('remoteWipeCheck_ trigger installed — polls every 5 min.');
}

// ── Poll wipe queue and execute pending requests ──────────────────
// Runs every 5 min via time trigger in Hub script (Hub owner context)
function remoteWipeCheck_() {
  try {
    const ss         = SpreadsheetApp.openById(MASTER_HUB_ID);
    const wipeSheet  = ss.getSheetByName(WIPE_QUEUE_SHEET);
    if (!wipeSheet) return; // queue not set up yet

    const data = wipeSheet.getDataRange().getValues();
    let processed = 0;

    for (let i = 1; i < data.length; i++) {
      const status = (data[i][WIPE_COL_STATUS - 1] || '').toString().trim();
      if (status !== 'PENDING') continue;

      const targetEmail = (data[i][WIPE_COL_EMAIL   - 1] || '').toString().trim();
      const wipeType    = (data[i][WIPE_COL_TYPE    - 1] || '').toString().trim();
      const clientName  = (data[i][WIPE_COL_NAME    - 1] || '').toString().trim();
      const crmFileId   = (data[i][WIPE_COL_FILE_ID - 1] || '').toString().trim();

      Logger.log('remoteWipeCheck_: processing ' + wipeType + ' wipe for ' + targetEmail);

      // Mark as IN_PROGRESS immediately to prevent duplicate execution
      wipeSheet.getRange(i + 1, WIPE_COL_STATUS).setValue('IN_PROGRESS');
      SpreadsheetApp.flush();

      let result = '';
      try {
        if (wipeType === 'FULL') {
          result = executeFullWipe_(ss);
        } else if (wipeType === 'TARGETED') {
          result = executeTargetedWipe_(ss, clientName, crmFileId);
        } else {
          result = 'Unknown wipe type: ' + wipeType;
        }
        wipeSheet.getRange(i + 1, WIPE_COL_STATUS).setValue('DONE');
      } catch (wipeErr) {
        result = 'ERROR: ' + wipeErr.message;
        wipeSheet.getRange(i + 1, WIPE_COL_STATUS).setValue('ERROR');
      }

      wipeSheet.getRange(i + 1, WIPE_COL_RESULT).setValue(result);
      wipeSheet.getRange(i + 1, WIPE_COL_TIME).setValue(new Date());
      SpreadsheetApp.flush();
      processed++;
    }

    if (processed > 0) Logger.log('remoteWipeCheck_: processed ' + processed + ' request(s)');
  } catch (e) {
    Logger.log('remoteWipeCheck_ error: ' + e.message);
  }
}

// ── Full reset: wipe ALL data from Hub + ALL CRM spokes ───────────
function executeFullWipe_(hubSS) {
  const log    = [];
  const errors = [];

  // Get Team Members data BEFORE clearing anything
  const membSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
  const membData  = membSheet ? membSheet.getDataRange().getValues() : [];

  // Collect all CRM File IDs first
  const allFileIds = [];
  for (let i = 1; i < membData.length; i++) {
    const fid  = (membData[i][MC_COL_FILE_ID - 1] || '').toString().trim();
    const name = (membData[i][MC_COL_NAME    - 1] || '').toString().trim();
    if (fid) allFileIds.push({ fid, name });
  }

  // Wipe Hub sheets (skip Team Members and Activity Log headers)
  const skipSheets = new Set([LOG_SHEET_NAME, WIPE_QUEUE_SHEET]);
  hubSS.getSheets().forEach(s => {
    if (skipSheets.has(s.getName())) return;
    try {
      // Remove sheet protection temporarily
      const protections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      const rangeProtections = s.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      [...protections, ...rangeProtections].forEach(p => {
        try { if (p.canEdit()) p.remove(); } catch (_) {}
      });

      if (s.getLastRow() > 1) {
        s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
        log.push('Hub "' + s.getName() + '": cleared');
      }
    } catch (e) { errors.push('Hub "' + s.getName() + '": ' + e.message); }
  });

  // Wipe ALL CRM spokes
  allFileIds.forEach(({ fid, name }) => {
    try {
      const crmSS = SpreadsheetApp.openById(fid);
      crmSS.getSheets().forEach(s => {
        try {
          // Remove protection in CRM too
          const protections = s.getProtections(SpreadsheetApp.ProtectionType.SHEET);
          const rangeProtections = s.getProtections(SpreadsheetApp.ProtectionType.RANGE);
          [...protections, ...rangeProtections].forEach(p => {
            try { if (p.canEdit()) p.remove(); } catch (_) {}
          });

          if (s.getLastRow() > 1) {
            s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
          }
        } catch (_) {}
      });
      log.push('CRM "' + name + '": all sheets cleared');
    } catch (e) { errors.push('CRM "' + name + '" (' + fid + '): ' + e.message); }
  });

  // Log to Activity Log
  try {
    const logSheet = hubSS.getSheetByName(LOG_SHEET_NAME);
    if (logSheet) {
      logSheet.appendRow([new Date(), 'FULL DATA WIPE', 'Remote wipe executed via License sheet', 'System']);
    }
  } catch (_) {}

  return (errors.length === 0 ? 'SUCCESS' : 'PARTIAL') +
    '\nCompleted:\n' + log.map(l => '  - ' + l).join('\n') +
    (errors.length ? '\nErrors:\n' + errors.map(e => '  - ' + e).join('\n') : '');
}

// ── Targeted wipe: clear only one client's data ───────────────────
function executeTargetedWipe_(hubSS, clientName, crmFileId) {
  const log    = [];
  const errors = [];

  // Helper to clear rows matching client name in a sheet
  const clearByName = (sheet, colIdx) => {
    if (!sheet) return 0;
    // Remove protection first
    [...sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET),
     ...sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)].forEach(p => {
      try { if (p.canEdit()) p.remove(); } catch (_) {}
    });
    const data    = sheet.getDataRange().getValues();
    const nameLow = clientName.toLowerCase().trim();
    let   count   = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if ((data[i][colIdx - 1] || '').toString().toLowerCase().trim() === nameLow) {
        sheet.deleteRow(i + 1);
        count++;
      }
    }
    return count;
  };

  // Wipe Hub sheets by client name
  [
    { name: ORDER_LEDGER_SHEET_NAME,   col: OL_COL_CLIENT_NAME },
    { name: CLIENT_SUMMARY_SHEET_NAME, col: 1 },
    { name: CLIENT_DUES_SHEET_NAME,    col: 1 },
    { name: DASHBOARD_SHEET_NAME,      col: 1 },
  ].forEach(({ name, col }) => {
    try {
      const s = hubSS.getSheetByName(name);
      const n = clearByName(s, col);
      log.push('Hub "' + name + '": ' + n + ' row(s) cleared');
    } catch (e) { errors.push('Hub "' + name + '": ' + e.message); }
  });

  // Wipe Team Members row for this client
  try {
    const membSheet = hubSS.getSheetByName(MEMBERS_SHEET_NAME);
    if (membSheet) {
      [...membSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET),
       ...membSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)].forEach(p => {
        try { if (p.canEdit()) p.remove(); } catch (_) {}
      });
      const mData = membSheet.getDataRange().getValues();
      for (let i = mData.length - 1; i >= 1; i--) {
        if ((mData[i][MC_COL_NAME - 1] || '').toString().toLowerCase().trim()
            === clientName.toLowerCase().trim()) {
          membSheet.getRange(i + 1, 1, 1, membSheet.getLastColumn()).clearContent();
          log.push('Team Members: row cleared');
        }
      }
    }
  } catch (e) { errors.push('Team Members: ' + e.message); }

  // Wipe CRM spoke
  if (crmFileId) {
    try {
      const crmSS = SpreadsheetApp.openById(crmFileId);
      crmSS.getSheets().forEach(s => {
        try {
          [...s.getProtections(SpreadsheetApp.ProtectionType.SHEET),
           ...s.getProtections(SpreadsheetApp.ProtectionType.RANGE)].forEach(p => {
            try { if (p.canEdit()) p.remove(); } catch (_) {}
          });
          if (s.getLastRow() > 1) {
            s.getRange(2, 1, s.getLastRow() - 1, s.getLastColumn()).clearContent();
            log.push('CRM "' + s.getName() + '": cleared');
          }
        } catch (e) { errors.push('CRM "' + s.getName() + '": ' + e.message); }
      });
    } catch (e) { errors.push('CRM file: ' + e.message); }
  }

  // Log
  try {
    const logSheet = hubSS.getSheetByName(LOG_SHEET_NAME);
    if (logSheet) {
      logSheet.appendRow([new Date(), 'TARGETED DATA WIPE',
        'Client: ' + clientName + ' | CRM: ' + crmFileId, 'System']);
    }
  } catch (_) {}

  return (errors.length === 0 ? 'SUCCESS' : 'PARTIAL') +
    '\nCompleted:\n' + log.map(l => '  - ' + l).join('\n') +
    (errors.length ? '\nErrors:\n' + errors.map(e => '  - ' + e).join('\n') : '');
}

// ══════════════════════════════════════════════════════════════════
// FEATURE 1: AUTO-SCROLL TO LAST ROW ON OPEN
// Applies to Order Ledger CRM sheet only.
// When the file opens, cursor jumps to the last data row so user
// never needs to scroll down manually.
// ══════════════════════════════════════════════════════════════════

function scrollToLastRow_() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORDER_LEDGER_CRM_NAME);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // no data yet

    // Activate the sheet and scroll to last row
    sheet.activate();
    sheet.setActiveRange(sheet.getRange(lastRow, 1));

    // Freeze top so header stays visible while scrolled down
    if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);

    Logger.log('scrollToLastRow_: jumped to row ' + lastRow + ' in ' + ORDER_LEDGER_CRM_NAME);
  } catch (e) {
    Logger.log('scrollToLastRow_ error: ' + e.message);
  }
}

// ══════════════════════════════════════════════════════════════════
// FEATURE 2: FINANCIAL YEAR ARCHIVE
//
// India FY: Apr 1 – Mar 31
// Archive sheet naming: "Archive FY2024-25" (covers Apr 2024–Mar 2025)
//
// What it does:
//   1. Identifies rows in Order Ledger CRM older than 1 year
//      (based on Inv Date column)
//   2. Copies them to "Archive FY20XX-XX" sheet (creates if needed)
//   3. Deletes the archived rows from Order Ledger CRM
//   4. Also archives matching rows from client sub-tabs
//   5. Logs the archive run to Activity Log
//
// Auto trigger: April 1st each year
// Manual: CRM Tools → Archive → Archive data now
// ══════════════════════════════════════════════════════════════════

// ── Determine current India FY label ─────────────────────────────
// e.g. if today is May 2026 → FY being archived = 2024-25
//      if today is Feb 2026 → FY being archived = 2023-24
function getArchiveFYLabel_(forDate) {
  const d    = forDate || new Date();
  const year = d.getFullYear();
  const month = d.getMonth() + 1; // 1-based

  // Rows to archive = rows from the FY that ended last March 31
  // If current month >= April → last FY ended this March
  // If current month < April  → last FY ended previous March
  let fyStartYear;
  if (month >= 4) {
    fyStartYear = year - 1; // FY that ended March of THIS year
  } else {
    fyStartYear = year - 2; // FY that ended March of PREVIOUS year
  }
  const fyEndYear = fyStartYear + 1;
  return 'Archive FY' + fyStartYear + '-' + String(fyEndYear).slice(-2);
}

// ── Date cutoff: rows with Inv Date before this date get archived ─
function getArchiveCutoffDate_() {
  const now   = new Date();
  const month = now.getMonth() + 1;
  const year  = now.getFullYear();

  // Archive rows older than April 1st of last FY start
  // i.e. keep current FY + previous FY, archive anything before that
  let cutoffYear;
  if (month >= 4) {
    cutoffYear = year - 1; // keep from Apr 1 last year onwards
  } else {
    cutoffYear = year - 2;
  }
  return new Date(cutoffYear, 3, 1); // April 1st (month index 3)
}

// ── Manual menu entry ─────────────────────────────────────────────
function menuArchiveNow_() {
  const ui = SpreadsheetApp.getUi();
  const cutoff  = getArchiveCutoffDate_();
  const fyLabel = getArchiveFYLabel_();

  const confirm = ui.alert(
    'Archive Old Data',
    'This will move orders with Inv Date before ' +
    Utilities.formatDate(cutoff, Session.getScriptTimeZone(), 'dd MMM yyyy') +
    ' to sheet: "' + fyLabel + '"\n\n' +
    'These rows will be REMOVED from Order Ledger CRM and client sub-tabs.\n' +
    'The archive sheet will be created if it does not exist.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const result = runArchive_();
  ui.alert(
    result.success ? 'Archive Complete' : 'Archive Completed with Errors',
    result.summary,
    ui.ButtonSet.OK
  );
}

// ── Auto trigger entry point (April 1st) ─────────────────────────
function autoArchiveOnApril1_() {
  const now   = new Date();
  const month = now.getMonth() + 1;
  const day   = now.getDate();
  // Only run on April 1st (trigger fires daily, check date)
  if (month !== 4 || day !== 1) return;
  Logger.log('autoArchiveOnApril1_: April 1st detected — running archive');
  runArchive_();
}

// ── Setup auto archive trigger (run once) ─────────────────────────
function setupArchiveTrigger_() {
  // Remove existing archive triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoArchiveOnApril1_') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Daily trigger — function checks if it's April 1st
  ScriptApp.newTrigger('autoArchiveOnApril1_')
    .timeBased().everyDays(1).atHour(1).create(); // runs at 1 AM daily
  try {
    SpreadsheetApp.getUi().alert(
      'Auto Archive Trigger Set!\n\n' +
      'Archive will run automatically every April 1st at 1 AM.\n' +
      'You can also run it manually anytime from CRM Tools → Archive.'
    );
  } catch (_) {}
  Logger.log('setupArchiveTrigger_: daily trigger installed.');
}

// ── Core archive function ─────────────────────────────────────────
function runArchive_() {
  const log    = [];
  const errors = [];
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const cutoffDate  = getArchiveCutoffDate_();
    const fyLabel     = getArchiveFYLabel_();
    const olSheet     = ss.getSheetByName(ORDER_LEDGER_CRM_NAME);

    if (!olSheet) throw new Error('Order Ledger CRM sheet not found.');

    log.push('Archive label: ' + fyLabel);
    log.push('Cutoff date: ' + Utilities.formatDate(
      cutoffDate, Session.getScriptTimeZone(), 'dd MMM yyyy'));

    // ── Get or create archive sheet ───────────────────────────────
    let archiveSheet = ss.getSheetByName(fyLabel);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(fyLabel);
      // Copy header row from Order Ledger CRM
      const header = olSheet.getRange(1, 1, 1, olSheet.getLastColumn()).getValues();
      archiveSheet.getRange(1, 1, 1, header[0].length).setValues(header);
      archiveSheet.setFrozenRows(1);
      // Style header same as Order Ledger CRM
      archiveSheet.getRange(1, 1, 1, header[0].length)
        .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
      log.push('Created archive sheet: ' + fyLabel);
    }

    // ── Scan Order Ledger CRM for old rows ────────────────────────
    const olData    = olSheet.getDataRange().getValues();
    const rowsToArchive = []; // {rowIndex, rowData, piNo}
    const piNosArchived = new Set();

    for (let i = 1; i < olData.length; i++) {
      const invDateRaw = olData[i][OL_COL_INV_DATE - 1];
      if (!invDateRaw) continue; // skip rows with no Inv Date

      const invDate = invDateRaw instanceof Date
        ? invDateRaw
        : new Date(invDateRaw);

      if (isNaN(invDate.getTime())) continue; // skip unparseable dates

      if (invDate < cutoffDate) {
        rowsToArchive.push({
          rowIndex: i + 1, // 1-based sheet row
          rowData:  olData[i],
          piNo:     (olData[i][OL_COL_PI_NO - 1] || '').toString().trim()
        });
        if (olData[i][OL_COL_PI_NO - 1]) {
          piNosArchived.add(olData[i][OL_COL_PI_NO - 1].toString().trim());
        }
      }
    }

    if (rowsToArchive.length === 0) {
      return {
        success: true,
        summary: 'No rows older than cutoff date found.\nNothing to archive.'
      };
    }

    log.push('Rows to archive from Order Ledger CRM: ' + rowsToArchive.length);

    // ── Copy rows to archive sheet ────────────────────────────────
    const archiveData = rowsToArchive.map(r => r.rowData);
    const archiveLastRow = archiveSheet.getLastRow();
    archiveSheet
      .getRange(archiveLastRow + 1, 1, archiveData.length, archiveData[0].length)
      .setValues(archiveData);
    log.push('Copied ' + archiveData.length + ' rows to ' + fyLabel);

    // ── Delete archived rows from Order Ledger CRM (bottom-up) ───
    // Sort descending by row index so deletions don't shift indices
    rowsToArchive.sort((a, b) => b.rowIndex - a.rowIndex);
    rowsToArchive.forEach(r => {
      try { olSheet.deleteRow(r.rowIndex); } catch (e) {
        errors.push('Delete row ' + r.rowIndex + ': ' + e.message);
      }
    });
    log.push('Deleted ' + rowsToArchive.length + ' rows from Order Ledger CRM');

    // ── Archive matching rows from client sub-tabs ────────────────
    const systemSheets = new Set([
      ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
      CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME, fyLabel
    ]);

    ss.getSheets().forEach(clientSheet => {
      const sheetName = clientSheet.getName();
      // Skip system sheets and existing archive sheets
      if (systemSheets.has(sheetName)) return;
      if (sheetName.startsWith('Archive FY')) return;

      try {
        const clientData = clientSheet.getDataRange().getValues();
        const clientRowsToArchive = [];

        for (let i = 1; i < clientData.length; i++) {
          const piNo = (clientData[i][SUB_COL_PI_NO - 1] || '').toString().trim();
          if (piNo && piNosArchived.has(piNo)) {
            clientRowsToArchive.push(i + 1); // 1-based
          }
        }

        if (clientRowsToArchive.length > 0) {
          // Delete bottom-up
          clientRowsToArchive.sort((a, b) => b - a);
          clientRowsToArchive.forEach(rowIdx => {
            try { clientSheet.deleteRow(rowIdx); } catch (_) {}
          });
          log.push('Client tab "' + sheetName + '": removed ' +
                   clientRowsToArchive.length + ' archived row(s)');
        }
      } catch (e) {
        errors.push('Client tab "' + sheetName + '": ' + e.message);
      }
    });

    // ── Log to Activity Log ───────────────────────────────────────
    try {
      const logSheet = ss.getSheetByName(LOG_SHEET_NAME);
      if (logSheet) {
        logSheet.appendRow([
          new Date(), 'ARCHIVE',
          fyLabel + ' — ' + rowsToArchive.length + ' orders archived',
          Session.getActiveUser().getEmail() || 'System'
        ]);
      }
    } catch (_) {}

    // ── Scroll to last row after archive ─────────────────────────
    try { scrollToLastRow_(); } catch (_) {}

    // ── Archive Hub Order Ledger (runs from CRM spoke, targets Hub) ──
    // Hub accumulates rows from ALL CRM spokes so needs its own archive.
    // Only rows whose PI No was archived from this CRM get moved in Hub too.
    try {
      if (piNosArchived.size > 0) {
        const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
        const hubOL    = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
        const hubCS    = hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME);
        const hubDues  = hubSS.getSheetByName(CLIENT_DUES_SHEET_NAME);

        // Get or create archive sheet in Hub
        let hubArchive = hubSS.getSheetByName(fyLabel);
        if (!hubArchive) {
          hubArchive = hubSS.insertSheet(fyLabel);
          if (hubOL) {
            const hubHeader = hubOL.getRange(1, 1, 1, hubOL.getLastColumn()).getValues();
            hubArchive.getRange(1, 1, 1, hubHeader[0].length).setValues(hubHeader);
            hubArchive.setFrozenRows(1);
            hubArchive.getRange(1, 1, 1, hubHeader[0].length)
              .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold');
          }
          log.push('Hub: created archive sheet ' + fyLabel);
        }

        // Archive Hub Order Ledger rows matching archived PI Nos
        if (hubOL) {
          const hubData      = hubOL.getDataRange().getValues();
          const hubRowsToArc = [];
          const hubArcData   = [];
          for (let i = 1; i < hubData.length; i++) {
            const hPiNo = (hubData[i][OL_COL_PI_NO - 1] || '').toString().trim();
            if (piNosArchived.has(hPiNo)) {
              hubRowsToArc.push(i + 1);
              hubArcData.push(hubData[i]);
            }
          }
          if (hubArcData.length > 0) {
            const hubArcLastRow = hubArchive.getLastRow();
            hubArchive.getRange(hubArcLastRow + 1, 1, hubArcData.length, hubArcData[0].length)
              .setValues(hubArcData);
            hubRowsToArc.sort((a, b) => b - a);
            hubRowsToArc.forEach(r => { try { hubOL.deleteRow(r); } catch (_) {} });
            log.push('Hub Order Ledger: archived ' + hubArcData.length + ' row(s)');
          }
        }

        // Remove archived client from Hub Client Summary if all their orders are gone
        if (hubCS) {
          const csData     = hubCS.getDataRange().getValues();
          const clientName = ss.getName(); // CRM spoke name = client name
          for (let i = csData.length - 1; i >= 1; i--) {
            const csClient = (csData[i][0] || '').toString().trim().toLowerCase();
            if (csClient === clientName.toLowerCase()) {
              // Only remove if ALL their orders are now archived (no remaining rows in Hub OL)
              const hubOLFresh = hubOL ? hubOL.getDataRange().getValues() : [];
              const hasRemaining = hubOLFresh.slice(1).some(r =>
                (r[OL_COL_CLIENT_NAME - 1] || '').toString().trim().toLowerCase()
                === clientName.toLowerCase()
              );
              if (!hasRemaining) {
                hubCS.deleteRow(i + 1);
                log.push('Hub Client Summary: removed ' + clientName + ' (no remaining orders)');
              }
            }
          }
        }

        // Log to Hub Activity Log
        try {
          const hubLog = hubSS.getSheetByName(LOG_SHEET_NAME);
          if (hubLog) {
            hubLog.appendRow([
              new Date(), 'ARCHIVE',
              fyLabel + ' — ' + piNosArchived.size + ' PI Nos from ' + ss.getName(),
              Session.getActiveUser().getEmail() || 'System'
            ]);
          }
        } catch (_) {}
      }
    } catch (hubArchErr) {
      errors.push('Hub archive: ' + hubArchErr.message);
    }

  } catch (fatal) {
    errors.push('FATAL: ' + fatal.message);
  }

  const success = errors.length === 0;
  const summary =
    '\nCompleted:\n' + log.map(l => '  - ' + l).join('\n') +
    (errors.length ? '\n\nErrors:\n' + errors.map(e => '  - ' + e).join('\n') : '');

  Logger.log('runArchive_: ' + summary);
  return { success, summary };
}

// ══════════════════════════════════════════════════════════════════
// FIX 4: ASYNC HUB SYNC QUEUE
// Problem: Every edit triggers an immediate sync to Hub (network call
// to another spreadsheet). With 5+ CRM spokes this causes lock
// conflicts and slows edits to 8-25 seconds.
//
// Solution: Writes are queued in a local "_SyncQueue" sheet in the
// CRM file. A 1-minute time trigger processes the queue in batches,
// syncing to Hub asynchronously. Edits feel instant to the user.
//
// Falls back to direct sync if queue sheet doesn't exist (backwards
// compatible — no setup required to use the CRM).
// ══════════════════════════════════════════════════════════════════

const SYNC_QUEUE_SHEET   = '_SyncQueue';
const SQ_COL_UUID        = 1;  // A
const SQ_COL_CLIENT      = 2;  // B
const SQ_COL_PI_NO       = 3;  // C
const SQ_COL_STATUS      = 4;  // D — PENDING / DONE / ERROR
const SQ_COL_TIMESTAMP   = 5;  // E
const SQ_COL_RETRY       = 6;  // F — retry count

// ── Setup sync queue sheet (run once per CRM) ─────────────────────
function setupSyncQueue_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sq = ss.getSheetByName(SYNC_QUEUE_SHEET);
  if (!sq) {
    sq = ss.insertSheet(SYNC_QUEUE_SHEET);
    sq.hideSheet();
    sq.getRange(1, 1, 1, 6).setValues([[
      'uuid', 'client_name', 'pi_no', 'status', 'timestamp', 'retry_count'
    ]]);
    sq.setFrozenRows(1);
    Logger.log('setupSyncQueue_: _SyncQueue sheet created.');
  }
  // Install 1-minute batch sync trigger
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processSyncQueue_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processSyncQueue_').timeBased().everyMinutes(1).create();
  Logger.log('setupSyncQueue_: 1-min processSyncQueue_ trigger installed.');
}

// ── Add a row to the sync queue ───────────────────────────────────
function enqueueSyncRow_(uuid, clientName, piNo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sq = ss.getSheetByName(SYNC_QUEUE_SHEET);
    if (!sq) return false; // queue not set up — caller falls back to direct sync

    // Check if already queued (avoid duplicate entries)
    const data = sq.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const qUuid   = (data[i][SQ_COL_UUID   - 1] || '').toString().trim();
      const qStatus = (data[i][SQ_COL_STATUS - 1] || '').toString().trim();
      if (qUuid === uuid && qStatus === 'PENDING') {
        Logger.log('enqueueSyncRow_: already queued UUID ' + uuid + ' — skipping');
        return true;
      }
    }
    sq.appendRow([uuid, clientName, piNo, 'PENDING', new Date(), 0]);
    Logger.log('enqueueSyncRow_: queued UUID ' + uuid + ' client=' + clientName);
    return true;
  } catch (e) {
    Logger.log('enqueueSyncRow_ error: ' + e.message);
    return false;
  }
}

// ── Process sync queue — runs every 1 minute via time trigger ─────
function processSyncQueue_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sq = ss.getSheetByName(SYNC_QUEUE_SHEET);
  if (!sq) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('processSyncQueue_: could not get lock — will retry next minute');
    return;
  }

  try {
    const data = sq.getDataRange().getValues();
    let processed = 0;

    for (let i = 1; i < data.length; i++) {
      const status     = (data[i][SQ_COL_STATUS  - 1] || '').toString().trim();
      const retryCount = Number(data[i][SQ_COL_RETRY - 1]) || 0;
      if (status !== 'PENDING') continue;
      if (retryCount >= 3) {
        // Too many retries — mark failed
        sq.getRange(i + 1, SQ_COL_STATUS).setValue('FAILED');
        continue;
      }

      const uuid       = (data[i][SQ_COL_UUID   - 1] || '').toString().trim();
      const clientName = (data[i][SQ_COL_CLIENT  - 1] || '').toString().trim();
      const piNo       = (data[i][SQ_COL_PI_NO   - 1] || '').toString().trim();

      sq.getRange(i + 1, SQ_COL_STATUS).setValue('IN_PROGRESS');
      SpreadsheetApp.flush();

      try {
        // Find the row in Order Ledger CRM and push to Hub
        const mainSheet = ss.getSheetByName(ORDER_LEDGER_CRM_NAME);
        if (!mainSheet) throw new Error('Order Ledger CRM not found');

        const mainData = mainSheet.getDataRange().getValues();
        let targetRow  = -1;
        for (let j = 1; j < mainData.length; j++) {
          const rowUuid = (mainData[j][CRM_COL_UUID - 1] || '').toString().trim();
          if (rowUuid === uuid) { targetRow = j + 1; break; }
        }

        if (targetRow === -1) {
          Logger.log('processSyncQueue_: UUID ' + uuid + ' not found in OL CRM — marking done');
          sq.getRange(i + 1, SQ_COL_STATUS).setValue('DONE');
        } else {
          // Push full row to Hub
          const userEmail = (mainData[targetRow - 1][CRM_COL_CRM_OWNER - 1] || '').toString().trim();
          pushFullRowToHub_(ss, mainSheet, targetRow, mainData[targetRow - 1], userEmail);
          sq.getRange(i + 1, SQ_COL_STATUS).setValue('DONE');
          Logger.log('processSyncQueue_: synced UUID ' + uuid + ' client=' + clientName);
          processed++;
        }
      } catch (rowErr) {
        Logger.log('processSyncQueue_: error on UUID ' + uuid + ': ' + rowErr.message);
        sq.getRange(i + 1, SQ_COL_STATUS).setValue('PENDING');
        sq.getRange(i + 1, SQ_COL_RETRY).setValue(retryCount + 1);
      }

      // Limit to 10 rows per run to stay within execution time
      if (processed >= 10) break;
    }

    // Clean up old DONE/FAILED rows older than 7 days
    const cutoff = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
    const freshData = sq.getDataRange().getValues();
    for (let i = freshData.length - 1; i >= 1; i--) {
      const rowStatus = (freshData[i][SQ_COL_STATUS - 1] || '').toString().trim();
      const rowTime   = freshData[i][SQ_COL_TIMESTAMP - 1];
      if ((rowStatus === 'DONE' || rowStatus === 'FAILED') &&
          rowTime instanceof Date && rowTime < cutoff) {
        sq.deleteRow(i + 1);
      }
    }

    if (processed > 0) Logger.log('processSyncQueue_: processed ' + processed + ' row(s)');
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════════
// FIX 5: READ OPTIMIZATION — Sheet Data Cache
// Problem: 96 full getDataRange().getValues() calls — every edit
// re-reads entire sheets even when nothing changed.
//
// Solution: In-memory cache per execution. Within a single trigger
// execution, sheet data is read once and reused. Cache is scoped
// to the execution (not persisted) so always fresh per edit.
// ══════════════════════════════════════════════════════════════════

const _sheetDataCache_ = {}; // execution-scoped in-memory cache

// ── Cached sheet read — use instead of sheet.getDataRange().getValues() ──
// key = unique cache key (e.g. 'hub_order_ledger', 'crm_main')
// Returns values array. Pass forceRefresh=true to bypass cache.
function getCachedSheetData_(sheet, key, forceRefresh) {
  if (!forceRefresh && _sheetDataCache_[key]) {
    return _sheetDataCache_[key];
  }
  const data = sheet.getDataRange().getValues();
  _sheetDataCache_[key] = data;
  return data;
}

// ── Invalidate a specific cache key (call after writes) ──────────
function invalidateSheetCache_(key) {
  delete _sheetDataCache_[key];
}

// ── Invalidate all sheet caches (call at start of new trigger) ───
function invalidateAllSheetCaches_() {
  Object.keys(_sheetDataCache_).forEach(k => delete _sheetDataCache_[k]);
}

// ══════════════════════════════════════════════════════════════════
// ONE-TIME FIX: Remove date formatting from Hub Order Ledger col I
// (Dev Charges). The old PATH 2 bug wrote Due Date to col I and
// applied dd/mm/yyyy formatting. Now values like 110 (Dev Charges)
// display as dates (19/04/1900). Fix: reset col I to number format.
// The old code used HUB_COL_DUE_DATE=9 (col I) instead of
// OL_COL_DUE_DATE=19 (col S) — writing dates into Dev Charges col.
//
// This function:
//   1. Clears any Date values from col I (Dev Charges) in Hub OL
//   2. Clears col I in CRM Order Ledger CRM (Due Date was also
//      written there incorrectly as CRM_COL_DUE_DATE=9)
//   3. Re-triggers full sync for each affected row to restore
//      correct Dev Charges values from sub-tabs
// Run ONCE from Apps Script editor after pasting the fixed code.
// ══════════════════════════════════════════════════════════════════
function fixCorruptedDevChargesCol() {
  const log    = [];
  const errors = [];

  try {
    const hubSS = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubOL = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubOL) throw new Error('Order Ledger sheet not found in Hub.');

    const lastRow = hubOL.getLastRow();
    if (lastRow < 2) {
      log.push('No data rows found — nothing to fix.');
    } else {
      // ── Step 1: Reset col I (Dev Charges) number format ──────
      // Old PATH 2 bug applied dd/mm/yyyy date format to col I.
      // This causes numeric values like 110 to display as 19/04/1900.
      // Fix: reset to currency number format.
      hubOL.getRange(2, OL_COL_DEV_CHARGES, lastRow - 1, 1)
           .setNumberFormat('#,##0.00');
      log.push('Hub Order Ledger col I: reset to number format (#,##0.00)');

      // ── Step 2: Also reset col S (Due Date) to date format ───
      // Ensure col S has correct date formatting after our fix.
      hubOL.getRange(2, OL_COL_DUE_DATE, lastRow - 1, 1)
           .setNumberFormat('dd/mm/yyyy');
      log.push('Hub Order Ledger col S: reset to date format (dd/mm/yyyy)');

      // ── Step 3: Clear any Date objects in col I ───────────────
      // If actual Date objects were stored (not just formatting issue),
      // clear those cells so they can be re-populated by sync.
      const hubData = hubOL.getRange(2, OL_COL_DEV_CHARGES, lastRow - 1, 1).getValues();
      let clearedCount = 0;
      for (let i = 0; i < hubData.length; i++) {
        if (hubData[i][0] instanceof Date) {
          hubOL.getRange(i + 2, OL_COL_DEV_CHARGES).clearContent();
          clearedCount++;
        }
      }
      if (clearedCount > 0) {
        log.push('Cleared ' + clearedCount + ' actual Date objects from col I');
      }

      SpreadsheetApp.flush();
    }

    // ── Step 4: Log to Activity Log ──────────────────────────
    try {
      const logSheet = hubSS.getSheetByName(LOG_SHEET_NAME);
      if (logSheet) {
        logSheet.appendRow([
          new Date(), 'FORMAT FIX',
          'Reset Hub Order Ledger col I (Dev Charges) number format',
          Session.getActiveUser().getEmail() || 'System'
        ]);
      }
    } catch (_) {}

  } catch (fatal) {
    errors.push('FATAL: ' + fatal.message);
  }

  const summary =
    'Steps completed:\n' + log.map(l => '  - ' + l).join('\n') +
    (errors.length ? '\n\nErrors:\n' + errors.map(e => '  - ' + e).join('\n') : '') +
    '\n\nNext step: Re-enter Dev Charges in each client sub-tab\n' +
    'to push correct values to Hub (or wait for next full-row sync).';

  Logger.log('fixCorruptedDevChargesCol:\n' + summary);

  try {
    SpreadsheetApp.getUi().alert(
      errors.length === 0 ? '✅ Format Fix Complete' : '⚠️ Fix Completed with Errors',
      summary,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (_) {}
}
function runReloadConfigManually() {
  menuReloadConfig_();
}