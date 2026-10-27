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
const MEMBERS_SHEET_NAME  = 'CRM Members';

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
const CREDIT_SHEET_NAME           = 'Client Credit';
const LOG_SHEET_NAME              = 'Activity Log';

// ── Sheet Name Constants — CRM Spoke ─────────────────────────
const MAIN_SHEET_NAME             = 'Main';               // legacy name kept for sync logic compatibility
const CRM_ORDER_LEDGER_SHEET      = 'Order Ledger CRM';   // NEW v14.0 — replaces 'Main' display name
const CRM_DUES_SHEET              = 'Client Dues';        // NEW v14.0 — CRM-filtered dues view

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
const HUB_TOTAL_COLS          = 18;

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
const HUB_ALLOWED_PUSH_COLS = [
  ...HUB_ACCOUNTS_COLS, ...HUB_LOGISTICS_COLS,
  HUB_COL_APPROVED_BY, HUB_COL_APPROVAL_REQ, HUB_COL_DISPATCH_STATUS
];

const HUB_TO_CRM_MAIN = {
  [HUB_COL_PAY_RECEIVED]:    CRM_COL_PAY_RECEIVED,
  [HUB_COL_PAY_DATE]:        CRM_COL_PAY_DATE,
  [HUB_COL_APPROVED_BY]:     CRM_COL_APPROVED_BY,
  [HUB_COL_APPROVAL_REQ]:    CRM_COL_APPROVAL_REQ,
  [HUB_COL_DISPATCH_STATUS]: CRM_COL_DISPATCH_STATUS,
};

const SUB_TO_HUB = {
  [SUB_COL_AWB_NO]:    HUB_COL_AWB_NO,
  [SUB_COL_DISP_DATE]: HUB_COL_DISP_DATE,
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


// ============================================================
// setupMembersSheet()  — v14.0: dropdown updated (Sales Manager removed)
// ============================================================
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── 1. Master Hub sheet (core — unchanged) ────────────────
    let mainSheet = ss.getSheetByName(HUB_MAIN_SHEET_NAME);
    if (!mainSheet) mainSheet = ss.insertSheet(HUB_MAIN_SHEET_NAME);

    const hubHeaders = [
      'UUID', 'Customer Name', 'CRM Owner (Gmail)',
      'PI No', 'Inv No', 'Inv Date', 'Inv Value',
      'Payment Status', 'Due Date', 'Payment Received', 'Payment Date',
      'Dispatch Status', 'Due Balance',
      'Approval Required', 'Approved By', 'Remarks',
      'AWB No', 'Dispatched Date'
    ];
    mainSheet.getRange(1, 1, 1, hubHeaders.length).setValues([hubHeaders]);
    mainSheet.getRange(1, 1, 1, hubHeaders.length)
      .setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF').setFontSize(11);
    mainSheet.hideColumns(HUB_COL_UUID);
    mainSheet.setFrozenRows(1);

    const hubWidths = {
      [HUB_COL_CUSTOMER]: 180,        [HUB_COL_CRM_OWNER]: 200,
      [HUB_COL_PI_NO]: 100,           [HUB_COL_INV_NO]: 80,
      [HUB_COL_INV_VALUE]: 100,       [HUB_COL_PAY_STATUS]: 120,
      [HUB_COL_DISPATCH_STATUS]: 130, [HUB_COL_DUE_BALANCE]: 110,
      [HUB_COL_APPROVAL_REQ]: 150,    [HUB_COL_AWB_NO]: 120,
      [HUB_COL_DISP_DATE]: 130
    };
    Object.entries(hubWidths).forEach(([c, w]) => mainSheet.setColumnWidth(Number(c), w));

    mainSheet.getRange(2, HUB_COL_APPROVAL_REQ, 1000, 1)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(['Yes', 'No', 'Approved'], true).build()
      );

    // ── Protections — Hub Master Hub sheet ────────────────────
    mainSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

    const ownerEmail       = Session.getEffectiveUser().getEmail();
    const config           = getCrmConfig_();
    const adminEditors     = [ownerEmail, ...config.adminEmails].filter((v, i, a) => a.indexOf(v) === i);
    const logisticsEditors = [ownerEmail, ...config.logisticsEmails].filter((v, i, a) => a.indexOf(v) === i);
    const accountsEditors  = [ownerEmail, ...config.accountsEmails ].filter((v, i, a) => a.indexOf(v) === i);
    // v14.0: approvalEditors = Admin + Sales (merged role)
    const approvalEditors  = [...new Set([
      ownerEmail, ...config.adminEmails, ...config.salesEmails
    ])];

    function applyProt_(range, desc, extraEditors) {
      const p = range.protect().setDescription(desc);
      p.removeEditors(p.getEditors());
      p.addEditors([...new Set([...adminEditors, ...(extraEditors || [])])]);
      p.setWarningOnly(false);
    }

    applyProt_(mainSheet.getRange(2, 2, 1000, 8),  'Admin only — core order columns B–I');
    applyProt_(mainSheet.getRange(2, HUB_COL_PAY_RECEIVED, 1000, 2),
               'Accounts only — Payment Received & Payment Date (J–K)', accountsEditors);
    applyProt_(mainSheet.getRange(2, 12, 1000, 2), 'Admin only — formula columns L–M');
    applyProt_(mainSheet.getRange(2, HUB_COL_APPROVAL_REQ, 1000, 3),
               'Sales — Approval Required, Approved By, Remarks (N–P)', approvalEditors);
    applyProt_(mainSheet.getRange(2, HUB_COL_AWB_NO, 1000, 2),
               'Logistics only — AWB No & Dispatched Date (Q–R)', logisticsEditors);

    const hubCfRange = mainSheet.getRange(2, 1, 1000, hubHeaders.length);
    mainSheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(I2<TODAY(),H2<>"Received")`)
        .setBackground('#FFCDD2').setRanges([hubCfRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=N2="Yes"`)
        .setBackground('#FFE0B2').setRanges([hubCfRange]).build(),
    ]);

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
function setupHubOrderLedgerSheet_(ss, adminEditors, logisticsEditors, accountsEditors) {
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
    'Due Balance'       // T  OL_COL_DUE_BALANCE
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
  sheet.getRange(2, OL_COL_PAY_STATUS, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Partial', 'Received'], true)
      .setAllowInvalid(false).build()
  );

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

  // Conditional formatting
  const cfRange = sheet.getRange(2, 1, 1000, headers.length);
  sheet.setConditionalFormatRules([
    // Overdue & unpaid — red
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(S2<TODAY(),R2<>"Received")`)
      .setBackground('#FFCDD2').setRanges([cfRange]).build(),
    // Partial payment — orange
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=R2="Partial"`)
      .setBackground('#FFE0B2').setRanges([cfRange]).build(),
    // Fully received — green
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=R2="Received"`)
      .setBackground('#E8F5E9').setRanges([cfRange]).build(),
  ]);

  // Protections: Order Ledger is admin-managed (sync-populated)
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());

  // Accounts can edit Payment cols (P–Q)
  const payProt = sheet.getRange(2, OL_COL_PAY_RECEIVED, 1000, 2)
    .protect().setDescription('Accounts — Payment Received & Date (P–Q)');
  payProt.removeEditors(payProt.getEditors());
  payProt.addEditors([...new Set([...adminEditors, ...accountsEditors])]);
  payProt.setWarningOnly(false);

  // Logistics can edit AWB No + Dispatch Date (M–N)
  const dispProt = sheet.getRange(2, OL_COL_AWB_NO, 1000, 2)
    .protect().setDescription('Logistics — AWB No & Dispatch Date (M–N)');
  dispProt.removeEditors(dispProt.getEditors());
  dispProt.addEditors([...new Set([...adminEditors, ...logisticsEditors])]);
  dispProt.setWarningOnly(false);

  // Remaining cols: Admin only
  const adminProt = sheet.getRange(2, OL_COL_CLIENT_NAME, 1000, OL_COL_ORDER_TYPE - 1)
    .protect().setDescription('Admin only — core order data B–F');
  adminProt.removeEditors(adminProt.getEditors());
  adminProt.addEditors(adminEditors);
  adminProt.setWarningOnly(false);

  // Note on usage
  sheet.getRange(1, OL_COL_UUID)
    .setNote('Order Ledger — one row per order. Auto-populated by sync scripts. Do not add rows manually.');

  Logger.log('setupHubOrderLedgerSheet_: "' + HUB_ORDER_LEDGER_SHEET + '" ready.');
}


// ============================================================
// setupHubClientSummarySheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Client Summary" sheet (one row per client).
// Columns: Client Name | CRM Owner | Total Orders | Total Invoiced |
//          Total Paid | Total Due | Credit Limit | Used Credit |
//          Available Credit | Credit Status | Last Order Date
// ============================================================
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


// ============================================================
// setupHubDuesSheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Client Dues" sheet (auto-populated, no manual entry).
// Shows only unpaid/partial orders, sorted by Due Date ascending.
// ============================================================
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


// ============================================================
// setupHubDashboardSheet_()
// NEW v14.0 — Internal helper called by setupHubSheet()
// Creates the "Dashboard" sheet (all clients, sorted by due balance desc).
// ============================================================
function setupHubDashboardSheet_(ss, adminEditors) {
  let sheet = ss.getSheetByName(HUB_DASHBOARD_SHEET);
  if (!sheet) sheet = ss.insertSheet(HUB_DASHBOARD_SHEET);

  const headers = [
    'Client Name',       // A  DB_COL_CLIENT_NAME
    'CRM Owner',         // B  DB_COL_CRM_OWNER
    'Total Due Balance', // C  DB_COL_TOTAL_DUE
    'Total Invoiced',    // D  DB_COL_TOTAL_INVOICED
    'Total Paid',        // E  DB_COL_TOTAL_PAID
    'Overdue Orders',    // F  DB_COL_OVERDUE_COUNT
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
    const adminEmails = [ownerEmail, ...config.adminEmails].filter((v, i, a) => a.indexOf(v) === i);

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

    // Col G: Last Updated — Sales + Admin
    const gProt = creditSheet.getRange(2, CC_COL_LAST_UPDATED, 999, 1)
      .protect().setDescription('Sales/Admin — Last Updated (G)');
    gProt.removeEditors(gProt.getEditors());
    gProt.addEditors(fullEditors);
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
function setupCRMSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── Part 1: Order Ledger CRM (formerly 'Main') ────────────
    let mainSheet = ss.getSheetByName(MAIN_SHEET_NAME);
    if (!mainSheet) mainSheet = ss.insertSheet(MAIN_SHEET_NAME);

    // Rename the tab to 'Order Ledger CRM' (display only — internal name stays 'Main')
    // Note: setName changes the tab label. All script references to MAIN_SHEET_NAME
    // will still find this sheet because getSheetByName() uses the internal name
    // only when there's no tab with the exact string — here we keep MAIN_SHEET_NAME
    // as the stored const and rename the visual tab for UX clarity.
    // IMPORTANT: After renaming, getSheetByName('Main') will FAIL because the tab
    // is now called 'Order Ledger CRM'. To preserve full backward compatibility of
    // all sync/trigger code, we do NOT rename here.  Instead we add a NOTE to the
    // header explaining the new name convention.
    // If you wish to rename the tab visually without breaking sync, do it manually
    // in the sheet and update MAIN_SHEET_NAME = 'Order Ledger CRM' in a future session.

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
// setupCRMDuesSheet_()
// NEW v14.0 — Internal helper called by setupCRMSheet()
// Creates the "Client Dues" sheet in the CRM spoke.
// This is a filtered/read-only view showing only this CRM user's
// orders that have outstanding balances.
// It is auto-populated by the Hub sync scripts — no manual entry.
// ============================================================
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


// ============================================================
// createClientTab()  — v14.0
// ─────────────────────────────────────────────────────────────
// CHANGE vs v13: Added "Due Date" header at col P (SUB_COL_DUE_DATE = 16).
// All column index references are unchanged (formulas use G:L which are
// Item Value through Discount — col indices unchanged).
// ============================================================
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
    const mainSheet = sheetFile.getSheetByName(MAIN_SHEET_NAME);
    if (!mainSheet) return clientSheet;

    let targetRow = mainSheetRow || null;

    if (!targetRow) {
      const crmDisplay = mainSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < crmDisplay.length; i++) {
        const cellVal = (crmDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim();
        if (cellVal.toLowerCase() === tabName.toLowerCase()) {
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


// ============================================================
// autoAddClientToCredit_()  — v14.0
// ─────────────────────────────────────────────────────────────
// Updated to use new CC_TOTAL_COLS (7) and CC_COL_UUID (8).
// Client Email (col F) defaults to blank — must be set manually
// or via a future onboarding flow.
// ============================================================
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


// ============================================================
// setCreditLimit_()  — v14.0
// Updated to use new CC_TOTAL_COLS (7) and CC_COL_UUID (8).
// Credit-request columns (old F–I) removed — no longer written.
// ============================================================
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
// invalidateConfigCache_()
// ============================================================
function invalidateConfigCache_() {
  try {
    CacheService.getScriptCache().remove('crm_config_v14');
    CacheService.getScriptCache().remove('crm_config_v3'); // remove old key too
    Logger.log('invalidateConfigCache_: cache cleared.');
  } catch (_) {}
}


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

// ============================================================
// END === SESSION 1 OUTPUT ===
// ============================================================
//
// MIGRATION NOTES FOR EXISTING DEPLOYMENTS:
//
//   1. Run "Setup -> Setup Hub Sheet" to create the 4 new Hub sheets.
//   2. Run "Setup -> Setup Spoke Sheet" on each CRM spoke to add the
//      Client Dues tab and update the CRM sheet headers.
//   3. Run "Setup -> Setup Credit Sheet" to apply the new column structure.
//      ⚠️  WARNING: The old columns F–I (Request New Limit, Request Reason,
//          Request Status, Approved By) will be REPLACED. Back up your
//          Client Credit sheet before running if those columns have data.
//   4. In CRM Members sheet, change any 'Sales Manager' role rows to 'Sales'.
//      The script handles both spellings in getCrmConfig_() for safety, but
//      the dropdown no longer offers 'Sales Manager' as a choice.
//   5. The CC_COL_UUID constant changed from 11 to 8. The onCreditRequestEdit()
//      trigger reads this column index when stamping UUID. Update that function
//      in Session 2 to replace the inline `const CC_COL_UUID = 11` with
//      the global `CC_COL_UUID = 8` from this file.
//   6. SUB_COL_DUE_DATE (col 16) is new. The EDITABLE_SUB_COLS array in
//      onCRMSubTabEdit() should be updated in Session 2 to include it.
//
// ============================================================
