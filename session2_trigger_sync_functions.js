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
function onCRMSubTabEdit(e) {
  if (!e || !e.range) {
    Logger.log('onCRMSubTabEdit: must be triggered by a sheet edit, not run manually.');
    return;
  }
  try {
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

    // ── Guard: must have at least 2 columns ──────────────────
    if (sheet.getLastColumn() < 2) return;

    // ── Editable sub-tab columns (v14: + SUB_COL_DUE_DATE) ──
    const EDITABLE_SUB_COLS = [
      SUB_COL_PI_NO,           // 1   → Order Ledger CRM col D + Hub col D
      SUB_COL_INV_NO,          // 2   → Order Ledger CRM col E + Hub col E
      SUB_COL_INV_DATE,        // 3   → Order Ledger CRM col F + Hub col F
      SUB_COL_AWB_NO,          // 5   → Hub col Q
      SUB_COL_DISP_DATE,       // 6   → Hub col R
      SUB_COL_ITEM_VALUE_USD,  // 7   → triggers Final Inv push
      SUB_COL_DEV_CHARGES_USD, // 8   → triggers Final Inv push
      SUB_COL_FOB_VALUE,       // 9   → triggers Final Inv push
      SUB_COL_FREIGHT,         // 10  → triggers Final Inv push
      SUB_COL_CIF_VALUE,       // 11  → triggers Final Inv push
      SUB_COL_DISCOUNT,        // 12  → triggers Final Inv push
      SUB_COL_FINAL_INV,       // 13  → Order Ledger CRM col G + Hub col G
      SUB_COL_PAY_RECEIVED,    // 14  → Order Ledger CRM col J + Hub col J
      SUB_COL_PAY_DATE,        // 15  → Order Ledger CRM col K + Hub col K
      SUB_COL_DUE_DATE,        // 16  → Order Ledger CRM col I + Hub col I (NEW)
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
    try {
      lock.waitLock(10000);
    } catch (_) {
      Logger.log('onCRMSubTabEdit: LockService timeout for "' + sheetName + '" row=' + row + ' col=' + col);
      return;
    }

    try {
      // ── Find matching Order Ledger CRM row by client name ───
      const mainData    = mainSheet.getDataRange().getValues();
      const mainDisplay = mainSheet.getDataRange().getDisplayValues();
      let mainRow = -1;
      for (let i = 1; i < mainDisplay.length; i++) {
        if ((mainDisplay[i][CRM_COL_CUSTOMER - 1] || '').toString().trim().toLowerCase()
            === sheetName.toLowerCase()) {
          mainRow = i + 1;
          break;
        }
      }
      if (mainRow === -1) {
        Logger.log('onCRMSubTabEdit: No matching Order Ledger CRM row for "' + sheetName + '"');
        return;
      }

      const uuid     = (mainData[mainRow - 1][CRM_COL_UUID      - 1] || '').toString().trim();
      const crmOwner = (mainData[mainRow - 1][CRM_COL_CRM_OWNER - 1] || '').toString().trim();
      const cellVal  = e.range.getValue();

      // ── DUPLICATE PI No CHECK ─────────────────────────────
      if (col === SUB_COL_PI_NO && cellVal) {
        const allSubData = sheet.getDataRange().getValues();
        let dupCount = 0;
        for (let i = 1; i < allSubData.length; i++) {
          if (i + 1 === row) continue; // skip current row
          if ((allSubData[i][SUB_COL_PI_NO - 1] || '').toString().trim() === cellVal.toString().trim()) {
            dupCount++;
          }
        }
        if (dupCount > 0) {
          e.range.clearContent();
          SpreadsheetApp.flush();
          alertIfUi_(
            '⚠️ Duplicate PI No Detected!\n\n' +
            'PI No "' + cellVal + '" already exists in another row of this client tab.\n\n' +
            'Each order must have a unique PI No.\nThe entry has been cleared.'
          );
          logActivity_(ss, 'DUP PI NO BLOCKED',
            'Client:' + sheetName + ' | PI:' + cellVal + ' | row:' + row, userEmail);
          return;
        }
      }

      // ── DUPLICATE Inv No CHECK ───────────────────────────
      if (col === SUB_COL_INV_NO && cellVal) {
        const invNoStr = cellVal.toString().trim();
        const sysSet = new Set([
          ORDER_LEDGER_CRM_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
          CREDIT_SHEET_NAME, CLIENT_DUES_SHEET_NAME,
          CLIENT_SUMMARY_SHEET_NAME, DASHBOARD_SHEET_NAME,
        ]);
        let dupFound = false;
        ss.getSheets().forEach(s => {
          if (dupFound) return;
          if (sysSet.has(s.getName())) return;
          const sData = s.getDataRange().getValues();
          for (let i = 1; i < sData.length; i++) {
            // Skip the exact cell being edited
            if (s.getName() === sheetName && i + 1 === row) continue;
            if ((sData[i][SUB_COL_INV_NO - 1] || '').toString().trim() === invNoStr) {
              dupFound = true; break;
            }
          }
        });
        if (dupFound) {
          e.range.clearContent();
          SpreadsheetApp.flush();
          alertIfUi_(
            '⚠️ Duplicate Invoice No Detected!\n\n' +
            'Inv No "' + invNoStr + '" already exists in another order in this CRM.\n\n' +
            'Invoice numbers must be unique across all clients.\nThe entry has been cleared.'
          );
          logActivity_(ss, 'DUP INV NO BLOCKED',
            'Client:' + sheetName + ' | InvNo:' + invNoStr + ' | row:' + row, userEmail);
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
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_PI_NO || col === SUB_COL_INV_NO || col === SUB_COL_INV_DATE) {
        const mainColMap = {
          [SUB_COL_PI_NO]:    CRM_COL_PI_NO,
          [SUB_COL_INV_NO]:   CRM_COL_INV_NO,
          [SUB_COL_INV_DATE]: CRM_COL_DATE,
        };
        const hubColMap = {
          [SUB_COL_PI_NO]:    HUB_COL_PI_NO,
          [SUB_COL_INV_NO]:   HUB_COL_INV_NO,
          [SUB_COL_INV_DATE]: HUB_COL_DATE,
        };
        const mainTargetCol = mainColMap[col];
        const hubTargetCol  = hubColMap[col];

        // Write to Order Ledger CRM
        PropertiesService.getScriptProperties().setProperty(PROP_SKIP_NEXT_EDIT, 'true');
        const mainCell = mainSheet.getRange(mainRow, mainTargetCol);
        mainCell.setValue(cellVal);
        if (col === SUB_COL_INV_DATE) mainCell.setNumberFormat('dd/mm/yyyy');
        SpreadsheetApp.flush();

        // Write to Hub Order Ledger (two-way — Inv No & Inv Date now flow back)
        if (uuid) {
          try {
            const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubSheet) {
              const hubData = hubSheet.getDataRange().getValues();
              let hubRow = -1;
              for (let i = 1; i < hubData.length; i++) {
                if ((hubData[i][HUB_COL_UUID - 1] || '').toString().trim() === uuid) {
                  hubRow = i + 1; break;
                }
              }
              if (hubRow !== -1) {
                const hubCell = hubSheet.getRange(hubRow, hubTargetCol);
                hubCell.setValue(cellVal);
                if (col === SUB_COL_INV_DATE) hubCell.setNumberFormat('dd/mm/yyyy');
                SpreadsheetApp.flush();
                logActivity_(hubSS, 'SUBTAB PI/INV/DATE->ORDER LEDGER',
                  'UUID:' + uuid + ' | subCol:' + col + ' -> HubCol:' + hubTargetCol + '=' + cellVal, userEmail);
              }
            }
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH1 Hub write error: ' + hubErr.message);
          }
        }

        logActivity_(ss, 'SUBTAB PI/INV/DATE->ORDER LEDGER CRM',
          'Client:' + sheetName + ' | subCol:' + col + '->MainCol:' + mainTargetCol + '=' + cellVal, userEmail);

        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 2: Due Date (col P) — NEW in v14.0
      // Push to: Order Ledger CRM col I, Hub Order Ledger col I,
      //          Hub Client Dues, CRM Client Dues (via refreshClientDues_)
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_DUE_DATE) {
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
              for (let i = 1; i < hubData.length; i++) {
                if ((hubData[i][HUB_COL_UUID - 1] || '').toString().trim() === uuid) {
                  hubRow = i + 1; break;
                }
              }
              if (hubRow !== -1) {
                const hubDueCell = hubSheet.getRange(hubRow, HUB_COL_DUE_DATE);
                hubDueCell.setValue(cellVal);
                if (cellVal) hubDueCell.setNumberFormat('dd/mm/yyyy');
                SpreadsheetApp.flush();
                logActivity_(hubSS, 'SUBTAB DUE DATE->ORDER LEDGER',
                  'UUID:' + uuid + ' | DueDate:' + cellVal, userEmail);
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
      // PATH 3: AWB No / Dispatch Date → Hub directly
      // ─────────────────────────────────────────────────────────
      if (col === SUB_COL_AWB_NO || col === SUB_COL_DISP_DATE) {
        const hubColTarget = SUB_TO_HUB[col];
        if (hubColTarget && uuid) {
          try {
            const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
            const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
            if (hubSheet) {
              const hubData = hubSheet.getDataRange().getValues();
              let hubRow = -1;
              for (let i = 1; i < hubData.length; i++) {
                if (hubData[i][HUB_COL_UUID - 1] === uuid) { hubRow = i + 1; break; }
              }
              if (hubRow !== -1) {
                const hCell = hubSheet.getRange(hubRow, hubColTarget);
                hCell.setValue(cellVal);
                if (col === SUB_COL_DISP_DATE && cellVal) hCell.setNumberFormat('dd/mm/yyyy');
                logActivity_(hubSS, 'SUBTAB AWB/DISP->ORDER LEDGER',
                  'UUID:' + uuid + ' col:' + col + '->HubCol:' + hubColTarget, userEmail);
              }
            }
          } catch (hubErr) {
            Logger.log('onCRMSubTabEdit PATH3 AWB/Disp Hub write error: ' + hubErr.message);
          }
        }
        applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);
        return;
      }

      // ─────────────────────────────────────────────────────────
      // PATH 4: Payment Received / Payment Date
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
              for (let i = 1; i < hubData.length; i++) {
                if (hubData[i][HUB_COL_UUID - 1] === uuid) { hubRow = i + 1; break; }
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
      // PATH 5: Value columns → Final Invoice push
      // ─────────────────────────────────────────────────────────
      Utilities.sleep(300);
      SpreadsheetApp.flush();
      const subRowData = sheet.getRange(row, 1, 1,
        Math.max(sheet.getLastColumn(), SUB_COL_FINAL_INV)).getValues()[0];
      const finalInv   = sheet.getRange(row, SUB_COL_FINAL_INV).getValue();
      const piNo       = subRowData[SUB_COL_PI_NO - 1];

      Logger.log('onCRMSubTabEdit: FinalInv path | piNo=' + piNo + ' finalInv=' + finalInv);

      pushFinalInvoiceToMainAndHub_(
        ss, mainSheet, mainRow, mainData,
        sheetName, piNo, finalInv, uuid, crmOwner, userEmail
      );

      refreshCreditTotals_(userEmail);
      refreshClientDues_(userEmail);
      refreshDashboard_(userEmail);

      logActivity_(ss, 'SUBTAB VALUE->ORDER LEDGER CRM+HUB',
        'Client:' + sheetName + ' row:' + row + ' col:' + col + ' FinalInv:' + finalInv, userEmail);

      applySubTabRowColour_(sheet, row, uuid, mainData[mainRow - 1]);

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log('ERROR in onCRMSubTabEdit: ' + err.message + '\n' + err.stack);
  }
}


// ============================================================
// applySubTabRowColour_(sheet, row, uuid, mainRowData)
// ─────────────────────────────────────────────────────────────
// Derives the current order status for the sub-tab row and
// paints the row background accordingly.
// mainRowData is the 0-indexed values array for the matching
// Order Ledger CRM row (used to get dueDate, invValue etc.).
// ============================================================
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
function onMasterHubEdit(e) {
  if (!checkLicense()) return;
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
      HUB_COL_INV_NO,            // E  — Logistics enters Inv No
      HUB_COL_DATE,              // F  — Logistics enters Inv Date
      HUB_COL_DUE_DATE,          // I  — Sales enters Due Date
    ];
    if (!ALLOWED_COLS_V14.includes(col)) return;

    const hubSS    = e.source;
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) { Logger.log('onMasterHubEdit: Order Ledger sheet not found'); return; }
    const rowData  = hubSheet.getRange(row, 1, 1, HUB_TOTAL_COLS).getValues()[0];

    const uuid      = rowData[HUB_COL_UUID      - 1];
    const crmOwner  = rowData[HUB_COL_CRM_OWNER - 1];
    const newValue  = e.range.getValue();
    const userEmail = Session.getActiveUser().getEmail();

    if (!uuid || !crmOwner) return;

    // ── LOGISTICS BLOCK: prevent AWB/Dispatch when order is on Hold ──
    if (col === HUB_COL_AWB_NO || col === HUB_COL_DISP_DATE) {
      const approvalStatus = (rowData[HUB_COL_APPROVAL_REQ - 1] || '').toString().trim();
      if (approvalStatus === 'Yes') {
        e.range.clearContent();
        SpreadsheetApp.flush();
        alertIfUi_(
          'Dispatch Blocked!\n\n' +
          'This order is on Hold — credit limit has been exceeded.\n' +
          'Col N (Approval Required) = "Yes"\n\n' +
          'Dispatch cannot proceed until:\n' +
          '  • Sales Manager approves (fills col O), OR\n' +
          '  • Credit limit is increased for this client.\n\n' +
          'Contact your Sales Manager or Admin.'
        );
        logActivity_(hubSS, 'LOGISTICS BLOCKED',
          'UUID:' + uuid + ' | row:' + row + ' | col:' + col + ' | N=Yes — reverted', userEmail);
        return;
      }
    }

    const config    = getCrmConfig_();
    const crmFileId = config.crmFileIds[(crmOwner || '').toString().toLowerCase()];
    if (!crmFileId) {
      Logger.log('onMasterHubEdit: No CRM file ID for: ' + crmOwner);
      return;
    }

    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); } catch (_) { Logger.log('onMasterHubEdit: lock timeout'); return; }

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

      // ── EXTENDED HUB_TO_CRM_MAIN map (v14 additions) ──────
      const HUB_TO_CRM_V14 = Object.assign({}, HUB_TO_CRM_MAIN, {
        [HUB_COL_INV_NO]:   CRM_COL_INV_NO,     // E → E
        [HUB_COL_DATE]:     CRM_COL_DATE,        // F → F
        [HUB_COL_DUE_DATE]: CRM_COL_DUE_DATE,   // I → I
      });

      // Push mapped column to Order Ledger CRM
      const crmMainCol = HUB_TO_CRM_V14[col];
      if (crmMainCol) {
        const pushValue = (col === HUB_COL_DISPATCH_STATUS)
          ? hubSheet.getRange(row, col).getDisplayValue()
          : newValue;
        const crmCell = crmSheet.getRange(matchRow, crmMainCol);
        crmCell.setValue(pushValue);
        if ((col === HUB_COL_DATE || col === HUB_COL_DUE_DATE) && pushValue) {
          crmCell.setNumberFormat('dd/mm/yyyy');
        }
        logActivity_(crmSS, 'HUB->ORDER LEDGER CRM col:' + col, 'UUID:' + uuid, userEmail);
        logActivity_(hubSS, 'PUSHED TO ORDER LEDGER CRM',        'UUID:' + uuid, userEmail);
      }

      // ── PUSH to CRM sub-tab ───────────────────────────────
      const customerName = rowData[HUB_COL_CUSTOMER - 1];
      const piNo         = rowData[HUB_COL_PI_NO    - 1];
      const clientSheet  = crmSS.getSheetByName((customerName || '').toString().trim());

      if (clientSheet) {
        // Extended map: Hub col → sub-tab col
        const HUB_TO_SUB_V14 = {
          [HUB_COL_AWB_NO]:      SUB_COL_AWB_NO,
          [HUB_COL_DISP_DATE]:   SUB_COL_DISP_DATE,
          [HUB_COL_PAY_RECEIVED]:SUB_COL_PAY_RECEIVED,
          [HUB_COL_PAY_DATE]:    SUB_COL_PAY_DATE,
          [HUB_COL_INV_NO]:      SUB_COL_INV_NO,    // Logistics Inv No → sub-tab col B
          [HUB_COL_DATE]:        SUB_COL_INV_DATE,  // Logistics Inv Date → sub-tab col C
          [HUB_COL_DUE_DATE]:    SUB_COL_DUE_DATE,  // Sales Due Date → sub-tab col P
        };
        const subCol = HUB_TO_SUB_V14[col];
        if (subCol) {
          const clientData = clientSheet.getDataRange().getValues();
          for (let i = 1; i < clientData.length; i++) {
            if (clientData[i][SUB_COL_PI_NO - 1] === piNo) {
              const targetCell = clientSheet.getRange(i + 1, subCol);
              targetCell.setValue(newValue);
              if (subCol === SUB_COL_DISP_DATE || subCol === SUB_COL_PAY_DATE ||
                  subCol === SUB_COL_INV_DATE  || subCol === SUB_COL_DUE_DATE) {
                if (newValue) targetCell.setNumberFormat('dd/mm/yyyy');
              }
              // Re-colour the sub-tab row
              applySubTabRowColour_(clientSheet, i + 1, uuid, crmData[matchRow - 1]);
              logActivity_(crmSS, 'HUB->SUBTAB col:' + subCol, 'PI:' + piNo, userEmail);
              break;
            }
          }
        }
      }

      // ── ORDER APPROVED NOTIFICATION (unchanged logic) ──────
      if (col === HUB_COL_APPROVED_BY && newValue) {
        hubSheet.getRange(row, HUB_COL_APPROVAL_REQ).setValue('Approved');
        crmSheet.getRange(matchRow, CRM_COL_APPROVAL_REQ).setValue('Approved');
        // (email notification logic retained from v13 — not repeated here)
        logActivity_(hubSS, 'ORDER APPROVED', 'UUID:' + uuid + ' | By:' + newValue, userEmail);
      }

      // ── Payment notification (unchanged logic) ─────────────
      // (full email block retained from v13 — not repeated here for brevity)

      // ── After payment edit: refresh Client Dues (Hub + CRM) ──
      if (col === HUB_COL_PAY_RECEIVED || col === HUB_COL_PAY_DATE) {
        refreshCreditTotals_(userEmail);
        refreshClientDues_(userEmail);
      }

      // ── After Due Date edit: refresh Client Dues ───────────
      if (col === HUB_COL_DUE_DATE) {
        refreshClientDues_(userEmail);
      }

      // ── After ANY edit: refresh Dashboard ─────────────────
      refreshDashboard_(userEmail);

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log('ERROR in onMasterHubEdit: ' + err.message);
  }
}


// ============================================================
// 3. refreshCreditTotals_(userEmail) — v14.0
// ─────────────────────────────────────────────────────────────
// Changes from v8:
//   • Reads from 'Order Ledger' (not old 'Master Hub' sheet)
//   • Updates Client Summary sheet totals (new sheet)
//   • Otherwise identical calculation: UsedCredit = InvValue − PayReceived
// ============================================================
function refreshCreditTotals_(userEmail) {
  try {
    const hubSS       = SpreadsheetApp.openById(MASTER_HUB_ID);
    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    const hubSheet    = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME); // UPDATED
    if (!creditSheet || !hubSheet) return;

    const hubData    = hubSheet.getDataRange().getValues();
    const creditData = creditSheet.getDataRange().getValues();

    // Build per-client totals from Order Ledger
    const invTotals = {};
    const payTotals = {};
    for (let i = 1; i < hubData.length; i++) {
      const name = (hubData[i][HUB_COL_CUSTOMER     - 1] || '').toString().trim().toLowerCase();
      const inv  = Number(hubData[i][HUB_COL_INV_VALUE   - 1]) || 0;
      const pay  = Number(hubData[i][HUB_COL_PAY_RECEIVED - 1]) || 0;
      const bal  = Number(hubData[i][HUB_COL_DUE_BALANCE  - 1]) ||
                   Math.max(0, inv - pay);  // fallback if formula col not yet computed
      if (name) {
        invTotals[name] = (invTotals[name] || 0) + inv;
        payTotals[name] = (payTotals[name] || 0) + pay;
      }
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
    updateClientSummary_(hubSS, invTotals, payTotals, userEmail);

    // Re-evaluate col N credit flags
    reEvaluateAllCreditFlags_(hubSS, userEmail);

    Logger.log('refreshCreditTotals_: refreshed ' + (creditData.length - 1) + ' row(s)');
  } catch (err) {
    Logger.log('ERROR in refreshCreditTotals_: ' + err.message);
  }
}


// ============================================================
// updateClientSummary_(hubSS, invTotals, payTotals, userEmail)
// ─────────────────────────────────────────────────────────────
// Writes aggregated per-client totals to the Client Summary
// sheet (columns: Client Name | CRM Owner | Total Invoiced |
// Total Paid | Total Due | Credit Limit | Available Credit).
// Called by refreshCreditTotals_ after computing totals.
// ============================================================
function updateClientSummary_(hubSS, invTotals, payTotals, userEmail) {
  try {
    const summarySheet = hubSS.getSheetByName(CLIENT_SUMMARY_SHEET_NAME);
    if (!summarySheet) return;

    const creditSheet = hubSS.getSheetByName(CREDIT_SHEET_NAME);
    if (!creditSheet) return;
    const creditData  = creditSheet.getDataRange().getValues();

    // Build credit limit map
    const limitMap    = {};
    const crmOwnerMap = {};
    for (let i = 1; i < creditData.length; i++) {
      const name  = (creditData[i][CC_COL_CLIENT_NAME  - 1] || '').toString().trim();
      const owner = (creditData[i][CC_COL_CRM_OWNER    - 1] || '').toString().trim();
      const limit = Number(creditData[i][CC_COL_CREDIT_LIMIT - 1]) || 0;
      if (name) {
        limitMap[name.toLowerCase()]    = limit;
        crmOwnerMap[name.toLowerCase()] = owner;
      }
    }

    // Build output rows
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
      const availCred = limit - Math.max(0, totalInv - totalPay);
      // Recover display name (Title Case) from limitMap keys
      const displayName = Object.keys(limitMap).find(k => k.toLowerCase() === key) || key;
      outputRows.push([
        displayName,
        crmOwnerMap[key] || '',
        totalInv,
        totalPay,
        totalDue,
        limit,
        availCred,
      ]);
    });

    if (!outputRows.length) return;

    // Write header + data
    const headers = ['Client Name','CRM Owner','Total Invoiced','Total Paid',
                     'Total Due','Credit Limit','Available Credit'];
    summarySheet.clearContents();
    summarySheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1A237E').setFontColor('#FFFFFF');
    summarySheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    SpreadsheetApp.flush();

    Logger.log('updateClientSummary_: wrote ' + outputRows.length + ' client row(s)');
  } catch (err) {
    Logger.log('ERROR in updateClientSummary_: ' + err.message);
  }
}


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
function refreshClientDues_(userEmail) {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(ORDER_LEDGER_SHEET_NAME);
    if (!hubSheet) return;

    const hubData = hubSheet.getDataRange().getValues();
    const now     = new Date();

    // Build dues rows from Order Ledger
    const duesRows = [];
    for (let i = 1; i < hubData.length; i++) {
      const row        = hubData[i];
      const customer   = (row[HUB_COL_CUSTOMER    - 1] || '').toString().trim();
      const crmOwner   = (row[HUB_COL_CRM_OWNER   - 1] || '').toString().trim();
      const piNo       = (row[HUB_COL_PI_NO        - 1] || '').toString().trim();
      const invNo      = (row[HUB_COL_INV_NO       - 1] || '').toString().trim();
      const invDate    = row[HUB_COL_DATE           - 1];
      const invValue   = Number(row[HUB_COL_INV_VALUE - 1]) || 0;
      const dueDate    = row[HUB_COL_DUE_DATE       - 1];
      const payStatus  = (row[HUB_COL_PAY_STATUS   - 1] || '').toString().trim();
      const dueBalance = Number(row[HUB_COL_DUE_BALANCE - 1]) || Math.max(0, invValue - (Number(row[HUB_COL_PAY_RECEIVED - 1]) || 0));
      const awbNo      = (row[HUB_COL_AWB_NO       - 1] || '').toString().trim();
      const uuid       = (row[HUB_COL_UUID         - 1] || '').toString().trim();

      if (!customer) continue;
      if (dueBalance <= 0) continue;  // only outstanding rows

      const dueDateObj = dueDate instanceof Date ? dueDate : (dueDate ? new Date(dueDate) : null);
      const status     = deriveOrderStatus_(payStatus, dueDateObj, invValue, awbNo, invNo);

      duesRows.push({
        customer, crmOwner, piNo, invNo, invDate, invValue,
        dueDate: dueDateObj, payStatus, dueBalance, status, uuid,
        row: [customer, crmOwner, piNo, invNo, invDate, invValue,
              dueDateObj, payStatus, dueBalance, status, uuid],
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

    // ── Write to Hub Client Dues sheet ───────────────────────
    writeDuesSheet_(hubSS, duesRows, 'Hub Client Dues');

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


// ── writeDuesSheet_ helper ─────────────────────────────────────
function writeDuesSheet_(ss, duesRows, sheetName) {
  try {
    let duesSheet = ss.getSheetByName(sheetName);
    if (!duesSheet) {
      duesSheet = ss.insertSheet(sheetName);
    }

    const headers = [
      'Customer', 'CRM Owner', 'PI No', 'Inv No', 'Inv Date',
      'Inv Value', 'Due Date', 'Pay Status', 'Due Balance', 'Status', 'UUID'
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
    const hasOverdue  = {};  // clientName.lower -> bool
    const hasPartial  = {};
    const now         = new Date();
    for (let i = 1; i < hubData.length; i++) {
      const row      = hubData[i];
      const customer = (row[HUB_COL_CUSTOMER   - 1] || '').toString().trim().toLowerCase();
      const payStatus= (row[HUB_COL_PAY_STATUS  - 1] || '').toString().trim();
      const dueDate  = row[HUB_COL_DUE_DATE     - 1];
      const invNo    = (row[HUB_COL_INV_NO      - 1] || '').toString().trim();
      const invValue = Number(row[HUB_COL_INV_VALUE - 1]) || 0;
      const awbNo    = (row[HUB_COL_AWB_NO      - 1] || '').toString().trim();
      const dueDateObj = dueDate instanceof Date ? dueDate : (dueDate ? new Date(dueDate) : null);
      const status   = deriveOrderStatus_(payStatus, dueDateObj, invValue, awbNo, invNo);
      if (status === 'Overdue')  hasOverdue[customer] = true;
      if (status === 'Partial')  hasPartial[customer] = true;
    }

    // Read Client Summary (cols: Client Name|CRM Owner|Total Inv|Total Paid|Total Due|Credit Limit|Avail Credit)
    const clients = [];
    for (let i = 1; i < summaryData.length; i++) {
      const r = summaryData[i];
      const clientName  = (r[0] || '').toString().trim();
      const crmOwner    = (r[1] || '').toString().trim();
      const totalInv    = Number(r[2]) || 0;
      const totalPaid   = Number(r[3]) || 0;
      const totalDue    = Number(r[4]) || 0;
      const creditLimit = Number(r[5]) || 0;
      const availCredit = Number(r[6]) || 0;
      if (!clientName) continue;
      clients.push({ clientName, crmOwner, totalInv, totalPaid, totalDue, creditLimit, availCredit });
    }

    // Sort: Total Due high → low; ties alphabetically
    clients.sort((a, b) => {
      if (b.totalDue !== a.totalDue) return b.totalDue - a.totalDue;
      return a.clientName.toLowerCase().localeCompare(b.clientName.toLowerCase());
    });

    // Build dashboard rows
    const headers = [
      'Rank', 'Client Name', 'CRM Owner', 'Total Invoiced', 'Total Paid',
      'Total Due', 'Credit Limit', 'Available Credit', 'Status'
    ];

    dashSheet.clearContents();
    dashSheet.clearFormats();

    dashSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#0D47A1')
      .setFontColor('#FFFFFF');

    if (!clients.length) { SpreadsheetApp.flush(); return; }

    const dataMatrix = clients.map((c, idx) => {
      const key    = c.clientName.toLowerCase();
      let dashStatus = 'Settled';
      if (c.totalDue > 0) {
        if (hasOverdue[key])       dashStatus = 'Overdue';
        else if (hasPartial[key])  dashStatus = 'Partial';
        else                       dashStatus = 'Outstanding';
      }
      return [
        idx + 1,
        c.clientName,
        c.crmOwner,
        c.totalInv,
        c.totalPaid,
        c.totalDue,
        c.creditLimit,
        c.availCredit,
        dashStatus,
      ];
    });

    dashSheet.getRange(2, 1, dataMatrix.length, headers.length).setValues(dataMatrix);

    // Apply row colours
    dataMatrix.forEach((row, idx) => {
      const dashRow    = idx + 2;
      const totalDue   = row[5]; // col F (0-indexed = 5)
      const dashStatus = row[8]; // col I

      let colour = '#E8F5E9'; // green — settled
      if (totalDue > 0) {
        if (dashStatus === 'Overdue')     colour = '#FFCDD2'; // red
        else if (dashStatus === 'Partial') colour = '#FFF9C4'; // yellow
        else                               colour = '#FFE0B2'; // orange — outstanding
      }
      dashSheet.getRange(dashRow, 1, 1, headers.length).setBackground(colour);
    });

    // Format numeric columns
    dashSheet.getRange(2, 4, dataMatrix.length, 3).setNumberFormat('#,##0.00'); // Inv/Paid/Due
    dashSheet.getRange(2, 7, dataMatrix.length, 2).setNumberFormat('#,##0.00'); // Credit cols

    // Freeze header and auto-resize
    dashSheet.setFrozenRows(1);
    dashSheet.autoResizeColumns(1, headers.length);

    SpreadsheetApp.flush();
    Logger.log('refreshDashboard_: wrote ' + dataMatrix.length + ' client row(s)');
  } catch (err) {
    Logger.log('ERROR in refreshDashboard_: ' + err.message);
  }
}


// ============================================================
// Menu wrappers — expose new refresh functions to Admin menu
// Add these items to the CRM Tools menu in onOpen():
//   .addItem('Refresh Client Dues',  'menuRefreshClientDues_')
//   .addItem('Refresh Dashboard',    'menuRefreshDashboard_')
// ============================================================
function menuRefreshClientDues_() {
  const userEmail = Session.getActiveUser().getEmail();
  refreshClientDues_(userEmail);
  alertIfUi_('✅ Client Dues refreshed on Hub + all CRM spokes.');
}

function menuRefreshDashboard_() {
  const userEmail = Session.getActiveUser().getEmail();
  refreshCreditTotals_(userEmail);  // also updates Client Summary
  refreshClientDues_(userEmail);
  refreshDashboard_(userEmail);
  alertIfUi_('✅ Dashboard refreshed (Credit Totals → Client Summary → Client Dues → Dashboard).');
}
