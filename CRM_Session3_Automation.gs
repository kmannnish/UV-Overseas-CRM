// ============================================================
// === SESSION 3 OUTPUT ===
// CRM AUTOMATION & ALERT FUNCTIONS
// Paste this block AFTER the existing code in your Apps Script.
// Do NOT modify any existing functions.
// Depends on: all constants, getCrmConfig_(), logActivity_(),
//             alertIfUi_() already defined in the master file.
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


// ============================================================
// buildAutomationMenu_()
// ─────────────────────────────────────────────────────────────
// Called by toggleReminders_() and onOpen() to add the
// Automation sub-menu with the current ON/OFF label.
// Safe to call multiple times — addToUi() appends or updates.
// ============================================================
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
function runDailyReminders_() {
  try {
    const hubSS    = SpreadsheetApp.openById(MASTER_HUB_ID);
    const hubSheet = hubSS.getSheetByName(HUB_MAIN_SHEET_NAME);
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


// ── Email helpers for runDailyReminders_ ─────────────────────

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

// Shared HTML email template for reminder/warning/due-today/escalation
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

// Safe GmailApp.sendEmail wrapper — logs errors without throwing
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
function menuSendOrderSummary_() {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const clientName  = activeSheet.getName();

    // Guard: don't run on system sheets
    const SYSTEM = new Set([
      MAIN_SHEET_NAME, LOG_SHEET_NAME, MEMBERS_SHEET_NAME,
      CREDIT_SHEET_NAME, HUB_MAIN_SHEET_NAME
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
    const hubSheet  = hubSS.getSheetByName(HUB_MAIN_SHEET_NAME);
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
function applySubTabRowColour_(sheet, row, rowData) {
  try {
    if (!sheet || !row || !rowData) return;

    const COLOURS = {
      created   : '#FFFFFF',
      invoiced  : '#E3F2FD',
      dispatched: '#FFF3E0',
      pending   : '#FFF9C4',
      overdue   : '#FFCDD2',
      partial   : '#FFFFF0',
      paid      : '#E8F5E9',
    };

    // Extract relevant fields (0-based indices into sub-tab row)
    const piNo      = (rowData[SUB_COL_PI_NO   - 1] || '').toString().trim();
    const invNo     = (rowData[SUB_COL_INV_NO  - 1] || '').toString().trim();
    const awbNo     = (rowData[SUB_COL_AWB_NO  - 1] || '').toString().trim();
    const dispDate  = rowData[SUB_COL_DISP_DATE - 1];
    const payRec    = Number(rowData[SUB_COL_PAY_RECEIVED - 1]) || 0;

    // Final Invoice is a formula cell — read it live from the sheet
    const finalInv  = Number(sheet.getRange(row, SUB_COL_FINAL_INV).getValue()) || 0;

    // Overdue sentinel — caller may inject rowData[15] = 'OVERDUE'
    const isOverdue = rowData.length > 15
      && (rowData[15] || '').toString().toUpperCase() === 'OVERDUE';

    // Determine state
    let colour;
    if (!piNo) {
      return; // Empty row — do nothing
    } else if (isOverdue) {
      colour = COLOURS.overdue;
    } else if (finalInv > 0 && payRec >= finalInv) {
      colour = COLOURS.paid;
    } else if (finalInv > 0 && payRec > 0 && payRec < finalInv) {
      colour = COLOURS.partial;
    } else if (finalInv > 0 && payRec === 0 && (awbNo || dispDate)) {
      colour = COLOURS.pending;
    } else if (awbNo || (dispDate instanceof Date)) {
      colour = COLOURS.dispatched;
    } else if (invNo) {
      colour = COLOURS.invoiced;
    } else {
      colour = COLOURS.created;
    }

    const numCols = sheet.getLastColumn() || 15;
    sheet.getRange(row, 1, 1, numCols).setBackground(colour);

  } catch (err) {
    Logger.log('ERROR in applySubTabRowColour_: ' + err.message);
  }
}


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

// ── Companion reader: retrieve soft-due threshold for a UUID ─
// Returns a Date or null.  Used by runDailyReminders_() rule (d)
// to check "no Due Date, invoice date > 30 days ago".
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
// END OF SESSION 3 OUTPUT
// ============================================================
// INTEGRATION CHECKLIST:
//
// 1. Paste this block AFTER the last line of your existing
//    Master_and_CRM_Code_v3 script (no edits needed there).
//
// 2. Add buildAutomationMenu_() call inside your existing
//    onOpen() function so the ⚡ Automation menu appears:
//
//      function onOpen() {
//        // ... existing menu code ...
//        buildAutomationMenu_();   // ← add this line
//      }
//
// 3. In onCRMSubTabEdit, when col === SUB_COL_INV_DATE fires,
//    add these two calls after the existing sync logic:
//
//      // Colour the row
//      const subRowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
//      applySubTabRowColour_(sheet, row, subRowData);
//
//      // Store soft due-date threshold (does NOT write to sheet)
//      autoSetDueDate_(uuid, cellVal);
//
// 4. In onCRMSubTabEdit, when col === SUB_COL_PI_NO or
//    col === SUB_COL_INV_NO, add duplicate check:
//
//      const dupCheck = checkDuplicateEntry_(sheet, row, col, cellVal);
//      if (dupCheck.isDuplicate) {
//        SpreadsheetApp.getActiveSpreadsheet().toast(
//          `⚠️ Duplicate value "${cellVal}" already exists in row ${dupCheck.existingRow}. Please verify.`,
//          'Duplicate Detected', 10
//        );
//        Logger.log('checkDuplicateEntry_: duplicate at row ' + dupCheck.existingRow);
//      }
//
// 5. Trigger toggleReminders_() from menu to install the daily
//    9AM time trigger.  No manual ScriptApp call needed.
//
// 6. The ⚡ Automation > Send Order Summary Draft menu item
//    works on any CRM Spoke spreadsheet (not on the Hub).
//    Make sure the Hub's Client Credit sheet has client email
//    addresses in column F for the draft To: field to populate.
// ============================================================
