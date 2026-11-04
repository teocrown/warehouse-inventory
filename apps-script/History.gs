/**
 * Crown Elite Warehouse Inventory: ticket history (view only)
 *
 * Reads every Northside ticket email since Dec 26 2025 into a "History" sheet
 * so the owner can look up old tickets by P.O., date and total, and open the PDF.
 * NOTHING here touches stock: no Tickets, TicketLines or Movements rows are written.
 *
 * This is its own file in the Apps Script project (History.gs).
 * One time: pick "historySetup" in the function menu and click Run.
 * It keeps going on its own every few minutes until every email is read.
 */

const HIST_FROM = '2025/12/24';            // first email date to read
const HIST_FIRST_TICKET = '2025-12-26';    // first ticket date kept (start of the Feb 10 statement)
const HIST_QUERY = 'from:noreply@northsideplumbing.com has:attachment';
const HIST_HEAD = ['Key', 'Kind', 'Ticket', 'Date', 'PO', 'ShipVia', 'Total', 'Link', 'Note', 'At'];

function histSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('History');
  if (!sh) sh = ss.insertSheet('History');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HIST_HEAD.length).setValues([HIST_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:J').setNumberFormat('@');
  }
  return sh;
}
function histRows_() {
  const sh = histSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HIST_HEAD.length).getDisplayValues().map(function (r) {
    const o = {}; HIST_HEAD.forEach(function (h, i) { o[h] = r[i]; }); return o;
  });
}
function histStopTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'historyContinue') ScriptApp.deleteTrigger(t);
  });
}

/* ---------- one time setup ---------- */
function historySetup() {
  histSheet_();
  historyContinue();
}

/* ---------- reads emails in batches, picks up where it left off ---------- */
function historyContinue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  const started = Date.now();
  let finished = true;
  try {
    const sh = histSheet_();
    const seen = {}, have = {};
    histRows_().forEach(function (r) { seen[r.Key.split('#')[0]] = true; if (r.Ticket) have[r.Kind + r.Ticket] = true; });
    const out = [];
    let start = 0;
    outer:
    while (true) {
      const threads = GmailApp.search(HIST_QUERY + ' after:' + HIST_FROM, start, 100);
      if (!threads.length) break;
      for (let ti = 0; ti < threads.length; ti++) {
        const msgs = threads[ti].getMessages();
        for (let mi = 0; mi < msgs.length; mi++) {
          const msg = msgs[mi];
          const id = msg.getId();
          if (seen[id]) continue;
          if (Date.now() - started > 4.5 * 60 * 1000) { finished = false; break outer; }
          seen[id] = true;
          if (msg.getFrom().toLowerCase().indexOf('northsideplumbing.com') < 0) continue;
          const link = 'https://mail.google.com/mail/u/?authuser=teo@crowneliteplumbing.com#all/' + id;
          const atts = msg.getAttachments().filter(function (a) {
            return /\.pdf$/i.test(a.getName()) || a.getContentType() === 'application/pdf';
          });
          if (!atts.length) out.push([id + '#0', 'NONE', '', '', '', '', '', link, 'No PDF', stamp_()]);
          atts.forEach(function (att, ai) {
            const key = id + '#' + ai;
            try {
              const L = autoText_(att.getBytes());
              const all = L.join('\n');
              if (/S\s*T\s*A\s*T\s*E\s*M\s*E\s*N\s*T/.test(all) && !/PDF INVOICE/.test(all) && !/CM\s*#\s*\d/.test(all)) {
                out.push([key, 'STATEMENT', '', '', '', '', '', link, 'Monthly statement', stamp_()]);
                return;
              }
              const doc = autoParse_(L);
              const t = doc.ticket || {};
              if ((doc.kind !== 'INVOICE' && doc.kind !== 'CREDIT') || !t.no) {
                out.push([key, 'OTHER', '', '', '', '', '', link, 'Not a ticket', stamp_()]);
                return;
              }
              if (t.date && t.date < HIST_FIRST_TICKET) {
                out.push([key, 'EARLY', t.no, t.date, '', '', '', link, 'Before Dec 26', stamp_()]);
                return;
              }
              if (have[doc.kind + t.no]) {
                out.push([key, 'DUPLICATE', t.no, t.date, '', '', '', link, 'Same ticket emailed again', stamp_()]);
                return;
              }
              have[doc.kind + t.no] = true;
              const total = doc.kind === 'CREDIT' ? -Math.abs(t.total) : t.total;
              out.push([key, doc.kind, t.no, t.date, clean_(t.po), clean_(t.shipVia), String(total), link, '', stamp_()]);
            } catch (err) {
              out.push([key, 'UNREAD', '', '', '', '', '', link, 'Could not read: ' + err.message, stamp_()]);
            }
          });
          if (out.length >= 50) { histWrite_(sh, out.splice(0)); }
        }
      }
      if (threads.length < 100) break;
      start += 100;
    }
    histWrite_(sh, out);
  } finally {
    lock.releaseLock();
  }
  histStopTriggers_();
  if (!finished) ScriptApp.newTrigger('historyContinue').timeBased().after(60 * 1000).create();
}
function histWrite_(sh, rows) {
  if (!rows.length) return;
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, rows.length, HIST_HEAD.length).setNumberFormat('@').setValues(rows);
}

/* ---------- app action (owner PIN) ---------- */
ACTIONS.historyList = function (p) {
  owner_(p);
  const running = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'historyContinue'; });
  const rows = histRows_().filter(function (r) { return r.Kind === 'INVOICE' || r.Kind === 'CREDIT' || r.Kind === 'UNREAD'; })
    .map(function (r) { return { kind: r.Kind, ticket: r.Ticket, date: r.Date, po: r.PO, ship: r.ShipVia, total: r.Total, link: r.Link, note: r.Note }; });
  return { running: running, rows: rows };
};
