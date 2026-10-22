/**
 * Crown Elite Warehouse Inventory: automatic Northside tickets
 *
 * Every hour this reads new Northside emails (invoices and credit memos),
 * reads the attached PDF, and:
 *  - Invoice whose math checks out: added to inventory like "Add ticket" does.
 *  - Invoice whose math does not check out: held in Needs review (Owner tab).
 *  - Credit memo whose items are all found: logged as Outgoing (returned to Northside).
 *  - Credit memo with an item it can't find: held in Needs review.
 * Tickets already in the app are skipped, so nothing is ever added twice.
 *
 * This block sits at the end of Code.gs (after the starting data).
 * One time: pick "autoSetup" in the function menu and click Run.
 */

const AUTO_START = '2026/10/13';           // first email date to look at
const AUTO_QUERY = 'from:noreply@northsideplumbing.com has:attachment';
const AUTO_FOLDER = '1J0SAIZ6S98ze9-Ns-vZ1HtDDobjXn3qF'; // "Northside Invoices" Drive folder
HEAD.AutoLog = ['Key', 'Kind', 'Ticket', 'Date', 'Status', 'Reason', 'Total', 'Data', 'MessageId', 'At', 'By'];

/* ---------- one time setup ---------- */
function autoSetup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('AutoLog');
  if (!sh) sh = ss.insertSheet('AutoLog');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.AutoLog.length).setValues([HEAD.AutoLog]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'autoImportNorthside') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoImportNorthside').timeBased().everyHours(1).create();
  autoImportNorthside();
}

/* ---------- the hourly job ---------- */
function autoImportNorthside() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const seen = {};
    rows_('AutoLog').forEach(function (r) { seen[r.Key] = true; });
    const threads = GmailApp.search(AUTO_QUERY + ' after:' + AUTO_START, 0, 100);
    threads.forEach(function (th) {
      th.getMessages().forEach(function (msg) {
        if (msg.getFrom().toLowerCase().indexOf('northsideplumbing.com') < 0) return;
        msg.getAttachments().forEach(function (att, ai) {
          if (!/\.pdf$/i.test(att.getName()) && att.getContentType() !== 'application/pdf') return;
          const key = msg.getId() + '#' + ai;
          if (seen[key]) return;
          seen[key] = true;
          try {
            autoHandle_(key, msg, att);
          } catch (err) {
            autoLog_(key, 'UNKNOWN', '', '', 'REVIEW', 'Could not read this PDF: ' + err.message, '', null, msg.getId());
          }
        });
      });
    });
    autoFixLinks_();
  } finally {
    lock.releaseLock();
  }
}

function autoHandle_(key, msg, att) {
  const doc = autoParse_(autoText_(att.getBytes()));
  const t = doc.ticket;
  if (doc.kind === 'INVOICE') {
    if (rows_('Tickets').some(function (r) { return norm_(r.Ticket) === norm_(t.no); })) {
      return autoLog_(key, 'INVOICE', t.no, t.date, 'SKIPPED', 'Already in the app', t.total, null, msg.getId());
    }
    t.pdf = autoPdfLink_(t.no, msg.getId());
    const bad = autoCheck_(doc);
    if (bad) return autoLog_(key, 'INVOICE', t.no, t.date, 'REVIEW', bad, t.total, doc, msg.getId());
    ACTIONS.importTicket({ pin: getSetting_('OwnerPIN'), by: 'Auto', ticket: t, lines: doc.lines });
    return autoLog_(key, 'INVOICE', t.no, t.date, 'ADDED', doc.lines.length + ' lines', t.total, null, msg.getId());
  }
  if (doc.kind === 'CREDIT') {
    if (autoCreditDone_(t.no)) {
      return autoLog_(key, 'CREDIT', t.no, t.date, 'SKIPPED', 'Already taken out', t.total, null, msg.getId());
    }
    t.pdf = autoPdfLink_(t.no, msg.getId());
    const bad = autoCheck_(doc);
    const missing = [];
    doc.lines.forEach(function (l) { l.item = autoFindCreditItem_(l); if (!l.item) missing.push(l.desc); });
    if (bad || missing.length) {
      const why = bad || ('Item not found in inventory: ' + missing.join('; '));
      return autoLog_(key, 'CREDIT', t.no, t.date, 'REVIEW', why, t.total, doc, msg.getId());
    }
    autoCreditOut_(t, doc.lines, 'Auto');
    return autoLog_(key, 'CREDIT', t.no, t.date, 'ADDED', doc.lines.length + ' lines returned', t.total, null, msg.getId());
  }
  autoLog_(key, 'UNKNOWN', '', '', 'REVIEW', 'Not an invoice or credit memo', '', null, msg.getId());
}

/* ---------- reading the PDF ---------- */
// Northside PDFs are plain typed text (Courier), one line per "( ... )'" in the page streams.
function autoText_(bytes) {
  const raw = bytes.map(function (b) { return String.fromCharCode(b & 255); }).join('');
  const out = [];
  const re = /\(((?:\\.|[^\\)])*)\)'/g;
  let m;
  while ((m = re.exec(raw))) {
    out.push(m[1].replace(/\\([()\\])/g, '$1').replace(/\r/g, ''));
  }
  if (!out.length) throw new Error('no text in PDF (scanned or new format)');
  return out;
}

function autoNum_(s) { return Number(String(s || '').replace(/,/g, '')) || 0; }
function autoDate_(s) {
  const p = String(s || '').split('/');
  if (p.length !== 3) return '';
  const y = p[2].length === 2 ? '20' + p[2] : p[2];
  return y + '-' + ('0' + p[0]).slice(-2) + '-' + ('0' + p[1]).slice(-2);
}
function autoPick_(lines, re) {
  for (let i = 0; i < lines.length; i++) { const m = re.exec(lines[i]); if (m) return m[1].trim(); }
  return '';
}

function autoParse_(L) {
  const all = L.join('\n');
  const credit = /CREDIT MEMO/.test(all);
  const invoice = /PDF INVOICE/.test(all);
  const t = {
    no: autoPick_(L, credit ? /CM\s*#\s*(\d+)/ : /INVOICE\s*#\s*(\d+)/),
    date: autoDate_(autoPick_(L, /(?:INVOICE )?DATE\s+(\d{1,2}\/\d{1,2}\/\d{2,4})/)),
    po: autoPick_(L, /P\.?\s?O\.?\s*#\s*(.*?)\s*$/),
    shipVia: autoPick_(L, /SHIP VIA:\s*(.*?)\s*$/),
    pick: autoPick_(L, /PICK TI(?:C)?K\s*#\s*(\S+)/),
    subtotal: autoNum_(autoPick_(L, /SUB TOTAL\s+(-?[\d,]+\.\d\d)/)),
    tax: autoNum_(autoPick_(L, /SALES TAX\s+(-?[\d,]+\.\d\d)/)),
    freight: autoNum_(autoPick_(L, /\bFREIGHT\s+(-?[\d,]+\.\d\d)/)),
    total: autoNum_(autoPick_(L, credit ? /CREDIT MEMO TOTAL\s+(-?[\d,]+\.\d\d)/ : /INVOICE TOTAL\s+(-?[\d,]+\.\d\d)/))
  };
  const lines = [];
  if (invoice && !credit) {
    const re = /^\s*(-?\d[\d,]*(?:\.\d+)?)\s{2,}(\S.*?)\s+PER\s+([A-Z]+)\s+(-?[\d,]*\.\d+)\s+(-?[\d,]*\.\d+)\s*$/;
    L.forEach(function (s) {
      const m = re.exec(s);
      if (m) lines.push({ qty: autoNum_(m[1]), desc: clean_(m[2]), uom: m[3], price: autoNum_(m[4]), ext: autoNum_(m[5]) });
    });
    return { kind: 'INVOICE', ticket: t, lines: lines };
  }
  if (credit) {
    // Credit memos print each item on 2 lines: qty, code, price, ext / then the description and UOM.
    const re = /^\s*(-?\d[\d,]*(?:\.\d+)?)\s+(?:(-?\d+)\s+)?(\S+)\s+(-?[\d,]*\.\d+)\s+(-?[\d,]*\.\d+)\s*$/;
    for (let i = 0; i < L.length; i++) {
      const m = re.exec(L[i]);
      if (!m) continue;
      const next = String(L[i + 1] || '');
      const um = /^(.*?)\s{2,}([A-Z]{1,4})\s*$/.exec(next);
      lines.push({ qty: Math.abs(autoNum_(m[1])), code: m[3], desc: clean_(um ? um[1] : next), uom: um ? um[2] : 'EA',
        price: autoNum_(m[4]), ext: autoNum_(m[5]) });
    }
    return { kind: 'CREDIT', ticket: t, lines: lines };
  }
  return { kind: 'UNKNOWN', ticket: t, lines: [] };
}

function autoCheck_(doc) {
  const t = doc.ticket;
  if (!t.no) return 'Ticket number not found on the PDF';
  if (!t.date) return 'Ticket date not found on the PDF';
  if (!doc.lines.length) return 'No item lines found on the PDF';
  const sum = doc.lines.reduce(function (a, l) { return a + l.ext; }, 0);
  const sub = doc.kind === 'CREDIT' ? -Math.abs(t.subtotal) : t.subtotal;
  if (Math.abs(Math.abs(sum) - Math.abs(sub)) > 0.02) {
    return 'Lines add up to ' + Math.abs(sum).toFixed(2) + ' but the subtotal is ' + Math.abs(sub).toFixed(2);
  }
  if (Math.abs(t.freight) > 0.004) return 'Ticket has a freight charge of ' + t.freight.toFixed(2);
  if (Math.abs(t.subtotal + t.tax + t.freight - t.total) > 0.02) return 'Subtotal plus tax does not equal the total';
  return '';
}

/* ---------- credits ---------- */
function autoFindCreditItem_(l) {
  let item = findItem_(l.desc) || findItem_(l.code);
  if (item) return item;
  const code = norm_(l.code).replace(/^SC/, '');
  const hit = rows_('TicketLines').filter(function (r) {
    return code && norm_(r.NorthsideDescription).split(' ')[0] === code;
  })[0];
  return hit ? hit.Item : '';
}
function autoCreditDone_(no) {
  return rows_('Movements').some(function (r) { return r.Type === 'OUT' && norm_(r.Ticket) === norm_(no) && /credit memo/i.test(r.Note); });
}
function autoCreditOut_(t, lines, by) {
  const mv = lines.map(function (l) {
    return [id_(), t.date || today_(), 'OUT', l.item, l.qty, t.no, 'NORTHSIDE', '',
      'Returned to Northside, credit memo ' + t.no, by, stamp_()];
  });
  append_('Movements', mv);
}

/* ---------- PDF links ---------- */
function autoPdfLink_(no, msgId) {
  try {
    const it = DriveApp.getFolderById(AUTO_FOLDER).searchFiles("title contains '" + no + "'");
    if (it.hasNext()) return it.next().getUrl();
  } catch (e) {}
  return 'https://mail.google.com/mail/u/?authuser=teo@crowneliteplumbing.com#all/' + msgId;
}
// The Invoice Saver copies PDFs to Drive once a day; swap email links for the Drive copy once it is there.
function autoFixLinks_() {
  const sh = sheet_('Tickets');
  const last = sh.getLastRow();
  if (last < 2) return;
  const col = HEAD.Tickets.indexOf('PdfLink') + 1;
  const nos = sh.getRange(2, 1, last - 1, 1).getValues();
  const links = sh.getRange(2, col, last - 1, 1).getValues();
  links.forEach(function (v, i) {
    if (String(v[0]).indexOf('mail.google.com') < 0) return;
    const url = autoPdfLink_(String(nos[i][0]), '');
    if (url.indexOf('drive.google.com') >= 0) sh.getRange(i + 2, col).setNumberFormat('@').setValue(url);
  });
}

/* ---------- log ---------- */
function autoLog_(key, kind, no, date, status, reason, total, doc, msgId) {
  append_('AutoLog', [[key, kind, no, date, status, reason, total, doc ? JSON.stringify(doc) : '', msgId, stamp_(), 'Auto']]);
}
function autoSetStatus_(key, status, by) {
  const sh = sheet_('AutoLog');
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Nothing to update');
  const keys = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(key)) {
      sh.getRange(i + 2, HEAD.AutoLog.indexOf('Status') + 1).setValue(status);
      sh.getRange(i + 2, HEAD.AutoLog.indexOf('By') + 1).setValue(clean_(by));
      return;
    }
  }
  throw new Error('That entry was not found');
}

/* ---------- app actions (owner PIN) ---------- */
WRITES.push('autoResolve', 'autoCredit');

ACTIONS.autoLog = function (p) {
  owner_(p);
  return rows_('AutoLog').slice(-200).map(function (r) {
    let data = null;
    if (r.Status === 'REVIEW' && r.Data) { try { data = JSON.parse(r.Data); } catch (e) {} }
    return { key: r.Key, kind: r.Kind, ticket: r.Ticket, date: r.Date, status: r.Status, reason: r.Reason,
      total: r.Total, data: data, link: r.Status === 'REVIEW' && r.MessageId ? autoPdfLink_(r.Ticket || 'x', r.MessageId) : '', at: r.At };
  });
};
// Mark a review entry as handled (after adding it by hand, or to ignore it).
ACTIONS.autoResolve = function (p) {
  owner_(p);
  autoSetStatus_(p.key, p.status === 'DISMISSED' ? 'DISMISSED' : 'FIXED', p.by);
  return true;
};
// Credit memo from Needs review, after the owner picked the inventory item for each line.
ACTIONS.autoCredit = function (p) {
  owner_(p);
  const t = p.ticket || {};
  if (!clean_(t.no)) throw new Error('Credit memo number is missing');
  if (autoCreditDone_(t.no)) throw new Error('Credit memo ' + t.no + ' was already taken out');
  const lines = (p.lines || []).map(function (l) { return { item: mustItem_(l.item), qty: mustQty_(l.qty) }; });
  if (!lines.length) throw new Error('No lines to take out');
  autoCreditOut_({ no: clean_(t.no), date: clean_(t.date) }, lines, p.by);
  if (p.key) autoSetStatus_(p.key, 'FIXED', p.by);
  return { lines: lines.length };
};
// "Check email now" button.
ACTIONS.autoRun = function (p) {
  owner_(p);
  autoImportNorthside();
  return true;
};

/* ---------- outgoing keeps the date picked in the app ---------- */
ACTIONS.stockOut = function (p) {
  const item = mustItem_(p.item);
  const qty = mustQty_(p.qty);
  if (!clean_(p.person)) throw new Error('Pick who is taking it');
  const person = ensure_('People', p.person);
  const project = clean_(p.project) ? ensure_('Projects', p.project) : '';
  const date = /^\d{4}-\d{2}-\d{2}$/.test(clean_(p.date)) ? clean_(p.date) : today_();
  append_('Movements', [[id_(), date, 'OUT', item, qty, clean_(p.ticket), person, project, clean_(p.note), clean_(p.by), stamp_()]]);
  return { item: item, onHand: round_(onHand_()[item] || 0) };
};

/* ---------- dry run: reads emails, changes nothing, writes the results to the log view ---------- */
function autoTest() {
  const threads = GmailApp.search(AUTO_QUERY + ' after:' + AUTO_START, 0, 100);
  threads.forEach(function (th) {
    th.getMessages().forEach(function (msg) {
      msg.getAttachments().forEach(function (att) {
        try {
          const doc = autoParse_(autoText_(att.getBytes()));
          Logger.log(doc.kind + ' ' + doc.ticket.no + ' ' + doc.ticket.date + ' lines=' + doc.lines.length +
            ' total=' + doc.ticket.total + ' check=' + (autoCheck_(doc) || 'OK'));
        } catch (e) { Logger.log('ERROR ' + att.getName() + ': ' + e.message); }
      });
    });
  });
}
