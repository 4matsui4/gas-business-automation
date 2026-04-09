/**** Configuration ****/
/* Column assignments:
   D  = Organisation name
   AB = 申込書A 完了日（締切1）
   AC = 申込書B 完了日（締切2）
   AD = 申込書A 完了状況（完了語で判定）
   AE = 申込書B 完了状況（完了語で判定）
   AF = Per-organisation input sheet URL (submission management book URL)
   AG = Data storage drive URL (parent folder URL)
   AH~AQ = Contact emails (variable number of recipients)
*/
const CFG = {
  MASTER_TAB_NAME:   'Master(完成版)',
  MASTER_SHEET_NAME: 'Master(完成版)',

  // Per-organisation book common settings
  ORGANISATION_TAB_NAME:   '提出管理シート',
  FORM_A_TAB_NAME:   '申込書A',

  // Master column definitions
  MASTER_COL: {
    ORGANISATION:    'D',
    FORM1_DEADLINE:  'AB',  // 申込書A 完了日（締切1）
    FORM2_DEADLINE:  'AC',  // 申込書B 完了日（締切2）
    STATUS1:         'AD',  // 申込書A 完了状況
    STATUS2:         'AE',  // 申込書B 完了状況
    ORGANISATION_BOOK_URL: 'AF',  // Per-organisation submission management book URL
    FOLDER_URL:      'AG',  // Parent folder URL
    MAIL_FROM:       'AH',  // Mail TO start (first)
    MAIL_TO:         'AJ',  // Mail CC end (contact list ③)
    REMIND_LOG1:     'AU',  // Form 1 reminder send log
    REMIND_LOG2:     'AV'   // Form 2 reminder send log
  },

  // Per-organisation submission management sheet (fixed)
  ORGANISATION_COL: {
    NAME:         'A',
    DOC_URL:      'B',  // URL / URL|Sheet!Range
    DUE:          'C',
    STATUS:       'D',
    DONE_DATE:    'E',
    PDF_URL:      'F',
    LAST_UPDATED: 'G'
  },

  // Permission role
  PERMISSION_ROLE: 'editor',

  // Reminder hour (14 for production; use separate trigger for testing)
  REMIND_HOUR: 14
};

const K = {
  MANAGED_PREFIX: 'managed:'  // Record prefix for "already notified by this script"
};

// Optional: set a fixed organisation book ID to process (leave empty for normal operation)
const ORGANISATION_BOOK_ID = '';

// ===== Admin Master link (used for Slack button) =====
const MASTER_SHEET_LINK = 'YOUR_MASTER_SHEET_LINK';

// ===== Notification timing =====
// Organisation email: pre-deadline [14,7,5,3] / overdue [-1,-2,-3,-4]
// Admin Slack: pre-deadline [5,3,1] (batch) / overdue [-1,-2,-3,-4] (per organisation)
const PRE_EMAIL_OFFSETS   = [14, 7, 5, 3];
const POST_EMAIL_OFFSETS  = [-1, -2, -3, -4];
const PRE_ADMIN_OFFSETS   = [5, 3, 1];
const POST_ADMIN_OFFSETS  = [-1, -2, -3, -4];

// Compatible syntax for Rhino engine
CFG.REMIND_OFFSETS = (function () {
  var out = []
    .concat(PRE_EMAIL_OFFSETS, POST_EMAIL_OFFSETS, PRE_ADMIN_OFFSETS, POST_ADMIN_OFFSETS);
  var seen = {};
  var res = [];
  for (var i = 0; i < out.length; i++) {
    var k = String(out[i]);
    if (!seen[k]) { seen[k] = 1; res.push(out[i]); }
  }
  return res;
})();

// ===== Admin email (optional; leave as empty arrays if not needed) =====
const ADMIN_EMAIL = {
  TO: [],
  CC: []
};

/*** Slack ***/
const SLACK = {
  WEBHOOK_URL: '',       // Set your Slack Incoming Webhook URL here
  DEFAULT_CHANNEL: ''
};

/*** Completion keywords / Stop keywords ***/
// Detected as "complete" if status contains any of these words (partial match)
const COMPLETE_WORDS = ['完', '完了', '済', '〆', '終了', '提出済'];
// Reminder stops if status contains any of these words (partial match)
const STOP_STATUSES  = ['停止', 'キャンセル', '中止', '通知停止'];

/*** Reminder settings (4 days before / 1 day before, runs at 7:00) ***/
const REMINDER_OFFSETS = [4, 1];
const ENABLE_OVERDUE   = false;          // Overdue emails OFF (scaffold kept for future use)
const OVERDUE_OFFSETS  = [-1, -2, -3, -4];
const REMIND_LOG_COL   = { form1: 'AU', form2: 'AV' };

try { CFG.REMIND_HOUR = 7; } catch (_) {}

/**** Menu ****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('運用メニュー')
    .addSubMenu(
      ui.createMenu('権限')
        .addItem('全行の権限同期（マスター）', 'syncAllPermissions')
        .addSeparator()
        .addItem('権限付与の定期同期（5分）', 'createMailSyncTrigger')
        .addItem('権限付与の定期同期（停止）', 'deleteMailSyncTriggers')
        .addSeparator()
        .addItem('権限付与者へ一斉メール（未完のみ）', 'sendGrantMailsBulk')
        .addItem('権限付与者へメール（この行）', 'sendGrantMailForActiveRow')
    )
    .addSubMenu(
      ui.createMenu('提出管理')
        .addItem('PDF化してF/G列に反映（単一ブック／このブック）', 'exportAllDocsToPdf')
        .addItem('（この行だけ）PDF化して各組織シートに反映', 'exportSelectedOrganisationDocsToPdf')
        .addItem('（全組織）PDF化（スマート：不足・更新のみ）', 'exportAllOrganisationsDocsSmart_')
        .addItem('（全組織）PDF化（強制フル再生成）', 'exportAllOrganisationsDocsForce_')
    )
    .addSubMenu(
      ui.createMenu('リマインド')
        .addItem('時間トリガー作成（毎日）', 'createDailyTrigger')
        .addItem('時間トリガー削除', 'deleteDailyTriggers')
        .addItem('完了検知トリガー作成（15分）', 'createCompletionTrigger')
        .addItem('完了検知トリガー削除', 'deleteCompletionTriggers')
    )
    .addSeparator()
    .addItem('初期設定（認可）', 'firstSetup')
    .addToUi();
}

function firstSetup() {
  SpreadsheetApp.getActiveSpreadsheet().getSheets();
  Browser.msgBox('初期設定OK。必要に応じて「時間トリガー作成」を実行してください。');
}

// Periodic sync trigger for email columns (via IMPORTRANGE)
function createMailSyncTrigger() {
  deleteMailSyncTriggers();
  ScriptApp.newTrigger('syncAllPermissions').timeBased().everyMinutes(5).create();
  Browser.msgBox('5分ごとの権限同期トリガーを作成しました');
}
function deleteMailSyncTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(function(t){ return t.getHandlerFunction() === 'syncAllPermissions'; })
    .forEach(function(t){ ScriptApp.deleteTrigger(t); });
  Browser.msgBox('権限同期トリガーを削除しました');
}

// Master spreadsheet ID — replace with your own
const MASTER_SPREADSHEET_ID = 'YOUR_MASTER_SPREADSHEET_ID';

function getMasterTab_() {
  const ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const sh = ss.getSheetByName(CFG.MASTER_TAB_NAME);
  if (!sh) throw new Error('マスタータブが見つかりません（Master(完成版)）');
  return sh;
}

/**** Rate limit mitigation settings ****/
const SYNC_CAP = {
  MAX_ADDS_PER_RUN: 20,
  MIN_SLEEP_MS: 600,
  RETRY_MAX: 6
};

function sleepMs_(ms) {
  const jitter = Math.floor(Math.random() * 60);
  Utilities.sleep(ms + jitter);
}

function getEditorsWithRetry_(folder) {
  let wait = 300;
  for (let i = 0; i <= SYNC_CAP.RETRY_MAX; i++) {
    try {
      return folder.getEditors().map(u => (u.getEmail() || '').toLowerCase());
    } catch (e) {
      const msg = String(e);
      if (i === SYNC_CAP.RETRY_MAX) throw new Error(`getEditors retry exceeded | ${msg}`);
      sleepMs_(wait); wait *= 2;
    }
  }
  return [];
}

function addEditorWithRetry_(folder, email) {
  let wait = 400;
  for (let i = 0; i <= SYNC_CAP.RETRY_MAX; i++) {
    try {
      if (CFG.PERMISSION_ROLE === 'viewer') folder.addViewer(email);
      else folder.addEditor(email);
      return { ok: true };
    } catch (e) {
      const msg = String(e);
      const transient = /Limit Exceeded|rate|quota|internal|backend|temporar/i.test(msg);
      if (!transient || i === SYNC_CAP.RETRY_MAX) return { ok: false, err: msg };
      sleepMs_(wait); wait *= 2;
    }
  }
  return { ok: false, err: 'unknown' };
}

function extractFolderIdFromUrl_(url) {
  const m = (url || '').match(/https?:\/\/drive\.google\.com\/drive\/folders\/([-\w]{25,})/);
  return m ? m[1] : '';
}

/**** onEdit (optional log & Form 1 helper) ****/
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName(), row = e.range.getRow(), col = e.range.getColumn();

  if (name === '編集ログ') return;
  if (name === CFG.FORM_A_TAB_NAME && (col === 2 || col === 3)) return;

  if (name === '入力シート' && col === 1 && row >= 2) {
    sh.getRange(row, 2).setValue(new Date()); return;
  }

  if (name === CFG.YOSHIKI1_TAB_NAME && col === 1 && row >= 2) {
    const now = new Date();
    const actorDisplay = tryGetLastActorDisplayName_(SpreadsheetApp.getActive().getId()) || '';
    const fallbackUser = (sh.getRange('C5').getValue() || '').toString().trim();
    sh.getRange(row, 2).setValue(now);
    sh.getRange(row, 3).setValue(actorDisplay || fallbackUser || '-');

    const mgmt = SpreadsheetApp.getActive().getSheetByName(CFG.ORGANISATION_TAB_NAME);
    if (mgmt) mgmt.getRange(`${CFG.ORGANISATION_COL.LAST_UPDATED}2`).setValue(now);
    return;
  }

  try {
    const log = getOrCreateLogSheet_();
    const ts = new Date();
    const a1 = e.range.getA1Notation();
    const oldV = (typeof e.oldValue !== 'undefined') ? e.oldValue : '';
    const newV = (typeof e.value    !== 'undefined') ? e.value    : '';
    const actor = tryGetLastActorDisplayName_(SpreadsheetApp.getActive().getId()) || '-';
    const rows = e.range.getNumRows(), cols = e.range.getNumColumns();
    const multi = (rows > 1 || cols > 1) ? `(${rows}x${cols})` : '';
    log.appendRow([ts, actor, name, a1 + multi, oldV, newV]);
  } catch (lgErr) { console.warn('logErr:', lgErr); }
}

/**** Permission sync entry point (per-row, capped, lock-protected) ****/
function syncAllPermissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const sh = getMasterTab_();
    const last = sh.getLastRow();

    const props = PropertiesService.getScriptProperties();
    const START_ROW_KEY = 'perm_sync_next_row';
    let r = parseInt(props.getProperty(START_ROW_KEY) || '2', 10);

    let adds = 0;
    while (r <= last) {
      const budget = SYNC_CAP.MAX_ADDS_PER_RUN - adds;
      if (budget <= 0) break;

      try {
        const made = syncPermForRow_capped_(sh, r, budget);
        adds += made;
      } catch (e) {
        setSendMeta_(sh, r, `${fmtNow_()} | rowErr | ${String(e)}`);
        sleepMs_(120);
      }
      r++;
    }

    props.setProperty(START_ROW_KEY, (r > last ? 2 : r).toString());
    SpreadsheetApp.getActive().toast(`権限同期 追加:${adds}件 / 次回開始行:${(r>last?2:r)}`);
  } finally {
    lock.releaseLock();
  }
}

/**** Per-row permission sync with cap and quarantine support ****/
function syncPermForRow_capped_(sh, r, budget) {
  if (budget <= 0) return 0;

  const folderUrl = sh.getRange(`${CFG.MASTER_COL.FOLDER_URL}${r}`).getDisplayValue();
  if (!folderUrl) { setSendMeta_(sh, r, `${fmtNow_()} | noFolderUrl`); return 0; }
  const folderId = extractFolderIdFromUrl_(folderUrl);
  if (!folderId) { setSendMeta_(sh, r, `${fmtNow_()} | badFolderUrl`); return 0; }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    setSendMeta_(sh, r, `${fmtNow_()} | getFolderByIdErr | ${String(e)}`);
    return 0;
  }

  let currentEditors = [];
  try {
    currentEditors = getEditorsWithRetry_(folder).map(e => e.toLowerCase());
  } catch (e) {
    setSendMeta_(sh, r, `${fmtNow_()} | getEditorsErr | ${String(e)}`);
    return 0;
  }

  const emails = readRowEmailsDisplay_(sh, r).map(e => e.toLowerCase());
  if (emails.length === 0) return 0;

  const props = PropertiesService.getScriptProperties();
  const managedKey = K.MANAGED_PREFIX + folderId;
  const managed = new Set(JSON.parse(props.getProperty(managedKey) || '[]').map(e => e.toLowerCase()));
  const target  = new Set(emails);

  const toNotifyBase = emails.filter(e => !managed.has(e));
  const quarantined  = loadQuarantineSet_(folderId);

  const toAddAll = emails
    .filter(e => !currentEditors.includes(e))
    .filter(e => !quarantined.has(e));

  const addFailures = [];
  let added = 0;
  for (const addr of toAddAll) {
    if (added >= budget) break;
    const res = addEditorWithRetry_(folder, addr);
    if (res.ok) {
      added++;
      sleepMs_(SYNC_CAP.MIN_SLEEP_MS);
    } else {
      const kind = classifyAddError_(res.err);
      addFailures.push({ email: addr, err: res.err, kind });
      if (kind === 'permanent') {
        quarantined.add(addr);
      } else if (/Limit Exceeded|rate|quota/i.test(res.err)) {
        sleepMs_(400);
      }
    }
  }

  saveQuarantineSet_(folderId, quarantined);

  const noticeList = toNotifyBase.filter(e => !addFailures.some(f => f.email === e));
  if (noticeList.length > 0) {
    try {
      const organisation = sh.getRange(`${CFG.MASTER_COL.ORGANISATION}${r}`).getDisplayValue() || '(組織名未設定)';
      const bookUrl      = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
      const form1DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM1_DEADLINE}${r}`).getDisplayValue() || '').toString();
      const form2DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM2_DEADLINE}${r}`).getDisplayValue() || '').toString();

      sendGrantNoticeCapped_(
        noticeList,
        { organisation, folderUrl, bookUrl, form1DueDisp, form2DueDisp },
        sh, r
      );
      noticeList.forEach(e => managed.add(e));
      props.setProperty(managedKey, JSON.stringify(Array.from(managed)));
    } catch (e) {
      setSendMeta_(sh, r, `${fmtNow_()} | grantNoticeWarn | ${String(e)}`);
    }
  }

  const toRemove = Array.from(managed).filter(e => !target.has(e));
  for (const addr of toRemove) {
    try { folder.removeEditor(addr); } catch (_) {}
    managed.delete(addr);
  }
  props.setProperty(managedKey, JSON.stringify(Array.from(managed)));

  if (addFailures.length) {
    const reasons = addFailures.map(f => `${f.email} -> ${f.err}`).join(' ; ');
    setSendMeta_(sh, r, `${fmtNow_()} | addFail:${addFailures.length} | ${reasons}`);
  }

  return added;
}

/**** Mail send cap (per run) + daily block ****/
const MAIL_CAP = { MAX_MAILS_PER_RUN: 20 };
var __mailSentThisRun = 0;

function isEmailBlocked_() {
  const props = PropertiesService.getScriptProperties();
  const until = parseInt(props.getProperty('mail_block_until') || '0', 10);
  return Date.now() < until;
}

function markEmailQuotaExceeded_(sh, row, note) {
  const props = PropertiesService.getScriptProperties();
  const until = Date.now() + 25 * 60 * 60 * 1000;
  props.setProperty('mail_block_until', String(until));
  if (sh && row) setSendMeta_(sh, row, `${fmtNow_()} | emailQuotaExceeded | ${note || ''}`);
}

function sendGrantNoticeCapped_(recipients, ctx, sh, row) {
  if (isEmailBlocked_()) {
    if (sh && row) setSendMeta_(sh, row, `${fmtNow_()} | emailSkip | blocked(today)`);
    return { sent: false, skipped: true, reason: 'blocked' };
  }
  if (__mailSentThisRun >= MAIL_CAP.MAX_MAILS_PER_RUN) {
    if (sh && row) setSendMeta_(sh, row, `${fmtNow_()} | emailSkip | runCap`);
    return { sent: false, skipped: true, reason: 'runCap' };
  }

  try {
    sendGrantNotice_(recipients, ctx);
    __mailSentThisRun++;
    return { sent: true };
  } catch (e) {
    const msg = String(e);
    if (/Service invoked too many times for one day: email/i.test(msg)) {
      markEmailQuotaExceeded_(sh, row, msg);
      return { sent: false, skipped: true, reason: 'dailyQuota' };
    }
    if (sh && row) setSendMeta_(sh, row, `${fmtNow_()} | grantNoticeWarn | ${msg}`);
    return { sent: false, skipped: true, reason: 'other' };
  }
}

const DIAG_THROW_ON_ERROR = true;

/**** Diagnostic version of per-row permission sync ****/
function syncPermForRow_(sh, r) {
  const folderUrl = sh.getRange(`${CFG.MASTER_COL.FOLDER_URL}${r}`).getDisplayValue();
  if (!folderUrl) {
    setSendMeta_(sh, r, `${fmtNow_()} | noFolderUrl`);
    if (DIAG_THROW_ON_ERROR) throw new Error(`r=${r} | no folderUrl`);
    return;
  }

  const folderId = extractFolderIdFromUrl_(folderUrl);
  if (!folderId) {
    setSendMeta_(sh, r, `${fmtNow_()} | badFolderUrl | ${folderUrl}`);
    if (DIAG_THROW_ON_ERROR) throw new Error(`r=${r} | bad folderId from ${folderUrl}`);
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    setSendMeta_(sh, r, `${fmtNow_()} | getFolderByIdErr | id=${folderId} | ${String(e)}`);
    if (DIAG_THROW_ON_ERROR) throw new Error(`r=${r} | getFolderById failed | id=${folderId} | ${e}`);
    return;
  }

  let currentEditors = [];
  try {
    currentEditors = folder.getEditors().map(u => (u.getEmail() || '').toLowerCase());
  } catch (e) {
    setSendMeta_(sh, r, `${fmtNow_()} | getEditorsErr | id=${folderId} | ${String(e)}`);
    if (DIAG_THROW_ON_ERROR) throw new Error(`r=${r} | getEditors failed | id=${folderId} | ${e}`);
    return;
  }

  const emails = readRowEmailsDisplay_(sh, r).map(e => e.toLowerCase());
  if (emails.length === 0) return;

  const props = PropertiesService.getScriptProperties();
  const managedKey = K.MANAGED_PREFIX + folderId;
  const managed = new Set(JSON.parse(props.getProperty(managedKey) || '[]').map(e => e.toLowerCase()));
  const target  = new Set(emails);

  const toNotify = emails.filter(e => !managed.has(e));
  const toAdd    = emails.filter(e => !currentEditors.includes(e));

  const addFailures = [];
  for (const addr of toAdd) {
    try {
      CFG.PERMISSION_ROLE === 'viewer' ? folder.addViewer(addr) : folder.addEditor(addr);
      Utilities.sleep(120);
    } catch (err) {
      addFailures.push({ email: addr, err: String(err) });
    }
  }

  const noticeList = toNotify.filter(e => !addFailures.some(f => f.email === e));

  if (noticeList.length > 0) {
    try {
      const organisation = sh.getRange(`${CFG.MASTER_COL.ORGANISATION}${r}`).getDisplayValue() || '(組織名未設定)';
      const bookUrl      = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
      const form1DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM1_DEADLINE}${r}`).getDisplayValue() || '').toString();
      const form2DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM2_DEADLINE}${r}`).getDisplayValue() || '').toString();

      sendGrantNotice_(noticeList, { organisation, folderUrl, bookUrl, form1DueDisp, form2DueDisp });

      noticeList.forEach(e => managed.add(e));
      props.setProperty(managedKey, JSON.stringify(Array.from(managed)));
    } catch (e) {
      setSendMeta_(sh, r, `${fmtNow_()} | grantNoticeWarn | ${String(e)}`);
      console.warn('grant notice failed:', e);
    }
  }

  const toRemove = Array.from(managed).filter(e => !target.has(e));
  for (const addr of toRemove) {
    try { folder.removeEditor(addr); } catch (_) {}
    managed.delete(addr);
  }
  props.setProperty(managedKey, JSON.stringify(Array.from(managed)));

  if (addFailures.length) {
    const reasons = addFailures.map(f => `${f.email} -> ${f.err}`).join(' ; ');
    setSendMeta_(sh, r, `${fmtNow_()} | addFail:${addFailures.length} | ${reasons}`);
  }
}

/**** Drive Activity (optional) ****/
function tryGetLastActorDisplayName_(fileId) {
  try {
    if (!DriveActivity) return '';
    const request = { itemName: `items/${fileId}`, pageSize: 1, filter: 'detail.action_detail_case:(EDIT)' };
    const resp = DriveActivity.Activity.query(request);
    const acts = resp.activities || [];
    if (!acts.length) return '';
    const actors = acts[0].actors || [];
    if (!actors.length) return '';
    const p = actors[0].user?.knownUser?.personName || '';
    return actors[0].user.knownUser.isCurrentUser ? '自分' : (p || '');
  } catch (e) { return ''; }
}

/**** PDF Export ****/
function exportAllDocsToPdf() {
  const targetSS = ORGANISATION_BOOK_ID ? SpreadsheetApp.openById(ORGANISATION_BOOK_ID) : SpreadsheetApp.getActive();
  processOneOrganisationSheet_(targetSS);
  SpreadsheetApp.getActive().toast('PDF化を完了しました');
}

function exportSelectedOrganisationDocsToPdf() {
  const sh = getMasterTab_();
  const a = sh.getActiveCell();
  if (!a) { SpreadsheetApp.getActive().toast('セルを選択してください'); return; }
  const r = a.getRow();
  if (r < 2) { SpreadsheetApp.getActive().toast('データ行を選択してください'); return; }

  const url = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
  if (!url) { SpreadsheetApp.getActive().toast('この行の各組織ブックURL（AF）が空です'); return; }

  const id = extractId_(url);
  if (!id) { SpreadsheetApp.getActive().toast('各組織ブックURLからIDを取得できませんでした'); return; }

  const ss = SpreadsheetApp.openById(id);
  processOneOrganisationSheet_(ss);
  SpreadsheetApp.getActive().toast(`PDF化完了：${ss.getName()}`);
}

function processOneOrganisationSheet_(targetSS) {
  const tab = targetSS.getSheetByName(CFG.ORGANISATION_TAB_NAME);
  if (!tab) throw new Error(`提出管理シートが見つかりません（${targetSS.getName()}）`);

  const pIt = DriveApp.getFileById(targetSS.getId()).getParents();
  if (!pIt.hasNext()) throw new Error(`親フォルダが見つかりません（${targetSS.getName()}）`);
  const parentFolder = pIt.next();

  const lastRow = tab.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const cell = tab.getRange(`${CFG.ORGANISATION_COL.DOC_URL}${r}`);
    const cellText = (cell.getDisplayValue() || '').toString().trim();
    const cellUrl  = getUrlFromCell_(cell);
    const target   = (cellUrl || cellText).trim();
    if (!target) continue;

    const { url, rangeA1 } = splitUrlRange_(target);
    const docId = extractId_(url);
    if (!docId) continue;

    deleteExistingPdfIfAny_(tab, r);

    const mime = DriveApp.getFileById(docId).getMimeType();
    const pdfBlob = makePdfBlobUniversal_(docId, mime, rangeA1);

    const organisationName = targetSS.getName();
    const formLabel  = (tab.getRange(`${CFG.ORGANISATION_COL.NAME}${r}`).getDisplayValue() || '提出書類');
    const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const safe = s => (s || '').replace(/[\\/:*?"<>|]/g, '');
    const name = `${safe(organisationName)}_${safe(formLabel)}_${ymd}.pdf`;

    const file = parentFolder.createFile(pdfBlob).setName(name);
    tab.getRange(`${CFG.ORGANISATION_COL.PDF_URL}${r}`).setValue(file.getUrl());
    tab.getRange(`${CFG.ORGANISATION_COL.LAST_UPDATED}${r}`).setValue(new Date());
  }
}

function deleteExistingPdfIfAny_(tab, row) {
  const url = (tab.getRange(`${CFG.ORGANISATION_COL.PDF_URL}${row}`).getDisplayValue() || '').toString().trim();
  const id = extractId_(url);
  if (!id) return;
  try {
    DriveApp.getFileById(id).setTrashed(true);
  } catch (e) {
    console.warn('既存PDFの削除に失敗:', e);
  }
}

function makePdfBlobUniversal_(fileId, mimeType, rangeA1) {
  const token = ScriptApp.getOAuthToken();
  const M = {
    SHEET: 'application/vnd.google-apps.spreadsheet',
    DOC:   'application/vnd.google-apps.document',
    SLIDE: 'application/vnd.google-apps.presentation'
  };

  let url = '';
  if (mimeType === M.SHEET) {
    const sss = SpreadsheetApp.openById(fileId);
    const firstGid = sss.getSheets()[0].getSheetId();
    url = `https://docs.google.com/spreadsheets/d/${fileId}/export?gid=${firstGid}&format=pdf&portrait=true&fitw=true&size=A4&gridlines=false&printtitle=false`;
    if (rangeA1) url += `&range=${encodeURIComponent(rangeA1)}`;
  } else if (mimeType === M.DOC) {
    url = `https://docs.google.com/document/d/${fileId}/export?format=pdf`;
  } else if (mimeType === M.SLIDE) {
    url = `https://docs.google.com/presentation/d/${fileId}/export/pdf`;
  } else {
    throw new Error('未対応のファイル種別です: ' + mimeType);
  }

  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true
  });
  return resp.getBlob().setContentType('application/pdf').setName('tmp.pdf');
}

function getUrlFromCell_(range) {
  const formula = range.getFormula();
  if (formula && /^=HYPERLINK\(/i.test(formula)) {
    const m = formula.match(/^=HYPERLINK\(\s*"([^"]+)"/i);
    if (m && m[1]) return m[1];
  }
  const rtv = range.getRichTextValue && range.getRichTextValue();
  if (rtv) {
    const u1 = rtv.getLinkUrl && rtv.getLinkUrl();
    if (u1) return u1;
    const runs = rtv.getRuns ? rtv.getRuns() : [];
    for (const run of runs) {
      const u = run.getLinkUrl && run.getLinkUrl();
      if (u) return u;
    }
  }
  const disp = range.getDisplayValue();
  if (/^https?:\/\/docs\.google\.com\//.test(disp)) return disp;
  return '';
}

function splitUrlRange_(text) {
  const [url, range] = (text || '').split('|');
  return { url: (url || '').trim(), rangeA1: (range || '').trim() || null };
}

/**** Daily Reminder Trigger ****/
function createDailyTrigger() {
  deleteDailyTriggers();
  ScriptApp.newTrigger('dailyReminder_').timeBased().everyDays(1).atHour(CFG.REMIND_HOUR).create();
  Browser.msgBox(`毎日${CFG.REMIND_HOUR}時のリマインドを設定しました`);
}
function deleteDailyTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyReminder_')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

const REMINDER_CAP = { MAX_SENDS_PER_RUN: 40, SLEEP_MS: 120 };

function dailyReminder_() {
  const sh = getMasterTab_();
  const last = sh.getLastRow();
  const todayYmd = toYmd_(new Date());
  const PRE_OFFSETS = [4, 1];

  const FORMS = [
    { key: 'form1', name: '申込書A', deadlineCol: CFG.MASTER_COL.FORM1_DEADLINE, statusCol: CFG.MASTER_COL.STATUS1, logCol: CFG.MASTER_COL.REMIND_LOG1 },
    { key: 'form2', name: '申込書B', deadlineCol: CFG.MASTER_COL.FORM2_DEADLINE, statusCol: CFG.MASTER_COL.STATUS2, logCol: CFG.MASTER_COL.REMIND_LOG2 }
  ];

  const PS = PropertiesService.getScriptProperties();
  const KEY = 'reminder_next_row';
  let r = Math.max(2, parseInt(PS.getProperty(KEY) || '2', 10));

  let sent = 0;
  let scanned = 0;

  while (r <= last && sent < REMINDER_CAP.MAX_SENDS_PER_RUN) {
    scanned++;

    const organisation = sh.getRange(`${CFG.MASTER_COL.ORGANISATION}${r}`).getDisplayValue() || '(組織名未設定)';
    const bookUrl = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
    const emails  = readRowEmails_(sh, r);

    if (emails.length) {
      for (const f of FORMS) {
        if (sent >= REMINDER_CAP.MAX_SENDS_PER_RUN) break;

        const dueVal = sh.getRange(`${f.deadlineCol}${r}`).getValue();
        const status = (sh.getRange(`${f.statusCol}${r}`).getDisplayValue() || '').toString();
        if (!dueVal) continue;

        const dueYmd = toYmd_(new Date(dueVal));
        if (!dueYmd) continue;

        const diff   = daysDiff_(todayYmd, dueYmd);
        const isStop = STOP_STATUSES.some(w => status.includes(w));
        const isDone = isDoneMark_(status);
        if (isDone || isStop) continue;

        if (PRE_OFFSETS.includes(diff)) {
          if (alreadySentAtOffset_(sh, r, f.logCol, diff)) continue;

          const subj = `【YOUR_PROGRAM_NAME】【${organisation}】${f.name} リマインド（締切まで残り${diff}日）`;
          const html =
            `${organisation} ご担当者さま<br><br>` +
            `標記の件につきまして、進捗状況はいかがでしょうか。<br>` +
            `本メールと行き違いで既にご対応済みの場合は、何卒ご容赦ください。<br>` +
            `未完了の方におかれましては、<b>【${dueYmd}】</b>までに作業をお願いいたします。<br><br>` +
            `対象申込書：${f.name}<br>` +
            `締切日：${dueYmd}<br>` +
            `現在の状況：${status || '-'}<br>` +
            `締切まで <b>残り ${diff}日</b> です。<br><br>` +
            `▼管理シート（共有ドライブへアクセス可能な方）<br>` +
            `${bookUrl ? `<a href="${bookUrl}">${bookUrl}</a>` : '(未設定)'}<br>` +
            `※直接入力の上、ステータスの「完了」への変更をお願いします。<br><br>` +
            `【提出方法について：共有ドライブへアクセスできない方】<br>` +
            `エクセルファイルに必要事項をご記入の上、以下のメールアドレス宛にご提出ください。<br>` +
            `（エクセルファイルは参画校サイトからもダウンロードできます。）<br>` +
            `【エクセル提出先】 <a href="mailto:info@your-domain.jp">info@your-domain.jp</a><br><br>` +
            `お忙しい折、お手数をおかけしますが、ご協力のほどよろしくお願い申し上げます。<br><br>` +
            `※本メールは自動送信です。ご不明点は事務局担当者へご連絡ください。`;

          try {
            MailApp.sendEmail({
              to: emails[0],
              cc: emails.slice(1).join(','),
              subject: subj,
              htmlBody: html
            });
            writeRemindLog_(sh, r, f.logCol, `d-${diff}`);
            sent++;
            Utilities.sleep(REMINDER_CAP.SLEEP_MS);
          } catch (e) {
            setSendMeta_(sh, r, `${fmtNow_()} | remindMailErr | ${String(e)}`);
          }
        }
      }
    }

    r++;
  }

  PS.setProperty(KEY, (r > last ? 2 : r).toString());
  SpreadsheetApp.getActive().toast(`リマインド: 送信${sent}件 / 走査${scanned}行 / 次回開始行${(r>last?2:r)}`);
}

function appendReminderLog_(sh, row, colLetter, text) {
  const colIndex = colToIndex_(colLetter);
  const cell = sh.getRange(row, colIndex);
  const current = (cell.getDisplayValue() || '').toString();
  const next = current ? (current + '\n' + text) : text;
  cell.setValue(next);
}

function dateAddDays_(dateObj, days) {
  const d = new Date(dateObj);
  d.setDate(d.getDate() + days);
  return d;
}

/**** Completion Detection + Slack Notification ****/
function createCompletionTrigger() {
  deleteCompletionTriggers();
  ScriptApp.newTrigger('checkCompletionAndNotify_').timeBased().everyMinutes(15).create();
  Browser.msgBox('完了検知トリガー（15分ごと）を作成しました');
}

function deleteCompletionTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkCompletionAndNotify_')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function sendSlack_(text, blocks) {
  if (!SLACK.WEBHOOK_URL) return;
  UrlFetchApp.fetch(SLACK.WEBHOOK_URL, {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ text, ...(SLACK.DEFAULT_CHANNEL ? { channel: SLACK.DEFAULT_CHANNEL } : {}), ...(blocks ? { blocks } : {}) }),
    muteHttpExceptions: true
  });
}

function exportOneRowByMaster_(row) {
  const sh = getMasterTab_();
  const url = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${row}`).getDisplayValue() || '').toString().trim();
  const id = extractId_(url); if (!id) return false;
  const ss = SpreadsheetApp.openById(id);
  processOneOrganisationSheet_(ss);
  return true;
}

function checkCompletionAndNotify_() {
  const sh = getMasterTab_(), last = sh.getLastRow(), props = PropertiesService.getScriptProperties();

  const h1 = sh.getRange(`${CFG.MASTER_COL.STATUS1}1`).getDisplayValue() || '申込書A';
  const h2 = sh.getRange(`${CFG.MASTER_COL.STATUS2}1`).getDisplayValue() || '申込書B';

  let sent = 0;
  for (let r = 2; r <= last; r++) {
    const organisation = sh.getRange(`${CFG.MASTER_COL.ORGANISATION}${r}`).getDisplayValue() || '(組織名未設定)';
    const s1 = sh.getRange(`${CFG.MASTER_COL.STATUS1}${r}`).getDisplayValue();
    const s2 = sh.getRange(`${CFG.MASTER_COL.STATUS2}${r}`).getDisplayValue();

    const isDone = isDoneMark_(s1) && isDoneMark_(s2);

    const flagKey = 'done_' + r;
    const already = props.getProperty(flagKey) === '1';

    if (isDone && !already) {
      try { exportOneRowByMaster_(r); } catch (e) { console.warn('完了直後PDF化失敗:', e); }

      const folderUrl = (sh.getRange(`${CFG.MASTER_COL.FOLDER_URL}${r}`).getDisplayValue() || '').toString();
      const d1 = sh.getRange(`${CFG.MASTER_COL.FORM1_DEADLINE}${r}`).getDisplayValue() || '';
      const d2 = sh.getRange(`${CFG.MASTER_COL.FORM2_DEADLINE}${r}`).getDisplayValue() || '';

      sendSlack_(
        `${organisation} の記入が完了しました`,
        [
          { type: "section", text: { type: "mrkdwn", text: `✅ *${organisation}* の記入が完了しました` } },
          { type: "section", text: { type: "mrkdwn", text: `• *${h1}*\n• *${h2}*` } },
          { type: "section", fields: [
            { type: "mrkdwn", text: `*最終更新（申込書A 完了日）*\n${d1 || '-'}` },
            { type: "mrkdwn", text: `*最終更新（申込書B 完了日）*\n${d2 || '-'}` }
          ]},
          ...(folderUrl ? [{ type: "section", text: { type: "mrkdwn", text: `<${folderUrl}|格納フォルダ>` } }] : [])
        ]
      );

      props.setProperty(flagKey, '1'); sent++;
    }
    if (!isDone && already) props.deleteProperty(flagKey);
  }
  SpreadsheetApp.getActive().toast(`完了検知: Slack送信 ${sent}件`);
}

function job_checkCompletionAndNotify() {
  checkCompletionAndNotify_();
}

function job_exportAllOrganisationsDocsSmart() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try { exportAllOrganisationsDocsSmart_(); } finally { lock.releaseLock(); }
}

/**** Smart bulk PDF export & Force full regeneration ****/
function exportAllOrganisationsDocsSmart_() {
  const master = getMasterTab_();
  const last = master.getLastRow();
  let done = 0, skip = 0;

  for (let r = 2; r <= last; r++) {
    const url = (master.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
    const id  = extractId_(url); if (!id) { skip++; continue; }
    const ss  = SpreadsheetApp.openById(id);
    const tab = ss.getSheetByName(CFG.ORGANISATION_TAB_NAME); if (!tab) { skip++; continue; }

    const pIt = DriveApp.getFileById(ss.getId()).getParents();
    if (!pIt.hasNext()) { skip++; continue; }
    const parentFolder = pIt.next();

    const lastRow = tab.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      const lastPdf = tab.getRange(`${CFG.ORGANISATION_COL.LAST_UPDATED}${i}`).getValue();
      const need = shouldRebuildRow_(tab, i, lastPdf instanceof Date ? lastPdf : null);
      if (!need) { skip++; continue; }

      deleteExistingPdfIfAny_(tab, i);

      const cell   = tab.getRange(`${CFG.ORGANISATION_COL.DOC_URL}${i}`);
      const text   = (cell.getDisplayValue() || '').toString().trim();
      const cellUrl= getUrlFromCell_(cell);
      const target = (cellUrl || text).trim();
      if (!target) { skip++; continue; }
      const { url:docUrl, rangeA1 } = splitUrlRange_(target);
      const docId = extractId_(docUrl); if (!docId) { skip++; continue; }

      const mime = DriveApp.getFileById(docId).getMimeType();
      const pdfBlob = makePdfBlobUniversal_(docId, mime, rangeA1);

      const organisationName = ss.getName();
      const formLabel  = (tab.getRange(`${CFG.ORGANISATION_COL.NAME}${i}`).getDisplayValue() || '提出書類');
      const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
      const safe = s => (s || '').replace(/[\\/:*?"<>|]/g, '');
      const name = `${safe(organisationName)}_${safe(formLabel)}_${ymd}.pdf`;

      const file = parentFolder.createFile(pdfBlob).setName(name);
      tab.getRange(`${CFG.ORGANISATION_COL.PDF_URL}${i}`).setValue(file.getUrl());
      tab.getRange(`${CFG.ORGANISATION_COL.LAST_UPDATED}${i}`).setValue(new Date());
      done++;
    }
  }
  SpreadsheetApp.getActive().toast(`スマートPDF化：更新/不足 ${done}件、スキップ ${skip}件`);
}

function shouldRebuildRow_(tab, row, lastPdfUpdated) {
  const cell = tab.getRange(`${CFG.ORGANISATION_COL.DOC_URL}${row}`);
  const cellText = (cell.getDisplayValue() || '').toString().trim();
  const cellUrl  = getUrlFromCell_(cell);
  const target   = (cellUrl || cellText).trim();
  if (!target) return false;

  const { url } = splitUrlRange_(target);
  const docId = extractId_(url); if (!docId) return false;
  try {
    const srcUpdated = DriveApp.getFileById(docId).getLastUpdated();
    return !lastPdfUpdated || srcUpdated > lastPdfUpdated;
  } catch (e) {
    return false;
  }
}

function exportAllOrganisationsDocsForce_() {
  const master = getMasterTab_();
  const last = master.getLastRow();
  let done = 0;

  for (let r = 2; r <= last; r++) {
    const url = (master.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
    const id  = extractId_(url); if (!id) continue;
    const ss  = SpreadsheetApp.openById(id);
    const tab = ss.getSheetByName(CFG.ORGANISATION_TAB_NAME); if (!tab) continue;

    const pIt = DriveApp.getFileById(ss.getId()).getParents();
    if (!pIt.hasNext()) continue;
    const parentFolder = pIt.next();

    const lastRow = tab.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
      deleteExistingPdfIfAny_(tab, i);

      const cell   = tab.getRange(`${CFG.ORGANISATION_COL.DOC_URL}${i}`);
      const text   = (cell.getDisplayValue() || '').toString().trim();
      const cellUrl= getUrlFromCell_(cell);
      const target = (cellUrl || text).trim();
      if (!target) continue;

      const { url:docUrl, rangeA1 } = splitUrlRange_(target);
      const docId = extractId_(docUrl); if (!docId) continue;

      const mime = DriveApp.getFileById(docId).getMimeType();
      const pdfBlob = makePdfBlobUniversal_(docId, mime, rangeA1);

      const organisationName = ss.getName();
      const formLabel  = (tab.getRange(`${CFG.ORGANISATION_COL.NAME}${i}`).getDisplayValue() || '提出書類');
      const ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
      const safe = s => (s || '').replace(/[\\/:*?"<>|]/g, '');
      const name = `${safe(organisationName)}_${safe(formLabel)}_${ymd}.pdf`;

      const file = parentFolder.createFile(pdfBlob).setName(name);
      tab.getRange(`${CFG.ORGANISATION_COL.PDF_URL}${i}`).setValue(file.getUrl());
      tab.getRange(`${CFG.ORGANISATION_COL.LAST_UPDATED}${i}`).setValue(new Date());
      done++;
    }
  }
  SpreadsheetApp.getActive().toast(`強制フル再生成：作成 ${done}件`);
}

/**** Notification flag helpers ****/
function alreadyNotified_(row, offset, formKey) {
  const props = PropertiesService.getScriptProperties();
  const key = `notified_${row}:${formKey}`;
  const set = new Set((props.getProperty(key) || '').split(',').filter(Boolean).map(Number));
  return set.has(offset);
}
function markNotified_(row, offset, formKey) {
  const props = PropertiesService.getScriptProperties();
  const key = `notified_${row}:${formKey}`;
  const set = new Set((props.getProperty(key) || '').split(',').filter(Boolean).map(Number));
  set.add(offset);
  props.setProperty(key, Array.from(set).join(','));
}
function clearNotified_(row, offset, formKey) {
  const props = PropertiesService.getScriptProperties();
  const key = `notified_${row}:${formKey}`;
  if (offset == null) { props.deleteProperty(key); return; }
  const set = new Set((props.getProperty(key) || '').split(',').filter(Boolean).map(Number));
  set.delete(offset);
  props.setProperty(key, Array.from(set).join(','));
}

/**** Utility functions ****/
function colToIndex_(letter) {
  let n = 0; for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
  return n;
}
function extractId_(url) {
  const m = (url || '').match(/[-\w]{25,}/);
  return m ? m[0] : '';
}
function toYmd_(d) {
  if (!(d instanceof Date) || isNaN(d)) return null;
  const y = d.getFullYear(), m = ('0'+(d.getMonth()+1)).slice(-2), day = ('0'+d.getDate()).slice(-2);
  return `${y}-${m}-${day}`;
}
function daysDiff_(fromYmd, toYmd) {
  const a = new Date(fromYmd+'T00:00:00'), b = new Date(toYmd+'T00:00:00');
  return Math.round((b-a)/(1000*60*60*24));
}
function readRowEmails_(sh, r) {
  const from = colToIndex_(CFG.MASTER_COL.MAIL_FROM), to = colToIndex_(CFG.MASTER_COL.MAIL_TO);
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return sh.getRange(r, from, 1, to-from+1).getValues()[0]
    .map(v => (v||'').toString().trim()).filter(v => emailRegex.test(v));
}
function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('編集ログ');
  if (!sh) {
    sh = ss.insertSheet('編集ログ');
    sh.getRange(1,1,1,6).setValues([['日時','編集者','シート','範囲','旧値','新値']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/**** Quarantine helpers ****/
function loadQuarantineSet_(folderId) {
  const key = `perm_quarantine:${folderId}`;
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  try { return new Set(JSON.parse(raw || '[]')); }
  catch (_) { return new Set(); }
}
function saveQuarantineSet_(folderId, set) {
  const key = `perm_quarantine:${folderId}`;
  PropertiesService.getScriptProperties()
    .setProperty(key, JSON.stringify(Array.from(set)));
}

function classifyAddError_(msg) {
  const s = String(msg || '');
  if (/Invalid argument/i.test(s)) return 'permanent';
  if (/notFound|File not found/i.test(s)) return 'permanent';
  if (/cannotShare|sharing.*disabled/i.test(s)) return 'permanent';
  return 'transient';
}

/**** Reminder log helpers (AU/AV columns) ****/
function parseRemindLog_(text) {
  const map = {};
  const s = (text || '').toString().trim();
  if (!s) return map;
  s.split(/\s*\|\s*/).forEach(token => {
    const m = token.match(/^(d-[1-4])\s*:\s*([\d-]{10}\s+[\d:]{8})$/);
    if (m) map[m[1]] = m[2];
  });
  return map;
}

function writeRemindLog_(sh, row, colLetter, offsetTag) {
  const col = colToIndex_(colLetter);
  const cur = (sh.getRange(row, col).getDisplayValue() || '').toString();
  const map = parseRemindLog_(cur);
  map[offsetTag] = fmtNow_();
  const out = ['d-4','d-1'].filter(k => map[k]).map(k => `${k}:${map[k]}`).join(' | ');
  sh.getRange(row, col).setValue(out || '');
}

function alreadySentAtOffset_(sh, row, colLetter, offset) {
  const tag = `d-${offset}`;
  const col = colToIndex_(colLetter);
  const cur = (sh.getRange(row, col).getDisplayValue() || '').toString();
  const map = parseRemindLog_(cur);
  return Boolean(map[tag]);
}

/**** Completion keyword check ****/
function isDoneMark_(val) {
  const s = (val || '').toString();
  return COMPLETE_WORDS.some(w => s.includes(w));
}

/**** Grant notification email ****/
function sendGrantNotice_(recipients, ctx) {
  if (!recipients || recipients.length === 0) return;

  const subj = `【${ctx.organisation}】提出物管理用フォルダへのアクセス権を付与しました`;
  const html =
    `${ctx.organisation} ご担当者さま<br><br>` +
    `提出物管理用のフォルダへのアクセス権を付与しました。<br>` +
    `▼提出用フォルダ：${ctx.folderUrl ? `<a href="${ctx.folderUrl}">${ctx.folderUrl}</a>` : '-' }<br>` +
    `▼管理シート（各組織ブック）：${ctx.bookUrl ? `<a href="${ctx.bookUrl}">${ctx.bookUrl}</a>` : '-' }<br>` +
    `<br>` +
    `【締切の目安】<br>` +
    `・申込書A：${ctx.form1DueDisp || '-'}<br>` +
    `・申込書B：${ctx.form2DueDisp || '-'}<br>` +
    `<br>` +
    `【作業の流れ】<br>` +
    `1) 管理シートの「ステータス」列をご確認の上、未完了の申込書からご対応ください。<br>` +
    `2) 完了後は「完了」へステータスを更新してください（自動判定されます）。<br>` +
    `3) PDFは自動生成/格納されます（時間差あり）。<br>` +
    `<br>` +
    `ご不明点があれば担当までご連絡ください。<br>` +
    `※本メールはシステムより自動送信しています。`;

  MailApp.sendEmail({
    to: recipients[0],
    cc: recipients.slice(1).join(','),
    subject: subj,
    htmlBody: html
  });
}

/**** Bulk grant mail sending ****/
function setSendStatus_(sh, row, status) {
  const arCol = colToIndex_('AR');
  sh.getRange(row, arCol).setValue(status || '');
}

function readRowEmailsDisplay_(sh, row) {
  const from = colToIndex_(CFG.MASTER_COL.MAIL_FROM);
  const to   = colToIndex_(CFG.MASTER_COL.MAIL_TO);
  const vals = sh.getRange(row, from, 1, to - from + 1).getDisplayValues()[0];
  const tokens = vals.flatMap(v =>
    (v || '').toString().split(/[,\n;\t 　、；<>]+/).map(s => s.trim()).filter(Boolean)
  );
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return Array.from(new Set(tokens.filter(s => emailRegex.test(s))));
}

function sendGrantMailForActiveRow() {
  const sh = getMasterTab_();
  const a = sh.getActiveCell();
  if (!a) { SpreadsheetApp.getActive().toast('セルを選択してください'); return; }
  const r = a.getRow();
  if (r < 2) { SpreadsheetApp.getActive().toast('データ行を選択してください'); return; }

  const result = sendGrantMailForRow_(sh, r, { force: true });
  SpreadsheetApp.getActive().toast(result.msg);
}

function sendGrantMailsBulk() {
  const sh = getMasterTab_();
  const last = sh.getLastRow();
  let ok = 0, err = 0, skip = 0;

  for (let r = 2; r <= last; r++) {
    const statusNow = (sh.getRange(`AR${r}`).getDisplayValue() || '').toString().trim();
    if (statusNow === '完了') {
      setSendMeta_(sh, r, `${fmtNow_()} | Skip(Already 完了)`);
      skip++;
      continue;
    }

    const res = sendGrantMailForRow_(sh, r, { force: false });
    if (res.sent) ok++; else if (res.errored) err++; else skip++;

    Utilities.sleep(300);
  }
  SpreadsheetApp.getActive().toast(`送信 完了:${ok} / エラー:${err} / スキップ:${skip}`);
}

function sendGrantMailForRow_(sh, r, opt) {
  const force = !!(opt && opt.force);
  const ts = fmtNow_();

  const emails = readRowEmailsDisplay_(sh, r);
  if (emails.length === 0) {
    setSendStatus_(sh, r, '');
    setSendMeta_(sh, r, `${ts} | No-Addr`);
    return { sent: false, errored: false, msg: `r=${r} アドレスなし（AK空欄/AL記録）` };
  }

  const akNow = (sh.getRange(`AK${r}`).getDisplayValue() || '').toString().trim();
  if (!force && akNow === '完了') {
    setSendMeta_(sh, r, `${ts} | Skip(Already 完了)`);
    return { sent: false, errored: false, msg: `r=${r} 既に完了のためスキップ` };
  }

  const organisation = sh.getRange(`${CFG.MASTER_COL.ORGANISATION}${r}`).getDisplayValue() || '(組織名未設定)';
  const folderUrl    = (sh.getRange(`${CFG.MASTER_COL.FOLDER_URL}${r}`).getDisplayValue() || '').toString().trim();
  const bookUrl      = (sh.getRange(`${CFG.MASTER_COL.ORGANISATION_BOOK_URL}${r}`).getDisplayValue() || '').toString().trim();
  const form1DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM1_DEADLINE}${r}`).getDisplayValue() || '').toString();
  const form2DueDisp = (sh.getRange(`${CFG.MASTER_COL.FORM2_DEADLINE}${r}`).getDisplayValue() || '').toString();
  const ctx = { organisation, folderUrl, bookUrl, form1DueDisp, form2DueDisp };

  const subj = `【${ctx.organisation}】提出物管理用フォルダへのアクセス権を付与しました`;
  const html =
    `${ctx.organisation} ご担当者さま<br><br>` +
    `提出物管理用のフォルダへのアクセス権を付与しました。<br>` +
    `▼提出用フォルダ：${ctx.folderUrl ? `<a href="${ctx.folderUrl}">${ctx.folderUrl}</a>` : '-' }<br>` +
    `▼管理シート（各組織ブック）：${ctx.bookUrl ? `<a href="${ctx.bookUrl}">${ctx.bookUrl}</a>` : '-' }<br>` +
    `<br>` +
    `【締切の目安】<br>` +
    `・申込書A：${ctx.form1DueDisp || '-'}<br>` +
    `・申込書B：${ctx.form2DueDisp || '-'}<br>` +
    `<br>` +
    `【作業の流れ】<br>` +
    `1) 管理シートの「ステータス」列をご確認の上、未完了の申込書からご対応ください。<br>` +
    `2) 完了後は「完了」へステータスを更新してください（自動判定されます）。<br>` +
    `3) PDFは自動生成/格納されます（時間差あり）。<br>` +
    `<br>` +
    `ご不明点があれば担当までご連絡ください。<br>` +
    `※本メールはシステムより自動送信しています。`;

  try {
    const to = emails[0];
    const cc = emails.slice(1).join(',');
    const result = sendHtmlEmail_(to, cc, subj, html);

    if (result.ok) {
      setSendStatus_(sh, r, '完了');
      const via = result.via + (result.warn ? ` | warn:${result.warn}` : '');
      setSendMeta_(sh, r, `${ts} | OK(${emails.length}) | ${via}`);
      return { sent: true, errored: false, msg: `r=${r} 送付完了（${emails.length}件）` };
    } else {
      setSendStatus_(sh, r, 'エラー');
      setSendMeta_(sh, r, `${ts} | ERROR | via:${result.via} | ${result.err || '-'}`);
      return { sent: false, errored: true, msg: `r=${r} エラー: ${result.err || '-'}` };
    }
  } catch (e) {
    setSendStatus_(sh, r, 'エラー');
    setSendMeta_(sh, r, `${ts} | EXCEPTION | ${e && e.message ? e.message : String(e)}`);
    return { sent: false, errored: true, msg: `r=${r} 例外: ${e && e.message ? e.message : e}` };
  }
}

/**** Timestamp formatter ****/
function fmtNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

/**** Send log helpers ****/
function setSendMeta_(sh, row, text) {
  const asCol = colToIndex_('AL');
  sh.getRange(row, asCol).setValue(text || '');
}

/**** Sender configuration (optional; default OFF) ****/
const SENDER = {
  USE_GMAILAPP: false,           // Set true to use GmailApp (enables alias sending)
  ALIAS: '',                     // Registered Gmail alias address (e.g. 'ops@example.org')
  NAME: 'YOUR_ORGANIZATION_NAME', // Display name for sender
  REPLY_TO: ''
};

/**** Unified HTML email sender (auto-switches Gmail/MailApp) ****/
function sendHtmlEmail_(to, cc, subject, html) {
  to = (to || '').trim();
  cc = (cc || '').trim();

  if (SENDER.USE_GMAILAPP) {
    try {
      const aliases = GmailApp.getAliases();
      const canAlias = SENDER.ALIAS && aliases && aliases.indexOf(SENDER.ALIAS) !== -1;

      const opts = {
        htmlBody: html,
        name: SENDER.NAME || undefined,
        replyTo: SENDER.REPLY_TO || undefined,
        cc: cc || undefined,
        from: canAlias ? SENDER.ALIAS : undefined
      };

      GmailApp.sendEmail(to, subject, '(HTMLメール対応クライアントでご確認ください)', opts);
      return { ok: true, via: canAlias ? ('GmailApp(alias:' + SENDER.ALIAS + ')') : 'GmailApp(default)' };
    } catch (e) {
      try {
        MailApp.sendEmail({ to, cc, subject, htmlBody: html, name: SENDER.NAME || undefined, replyTo: SENDER.REPLY_TO || undefined });
        return { ok: true, via: 'MailApp(fallback)', warn: e.message || String(e) };
      } catch (ee) {
        return { ok: false, via: 'MailApp(fallback)', err: ee.message || String(ee), warn: e.message || String(e) };
      }
    }
  }

  try {
    MailApp.sendEmail({ to, cc, subject, htmlBody: html, name: SENDER.NAME || undefined, replyTo: SENDER.REPLY_TO || undefined });
    return { ok: true, via: 'MailApp' };
  } catch (e) {
    return { ok: false, via: 'MailApp', err: e.message || String(e) };
  }
}

/**** Permission status check (writes to AO~AQ) ****/
function checkShareStatusAll() {
  const sh = getMasterTab_();
  const last = sh.getLastRow();
  for (let r = 2; r <= last; r++) {
    try { checkShareStatusForRow_(sh, r); } catch (e) { console.warn('check row fail:', r, e); }
    Utilities.sleep(50);
  }
  SpreadsheetApp.getActive().toast('権限チェックを完了しました（AO〜AQ を更新）');
}

function checkShareStatusForRow_(sh, r) {
  const folderUrl = (sh.getRange(`${CFG.MASTER_COL.FOLDER_URL}${r}`).getDisplayValue() || '').toString().trim();
  if (!folderUrl) return;
  const folderId = extractId_(folderUrl);
  if (!folderId) return;

  const fromCol = colToIndex_(CFG.MASTER_COL.MAIL_FROM);
  const toCol   = colToIndex_(CFG.MASTER_COL.MAIL_TO);
  const rawRow  = sh.getRange(r, fromCol, 1, toCol - fromCol + 1).getDisplayValues()[0];

  const can = buildGrantedEmailSet_(folderId, CFG.PERMISSION_ROLE);

  const outStart = colToIndex_('AO');
  const results = rawRow.map(cell => {
    const tokens = splitEmailsFromDisplayCell_(cell);
    if (tokens.length === 0) return '';
    const hits = tokens.filter(e => can.has(e.toLowerCase()));
    if (hits.length === 0) return '未共有';
    if (hits.length === tokens.length) return '済';
    return '一部';
  });

  sh.getRange(r, outStart, 1, results.length).setValues([results]);
}

function buildGrantedEmailSet_(folderId, roleMode) {
  const wantEditor = (roleMode || '').toLowerCase() !== 'viewer';
  const set = new Set();

  try {
    const folder = DriveApp.getFolderById(folderId);
    folder.getEditors().forEach(u => {
      const e = (u.getEmail() || '').toLowerCase();
      if (e) set.add(e);
    });
    if (!wantEditor) {
      folder.getViewers().forEach(u => {
        const e = (u.getEmail() || '').toLowerCase();
        if (e) set.add(e);
      });
    }
    try {
      const owner = folder.getOwner();
      if (owner) {
        const e = (owner.getEmail() || '').toLowerCase();
        if (e) set.add(e);
      }
    } catch (_) {}
  } catch (e) {
    console.warn('DriveApp list error:', e);
  }

  // Requires Drive API service enabled in Script Editor > Services
  try {
    if (typeof Drive !== 'undefined' && Drive.Permissions) {
      const resp = Drive.Permissions.list(folderId, {
        supportsAllDrives: true,
        fields: 'permissions(emailAddress,role,type,domain)'
      });
      const perms = (resp && resp.permissions) || [];
      perms.forEach(p => {
        const type = (p.type || '').toLowerCase();
        const role = (p.role || '').toLowerCase();
        const ok = wantEditor ? (role === 'writer' || role === 'fileorganizer' || role === 'organizer' || role === 'owner')
                              : (role === 'reader' || role === 'commenter' || role === 'writer' || role === 'fileorganizer' || role === 'organizer' || role === 'owner');
        if (!ok) return;
        if (type === 'user' || type === 'group') {
          const e = (p.emailAddress || '').toLowerCase();
          if (e) set.add(e);
        }
      });
    }
  } catch (e) {
    console.warn('Drive API list error:', e);
  }

  return set;
}

function splitEmailsFromDisplayCell_(v) {
  const s = (v || '').toString();
  const tokens = s.split(/[,\n;\t 　、；<>]+/).map(x => x.trim()).filter(Boolean);
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return tokens.filter(x => emailRegex.test(x)).map(x => x.toLowerCase());
}

/**** PDF kill switch ****/
const KILL = { PDF_MASTER_DISABLED_PROP: 'pdf_master_disabled' };

function isMasterPdfDisabled_() {
  return PropertiesService.getScriptProperties().getProperty(KILL.PDF_MASTER_DISABLED_PROP) === '1';
}

function stopMasterPdfNow() {
  PropertiesService.getScriptProperties().setProperty(KILL.PDF_MASTER_DISABLED_PROP, '1');

  const killHandlers = [
    'exportAllDocsToPdf',
    'exportSelectedOrganisationDocsToPdf',
    'exportAllOrganisationsDocsSmart_',
    'exportAllOrganisationsDocsForce_',
    'checkCompletionAndNotify_',
    'job_exportAllOrganisationsDocsSmart'
  ];
  ScriptApp.getProjectTriggers()
    .filter(t => killHandlers.includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));

  SpreadsheetApp.getActive().toast('マスター側PDF生成を停止しました（トリガー削除＋キルスイッチON）');
}

function resumeMasterPdf() {
  PropertiesService.getScriptProperties().deleteProperty(KILL.PDF_MASTER_DISABLED_PROP);
  SpreadsheetApp.getActive().toast('マスター側PDF生成の停止フラグを解除しました');
}

/**** Test helpers (remove or disable before production use) ****/
function createReminderTestTrigger5min() {
  deleteReminderTestTriggers();
  ScriptApp.newTrigger('dailyReminder_').timeBased().everyMinutes(5).create();
  Browser.msgBox('テスト：5分おきに dailyReminder_ を実行します');
}
function deleteReminderTestTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyReminder_')
    .forEach(t => ScriptApp.deleteTrigger(t));
  Browser.msgBox('テスト：5分おきトリガーを削除しました');
}

function testReminder() {
  dailyReminder_();
}

// Send a test email to yourself only — remove after verification
function testMailOnlyForMe() {
  const myEmail = 'your-email@example.com'; // Replace with your email address

  const organisation = 'テスト組織';
  const formName = '申込書B';
  const dueYmd   = '2026-02-28';
  const status   = '未完了';
  const diff     = '3';
  const bookUrl  = 'YOUR_TEST_SPREADSHEET_URL';

  const subj = `【YOUR_PROGRAM_NAME】【${organisation}】${formName} リマインド（締切まで残り${diff}日）`;
  const html =
    `${organisation} ご担当者さま<br><br>` +
    `標記の件につきまして、進捗状況はいかがでしょうか。<br>` +
    `本メールと行き違いで既にご対応済みの場合は、何卒ご容赦ください。<br>` +
    `未完了の方におかれましては、<b>【${dueYmd}】</b>までに作業をお願いいたします。<br><br>` +
    `対象申込書：${formName}<br>` +
    `締切日：${dueYmd}<br>` +
    `現在の状況：${status}<br>` +
    `締切まで <b>残り ${diff}日</b> です。<br><br>` +
    `▼管理シート（共有ドライブへアクセス可能な方）<br>` +
    `${bookUrl ? `<a href="${bookUrl}">${bookUrl}</a>` : '(未設定)'}<br>` +
    `※直接入力の上、ステータスの「完了」への変更をお願いします。<br><br>` +
    `【提出方法について：共有ドライブへアクセスできない方】<br>` +
    `エクセルファイルに必要事項をご記入の上、以下のメールアドレス宛にご提出ください。<br>` +
    `（エクセルファイルは参画校サイトからもダウンロードできます。）<br>` +
    `【エクセル提出先】 <a href="mailto:info@your-domain.jp">info@your-domain.jp</a><br><br>` +
    `お忙しい折、お手数をおかけしますが、ご協力のほどよろしくお願い申し上げます。<br><br>` +
    `※本メールは自動送信です。ご不明点は事務局担当者へご連絡ください。`;

  MailApp.sendEmail({ to: myEmail, subject: subj, htmlBody: html });
  Logger.log('Test email sent. Please check your inbox.');
}
