import API from './lib/api.js';
import { MergeEngine } from './lib/merge.js';

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  currentWizardStep: 1,
  selectedSpreadsheet: null,   // { id, name }
  selectedSheetName: null,
  emailColumn: null,
  sheetHeaders: [],
  sheetRows: [],                // array of row objects { Email, Name, ... }
  selectedDraft: null,          // { id, subject, htmlBody, textBody }
  mergedRecipients: [],         // array of { to, subject, htmlBody, vars, rowIndex }
  previewIndex: 0,
  sendStats: { sent: 0, failed: 0, skipped: 0, errors: [] },
  isSending: false,
  stopRequested: false,
  campaignStartTime: null,
  campaigns: [],                // saved from storage
  currentCampaignId: null,
  activeView: 'new-campaign',
  userInfo: null,
  settings: {},
};

// ── Init ───────────────────────────────────────────────────────────────────
async function init() {
  await loadSettings();
  await loadCampaigns();

  try {
    await API.getToken(false);
    state.userInfo = await API.getUserInfo();
    showUserInNav(state.userInfo);
  } catch {
    // Not signed in — try interactive
    try {
      await API.getToken(true);
      state.userInfo = await API.getUserInfo();
      showUserInNav(state.userInfo);
    } catch (err) {
      showToast('Please sign in via the extension popup first.', 'error');
    }
  }

  // Load send-as aliases
  loadSendAsAliases();

  // Load quota info
  updateQuotaView();

  // Set up URL param routing
  const params = new URLSearchParams(window.location.search);
  if (params.get('view') === 'campaigns') switchView('campaigns');

  loadCampaignsView();
  loadScheduledView();
}

function showUserInNav(user) {
  if (!user) return;
  const navUser = document.getElementById('nav-user');
  navUser.classList.remove('hidden');
  document.getElementById('nav-email').textContent = user.email;
  const avatar = document.getElementById('nav-avatar');
  if (user.picture) avatar.src = user.picture;
}

async function loadSendAsAliases() {
  try {
    const data = await API.getSendAsAliases();
    const sel = document.getElementById('send-as-select');
    sel.innerHTML = '';
    (data.sendAs || []).forEach(alias => {
      const opt = document.createElement('option');
      opt.value = alias.sendAsEmail;
      opt.textContent = alias.displayName
        ? `${alias.displayName} <${alias.sendAsEmail}>`
        : alias.sendAsEmail;
      if (alias.isDefault) opt.selected = true;
      sel.appendChild(opt);
    });
    // Pre-fill from name
    const selected = data.sendAs?.find(a => a.isDefault);
    if (selected?.displayName) {
      document.getElementById('from-name').value = selected.displayName;
    }
  } catch (err) {
    console.warn('Could not load aliases', err);
  }
}

// ── Settings ───────────────────────────────────────────────────────────────
function loadSettings() {
  return new Promise(resolve => {
    chrome.storage.local.get(['mailMergeSettings'], data => {
      state.settings = data.mailMergeSettings || {
        dailyQuota: 400,
        trackingServerUrl: '',
        defaultFromName: '',
        defaultReplyTo: '',
      };
      resolve();
    });
  });
}

function saveSettings(patch) {
  Object.assign(state.settings, patch);
  chrome.storage.local.set({ mailMergeSettings: state.settings });
}

function loadCampaigns() {
  return new Promise(resolve => {
    chrome.storage.local.get(['campaigns'], data => {
      state.campaigns = data.campaigns || [];
      resolve();
    });
  });
}

function saveCampaign(campaign) {
  const idx = state.campaigns.findIndex(c => c.id === campaign.id);
  if (idx >= 0) state.campaigns[idx] = campaign;
  else state.campaigns.unshift(campaign);
  chrome.storage.local.set({ campaigns: state.campaigns });
}

// ── View Switching ─────────────────────────────────────────────────────────
function switchView(viewName) {
  document.querySelectorAll('.content-view').forEach(v => v.classList.add('hidden'));
  document.getElementById(`view-${viewName}`).classList.remove('hidden');
  document.querySelectorAll('.sidebar-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.view === viewName);
  });
  state.activeView = viewName;
  if (viewName === 'campaigns') loadCampaignsView();
  if (viewName === 'scheduled') loadScheduledView();
  if (viewName === 'templates') loadTemplatesView();
  if (viewName === 'quota') updateQuotaView();
}

// ── Wizard Navigation ──────────────────────────────────────────────────────
function goToStep(step) {
  document.querySelectorAll('.wizard-step').forEach(s => s.classList.add('hidden'));
  document.getElementById(`step-${step}`).classList.remove('hidden');

  document.querySelectorAll('.step').forEach(el => {
    const n = parseInt(el.dataset.step);
    el.classList.remove('active', 'completed');
    if (n === step) el.classList.add('active');
    else if (n < step) el.classList.add('completed');
  });

  state.currentWizardStep = step;
}

// ── Step 1: Spreadsheet Search ─────────────────────────────────────────────
let sheetSearchTimer = null;
document.getElementById('sheet-search').addEventListener('input', e => {
  clearTimeout(sheetSearchTimer);
  const q = e.target.value.trim();
  if (!q) {
    hideDropdown('sheet-results');
    return;
  }
  // Handle pasted spreadsheet URLs
  const urlMatch = q.match(/spreadsheets\/d\/([\w-]+)/);
  if (urlMatch) {
    selectSpreadsheetById(urlMatch[1]);
    return;
  }
  sheetSearchTimer = setTimeout(() => searchSpreadsheets(q), 400);
});

document.getElementById('btn-browse-sheets').addEventListener('click', () => searchSpreadsheets(''));

async function searchSpreadsheets(query) {
  const resultEl = document.getElementById('sheet-results');
  resultEl.innerHTML = '<div class="dropdown-item"><span>Loading...</span></div>';
  resultEl.classList.remove('hidden');
  try {
    const data = query
      ? await API.searchSpreadsheets(query)
      : await API.listSpreadsheets();
    const files = data.files || [];
    if (!files.length) {
      resultEl.innerHTML = '<div class="dropdown-item"><span class="dropdown-item-name">No spreadsheets found</span></div>';
      return;
    }
    resultEl.innerHTML = '';
    files.forEach(file => {
      const item = document.createElement('div');
      item.className = 'dropdown-item';
      item.innerHTML = `<span>📊</span><div><div class="dropdown-item-name">${esc(file.name)}</div><div class="dropdown-item-sub">${esc(file.id)}</div></div>`;
      item.addEventListener('click', () => selectSpreadsheet(file));
      resultEl.appendChild(item);
    });
  } catch (err) {
    resultEl.innerHTML = `<div class="dropdown-item"><span class="dropdown-item-name" style="color:#d93025">Error: ${esc(err.message)}</span></div>`;
  }
}

async function selectSpreadsheetById(id) {
  try {
    const meta = await API.getSpreadsheetMeta(id);
    selectSpreadsheet({ id: meta.spreadsheetId, name: meta.properties.title });
  } catch (err) {
    showToast('Could not load spreadsheet: ' + err.message, 'error');
  }
}

async function selectSpreadsheet(file) {
  state.selectedSpreadsheet = file;
  hideDropdown('sheet-results');
  document.getElementById('sheet-search').value = '';

  document.getElementById('selected-sheet-name').textContent = file.name;
  document.getElementById('selected-sheet-id').textContent = file.id;
  document.getElementById('selected-sheet').classList.remove('hidden');

  // Load sheet tabs
  try {
    const meta = await API.getSpreadsheetMeta(file.id);
    const sheets = meta.sheets || [];
    const sel = document.getElementById('sheet-tab-select');
    sel.innerHTML = '<option value="">Select sheet tab...</option>';
    sheets.forEach(s => {
      const opt = document.createElement('option');
      opt.value = s.properties.title;
      opt.textContent = s.properties.title;
      sel.appendChild(opt);
    });
    document.getElementById('section-sheet-tab').classList.remove('hidden');
    if (sheets.length === 1) {
      sel.value = sheets[0].properties.title;
      sel.dispatchEvent(new Event('change'));
    }
  } catch (err) {
    showToast('Could not load sheet tabs: ' + err.message, 'error');
  }
  checkStep1Complete();
}

document.getElementById('btn-clear-sheet').addEventListener('click', () => {
  state.selectedSpreadsheet = null;
  state.selectedSheetName = null;
  state.sheetHeaders = [];
  state.sheetRows = [];
  document.getElementById('selected-sheet').classList.add('hidden');
  document.getElementById('section-sheet-tab').classList.add('hidden');
  document.getElementById('section-email-col').classList.add('hidden');
  checkStep1Complete();
});

document.getElementById('sheet-tab-select').addEventListener('change', async e => {
  const tab = e.target.value;
  if (!tab) return;
  state.selectedSheetName = tab;
  await loadSheetData();
  checkStep1Complete();
});

async function loadSheetData() {
  if (!state.selectedSpreadsheet || !state.selectedSheetName) return;
  try {
    const data = await API.getSheetValues(
      state.selectedSpreadsheet.id,
      `'${state.selectedSheetName}'`
    );
    const rows = data.values || [];
    if (!rows.length) { showToast('Sheet appears empty.', 'error'); return; }

    state.sheetHeaders = rows[0];
    state.sheetRows = rows.slice(1).map((row, i) => {
      const obj = { _rowIndex: i + 2 }; // 1-indexed, +1 for header
      state.sheetHeaders.forEach((h, j) => { obj[h] = row[j] || ''; });
      return obj;
    });

    // Populate email column selector
    const emailSel = document.getElementById('email-col-select');
    emailSel.innerHTML = '<option value="">Select column containing email addresses...</option>';
    state.sheetHeaders.forEach(h => {
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = h;
      // Auto-select columns likely containing emails
      if (/email/i.test(h)) opt.selected = true;
      emailSel.appendChild(opt);
    });

    document.getElementById('section-email-col').classList.remove('hidden');
    if (emailSel.value) updateRecipientCount();

    document.getElementById('section-email-col').classList.remove('hidden');
    checkStep1Complete();
  } catch (err) {
    showToast('Failed to load sheet data: ' + err.message, 'error');
  }
}

document.getElementById('email-col-select').addEventListener('change', e => {
  state.emailColumn = e.target.value;
  updateRecipientCount();
  checkStep1Complete();
});

function updateRecipientCount() {
  const col = document.getElementById('email-col-select').value;
  if (!col) return;
  const skipSent = document.getElementById('opt-skip-sent')?.checked ?? true;
  let count = 0, skipped = 0;
  state.sheetRows.forEach(row => {
    if (!row[col]) return;
    const statusCol = state.sheetHeaders.find(h => /email.?sent|merge.?status/i.test(h));
    if (skipSent && statusCol && row[statusCol]) { skipped++; return; }
    count++;
  });
  const hint = document.getElementById('recipient-count-hint');
  hint.textContent = `${count} recipient${count !== 1 ? 's' : ''} found${skipped ? ` (${skipped} already sent, will be skipped)` : ''}`;
}

// ── Draft Search ───────────────────────────────────────────────────────────
let allDrafts = [];
let draftSearchTimer = null;

document.getElementById('btn-refresh-drafts').addEventListener('click', loadDrafts);

document.getElementById('draft-search').addEventListener('input', e => {
  clearTimeout(draftSearchTimer);
  draftSearchTimer = setTimeout(() => filterDrafts(e.target.value.trim()), 300);
});

document.getElementById('draft-search').addEventListener('focus', () => {
  if (!allDrafts.length) loadDrafts();
  else filterDrafts(document.getElementById('draft-search').value);
});

async function loadDrafts() {
  const resultEl = document.getElementById('draft-results');
  resultEl.innerHTML = '<div class="dropdown-item">Loading drafts...</div>';
  resultEl.classList.remove('hidden');
  try {
    allDrafts = await API.listDrafts();
    filterDrafts('');
  } catch (err) {
    resultEl.innerHTML = `<div class="dropdown-item" style="color:#d93025">Error: ${esc(err.message)}</div>`;
  }
}

function filterDrafts(query) {
  const resultEl = document.getElementById('draft-results');
  const filtered = query
    ? allDrafts.filter(d => {
        const parsed = MergeEngine.parseDraftPayload(d.message);
        return parsed.subject.toLowerCase().includes(query.toLowerCase());
      })
    : allDrafts;

  if (!filtered.length) {
    resultEl.innerHTML = '<div class="dropdown-item">No drafts found. Create a Gmail draft first.</div>';
    resultEl.classList.remove('hidden');
    return;
  }

  resultEl.innerHTML = '';
  filtered.slice(0, 30).forEach(draft => {
    const parsed = MergeEngine.parseDraftPayload(draft.message);
    const item = document.createElement('div');
    item.className = 'dropdown-item';
    item.innerHTML = `
      <span>📧</span>
      <div>
        <div class="dropdown-item-name">${esc(parsed.subject || '(No subject)')}</div>
        <div class="dropdown-item-sub">${esc(draft.message.snippet || '').slice(0, 80)}</div>
      </div>`;
    item.addEventListener('click', () => selectDraft(draft, parsed));
    resultEl.appendChild(item);
  });
  resultEl.classList.remove('hidden');
}

function selectDraft(draft, parsed) {
  const p = parsed || MergeEngine.parseDraftPayload(draft.message);
  state.selectedDraft = {
    id: draft.id,
    subject: p.subject,
    htmlBody: p.htmlBody,
    textBody: p.textBody,
    snippet: draft.message.snippet || '',
  };
  hideDropdown('draft-results');
  document.getElementById('draft-search').value = '';
  document.getElementById('selected-draft-subject').textContent = p.subject || '(No subject)';
  document.getElementById('selected-draft-snippet').textContent = (draft.message.snippet || '').slice(0, 100);
  document.getElementById('selected-draft').classList.remove('hidden');
  checkStep1Complete();
}

document.getElementById('btn-clear-draft').addEventListener('click', () => {
  state.selectedDraft = null;
  document.getElementById('selected-draft').classList.add('hidden');
  checkStep1Complete();
});

function checkStep1Complete() {
  const ok = !!(
    state.selectedSpreadsheet &&
    state.selectedSheetName &&
    state.emailColumn &&
    state.selectedDraft
  );
  document.getElementById('btn-step1-next').disabled = !ok;
}

// ── Step 2: Options ─────────────────────────────────────────────────────────
document.querySelectorAll('input[name="send-time"]').forEach(radio => {
  radio.addEventListener('change', e => {
    document.getElementById('schedule-datetime').classList.toggle('hidden', e.target.value !== 'scheduled');
  });
});

document.getElementById('opt-track-opens').addEventListener('change', updateTrackingUI);
document.getElementById('opt-track-clicks').addEventListener('change', updateTrackingUI);
document.getElementById('opt-unsubscribe').addEventListener('change', e => {
  document.getElementById('unsubscribe-section').classList.toggle('hidden', !e.target.checked);
});

function updateTrackingUI() {
  const needsServer = document.getElementById('opt-track-opens').checked ||
                      document.getElementById('opt-track-clicks').checked;
  document.getElementById('tracking-server-section').classList.toggle('hidden', !needsServer);
  if (state.settings.trackingServerUrl) {
    document.getElementById('tracking-server-url').value = state.settings.trackingServerUrl;
  }
}

document.getElementById('link-tracking-setup').addEventListener('click', e => {
  e.preventDefault();
  openModal('modal-tracking-setup');
});

// ── Preview ────────────────────────────────────────────────────────────────
function buildMergedRecipients() {
  const emailCol = state.emailColumn;
  const skipSent = document.getElementById('opt-skip-sent').checked;
  const statusColName = state.sheetHeaders.find(h => /email.?sent|merge.?status/i.test(h)) || 'Email Sent';
  const fromName = document.getElementById('from-name').value.trim();
  const sendAs = document.getElementById('send-as-select').value;
  const replyTo = document.getElementById('reply-to').value.trim();
  const replyToName = document.getElementById('reply-to-name').value.trim();
  const ccField = document.getElementById('cc-field').value.trim();
  const bccField = document.getElementById('bcc-field').value.trim();
  const addUnsubscribe = document.getElementById('opt-unsubscribe').checked;
  const trackingUrl = document.getElementById('tracking-server-url')?.value.trim();
  const unsubText = document.getElementById('unsubscribe-text').value.trim();
  const subject = state.selectedDraft.subject;
  const htmlBody = state.selectedDraft.htmlBody;
  const textBody = state.selectedDraft.textBody;

  const recipients = [];
  let skippedCount = 0;

  state.sheetRows.forEach(row => {
    const email = (row[emailCol] || '').trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      if (email) recipients.push({ error: `Invalid email: ${email}`, rowIndex: row._rowIndex });
      return;
    }

    const sentStatus = row[statusColName];
    if (skipSent && sentStatus) {
      skippedCount++;
      return;
    }

    const vars = {};
    state.sheetHeaders.forEach(h => { vars[h] = row[h] || ''; });

    const mergedSubject = MergeEngine.process(subject, vars);
    const mergedHtml = MergeEngine.process(htmlBody || textBody, vars, {
      trackingPixelUrl: (document.getElementById('opt-track-opens').checked && trackingUrl)
        ? `${trackingUrl}?type=open&cid=${encodeURIComponent(state.currentCampaignId || 'campaign')}&email=${encodeURIComponent(email)}`
        : null,
      unsubscribeUrl: addUnsubscribe && trackingUrl
        ? `${trackingUrl}?type=unsubscribe&email=${encodeURIComponent(email)}`
        : null,
      unsubscribeText: unsubText,
    });
    const mergedText = MergeEngine.process(textBody, vars);

    const fromStr = fromName
      ? `${fromName} <${sendAs}>`
      : sendAs;
    const replyToStr = replyTo
      ? (replyToName ? `${replyToName} <${replyTo}>` : replyTo)
      : null;

    let cc = MergeEngine.process(ccField, vars);
    let bcc = MergeEngine.process(bccField, vars);

    recipients.push({
      to: email,
      from: fromStr,
      replyTo: replyToStr,
      cc: cc || null,
      bcc: bcc || null,
      subject: mergedSubject,
      htmlBody: mergedHtml,
      textBody: mergedText,
      vars,
      rowIndex: row._rowIndex,
    });
  });

  state.mergedRecipients = recipients;
  return { recipients, skippedCount };
}

function renderPreview() {
  const { recipients, skippedCount } = buildMergedRecipients();
  const valid = recipients.filter(r => !r.error);
  const errors = recipients.filter(r => r.error);

  document.getElementById('preview-total-count').textContent = `${valid.length} recipient${valid.length !== 1 ? 's' : ''}`;
  document.getElementById('preview-skip-count').textContent = `${skippedCount} skipped`;
  document.getElementById('preview-error-count').textContent = `${errors.length} error${errors.length !== 1 ? 's' : ''}`;

  const sendLabel = document.getElementById('send-count-label');
  sendLabel.textContent = valid.length;

  // Variables in template
  const allVars = MergeEngine.extractVariables(
    (state.selectedDraft.htmlBody || '') + ' ' + state.selectedDraft.subject
  );
  const varList = document.getElementById('variables-list');
  varList.innerHTML = '';
  allVars.forEach(v => {
    const chip = document.createElement('span');
    chip.className = `variable-chip ${state.sheetHeaders.includes(v) ? 'matched' : 'unmatched'}`;
    chip.textContent = `{{${v}}}`;
    chip.title = state.sheetHeaders.includes(v)
      ? `Matched to column "${v}"`
      : `No column named "${v}" found in sheet`;
    varList.appendChild(chip);
  });

  state.previewIndex = 0;
  showPreviewAt(0);
}

function showPreviewAt(idx) {
  const valid = state.mergedRecipients.filter(r => !r.error);
  if (!valid.length) return;
  idx = Math.max(0, Math.min(idx, valid.length - 1));
  state.previewIndex = idx;
  const r = valid[idx];

  document.getElementById('preview-counter').textContent = `${idx + 1} of ${valid.length}`;
  document.getElementById('preview-to').textContent = r.to;
  document.getElementById('preview-from').textContent = r.from || '';
  document.getElementById('preview-subject').textContent = r.subject || '(No subject)';

  const ccRow = document.getElementById('preview-cc-row');
  if (r.cc) { ccRow.style.display = 'flex'; document.getElementById('preview-cc').textContent = r.cc; }
  else { ccRow.style.display = 'none'; }

  const bodyEl = document.getElementById('preview-body');
  if (r.htmlBody) {
    // Sandboxed preview — disable scripts
    const shadow = bodyEl.attachShadow ? bodyEl : bodyEl;
    bodyEl.innerHTML = '';
    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'width:100%;border:none;min-height:200px;';
    iframe.sandbox = 'allow-same-origin';
    bodyEl.appendChild(iframe);
    setTimeout(() => {
      iframe.contentDocument.open();
      iframe.contentDocument.write(r.htmlBody);
      iframe.contentDocument.close();
      iframe.style.height = iframe.contentDocument.body.scrollHeight + 20 + 'px';
    }, 50);
  } else {
    bodyEl.innerHTML = `<pre style="white-space:pre-wrap;font-family:inherit">${esc(r.textBody || '')}</pre>`;
  }
}

document.getElementById('btn-prev-preview').addEventListener('click', () => showPreviewAt(state.previewIndex - 1));
document.getElementById('btn-next-preview').addEventListener('click', () => showPreviewAt(state.previewIndex + 1));

document.getElementById('btn-send-test').addEventListener('click', () => {
  const firstValid = state.mergedRecipients.find(r => !r.error);
  if (firstValid) {
    document.getElementById('test-email-addr').value = state.userInfo?.email || '';
  }
  openModal('modal-test-send');
});

document.getElementById('btn-confirm-test-send').addEventListener('click', async () => {
  const testEmail = document.getElementById('test-email-addr').value.trim();
  if (!testEmail) { showToast('Enter a test email address', 'error'); return; }
  closeModal('modal-test-send');

  const firstValid = state.mergedRecipients.find(r => !r.error);
  if (!firstValid) { showToast('No valid recipients to preview', 'error'); return; }

  try {
    const raw = MergeEngine.encodeEmail({
      to: testEmail,
      from: firstValid.from,
      replyTo: firstValid.replyTo,
      subject: `[TEST] ${firstValid.subject}`,
      htmlBody: firstValid.htmlBody,
      textBody: firstValid.textBody,
    });
    await API.sendEmail(raw);
    showToast(`Test email sent to ${testEmail}`, 'success');
  } catch (err) {
    showToast('Failed to send test: ' + err.message, 'error');
  }
});

// ── Confirm & Send ─────────────────────────────────────────────────────────
document.getElementById('btn-step3-send').addEventListener('click', () => {
  const valid = state.mergedRecipients.filter(r => !r.error);
  document.getElementById('confirm-send-message').textContent =
    `You are about to send ${valid.length} personalized email${valid.length !== 1 ? 's' : ''} from "${document.getElementById('from-name').value || document.getElementById('send-as-select').value}".`;

  document.getElementById('confirm-check-reviewed').checked = false;
  document.getElementById('confirm-check-recipients').checked = false;
  document.getElementById('btn-confirm-send').disabled = true;
  openModal('modal-confirm-send');
});

['confirm-check-reviewed', 'confirm-check-recipients'].forEach(id => {
  document.getElementById(id).addEventListener('change', () => {
    const both = document.getElementById('confirm-check-reviewed').checked &&
                 document.getElementById('confirm-check-recipients').checked;
    document.getElementById('btn-confirm-send').disabled = !both;
  });
});

document.getElementById('btn-confirm-send').addEventListener('click', () => {
  closeModal('modal-confirm-send');
  const sendTime = document.querySelector('input[name="send-time"]:checked').value;
  if (sendTime === 'scheduled') {
    scheduleMailMerge();
  } else {
    startMailMerge();
  }
});

// ── Sending ────────────────────────────────────────────────────────────────
async function startMailMerge() {
  goToStep(4);
  switchView('new-campaign');

  const valid = state.mergedRecipients.filter(r => !r.error);
  const total = valid.length;
  const batchSize = parseInt(document.getElementById('batch-size').value) || 50;
  const delay = parseInt(document.getElementById('send-delay').value) || 300;
  const updateSheet = document.getElementById('opt-update-sheet').checked;
  const statusColName = 'Email Sent';

  state.isSending = true;
  state.stopRequested = false;
  state.campaignStartTime = Date.now();
  state.sendStats = { sent: 0, failed: 0, skipped: 0, errors: [] };

  document.getElementById('prog-total').textContent = total;
  document.getElementById('sending-status').textContent = 'Sending...';

  const campaignId = `campaign_${Date.now()}`;
  state.currentCampaignId = campaignId;

  // Ensure "Email Sent" column exists in sheet
  let statusColIndex = -1;
  if (updateSheet) {
    statusColIndex = state.sheetHeaders.indexOf(statusColName);
    if (statusColIndex === -1) {
      statusColIndex = state.sheetHeaders.length;
      state.sheetHeaders.push(statusColName);
      // Write header
      await API.updateSheetValues(
        state.selectedSpreadsheet.id,
        `'${state.selectedSheetName}'!${colLetter(statusColIndex + 1)}1`,
        [[statusColName]]
      );
    }
  }

  for (let i = 0; i < valid.length; i++) {
    if (state.stopRequested) {
      addLog('⛔ Sending stopped by user.', 'info');
      break;
    }

    const r = valid[i];
    const pct = Math.round(((state.sendStats.sent + state.sendStats.failed + state.sendStats.skipped) / total) * 100);
    document.getElementById('progress-bar').style.width = pct + '%';
    document.getElementById('sending-status').textContent = `Sending ${i + 1} of ${total}...`;

    try {
      const raw = MergeEngine.encodeEmail({
        to: r.to,
        from: r.from,
        replyTo: r.replyTo,
        cc: r.cc,
        bcc: r.bcc,
        subject: r.subject,
        htmlBody: r.htmlBody,
        textBody: r.textBody,
      });

      await API.sendEmail(raw);
      state.sendStats.sent++;
      document.getElementById('prog-sent').textContent = state.sendStats.sent;
      addLog(`✓ Sent to ${r.to}`);

      // Update sheet status
      if (updateSheet && statusColIndex >= 0) {
        const cellRange = `'${state.selectedSheetName}'!${colLetter(statusColIndex + 1)}${r.rowIndex}`;
        await API.updateSheetValues(
          state.selectedSpreadsheet.id,
          cellRange,
          [[`Email Sent ${new Date().toLocaleString()}`]]
        ).catch(() => {}); // Non-fatal
      }

      // Increment daily counter
      incrementDailyCounter();
    } catch (err) {
      state.sendStats.failed++;
      document.getElementById('prog-failed').textContent = state.sendStats.failed;
      const errMsg = err.message || String(err);
      state.sendStats.errors.push({ email: r.to, error: errMsg });
      addLog(`✗ Failed to ${r.to}: ${errMsg}`, 'error');
    }

    // Rate limiting: pause between batches
    if ((i + 1) % batchSize === 0 && i + 1 < valid.length) {
      addLog(`Pausing between batch... (batch size: ${batchSize})`, 'info');
      await sleep(2000);
    } else if (delay > 0) {
      await sleep(delay);
    }
  }

  document.getElementById('progress-bar').style.width = '100%';
  state.isSending = false;

  // Save campaign record
  const duration = Math.round((Date.now() - state.campaignStartTime) / 1000);
  const campaignRecord = {
    id: campaignId,
    name: document.getElementById('campaign-name').value || 'Campaign ' + new Date().toLocaleDateString(),
    spreadsheetId: state.selectedSpreadsheet.id,
    spreadsheetName: state.selectedSpreadsheet.name,
    draftSubject: state.selectedDraft.subject,
    sentAt: new Date().toISOString(),
    stats: { ...state.sendStats, total, duration },
  };
  saveCampaign(campaignRecord);

  showReport(total, duration);
}

function scheduleMailMerge() {
  const dt = document.getElementById('schedule-date').value;
  if (!dt) { showToast('Please select a date and time', 'error'); return; }
  const fireAt = new Date(dt).getTime();
  if (fireAt <= Date.now()) { showToast('Schedule time must be in the future', 'error'); return; }

  const alarmName = `mailmerge_${Date.now()}`;
  const config = captureCurrentConfig();
  chrome.storage.local.get(['scheduledMerges'], data => {
    const scheduled = data.scheduledMerges || [];
    scheduled.push({ alarmName, fireAt, config, name: config.campaignName });
    chrome.storage.local.set({ scheduledMerges: scheduled });
  });

  chrome.alarms.create(alarmName, { when: fireAt });
  showToast(`Mail merge scheduled for ${new Date(dt).toLocaleString()}`, 'success');

  // Update scheduled count badge
  chrome.storage.local.get(['scheduledCount'], d => {
    chrome.storage.local.set({ scheduledCount: (d.scheduledCount || 0) + 1 });
  });

  // Go to scheduled view
  switchView('scheduled');
  loadScheduledView();
}

function captureCurrentConfig() {
  return {
    campaignName: document.getElementById('campaign-name').value,
    spreadsheetId: state.selectedSpreadsheet?.id,
    spreadsheetName: state.selectedSpreadsheet?.name,
    sheetName: state.selectedSheetName,
    emailColumn: state.emailColumn,
    draftId: state.selectedDraft?.id,
    fromName: document.getElementById('from-name').value,
    sendAs: document.getElementById('send-as-select').value,
    replyTo: document.getElementById('reply-to').value,
    replyToName: document.getElementById('reply-to-name').value,
    cc: document.getElementById('cc-field').value,
    bcc: document.getElementById('bcc-field').value,
    batchSize: parseInt(document.getElementById('batch-size').value) || 50,
    sendDelay: parseInt(document.getElementById('send-delay').value) || 300,
    skipSent: document.getElementById('opt-skip-sent').checked,
    updateSheet: document.getElementById('opt-update-sheet').checked,
    trackOpens: document.getElementById('opt-track-opens').checked,
    trackClicks: document.getElementById('opt-track-clicks').checked,
    addUnsubscribe: document.getElementById('opt-unsubscribe').checked,
    trackingServerUrl: document.getElementById('tracking-server-url')?.value || '',
    unsubscribeText: document.getElementById('unsubscribe-text').value,
  };
}

function showReport(total, duration) {
  goToStep(5);
  const { sent, failed, skipped, errors } = state.sendStats;
  document.getElementById('report-sent').textContent = sent;
  document.getElementById('report-failed').textContent = failed;
  document.getElementById('report-skipped').textContent = skipped;
  document.getElementById('report-duration').textContent = duration < 60 ? `${duration}s` : `${Math.round(duration / 60)}m`;

  const iconEl = document.getElementById('report-icon');
  const titleEl = document.getElementById('report-title');
  const subEl = document.getElementById('report-subtitle');

  if (failed === 0) {
    iconEl.className = 'report-icon success';
    iconEl.textContent = '✓';
    titleEl.textContent = 'Mail Merge Complete!';
    subEl.textContent = `Successfully sent ${sent} email${sent !== 1 ? 's' : ''}.`;
  } else if (sent > 0) {
    iconEl.className = 'report-icon partial';
    iconEl.textContent = '!';
    titleEl.textContent = 'Partially Complete';
    subEl.textContent = `${sent} sent, ${failed} failed.`;
  } else {
    iconEl.className = 'report-icon error';
    iconEl.textContent = '✕';
    titleEl.textContent = 'Sending Failed';
    subEl.textContent = 'No emails were sent. Check errors below.';
  }

  if (errors.length) {
    const errDiv = document.getElementById('report-errors');
    errDiv.classList.remove('hidden');
    const ul = document.getElementById('error-list');
    ul.innerHTML = '';
    errors.slice(0, 20).forEach(e => {
      const li = document.createElement('li');
      li.textContent = `${e.email}: ${e.error}`;
      ul.appendChild(li);
    });
  } else {
    document.getElementById('report-errors').classList.add('hidden');
  }
}

document.getElementById('btn-stop-sending').addEventListener('click', () => {
  state.stopRequested = true;
  document.getElementById('sending-status').textContent = 'Stopping after current email...';
  document.getElementById('btn-stop-sending').disabled = true;
});

document.getElementById('btn-open-sheet').addEventListener('click', () => {
  if (state.selectedSpreadsheet) {
    window.open(`https://docs.google.com/spreadsheets/d/${state.selectedSpreadsheet.id}`, '_blank');
  }
});

document.getElementById('btn-new-campaign').addEventListener('click', resetWizard);

// ── Progress Log ───────────────────────────────────────────────────────────
function addLog(text, type = '') {
  const log = document.getElementById('progress-log');
  const line = document.createElement('div');
  line.className = `log-line ${type}`;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${text}`;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

// ── Campaigns View ─────────────────────────────────────────────────────────
function loadCampaignsView() {
  const list = document.getElementById('campaigns-list');
  if (!state.campaigns.length) {
    list.innerHTML = `<div class="empty-state">
      <div class="empty-icon">📭</div>
      <p>No campaigns yet. Start your first mail merge!</p>
      <button class="btn btn-primary" onclick="document.querySelector('[data-view=new-campaign]').click()">Start Mail Merge</button>
    </div>`;
    return;
  }

  list.innerHTML = '';
  [...state.campaigns].sort((a, b) => new Date(b.sentAt) - new Date(a.sentAt)).forEach(c => {
    const card = document.createElement('div');
    card.className = 'campaign-card';
    card.innerHTML = `
      <div class="campaign-info">
        <div class="campaign-name">${esc(c.name)}</div>
        <div class="campaign-meta">
          ${esc(c.spreadsheetName)} · ${new Date(c.sentAt).toLocaleDateString()} ·
          Draft: ${esc(c.draftSubject || 'Unknown')}
        </div>
      </div>
      <div class="campaign-stats">
        <div class="campaign-stat"><span>${c.stats?.sent || 0}</span><label>Sent</label></div>
        <div class="campaign-stat"><span>${c.stats?.failed || 0}</span><label>Failed</label></div>
        <div class="campaign-stat"><span>${c.stats?.total || 0}</span><label>Total</label></div>
      </div>
      <button class="btn btn-sm btn-secondary" onclick="window.open('https://docs.google.com/spreadsheets/d/${c.spreadsheetId}','_blank')">View Sheet</button>`;
    list.appendChild(card);
  });
}

// ── Scheduled View ─────────────────────────────────────────────────────────
function loadScheduledView() {
  chrome.storage.local.get(['scheduledMerges'], data => {
    const scheduled = (data.scheduledMerges || []).filter(s => s.fireAt > Date.now());
    const list = document.getElementById('scheduled-list');
    if (!scheduled.length) {
      list.innerHTML = `<div class="empty-state"><div class="empty-icon">🗓</div><p>No scheduled campaigns.</p></div>`;
      return;
    }
    list.innerHTML = '';
    scheduled.forEach(s => {
      const card = document.createElement('div');
      card.className = 'campaign-card';
      card.innerHTML = `
        <div class="campaign-info">
          <div class="campaign-name">${esc(s.config?.campaignName || 'Scheduled Campaign')}</div>
          <div class="campaign-meta">Scheduled for ${new Date(s.fireAt).toLocaleString()} · Sheet: ${esc(s.config?.spreadsheetName || 'Unknown')}</div>
        </div>
        <button class="btn btn-sm btn-danger" data-alarm="${s.alarmName}">Cancel</button>`;
      card.querySelector('button').addEventListener('click', () => cancelScheduled(s.alarmName));
      list.appendChild(card);
    });
  });
}

function cancelScheduled(alarmName) {
  chrome.alarms.clear(alarmName, () => {
    chrome.storage.local.get(['scheduledMerges'], data => {
      const filtered = (data.scheduledMerges || []).filter(s => s.alarmName !== alarmName);
      chrome.storage.local.set({ scheduledMerges: filtered }, () => {
        loadScheduledView();
        showToast('Scheduled campaign cancelled');
      });
    });
  });
}

// ── Templates View ─────────────────────────────────────────────────────────
async function loadTemplatesView() {
  const list = document.getElementById('templates-list');
  list.innerHTML = '<div class="empty-state"><div class="empty-icon">⏳</div><p>Loading drafts...</p></div>';
  try {
    const drafts = await API.listDrafts();
    allDrafts = drafts;
    list.innerHTML = '';
    if (!drafts.length) {
      list.innerHTML = '<div class="empty-state"><div class="empty-icon">📝</div><p>No Gmail drafts found. Create a draft with {{Column}} placeholders.</p></div>';
      return;
    }
    drafts.forEach(draft => {
      const parsed = MergeEngine.parseDraftPayload(draft.message);
      const vars = MergeEngine.extractVariables((parsed.htmlBody || parsed.textBody) + ' ' + parsed.subject);
      const card = document.createElement('div');
      card.className = 'template-card';
      card.innerHTML = `
        <div class="template-subject">${esc(parsed.subject || '(No subject)')}</div>
        <div class="template-snippet">${esc(draft.message.snippet || '').slice(0, 120)}</div>
        <div class="template-vars">
          ${vars.map(v => `<span class="variable-chip matched">{{${esc(v)}}}</span>`).join('')}
        </div>`;
      card.addEventListener('click', () => {
        selectDraft(draft, parsed);
        switchView('new-campaign');
        showToast(`Template "${parsed.subject}" selected`);
      });
      list.appendChild(card);
    });
  } catch (err) {
    list.innerHTML = `<div class="empty-state"><div class="empty-icon">⚠</div><p>Error loading drafts: ${esc(err.message)}</p></div>`;
  }
}

document.getElementById('btn-refresh-all-drafts').addEventListener('click', loadTemplatesView);

// ── Quota View ─────────────────────────────────────────────────────────────
function updateQuotaView() {
  chrome.storage.local.get(['sentToday', 'lastResetDate', 'sendHistory'], data => {
    const today = new Date().toDateString();
    const sentToday = data.lastResetDate === today ? (data.sentToday || 0) : 0;
    const limit = state.settings.dailyQuota || 400;

    document.getElementById('quota-used').textContent = sentToday;
    document.getElementById('quota-limit').textContent = limit;
    document.getElementById('quota-gauge-fill').style.width = Math.min(100, (sentToday / limit) * 100) + '%';
    document.getElementById('daily-limit-input').value = limit;

    const badge = document.getElementById('quota-type-badge');
    badge.textContent = limit >= 2000 ? 'Google Workspace' : limit >= 1000 ? 'Workspace Basic' : 'Free Gmail';

    // Render 7-day history
    const history = data.sendHistory || [];
    const chart = document.getElementById('quota-history-chart');
    chart.innerHTML = '';
    const maxVal = Math.max(...history.map(h => h.count), 1);
    for (let i = 6; i >= 0; i--) {
      const date = new Date(Date.now() - i * 86400000);
      const entry = history.find(h => h.date === date.toDateString());
      const count = entry ? entry.count : 0;
      const wrap = document.createElement('div');
      wrap.className = 'history-bar-wrap';
      wrap.innerHTML = `
        <div class="history-bar" style="height:${Math.round((count / maxVal) * 64)}px" title="${count} sent"></div>
        <div class="history-label">${date.toLocaleDateString('en', { weekday: 'short' })}</div>`;
      chart.appendChild(wrap);
    }
  });
}

document.getElementById('btn-save-quota').addEventListener('click', () => {
  const val = parseInt(document.getElementById('daily-limit-input').value) || 400;
  saveSettings({ dailyQuota: val });
  updateQuotaView();
  showToast('Quota settings saved');
});

function incrementDailyCounter() {
  chrome.storage.local.get(['sentToday', 'lastResetDate', 'sendHistory'], data => {
    const today = new Date().toDateString();
    let sentToday = data.sentToday || 0;
    if (data.lastResetDate !== today) sentToday = 0;

    sentToday++;
    const history = data.sendHistory || [];
    const todayEntry = history.find(h => h.date === today);
    if (todayEntry) todayEntry.count = sentToday;
    else history.push({ date: today, count: sentToday });
    const trimmed = history.slice(-30); // keep 30 days

    chrome.storage.local.set({ sentToday, lastResetDate: today, sendHistory: trimmed });
  });
}

// ── Wizard Step Navigation ─────────────────────────────────────────────────
document.getElementById('btn-step1-next').addEventListener('click', () => {
  goToStep(2);
});

document.getElementById('btn-step2-back').addEventListener('click', () => goToStep(1));
document.getElementById('btn-step2-next').addEventListener('click', () => {
  goToStep(3);
  renderPreview();
});

document.getElementById('btn-step3-back').addEventListener('click', () => goToStep(2));

// ── Sidebar Navigation ─────────────────────────────────────────────────────
document.querySelectorAll('.sidebar-btn').forEach(btn => {
  btn.addEventListener('click', () => switchView(btn.dataset.view));
});

document.querySelectorAll('[data-view-trigger]').forEach(btn => {
  btn.addEventListener('click', () => switchView(btn.dataset.viewTrigger));
});

document.getElementById('btn-nav-settings').addEventListener('click', () => {
  chrome.runtime.openOptionsPage();
});

// ── Modal Helpers ──────────────────────────────────────────────────────────
function openModal(id) {
  document.getElementById('modal-overlay').classList.remove('hidden');
  document.getElementById(id).classList.remove('hidden');
}

function closeModal(id) {
  document.getElementById(id).classList.add('hidden');
  const anyOpen = document.querySelectorAll('.modal:not(.hidden)').length > 0;
  if (!anyOpen) document.getElementById('modal-overlay').classList.add('hidden');
}

document.querySelectorAll('.modal-close, [data-modal-close]').forEach(btn => {
  btn.addEventListener('click', () => {
    const id = btn.dataset.modal || btn.dataset.modalClose;
    closeModal(id);
  });
});

document.getElementById('modal-overlay').addEventListener('click', e => {
  if (e.target === document.getElementById('modal-overlay')) {
    document.querySelectorAll('.modal:not(.hidden)').forEach(m => closeModal(m.id));
  }
});

// ── Utility Helpers ────────────────────────────────────────────────────────
function hideDropdown(id) {
  document.getElementById(id).classList.add('hidden');
}

document.addEventListener('click', e => {
  if (!e.target.closest('.search-wrapper') && !e.target.closest('.dropdown-list')) {
    document.querySelectorAll('.dropdown-list').forEach(d => d.classList.add('hidden'));
  }
});

function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

let toastTimer = null;
function showToast(msg, type = '') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast ${type}`;
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.add('hidden'), 3500);
}

function resetWizard() {
  state.currentWizardStep = 1;
  state.selectedSpreadsheet = null;
  state.selectedSheetName = null;
  state.emailColumn = null;
  state.sheetHeaders = [];
  state.sheetRows = [];
  state.selectedDraft = null;
  state.mergedRecipients = [];
  state.sendStats = { sent: 0, failed: 0, skipped: 0, errors: [] };

  document.getElementById('campaign-name').value = '';
  document.getElementById('sheet-search').value = '';
  document.getElementById('draft-search').value = '';
  document.getElementById('selected-sheet').classList.add('hidden');
  document.getElementById('selected-draft').classList.add('hidden');
  document.getElementById('section-sheet-tab').classList.add('hidden');
  document.getElementById('section-email-col').classList.add('hidden');
  document.getElementById('btn-step1-next').disabled = true;
  document.getElementById('progress-log').innerHTML = '';

  goToStep(1);
  switchView('new-campaign');
}

// ── Boot ───────────────────────────────────────────────────────────────────
init();
