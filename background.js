/**
 * background.js — Service worker for Mail Merge for Gmail
 * Handles alarm-based scheduled sends and badge updates.
 */

import { GoogleAPI } from './lib/api.js';
import { MergeEngine } from './lib/merge.js';

// ── Alarm Handler (Scheduled Mail Merges) ──────────────────────────────────
chrome.alarms.onAlarm.addListener(async alarm => {
  if (!alarm.name.startsWith('mailmerge_')) return;

  const data = await storageGet(['scheduledMerges', 'mailMergeSettings']);
  const scheduled = data.scheduledMerges || [];
  const job = scheduled.find(s => s.alarmName === alarm.name);
  if (!job) return;

  // Remove from scheduled list
  const remaining = scheduled.filter(s => s.alarmName !== alarm.name);
  await storageSet({ scheduledMerges: remaining });
  updateScheduledBadge(remaining.filter(s => s.fireAt > Date.now()).length);

  // Notify user
  chrome.notifications.create(`${alarm.name}_start`, {
    type: 'basic',
    iconUrl: 'icons/icon48.svg',
    title: 'Mail Merge Starting',
    message: `Sending "${job.config?.campaignName || 'campaign'}"...`,
  });

  try {
    const result = await runScheduledMerge(job.config, data.mailMergeSettings || {});
    chrome.notifications.create(`${alarm.name}_done`, {
      type: 'basic',
      iconUrl: 'icons/icon48.svg',
      title: 'Mail Merge Complete',
      message: `Sent ${result.sent} emails. ${result.failed ? result.failed + ' failed.' : ''}`,
    });
  } catch (err) {
    chrome.notifications.create(`${alarm.name}_err`, {
      type: 'basic',
      iconUrl: 'icons/icon48.svg',
      title: 'Mail Merge Failed',
      message: err.message,
    });
  }
});

async function runScheduledMerge(config, settings) {
  const api = new GoogleAPI();
  await api.getToken(false);

  // Load sheet data
  const sheetData = await api.getSheetValues(config.spreadsheetId, `'${config.sheetName}'`);
  const rows = sheetData.values || [];
  if (rows.length < 2) throw new Error('Sheet is empty or has no data rows.');

  const headers = rows[0];
  const dataRows = rows.slice(1).map((row, i) => {
    const obj = { _rowIndex: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j] || ''; });
    return obj;
  });

  // Load draft
  const draft = await api.getDraft(config.draftId);
  const parsed = MergeEngine.parseDraftPayload(draft.message);

  let sent = 0, failed = 0, skipped = 0;
  const statusColName = 'Email Sent';
  let statusColIndex = headers.indexOf(statusColName);
  if (statusColIndex === -1) {
    statusColIndex = headers.length;
    await api.updateSheetValues(
      config.spreadsheetId,
      `'${config.sheetName}'!${colLetter(statusColIndex + 1)}1`,
      [[statusColName]]
    );
  }

  for (const row of dataRows) {
    const email = (row[config.emailColumn] || '').trim();
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;

    const sentStatus = row[statusColName];
    if (config.skipSent && sentStatus) { skipped++; continue; }

    const vars = {};
    headers.forEach(h => { vars[h] = row[h] || ''; });

    const mergedSubject = MergeEngine.process(parsed.subject, vars);
    const mergedHtml = MergeEngine.process(parsed.htmlBody || parsed.textBody, vars, {
      unsubscribeUrl: config.addUnsubscribe && config.trackingServerUrl
        ? `${config.trackingServerUrl}?type=unsubscribe&email=${encodeURIComponent(email)}`
        : null,
      unsubscribeText: config.unsubscribeText,
    });
    const mergedText = MergeEngine.process(parsed.textBody, vars);

    const fromStr = config.fromName ? `${config.fromName} <${config.sendAs}>` : config.sendAs;
    const replyToStr = config.replyTo
      ? (config.replyToName ? `${config.replyToName} <${config.replyTo}>` : config.replyTo)
      : null;

    try {
      const raw = MergeEngine.encodeEmail({
        to: email,
        from: fromStr,
        replyTo: replyToStr,
        cc: MergeEngine.process(config.cc || '', vars) || null,
        bcc: MergeEngine.process(config.bcc || '', vars) || null,
        subject: mergedSubject,
        htmlBody: mergedHtml,
        textBody: mergedText,
      });

      await api.sendEmail(raw);
      sent++;

      if (config.updateSheet) {
        const cellRange = `'${config.sheetName}'!${colLetter(statusColIndex + 1)}${row._rowIndex}`;
        await api.updateSheetValues(
          config.spreadsheetId, cellRange,
          [[`Email Sent ${new Date().toLocaleString()}`]]
        ).catch(() => {});
      }

      await incrementDailyCounter();
    } catch (err) {
      failed++;
      console.error(`Failed to send to ${email}:`, err.message);
    }

    if (config.sendDelay > 0) await sleep(config.sendDelay);
  }

  return { sent, failed, skipped };
}

// ── Extension Lifecycle ────────────────────────────────────────────────────
chrome.runtime.onInstalled.addListener(details => {
  if (details.reason === 'install') {
    chrome.storage.local.set({
      mailMergeSettings: {
        dailyQuota: 400,
        trackingServerUrl: '',
        defaultFromName: '',
        defaultReplyTo: '',
      },
      campaigns: [],
      scheduledMerges: [],
      sentToday: 0,
      sendHistory: [],
      lastResetDate: new Date().toDateString(),
    });
    // Open onboarding
    chrome.tabs.create({ url: chrome.runtime.getURL('app.html') });
  }
});

// Re-register alarms on service worker restart (they persist but listener needs re-attach)
chrome.runtime.onStartup.addListener(async () => {
  const data = await storageGet(['scheduledMerges']);
  const scheduled = (data.scheduledMerges || []).filter(s => s.fireAt > Date.now());
  updateScheduledBadge(scheduled.length);
});

// ── Badge ──────────────────────────────────────────────────────────────────
function updateScheduledBadge(count) {
  if (count > 0) {
    chrome.action.setBadgeText({ text: String(count) });
    chrome.action.setBadgeBackgroundColor({ color: '#1a73e8' });
  } else {
    chrome.action.setBadgeText({ text: '' });
  }
}

// ── Helpers ────────────────────────────────────────────────────────────────
function storageGet(keys) {
  return new Promise(resolve => chrome.storage.local.get(keys, resolve));
}

function storageSet(obj) {
  return new Promise(resolve => chrome.storage.local.set(obj, resolve));
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

async function incrementDailyCounter() {
  const data = await storageGet(['sentToday', 'lastResetDate', 'sendHistory']);
  const today = new Date().toDateString();
  let sentToday = data.sentToday || 0;
  if (data.lastResetDate !== today) sentToday = 0;
  sentToday++;

  const history = data.sendHistory || [];
  const todayEntry = history.find(h => h.date === today);
  if (todayEntry) todayEntry.count = sentToday;
  else history.push({ date: today, count: sentToday });

  await storageSet({ sentToday, lastResetDate: today, sendHistory: history.slice(-30) });
}
