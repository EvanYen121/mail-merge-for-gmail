'use strict';

const SETTINGS_KEY = 'mailMergeSettings';

const fields = [
  'default-from-name',
  'default-reply-to',
  'daily-quota',
  'default-delay',
  'default-batch-size',
  'tracking-server-url',
  'unsubscribe-text',
  'unsubscribe-style',
  'status-col-name',
];

function loadSettings() {
  chrome.storage.local.get([SETTINGS_KEY], data => {
    const s = data[SETTINGS_KEY] || {};
    document.getElementById('default-from-name').value = s.defaultFromName || '';
    document.getElementById('default-reply-to').value = s.defaultReplyTo || '';
    document.getElementById('daily-quota').value = s.dailyQuota || 400;
    document.getElementById('default-delay').value = s.defaultDelay !== undefined ? s.defaultDelay : 300;
    document.getElementById('default-batch-size').value = s.defaultBatchSize || 50;
    document.getElementById('tracking-server-url').value = s.trackingServerUrl || '';
    document.getElementById('unsubscribe-text').value = s.unsubscribeText || 'Unsubscribe from this list';
    document.getElementById('unsubscribe-style').value = s.unsubscribeStyle || 'standard';
    document.getElementById('status-col-name').value = s.statusColName || 'Email Sent';
  });
}

function saveSettings() {
  const s = {
    defaultFromName: document.getElementById('default-from-name').value.trim(),
    defaultReplyTo: document.getElementById('default-reply-to').value.trim(),
    dailyQuota: parseInt(document.getElementById('daily-quota').value) || 400,
    defaultDelay: parseInt(document.getElementById('default-delay').value) || 300,
    defaultBatchSize: parseInt(document.getElementById('default-batch-size').value) || 50,
    trackingServerUrl: document.getElementById('tracking-server-url').value.trim(),
    unsubscribeText: document.getElementById('unsubscribe-text').value.trim(),
    unsubscribeStyle: document.getElementById('unsubscribe-style').value,
    statusColName: document.getElementById('status-col-name').value.trim() || 'Email Sent',
  };
  chrome.storage.local.set({ [SETTINGS_KEY]: s }, () => showToast('Settings saved!'));
}

function loadAccount() {
  chrome.identity.getAuthToken({ interactive: false }, async token => {
    const row = document.getElementById('account-row');
    if (chrome.runtime.lastError || !token) {
      row.innerHTML = '<div class="user-info"><div class="name">Not signed in</div></div>';
      return;
    }
    try {
      const res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
        headers: { Authorization: `Bearer ${token}` }
      });
      const user = await res.json();
      row.innerHTML = `
        ${user.picture ? `<img class="avatar" src="${user.picture}" alt="">` : ''}
        <div class="user-info">
          <div class="name">${esc(user.name || '')}</div>
          <div class="email">${esc(user.email || '')}</div>
        </div>`;
    } catch {
      row.innerHTML = '<div class="user-info"><div class="name">Signed in</div></div>';
    }
  });
}

document.getElementById('btn-signin').addEventListener('click', () => {
  chrome.identity.getAuthToken({ interactive: true }, () => loadAccount());
});

document.getElementById('btn-signout').addEventListener('click', () => {
  chrome.identity.getAuthToken({ interactive: false }, token => {
    if (token) {
      fetch(`https://accounts.google.com/o/oauth2/revoke?token=${token}`);
      chrome.identity.removeCachedAuthToken({ token }, () => loadAccount());
    }
  });
});

document.getElementById('btn-save').addEventListener('click', saveSettings);

document.getElementById('btn-clear-history').addEventListener('click', () => {
  if (!confirm('Clear all campaign history? This cannot be undone.')) return;
  chrome.storage.local.set({ campaigns: [], sendHistory: [], sentToday: 0 }, () => {
    showToast('Campaign history cleared.');
  });
});

document.getElementById('btn-clear-all').addEventListener('click', () => {
  if (!confirm('Reset ALL extension data? This cannot be undone.')) return;
  chrome.storage.local.clear(() => showToast('All data cleared. Extension reset.'));
});

document.getElementById('link-setup-guide').addEventListener('click', e => {
  e.preventDefault();
  alert('To set up the tracking server:\n\n1. Go to script.google.com\n2. Create a new project\n3. Paste the code from tracker/apps-script.js\n4. Deploy as Web App (Execute as: Me, Access: Anyone)\n5. Copy the deployment URL and paste it here.');
});

let toastTimer;
function showToast(msg) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.display = 'block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.style.display = 'none'; }, 3000);
}

function esc(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

loadSettings();
loadAccount();
