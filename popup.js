'use strict';

const Views = {
  LOADING: 'view-loading',
  SIGNED_OUT: 'view-signed-out',
  SIGNED_IN: 'view-signed-in',
};

function showView(viewId) {
  document.querySelectorAll('.view').forEach(v => v.classList.add('hidden'));
  document.getElementById(viewId).classList.remove('hidden');
}

async function loadUserInfo(token) {
  const res = await fetch('https://www.googleapis.com/oauth2/v2/userinfo', {
    headers: { Authorization: `Bearer ${token}` }
  });
  return res.ok ? res.json() : null;
}

async function getStats() {
  return new Promise(resolve => {
    chrome.storage.local.get(['sentToday', 'scheduledCount', 'lastResetDate'], data => {
      const today = new Date().toDateString();
      if (data.lastResetDate !== today) {
        chrome.storage.local.set({ sentToday: 0, lastResetDate: today });
        resolve({ sentToday: 0, scheduledCount: data.scheduledCount || 0 });
      } else {
        resolve({
          sentToday: data.sentToday || 0,
          scheduledCount: data.scheduledCount || 0
        });
      }
    });
  });
}

async function init() {
  showView(Views.LOADING);

  chrome.identity.getAuthToken({ interactive: false }, async token => {
    if (chrome.runtime.lastError || !token) {
      showView(Views.SIGNED_OUT);
      return;
    }

    const userInfo = await loadUserInfo(token);
    if (!userInfo) {
      showView(Views.SIGNED_OUT);
      return;
    }

    document.getElementById('user-name').textContent = userInfo.name || '';
    document.getElementById('user-email').textContent = userInfo.email || '';
    if (userInfo.picture) {
      document.getElementById('user-avatar').src = userInfo.picture;
    }

    const stats = await getStats();
    document.getElementById('stat-sent-today').textContent = stats.sentToday;
    document.getElementById('stat-scheduled').textContent = stats.scheduledCount;

    chrome.storage.local.get(['dailyQuota'], data => {
      const quota = (data.dailyQuota || 400) - stats.sentToday;
      document.getElementById('stat-quota').textContent = Math.max(0, quota);
    });

    showView(Views.SIGNED_IN);
  });
}

document.getElementById('btn-signin').addEventListener('click', () => {
  showView(Views.LOADING);
  chrome.identity.getAuthToken({ interactive: true }, async token => {
    if (chrome.runtime.lastError || !token) {
      showView(Views.SIGNED_OUT);
      return;
    }
    await init();
  });
});

document.getElementById('btn-signout').addEventListener('click', () => {
  chrome.identity.getAuthToken({ interactive: false }, token => {
    if (token) {
      fetch(`https://accounts.google.com/o/oauth2/revoke?token=${token}`);
      chrome.identity.removeCachedAuthToken({ token }, () => {
        showView(Views.SIGNED_OUT);
      });
    } else {
      showView(Views.SIGNED_OUT);
    }
  });
});

document.getElementById('btn-open-app').addEventListener('click', () => {
  chrome.tabs.create({ url: chrome.runtime.getURL('app.html') });
  window.close();
});

document.getElementById('btn-view-campaigns').addEventListener('click', () => {
  chrome.tabs.create({ url: chrome.runtime.getURL('app.html?view=campaigns') });
  window.close();
});

document.getElementById('btn-settings').addEventListener('click', () => {
  chrome.runtime.openOptionsPage();
  window.close();
});

init();
