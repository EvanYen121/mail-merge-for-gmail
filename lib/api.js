'use strict';

export class GoogleAPI {
  constructor() {
    this.token = null;
  }

  async getToken(interactive = false) {
    return new Promise((resolve, reject) => {
      chrome.identity.getAuthToken({ interactive }, token => {
        if (chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
        } else {
          this.token = token;
          resolve(token);
        }
      });
    });
  }

  async request(url, options = {}) {
    if (!this.token) await this.getToken(true);
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${this.token}`,
        'Content-Type': 'application/json',
        ...(options.headers || {})
      }
    });
    if (res.status === 401) {
      // Token expired — remove and retry once
      await new Promise(r => chrome.identity.removeCachedAuthToken({ token: this.token }, r));
      this.token = null;
      await this.getToken(false);
      const retry = await fetch(url, {
        ...options,
        headers: {
          Authorization: `Bearer ${this.token}`,
          'Content-Type': 'application/json',
          ...(options.headers || {})
        }
      });
      if (!retry.ok) throw new Error(`API error ${retry.status}: ${await retry.text()}`);
      return retry.json();
    }
    if (!res.ok) {
      const body = await res.text();
      throw new Error(`API error ${res.status}: ${body}`);
    }
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  // ─── User Info ──────────────────────────────────────────────────────────────
  async getUserInfo() {
    return this.request('https://www.googleapis.com/oauth2/v2/userinfo');
  }

  async getSendAsAliases() {
    return this.request('https://gmail.googleapis.com/gmail/v1/users/me/settings/sendAs');
  }

  // ─── Drive / Sheets List ────────────────────────────────────────────────────
  async listSpreadsheets(pageToken = null) {
    let url = "https://www.googleapis.com/drive/v3/files?q=mimeType%3D'application%2Fvnd.google-apps.spreadsheet'&orderBy=modifiedTime%20desc&pageSize=50&fields=files(id%2Cname%2CmodifiedTime)%2CnextPageToken";
    if (pageToken) url += `&pageToken=${pageToken}`;
    return this.request(url);
  }

  async searchSpreadsheets(query) {
    const q = encodeURIComponent(`mimeType='application/vnd.google-apps.spreadsheet' and name contains '${query}'`);
    const url = `https://www.googleapis.com/drive/v3/files?q=${q}&orderBy=modifiedTime desc&pageSize=20&fields=files(id,name,modifiedTime)`;
    return this.request(url);
  }

  // ─── Sheets API ─────────────────────────────────────────────────────────────
  async getSpreadsheetMeta(spreadsheetId) {
    return this.request(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=spreadsheetId,properties,sheets.properties`);
  }

  async getSheetValues(spreadsheetId, range) {
    return this.request(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`);
  }

  async updateSheetValues(spreadsheetId, range, values) {
    return this.request(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`,
      { method: 'PUT', body: JSON.stringify({ values }) }
    );
  }

  async batchUpdateSheetValues(spreadsheetId, data) {
    return this.request(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
      { method: 'POST', body: JSON.stringify({ valueInputOption: 'USER_ENTERED', data }) }
    );
  }

  async appendSheetRow(spreadsheetId, range, values) {
    return this.request(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      { method: 'POST', body: JSON.stringify({ values: [values] }) }
    );
  }

  // ─── Gmail Drafts ───────────────────────────────────────────────────────────
  async listDrafts() {
    const data = await this.request('https://gmail.googleapis.com/gmail/v1/users/me/drafts?maxResults=100&fields=drafts(id,message(id,snippet))');
    const drafts = data.drafts || [];
    // Fetch details for each draft (subject + snippet)
    const detailed = await Promise.all(
      drafts.slice(0, 50).map(d =>
        this.request(`https://gmail.googleapis.com/gmail/v1/users/me/drafts/${d.id}?format=full`)
      )
    );
    return detailed;
  }

  async getDraft(draftId) {
    return this.request(`https://gmail.googleapis.com/gmail/v1/users/me/drafts/${draftId}?format=full`);
  }

  // ─── Gmail Send ─────────────────────────────────────────────────────────────
  async sendEmail(rawBase64) {
    return this.request(
      'https://gmail.googleapis.com/gmail/v1/users/me/messages/send',
      { method: 'POST', body: JSON.stringify({ raw: rawBase64 }) }
    );
  }

  async createDraft(rawBase64) {
    return this.request(
      'https://gmail.googleapis.com/gmail/v1/users/me/drafts',
      { method: 'POST', body: JSON.stringify({ message: { raw: rawBase64 } }) }
    );
  }

  // ─── Gmail Labels / Threads ─────────────────────────────────────────────────
  async listLabels() {
    return this.request('https://gmail.googleapis.com/gmail/v1/users/me/labels');
  }

  async getMessageThread(threadId) {
    return this.request(`https://gmail.googleapis.com/gmail/v1/users/me/threads/${threadId}?format=metadata`);
  }
}

export default new GoogleAPI();
