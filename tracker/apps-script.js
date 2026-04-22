/**
 * Mail Merge Tracking Server — Google Apps Script
 *
 * Deploy this as a Google Apps Script Web App:
 *   1. Go to https://script.google.com and create a new project.
 *   2. Paste this entire file as Code.gs.
 *   3. Click "Deploy" → "New deployment" → Type: "Web app".
 *   4. Execute as: Me | Who has access: Anyone.
 *   5. Copy the Web App URL and paste it into the extension's Settings
 *      under "Tracking Server URL".
 *
 * Endpoints:
 *   ?type=open&cid=CAMPAIGN_ID&email=EMAIL
 *      → Returns a 1x1 tracking pixel, records open in the spreadsheet.
 *   ?type=click&cid=CAMPAIGN_ID&email=EMAIL&url=ORIGINAL_URL
 *      → Redirects to the original URL and records click.
 *   ?type=status&cid=CAMPAIGN_ID
 *      → Returns JSON open/click counts for a campaign.
 *
 * The script writes to a "Tracking Log" sheet in a spreadsheet you specify.
 * Set TRACKING_SHEET_ID below to your spreadsheet's ID.
 */

// ─── CONFIGURATION ──────────────────────────────────────────────────────────
// No hardcoded sheet ID needed — the spreadsheet ID is passed as the `sid`
// URL parameter by the extension, so tracking logs go into the same sheet
// used for each campaign automatically.
const TRACKING_SHEET_NAME = 'Tracking Log';

// ─── MAIN HANDLER ───────────────────────────────────────────────────────────
function doGet(e) {
  const params = e.parameter || {};
  const type = params.type || '';
  const campaignId = params.cid || '';
  const email = decodeURIComponent(params.email || '');
  const redirectUrl = decodeURIComponent(params.url || '');
  const spreadsheetId = params.sid || '';  // ID of the campaign's own spreadsheet

  try {
    if (type === 'open') {
      recordEvent(spreadsheetId, campaignId, email, 'open');
      return trackingPixelResponse();
    }

    if (type === 'click') {
      recordEvent(spreadsheetId, campaignId, email, 'click', redirectUrl);
      return HtmlService.createHtmlOutput(redirectScript(redirectUrl));
    }


    if (type === 'status') {
      return ContentService
        .createTextOutput(JSON.stringify(getCampaignStats(spreadsheetId, campaignId)))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (err) {
    console.error('Tracking error:', err);
  }

  // Default: pixel
  return trackingPixelResponse();
}

// ─── EVENT LOGGING ───────────────────────────────────────────────────────────
function recordEvent(spreadsheetId, campaignId, email, eventType, url) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRACKING_SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Campaign ID', 'Email', 'Event', 'URL', 'IP']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date().toISOString(),
    campaignId,
    email,
    eventType,
    url || '',
    '',
  ]);
}


function getCampaignStats(spreadsheetId, campaignId) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(TRACKING_SHEET_NAME);
    if (!sheet) return { opens: 0, clicks: 0 };
    const data = sheet.getDataRange().getValues();
    let opens = 0, clicks = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === campaignId) {
        const ev = data[i][3];
        if (ev === 'open') opens++;
        else if (ev === 'click') clicks++;
      }
    }
    return { opens, clicks };
  } catch {
    return { opens: 0, clicks: 0 };
  }
}

// ─── RESPONSES ───────────────────────────────────────────────────────────────
function trackingPixelResponse() {
  // 1x1 transparent GIF
  const pixel = Utilities.base64Decode(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
  );
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON); // Apps Script can't return binary; pixel via noimg
}

function redirectScript(url) {
  const safe = url.replace(/"/g, '&quot;').replace(/</g, '&lt;');
  return `<!DOCTYPE html><html><head><script>window.location.href="${safe}";<\/script></head>
<body><a href="${safe}">Click here if not redirected</a></body></html>`;
}


