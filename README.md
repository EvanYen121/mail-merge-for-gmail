# Mail Merge for Gmail

A full-featured Chrome extension equivalent to YAMM (Yet Another Mail Merge). Send personalized bulk emails using Gmail drafts and Google Sheets.

---

## Features

| Feature | Supported |
|---|---|
| Gmail draft as email template | ✅ |
| `{{Column Name}}` merge fields | ✅ |
| Google Sheets as recipient list | ✅ |
| Send via Gmail API (your account) | ✅ |
| Write "Email Sent" status back to sheet | ✅ |
| Skip already-sent rows | ✅ |
| CC / BCC (static or from column) | ✅ |
| Send-as alias support | ✅ |
| Reply-To override | ✅ |
| Test email (send to yourself first) | ✅ |
| Preview before sending | ✅ |
| Scheduled sends | ✅ |
| Batch size + delay rate limiting | ✅ |
| Campaign history | ✅ |
| Daily quota tracking | ✅ |
| Email open tracking | ✅ (needs tracker server) |
| Link click tracking | ✅ (needs tracker server) |
| Unsubscribe link | ✅ (needs tracker server) |

---

## Installation

### 1. Create a Google Cloud Project & OAuth Client

1. Go to [console.cloud.google.com](https://console.cloud.google.com).
2. Create a new project (e.g. "Mail Merge Extension").
3. Enable these APIs:
   - **Gmail API**
   - **Google Sheets API**
   - **Google Drive API**
4. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**.
5. Application type: **Chrome Extension**.
6. Get your extension's ID from `chrome://extensions` (after loading unpacked).
7. Paste the extension ID and click Create.
8. Copy the **Client ID** (looks like `1234567890-abc.apps.googleusercontent.com`).

### 2. Update manifest.json

Open `manifest.json` and replace:
```json
"client_id": "YOUR_GOOGLE_OAUTH_CLIENT_ID.apps.googleusercontent.com"
```
with your actual Client ID.

### 3. Load the Extension

1. Open Chrome → `chrome://extensions`
2. Enable **Developer Mode** (top right)
3. Click **Load unpacked**
4. Select this `Mail merge` folder

### 4. Sign In

Click the Mail Merge icon in the Chrome toolbar → Sign in with Google → authorize all requested permissions.

---

## Usage

### Basic Mail Merge

1. **Create your template**: In Gmail, compose a **Draft** (don't send it).
   - Use `{{Column Name}}` anywhere in the subject or body to insert data.
   - Example: `Hi {{First Name}}, your order {{Order ID}} is ready!`

2. **Prepare your spreadsheet**: In Google Sheets:
   - Row 1 = headers (these become merge field names)
   - One column must contain email addresses (e.g., `Email`)
   - Example columns: `Email | First Name | Company | Order ID`

3. **Click the extension icon** → Start Mail Merge

4. **Step 1 – Setup**:
   - Search for or paste your spreadsheet URL
   - Select the sheet tab
   - Select the email column
   - Search for your Gmail draft

5. **Step 2 – Configure**:
   - Set From Name, Reply-To, CC/BCC
   - Choose Send Now or Schedule
   - Configure batch size / delay

6. **Step 3 – Preview**: Review the merged emails before sending.

7. **Send**: Confirm and send. Progress is shown in real time.

### Status Column

After sending, a column called **"Email Sent"** is added to your sheet with a timestamp for each sent row. Rows with this value set are skipped in future merges (if "Skip already sent" is checked).

---

## Email Open & Click Tracking (Optional)

Tracking requires a free server hosted on Google Apps Script.

### Setup Tracking Server

1. Go to [script.google.com](https://script.google.com) → New Project.
2. Copy the contents of `tracker/apps-script.js` into the editor.
3. Set `TRACKING_SHEET_ID` to the ID of a Google Sheet where you want logs stored.
4. Click **Deploy → New deployment → Web App**.
5. Set:
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click Deploy → copy the URL.
7. Paste the URL in **Settings → Tracking Server URL**.

### How Tracking Works

- **Open tracking**: A 1×1 invisible image is added to each email. When the recipient opens the email, the image loads from your Apps Script server, which logs the open.
- **Click tracking**: Links are rewritten to pass through your server before redirecting to the original URL.
- **Unsubscribe**: An unsubscribe link is appended to the email footer. Clicking it shows a confirmation page and logs the unsubscribe.

---

## Merge Field Reference

| Syntax | Description |
|---|---|
| `{{Column Name}}` | Replaced with the value from that column |
| `{{Email}}` | The recipient's email address |
| `{{First Name}}` | First name column |

Field names are **case-insensitive** and **whitespace-tolerant**:
`{{first name}}`, `{{First Name}}`, `{{ FIRST NAME }}` all work the same.

Unmatched fields (no column with that name) are replaced with an empty string.

---

## Sending Limits

| Account Type | Daily Limit |
|---|---|
| Gmail (free) | ~500 emails/day |
| Google Workspace Basic | ~2,000 emails/day |
| Google Workspace Business | ~2,000 emails/day |

The extension enforces a configurable self-imposed quota (default: 400/day) to stay safely below Google's limits.

---

## File Structure

```
Mail merge/
├── manifest.json          # Extension manifest (MV3)
├── popup.html/css/js      # Toolbar popup
├── app.html/css/js        # Full mail merge application
├── background.js          # Service worker (scheduled sends, alarms)
├── options.html/js        # Settings page
├── lib/
│   ├── api.js             # Google Sheets + Gmail API wrapper
│   └── merge.js           # Template engine + email encoder
├── icons/                 # SVG icons
└── tracker/
    └── apps-script.js     # Optional Google Apps Script tracking server
```

---

## Privacy

- All data is stored locally in your Chrome browser (`chrome.storage.local`).
- No data is sent to any third-party servers.
- Emails are sent directly from your Gmail account via Google's official API.
- The optional tracking server is hosted on **your own** Google account.
