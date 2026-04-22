'use strict';

/**
 * MergeEngine — processes Gmail draft templates against spreadsheet row data.
 * Supports {{ColumnName}} placeholders (case-insensitive, whitespace-tolerant).
 */
export class MergeEngine {
  /**
   * Replace all {{Variable}} placeholders in a string with row values.
   * @param {string} template
   * @param {Object} variables  { ColumnName: value, ... }
   * @param {Object} [opts]
   * @param {string} [opts.trackingPixelUrl]  If set, appended before </body>
   * @param {string} [opts.unsubscribeUrl]    If set, footer added before </body>
   * @param {string} [opts.unsubscribeText]   Footer text
   * @returns {string}
   */
  static process(template, variables, opts = {}) {
    let result = template;

    // Replace all {{Key}} occurrences (case-insensitive key matching)
    for (const [key, value] of Object.entries(variables)) {
      const regex = new RegExp(`\\{\\{\\s*${escapeRegex(key)}\\s*\\}\\}`, 'gi');
      result = result.replace(regex, value != null ? String(value) : '');
    }

    // Replace any leftover unmatched {{...}} with empty string
    result = result.replace(/\{\{[^}]*\}\}/g, '');

    // Inject tracking pixel
    if (opts.trackingPixelUrl) {
      const pixel = `<img src="${opts.trackingPixelUrl}" width="1" height="1" style="display:none" alt="">`;
      result = result.includes('</body>')
        ? result.replace('</body>', pixel + '</body>')
        : result + pixel;
    }

    // Inject unsubscribe footer
    if (opts.unsubscribeUrl) {
      const text = opts.unsubscribeText || 'Unsubscribe from this list';
      const footer = `<div style="margin-top:24px;padding-top:12px;border-top:1px solid #e0e0e0;font-size:11px;color:#888;font-family:sans-serif;">
  <a href="${opts.unsubscribeUrl}" style="color:#888;">${text}</a>
</div>`;
      result = result.includes('</body>')
        ? result.replace('</body>', footer + '</body>')
        : result + footer;
    }

    return result;
  }

  /**
   * Extract all {{Variable}} names from a template string.
   */
  static extractVariables(template) {
    const regex = /\{\{\s*([^}]+?)\s*\}\}/g;
    const vars = new Set();
    let match;
    while ((match = regex.exec(template)) !== null) {
      vars.add(match[1].trim());
    }
    return Array.from(vars);
  }

  /**
   * Parse a Gmail draft message payload into { subject, htmlBody, textBody, headers }.
   */
  static parseDraftPayload(message) {
    const headers = {};
    (message.payload?.headers || []).forEach(h => {
      headers[h.name.toLowerCase()] = h.value;
    });

    let htmlBody = '';
    let textBody = '';

    function walkParts(parts) {
      if (!parts) return;
      for (const part of parts) {
        if (part.mimeType === 'text/html' && part.body?.data) {
          htmlBody = decodeBase64Url(part.body.data);
        } else if (part.mimeType === 'text/plain' && part.body?.data) {
          textBody = decodeBase64Url(part.body.data);
        } else if (part.parts) {
          walkParts(part.parts);
        }
      }
    }

    if (message.payload?.body?.data) {
      const data = decodeBase64Url(message.payload.body.data);
      if (message.payload.mimeType === 'text/html') htmlBody = data;
      else textBody = data;
    }
    walkParts(message.payload?.parts);

    return { subject: headers['subject'] || '', htmlBody, textBody, headers };
  }

  /**
   * Encode a complete email as RFC 2822 base64url string for the Gmail API.
   */
  static encodeEmail({ to, from, replyTo, cc, bcc, subject, htmlBody, textBody, messageId, inReplyTo }) {
    const boundary = `----=_Part_${Math.random().toString(36).slice(2)}`;
    const lines = [];

    lines.push(`To: ${to}`);
    if (from) lines.push(`From: ${from}`);
    if (replyTo) lines.push(`Reply-To: ${replyTo}`);
    if (cc) lines.push(`Cc: ${cc}`);
    if (bcc) lines.push(`Bcc: ${bcc}`);
    lines.push(`Subject: ${encodeSubject(subject)}`);
    if (messageId) lines.push(`Message-ID: ${messageId}`);
    if (inReplyTo) lines.push(`In-Reply-To: ${inReplyTo}`);
    lines.push('MIME-Version: 1.0');

    if (htmlBody && textBody) {
      lines.push(`Content-Type: multipart/alternative; boundary="${boundary}"`);
      lines.push('');
      lines.push(`--${boundary}`);
      lines.push('Content-Type: text/plain; charset=UTF-8');
      lines.push('Content-Transfer-Encoding: quoted-printable');
      lines.push('');
      lines.push(textBody);
      lines.push(`--${boundary}`);
      lines.push('Content-Type: text/html; charset=UTF-8');
      lines.push('Content-Transfer-Encoding: quoted-printable');
      lines.push('');
      lines.push(htmlBody);
      lines.push(`--${boundary}--`);
    } else {
      const body = htmlBody || textBody || '';
      const mime = htmlBody ? 'text/html' : 'text/plain';
      lines.push(`Content-Type: ${mime}; charset=UTF-8`);
      lines.push('Content-Transfer-Encoding: base64');
      lines.push('');
      lines.push(btoa(unescape(encodeURIComponent(body))).replace(/(.{76})/g, '$1\n'));
    }

    const raw = lines.join('\r\n');
    return btoa(unescape(encodeURIComponent(raw)))
      .replace(/\+/g, '-')
      .replace(/\//g, '_')
      .replace(/=+$/, '');
  }
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function decodeBase64Url(data) {
  const base64 = data.replace(/-/g, '+').replace(/_/g, '/');
  try {
    return decodeURIComponent(
      atob(base64).split('').map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2)).join('')
    );
  } catch {
    return atob(base64);
  }
}

function encodeSubject(subject) {
  // RFC 2047 encode if non-ASCII chars present
  if (/[^\x00-\x7F]/.test(subject)) {
    return `=?UTF-8?B?${btoa(unescape(encodeURIComponent(subject)))}?=`;
  }
  return subject;
}

export default MergeEngine;
