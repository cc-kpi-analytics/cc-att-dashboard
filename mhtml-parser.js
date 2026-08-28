/* ============================================================
   MHTML parser — turns a saved "Webpage, Single File" (.mhtml)
   export of the Five9 Published Schedule Detail report into an
   array of { agent, date, type, duration } interval records.
   ============================================================ */

const MHTML_MONTH_NAMES = ['January','February','March','April','May','June',
                            'July','August','September','October','November','December'];

/**
 * Pull the text/html MIME part out of a raw .mhtml file and decode it
 * (mhtml wraps a normal webpage in a multipart/related MIME message,
 * usually with the HTML part quoted-printable encoded).
 */
function extractHtmlFromMhtml(raw) {
  const boundaryMatch = raw.match(/boundary="([^"]+)"/);
  if (!boundaryMatch) throw new Error('Not a recognizable .mhtml file (no MIME boundary found).');
  const boundary = boundaryMatch[1];

  const partDelim = '--' + boundary;
  const parts = raw.split(partDelim);
  let htmlPart = null;
  for (const part of parts) {
    if (/Content-Type:\s*text\/html/i.test(part)) { htmlPart = part; break; }
  }
  if (!htmlPart) throw new Error('No text/html section found inside the .mhtml file.');

  const headerEnd = htmlPart.search(/\r?\n\r?\n/);
  if (headerEnd === -1) throw new Error('Malformed .mhtml file (missing header/body separator).');
  const headers = htmlPart.slice(0, headerEnd);
  const body = htmlPart.slice(headerEnd).replace(/^\r?\n\r?\n/, '');

  const isQuotedPrintable = /Content-Transfer-Encoding:\s*quoted-printable/i.test(headers);
  return isQuotedPrintable ? decodeQuotedPrintable(body) : body;
}

/** Decode quoted-printable text (handles multi-byte UTF-8 sequences correctly). */
function decodeQuotedPrintable(str) {
  const joined = str.replace(/=\r\n/g, '').replace(/=\n/g, ''); // soft line breaks
  const bytes = [];
  for (let i = 0; i < joined.length; i++) {
    const c = joined[i];
    if (c === '=' && i + 2 < joined.length) {
      const hex = joined.slice(i + 1, i + 3);
      if (/^[0-9A-Fa-f]{2}$/.test(hex)) {
        bytes.push(parseInt(hex, 16));
        i += 2;
        continue;
      }
    }
    bytes.push(c.charCodeAt(0));
  }
  return new TextDecoder('utf-8').decode(new Uint8Array(bytes));
}

/** Parse a Five9 "Monday, 24 August 2026"-style date string into a JS Date. */
function parseScheduleDateString(str) {
  const commaIdx = str.indexOf(',');
  const rest = (commaIdx >= 0 ? str.slice(commaIdx + 1) : str).trim();
  const bits = rest.split(/\s+/);
  if (bits.length < 3) return null;
  const day = parseInt(bits[0], 10);
  const monthIdx = MHTML_MONTH_NAMES.findIndex(m => m.toLowerCase() === bits[1].toLowerCase());
  const year = parseInt(bits[2], 10);
  if (isNaN(day) || monthIdx === -1 || isNaN(year)) return null;
  return new Date(year, monthIdx, day);
}

/**
 * Parse the extracted HTML into interval records. Walks agent-name
 * markers and interval tables in document order — every interval table
 * belongs to whichever agent-name marker most recently preceded it, so
 * we don't need to reason about the page's nested nesting structure.
 */
function parseFive9ScheduleHtml(htmlString) {
  const doc = new DOMParser().parseFromString(htmlString, 'text/html');
  const nodes = doc.querySelectorAll('[id*="lblAgentName"], table[id*="IntervalReportGrid"]');
  const records = [];
  let currentAgent = null;

  nodes.forEach(node => {
    if (node.tagName === 'SPAN' && node.id.includes('lblAgentName')) {
      currentAgent = node.textContent.trim();
    } else if (node.tagName === 'TABLE') {
      if (!currentAgent) return;
      const rows = node.querySelectorAll('tr');
      rows.forEach(tr => {
        const cells = tr.querySelectorAll('td');
        if (cells.length < 5) return; // header row or malformed
        const dateStr = cells[0].textContent.trim();
        const typeStr = cells[3].textContent.trim();
        const durStr = cells[4].textContent.trim();
        const date = parseScheduleDateString(dateStr);
        if (!date) return;
        const duration = parseFloat(durStr) || 0;
        records.push({ agent: currentAgent, date, type: typeStr, duration });
      });
    }
  });

  return records;
}

/** Full pipeline: raw .mhtml text in, interval records out. */
function parseWeeklyMhtml(rawMhtmlText) {
  const html = extractHtmlFromMhtml(rawMhtmlText);
  return parseFive9ScheduleHtml(html);
}
