/** ARC registration and download page. Deploy as Apps Script web app.
 * Create a Google Sheet first, open Extensions > Apps Script, paste this code.
 * Set EPUB_URL and PDF_URL to actual downloadable file links.
 * Deploy > New deployment > Web app: Execute as Me; access Anyone.
 * Paste the deployed /exec URL into index.html form action.
 */
const EPUB_URL = 'REPLACE_WITH_EPUB_DOWNLOAD_URL';
const PDF_URL = 'REPLACE_WITH_PDF_DOWNLOAD_URL';
const SHEET_NAME = 'ARC Requests';

function doPost(e) {
  const p = (e && e.parameter) || {};
  const name = String(p.name || '').trim().slice(0,120);
  const email = String(p.email || '').trim().slice(0,254);
  const profile = String(p.profile || '').trim().slice(0,300);
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return HtmlService.createHtmlOutput('Please return to the form and enter a valid name and email address.');
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Submitted at','Name','Email','Profile']);
  // Prefix formula-leading characters to prevent spreadsheet formula injection.
  const safe = s => /^[=+\-@]/.test(s) ? "'" + s : s;
  sheet.appendRow([new Date(),safe(name),safe(email),safe(profile)]);
  const escapeHtml = s => String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const linksReady = EPUB_URL.startsWith('https://') && PDF_URL.startsWith('https://');
  const buttons = linksReady ? `<p><a href="${escapeHtml(EPUB_URL)}" target="_blank" rel="noopener">Download EPUB</a></p><p><a href="${escapeHtml(PDF_URL)}" target="_blank" rel="noopener">Download PDF</a></p>` : '<p>The downloads are not configured yet. Please contact the author.</p>';
  const html = `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>body{font:18px/1.6 system-ui;background:#06273d;color:white;text-align:center;padding:35px 18px}main{background:white;color:#142738;border-radius:18px;padding:30px;max-width:600px;margin:30px auto}a{display:inline-block;background:#ffdb22;color:#142738;font-weight:bold;text-decoration:none;padding:12px 22px;border-radius:8px}p{margin:20px 0}</style></head><body><main><h1>Thank you, ${escapeHtml(name)}!</h1><p>Your free review copy is ready.</p>${buttons}<p>There is no obligation to review. If you do, please share your honest opinion.</p></main></body></html>`;
  return HtmlService.createHtmlOutput(html).setTitle('Your Free Review Copy');
}
