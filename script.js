/*SpendLite (Beginner‑Annotated Edition)
File: script.js
Generated: 2025-11-10 02:22

WHAT THIS FILE IS
- Core application logic for SpendLite (parsing CSV, categorising transactions, computing totals, and rendering the UI).

BEGINNER MAP OF THE CODE
- SECTION 1: Constants (column indexes, localStorage keys, page size)
- SECTION 2: App state (arrays and variables that hold current data)
- SECTION 3..15: Utility helpers (parsing dates/amounts, sorting rules, exporting, rendering tables/pager, etc.)
- EVENT WIRING: Code that attaches click/input listeners to elements in index.html

KEY TAKEAWAYS
- Keep "pure helpers" (formatting, parsing) separate from "rendering" (DOM updates).
- Always sanitise and normalise input (e.g., parseAmount to strip commas/currency symbols).
- Derive UI from state: compute CURRENT_TXNS -> apply rules -> render tables.
- LocalStorage can persist user data between sessions, but always handle missing/invalid data safely.
- Pagination reduces DOM size for large tables and keeps UI snappy.

DEBUGGING CHECKLIST
- If rules won’t save: check localStorage keys and browser storage limits.
- If totals look wrong: log intermediate values (console.log) for the category totals and net.
- If month filtering fails: verify date parsing (parseDateSmart) recognises your bank’s formats.

CHANGE WITH CONFIDENCE
- Add new columns by extending table rendering only; keep parsing logic consistent.
- When changing rules format, provide a migration or keep backward compatibility.
*/

// ============================================================================
// SPENDLITE V6.6.28 - Personal Expense Tracker
// ============================================================================
// Changelog (2025-10-19):
// - NEW: Rules are alphabetized automatically on startup (after rules are loaded).
// - EXISTING: Rules are alphabetized every time a rule is added/updated.
// ============================================================================

// ============================================================================
/** BEGINNER_INLINE_DOCS:SECTION 1: CONSTANTS
* Section overview: constants drive structure and assumptions
* - COL indexes must match your CSV mapping — change carefully
* - LS_KEYS names versioned to avoid breaking old localStorage
* - PAGE_SIZE controls pagination and DOM size/performance
*/
// SECTION 1: CONSTANTS AND CONFIGURATION
// ============================================================================

/** BEGINNER_INLINE_DOCS:COL
* Map CSV columns -> your internal indexes.
* If your bank CSV changes column order, update these numbers.
*/
const COL = { 
  DATE: 2,
  DEBIT: 5,
  LONGDESC: 9
};

/** BEGINNER_INLINE_DOCS:PAGE_SIZE
* Control how many rows are visible per page.
* Larger = fewer pages but more DOM; smaller = snappier on low-end devices.
*/
const PAGE_SIZE = 10;

const LS_KEYS = { 
  RULES: 'spendlite_rules_v6626',
  FILTER: 'spendlite_filter_v6626',
  MONTH: 'spendlite_month_v6627',
  TXNS_COLLAPSED: 'spendlite_txns_collapsed_v7',
  TXNS_JSON: 'spendlite_txns_json_v7'
};

const SAMPLE_RULES = `# Rules format: KEYWORD => CATEGORY
`;

// ============================================================================
// SECTION 2: APP STATE
// ============================================================================

let CURRENT_TXNS = [];
let CURRENT_RULES = [];
let CURRENT_FILTER = null;
let MONTH_FILTER = "";
let CURRENT_PAGE = 1;

// ============================================================================
// SECTION 3: DATE HELPERS
// ============================================================================

function formatMonthLabel(ym) {
  if (!ym) return 'All months';
  const [y, m] = ym.split('-').map(Number);
  const date = new Date(y, m - 1, 1);
  return date.toLocaleString(undefined, { month: 'long', year: 'numeric' });
}

function friendlyMonthOrAll(label) {
  if (!label) return 'All months';
  if (/^\d{4}-\d{2}$/.test(label)) return formatMonthLabel(label);
  return String(label);
}

function forFilename(label) {
  return String(label).replace(/\s+/g, '_');
}

// ============================================================================
// SECTION 4: TEXT UTILS
// ============================================================================

function toTitleCase(str) {
  if (!str) return '';
  return String(str)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\b([a-z])/g, (m, p1) => p1.toUpperCase());
}

/** BEGINNER_INLINE_DOCS:parseAmount
* Purpose: Convert strings like "$1,234.56" to a Number 1234.56
* Why: CSVs from banks often include commas/currency symbols. We must strip them.
* Edge cases: Empty strings => 0; non-number => 0 (safe default).
*/
function parseAmount(s) {
  if (s == null) return 0;
  s = String(s).replace(/[^\d\-,.]/g, '').replace(/,/g, '');
  return Number(s) || 0;
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================================================
// SECTION 4b: RULES SORTING
// ============================================================================

/**
 * Alphabetize rules in #rulesBox by keyword (case-insensitive).
 * - Preserves comments (# ...) and blank lines by moving them to the top in their
 *   original relative order, followed by the sorted rules block.
 * - Normalizes "KEY => VALUE" to "KEY => VALUE" (1 space around arrow).
 * Returns true if a change was made.
 */
/** BEGINNER_INLINE_DOCS:sortRulesBox
* Purpose: Keep rules sorted and normalized
* Preservation: Keeps comments and blanks at top in original order
* Normalization: Forces 'KEY => VALUE' uppercased with single spaces.
*/
function sortRulesBox({silent = false} = {}) {
  const box = document.getElementById('rulesBox');
  if (!box) return false;
  const original = String(box.value || '');
  const lines = original.split(/\r?\n/);

  const comments = [];
  const ruleLines = [];

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) {
      comments.push(line);
      continue;
    }
    // split on first =>
    const parts = line.split(/=>/i);
    if (parts.length >= 2) {
      const keyword = parts[0].trim();
      const category = parts.slice(1).join('=>').trim(); // in case => appears inside
      if (keyword && category) {
        ruleLines.push(`${keyword.toUpperCase()} => ${category.toUpperCase()}`);
      }
    } else {
      // Not a valid rule line; keep as comment to avoid data loss
      comments.push(line);
    }
  }

  const sorted = ruleLines.sort((a, b) => {
    const ka = a.split(/=>/)[0].trim().toLowerCase();
    const kb = b.split(/=>/)[0].trim().toLowerCase();
    return ka.localeCompare(kb, undefined, { sensitivity: 'base' });
  });

  // Reassemble: comments (as-is), blank line if both parts exist, then sorted rules
  const parts = [];
  if (comments.length) parts.push(...comments);
  if (comments.length && sorted.length) parts.push('');
  if (sorted.length) parts.push(...sorted);

  const next = parts.join('\n');
  if (next !== original) {
    box.value = next;
    try { localStorage.setItem(LS_KEYS.RULES, box.value); } catch {}
    if (!silent) {
      try { box.dispatchEvent(new Event('input', { bubbles: true })); } catch {}
    }
    return true;
  }
  return false;
}

// ============================================================================
// SECTION 5: DATE PARSING (AU)
// ============================================================================

/** BEGINNER_INLINE_DOCS:parseDateSmart
* Purpose: Accept multiple AU-friendly date formats and return a Date
* Supported: 'YYYY-MM-DD', 'DD/MM/YYYY', '1 July 2025' (and variants)
* Tips: If your bank uses a different format, extend the regex branches here.
*/
function parseDateSmart(s) {
  if (!s) return null;
  const str = String(s).trim();
  let m;

  m = str.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);

  m = str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);

  const s2 = str.replace(/^\d{1,2}:\d{2}\s*(am|pm)\s*/i, '');
  m = s2.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?\s*(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December),?\s+(\d{4})/i);
  if (m) {
    const day = +m[1], monthName = m[2].toLowerCase(), y = +m[3];
    const monthMap = {january:0,february:1,march:2,april:3,may:4,june:5,july:6,august:7,september:8,october:9,november:10,december:11};
    const mi = monthMap[monthName];
    if (mi != null) return new Date(y, mi, day);
  }
  return null;
}

function yyyymm(d) { 
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`; 
}

function getFirstTxnMonth(txns = CURRENT_TXNS) {
  if (!txns.length) return null;
  const d = parseDateSmart(txns[0].date);
  if (!d || isNaN(d)) return null;
  return yyyymm(d);
}

// ============================================================================
/** BEGINNER_INLINE_DOCS:SECTION 6: CSV LOADING
* How CSV becomes app state (CURRENT_TXNS)
* - Papa.parse converts CSV to rows; we skip header row if needed
* - We normalise amounts and keep only necessary fields
* - Keep parsing tolerant — never crash on weird rows
*/
// SECTION 6: CSV LOADING
// ============================================================================

/** BEGINNER_INLINE_DOCS:loadCsvText
* Purpose: Turn raw CSV text into CURRENT_TXNS array
* Steps:
*  1) Parse CSV into rows (skip header row if first amount isn't numeric)
*  2) Extract Effective Date, Amount, Description using COL indexes
*  3) Push a {date, amount, description} object for valid rows
*  4) Save to localStorage and trigger initial rendering
* Safety: Ignore rows with missing fields; never throw.
*/
function loadCsvText(csvText) {
  const rows = Papa.parse(csvText.trim(), { skipEmptyLines: true }).data;
  const startIdx = rows.length && isNaN(parseAmount(rows[0][COL.DEBIT])) ? 1 : 0;
  const txns = [];
  for (let i = startIdx; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.length < 10) continue;
    const effectiveDate = r[COL.DATE] || '';
    const debit = parseAmount(r[COL.DEBIT]);
    const longDesc = (r[COL.LONGDESC] || '').trim();
    if ((effectiveDate || longDesc) && Number.isFinite(debit) && debit !== 0) {
       txns.push({ date: effectiveDate, amount: debit, description: longDesc });
    }
  }
  CURRENT_TXNS = txns;
  saveTxnsToLocalStorage();
  try { updateMonthBanner(); } catch {}
  rebuildMonthDropdown();
  applyRulesAndRender();
  return txns;
}

// ============================================================================
// SECTION 7: MONTH FILTERING
// ============================================================================

/** BEGINNER_INLINE_DOCS:rebuildMonthDropdown
* Purpose: Extract unique months from CURRENT_TXNS and populate the <select>
* UX: Keeps current selection if still valid; otherwise resets to 'All months'.
*/
function rebuildMonthDropdown() {
  const sel = document.getElementById('monthFilter');
  const months = new Set();
  for (const t of CURRENT_TXNS) {
    const d = parseDateSmart(t.date);
    if (d) months.add(yyyymm(d));
  }
  const list = Array.from(months).sort();
  const current = MONTH_FILTER;
  sel.innerHTML = `<option value="">All months</option>` + 
    list.map(m => `<option value="${m}">${formatMonthLabel(m)}</option>`).join('');
  sel.value = current && list.includes(current) ? current : "";
  updateMonthBanner();
}

/** BEGINNER_INLINE_DOCS:monthFilteredTxns
* Purpose: Filter CURRENT_TXNS by MONTH_FILTER (YYYY-MM)
* Note: If no month is set, return all transactions.
*/
function monthFilteredTxns() {
  if (!MONTH_FILTER) return CURRENT_TXNS;
  return CURRENT_TXNS.filter(t => {
    const d = parseDateSmart(t.date);
    return d && yyyymm(d) === MONTH_FILTER;
  });
}

// ============================================================================
/** BEGINNER_INLINE_DOCS:SECTION 8: RULES
* Rule engine and matching strategy
* - Rules are case-insensitive 'KEYWORD => CATEGORY' lines
* - We match whole tokens to avoid false positives
* - Keep rules simple; first match wins in many tools, we scan all
*/
// SECTION 8: RULES
// ============================================================================

/** BEGINNER_INLINE_DOCS:parseRules
* Purpose: Convert rules text into [{ keyword, category }] objects
* Format: 'KEYWORD => CATEGORY' per line (case-insensitive)
* Notes: Ignores comments (# ...) and blank lines; trims whitespace.
*/
function parseRules(text) {
  const lines = String(text || "").split(/\r?\n/);
  const rules = [];
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const parts = trimmed.split(/=>/i);
    if (parts.length >= 2) {
      const keyword = parts[0].trim().toLowerCase();
      const category = parts[1].trim().toUpperCase();
      if (keyword && category) rules.push({ keyword, category });
    }
  }
  return rules;
}

/** BEGINNER_INLINE_DOCS:matchesKeyword
* Purpose: Check if a transaction description matches a rule keyword
* Strategy: Token-based, whole-word-ish matches using regex boundaries
* Special: For 3-word keywords, allows any non-alnum between words (robust to punctuation).
*/
function matchesKeyword(descLower, keywordLower) {
  if (!keywordLower) return false;
  const text = String(descLower || '').toLowerCase();
  const tokens = String(keywordLower).toLowerCase().split(/\s+/).filter(Boolean);
  if (!tokens.length) return false;
  const delim = '[^A-Za-z0-9&._]';
  if (tokens.length === 3) {
    const safe = tokens.map(tok => tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
    const re = new RegExp(`(?:^|${delim})${safe[0]}(?:${delim})+${safe[1]}(?:${delim})+${safe[2]}(?:${delim}|$)`, 'i');
    return re.test(text);
  }
  return tokens.every(tok => {
    const safe = tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(`(?:^|${delim})${safe}(?:${delim}|$)`, 'i');
    return re.test(text);
  });
}

/** BEGINNER_INLINE_DOCS:categorise
* Purpose: Assign categories to txns using rules
* Flow: For each txn, scan rules; last match wins; can add special cases
* Example special case: if PETROL <= $2 treat as COFFEE (demo of post-filters).
*/
function categorise(txns, rules) {
  for (const t of txns) {
    const descLower = String(t.desc || t.description || "").toLowerCase();
    const amount = Math.abs(Number(t.amount || t.debit || 0));
    let matched = null;
    for (const r of rules) {
      if (matchesKeyword(descLower, r.keyword)) { matched = r.category; }
    }
    if (matched && String(matched).toUpperCase() === "PETROL" && amount <= 2) matched = "COFFEE";
    t.category = matched || "UNCATEGORISED";
  }
}

// ============================================================================
// SECTION 9: CATEGORY TOTALS
// ============================================================================

/** BEGINNER_INLINE_DOCS:computeCategoryTotals
* Purpose: Sum amounts by category and return sorted rows
* Output: { rows: [[category, total]...], grand: sum }
* Use: The rows feed the Category Totals table; grand computes percentages.
*/
function computeCategoryTotals(txns) {
  const byCat = new Map();
  for (const t of txns) {
    const cat = (t.category || 'UNCATEGORISED').toUpperCase();
    byCat.set(cat, (byCat.get(cat) || 0) + t.amount);
  }
  const rows = [...byCat.entries()].sort((a, b) => b[1] - a[1]);
  const grand = rows.reduce((acc, [, v]) => acc + v, 0);
  return { rows, grand };
}

/** BEGINNER_INLINE_DOCS:renderCategoryTotals
* Purpose: Paint the Category Totals table
* Tips: Keep string building local; use toTitleCase for display
* Interactivity: Each category name is a link that sets CURRENT_FILTER.
*/
function renderCategoryTotals(txns) {
  const { rows, grand } = computeCategoryTotals(txns);
  const totalsDiv = document.getElementById('categoryTotals');
  let html = '<table class="cats">';
  html += '<colgroup><col class="col-cat"><col class="col-total"><col class="col-pct"></colgroup>';
  html += '<thead><tr><th>Category</th><th class="num">Total</th><th class="num">%</th></tr></thead>';
  html += '<tbody>';
  for (const [cat, total] of rows) {
    const pct = grand ? (total / grand * 100) : 0;
    html += `<tr>
      <td><a class="catlink" data-cat="${escapeHtml(cat)}"><span class="category-name">${escapeHtml(toTitleCase(cat))}</span></a></td>
      <td class="num">${total.toFixed(2)}</td>
      <td class="num">${pct.toFixed(1)}%</td>
    </tr>`;
  }
  html += `</tbody>`;
  html += `<tfoot><tr><td>Total</td><td class="num">${grand.toFixed(2)}</td><td class="num">100%</td></tr></tfoot>`;
  html += '</table>';
  totalsDiv.innerHTML = html;

  totalsDiv.querySelectorAll('a.catlink').forEach(a => {
    a.addEventListener('click', () => {
      CURRENT_FILTER = a.getAttribute('data-cat');
      try { localStorage.setItem(LS_KEYS.FILTER, CURRENT_FILTER || ''); } catch {}
      updateFilterUI();
      CURRENT_PAGE = 1;
      renderTransactionsTable();
    });
  });
}

/** BEGINNER_INLINE_DOCS:renderMonthTotals
* Purpose: Show month summary (count, debit, credit, net)
* Dependency: getFilteredTxns(monthFilteredTxns())
* UX: Adds a friendly label and highlights current filter.
*/
function renderMonthTotals() {
  const txns = getFilteredTxns(monthFilteredTxns());
  let debit = 0, credit = 0, count = 0;
  for (const t of txns) {
    const amt = Number(t.amount) || 0;
    if (amt > 0) debit += amt; else credit += Math.abs(amt);
    count++;
  }
  const net = debit - credit;
  const el = document.getElementById('monthTotals');
  if (el) {
    const label = friendlyMonthOrAll(MONTH_FILTER);
    const cat = CURRENT_FILTER ? ` + category "${CURRENT_FILTER}"` : "";
    el.innerHTML = `Showing <span class="badge">${count}</span> transactions for <strong>${label}${cat}</strong> · ` +
                   `Debit: <strong>$${debit.toFixed(2)}</strong> · ` +
                   `Credit: <strong>$${credit.toFixed(2)}</strong> · ` +
                   `Net: <strong>$${net.toFixed(2)}</strong>`;
  }
}

// ============================================================================
/** BEGINNER_INLINE_DOCS:SECTION 10: MAIN RENDER
* Render pipeline overview
* - Load txns -> parse rules -> categorise -> render totals/table
* - State lives in CURRENT_* variables; UI derives from state
* - Minimise direct DOM mutations; batch via string building
*/
// SECTION 10: MAIN RENDER
// ============================================================================

/** BEGINNER_INLINE_DOCS:applyRulesAndRender
* Purpose: Central orchestrator for re-render
* Steps: Parse rules -> categorise -> render totals -> render table -> persist
* Option: keepPage=true to avoid jumping back to page 1 after edits.
*/
function applyRulesAndRender({keepPage = false} = {}) { 
  if (!keepPage) CURRENT_PAGE = 1;
  CURRENT_RULES = parseRules(document.getElementById('rulesBox').value);
  try { localStorage.setItem(LS_KEYS.RULES, document.getElementById('rulesBox').value); } catch {}
  const txns = monthFilteredTxns();
  categorise(txns, CURRENT_RULES);
  renderMonthTotals();
  renderCategoryTotals(txns);
  renderTransactionsTable(txns);
  saveTxnsToLocalStorage();
  try { updateMonthBanner(); } catch {}
}

// ============================================================================
// SECTION 11: TXN TABLE & PAGER
// ============================================================================

/** BEGINNER_INLINE_DOCS:getFilteredTxns
* Purpose: Apply the category filter (CURRENT_FILTER) on top of month filtering
* Returns: A narrowed array for rendering and summaries.
*/
function getFilteredTxns(txns) {
  if (!CURRENT_FILTER) return txns;
  return txns.filter(t => (t.category || 'UNCATEGORISED').toUpperCase() === CURRENT_FILTER);
}

/** BEGINNER_INLINE_DOCS:updateFilterUI
* Purpose: Show/hide the 'Show all' button and the current filter label.
*/
function updateFilterUI() {
  const label = document.getElementById('activeFilter');
  const btn = document.getElementById('clearFilterBtn');
  if (CURRENT_FILTER) {
    label.textContent = `— filtered by "${CURRENT_FILTER}"`;
    btn.style.display = '';
  } else {
    label.textContent = '';
    btn.style.display = 'none';
  }
}

/** BEGINNER_INLINE_DOCS:updateMonthBanner
* Purpose: Update the label near Section 2 heading with the active month.
*/
function updateMonthBanner() {
  const banner = document.getElementById('monthBanner');
  const label = friendlyMonthOrAll(MONTH_FILTER);
  banner.textContent = `— ${label}`;
}

/** BEGINNER_INLINE_DOCS:renderTransactionsTable
* Purpose: Paint the paged transactions table
* Pagination: Uses CURRENT_PAGE and PAGE_SIZE
* Performance: Builds HTML string once; attaches click handlers after.
*/
function renderTransactionsTable(txns = monthFilteredTxns()) {
  const filtered = getFilteredTxns(txns);
  const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
  if (CURRENT_PAGE > totalPages) CURRENT_PAGE = totalPages;
  if (CURRENT_PAGE < 1) CURRENT_PAGE = 1;
  const start = (CURRENT_PAGE - 1) * PAGE_SIZE;
  const pageItems = filtered.slice(start, start + PAGE_SIZE);
  const table = document.getElementById('transactionsTable');
  let html = '<tr><th>Date</th><th>Amount</th><th>Category</th><th>Description</th><th></th></tr>';
  pageItems.forEach((t) => {
    const idx = CURRENT_TXNS.indexOf(t);
    const cat = (t.category || 'UNCATEGORISED').toUpperCase();
    const displayCat = toTitleCase(cat);
    html += `<tr>
      <td>${escapeHtml(t.date)}</td>
      <td>${t.amount.toFixed(2)}</td>
      <td><span class="category-name">${escapeHtml(displayCat)}</span></td>
      <td>${escapeHtml(t.description)}</td>
      <td><button class="rule-btn" onclick="assignCategory(${idx})">+</button></td>
    </tr>`;
  });
  table.innerHTML = html;
  renderPager(totalPages);
}

/** BEGINNER_INLINE_DOCS:renderPager
* Purpose: Render First/Prev/Next/Last and numbered page buttons
* UX: Shows current page and total pages; supports mouse wheel paging
* Safety: Disables buttons at edges; ignores clicks to current page.
*/
function renderPager(totalPages) {
  const pager = document.getElementById('pager');
  if (!pager) return;
  const pages = totalPages || 1;
  const cur = CURRENT_PAGE;

  function pageButton(label, page, disabled = false, isActive = false) {
    const disAttr = disabled ? ' disabled' : '';
    const activeClass = isActive ? ' active' : '';
    return `<button class="page-btn${activeClass}" data-page="${page}"${disAttr}>${label}</button>`;
  }

  const windowSize = 5;
  let start = Math.max(1, cur - Math.floor(windowSize / 2));
  let end = Math.min(pages, start + windowSize - 1);
  start = Math.max(1, Math.min(start, end - windowSize + 1));

  let html = '';
  html += pageButton('First', 1, cur === 1);
  html += pageButton('Prev', Math.max(1, cur - 1), cur === 1);
  for (let p = start; p <= end; p++) html += pageButton(String(p), p, false, p === cur);
  html += pageButton('Next', Math.min(pages, cur + 1), cur === pages);
  html += pageButton('Last', pages, cur === pages);
  html += `<span style="margin-left:8px">Page ${cur} / ${pages}</span>`;
  pager.innerHTML = html;

  pager.querySelectorAll('button.page-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const page = Number(e.currentTarget.getAttribute('data-page'));
      if (!page || page === CURRENT_PAGE) return;
      CURRENT_PAGE = page;
      renderTransactionsTable();
    });
  });

  const table = document.getElementById('transactionsTable');
  if (table && !table._wheelBound) {
    table.addEventListener('wheel', (e) => {
      if (pages <= 1) return;
      if (e.deltaY > 0 && CURRENT_PAGE < pages) {
        CURRENT_PAGE++;
        renderTransactionsTable();
      } else if (e.deltaY < 0 && CURRENT_PAGE > 1) {
        CURRENT_PAGE--;
        renderTransactionsTable();
      }
    }, { passive: true });
    table._wheelBound = true;
  }
}

// ============================================================================
// SECTION 12: EXPORT/IMPORT
// ============================================================================

/** BEGINNER_INLINE_DOCS:exportTotals
* Purpose: Create a fixed-width text report for totals
* Formatting: Pads columns so the text exports neatly aligned
* File: Uses Blob + URL.createObjectURL to trigger a download.
*/
function exportTotals() {
  const txns = monthFilteredTxns();
  const { rows, grand } = computeCategoryTotals(txns);
  const label = friendlyMonthOrAll(MONTH_FILTER || getFirstTxnMonth(txns) || new Date());
  const header = `SpendLite Category Totals (${label})`;
  const catWidth = Math.max(8, ...rows.map(([cat]) => toTitleCase(cat).length), 'Category'.length);
  const amtWidth = 12;
  const pctWidth = 6;
  const lines = [];
  lines.push(header);
  lines.push('='.repeat(header.length));
  lines.push('Category'.padEnd(catWidth) + ' ' + 'Amount'.padStart(amtWidth) + ' ' + '%'.padStart(pctWidth));
  for (const [cat, total] of rows) {
    const pct = grand ? (total / grand * 100) : 0;
    lines.push(toTitleCase(cat).padEnd(catWidth) + ' ' + total.toFixed(2).padStart(amtWidth) + ' ' + (pct.toFixed(1) + '%').padStart(pctWidth));
  }
  lines.push('');
  lines.push('TOTAL'.padEnd(catWidth) + ' ' + grand.toFixed(2).padStart(amtWidth) + ' ' + '100%'.padStart(pctWidth));
  const blob = new Blob([lines.join('\n')], { type: 'text/plain' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `category_totals_${forFilename(label)}.txt`;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/** BEGINNER_INLINE_DOCS:exportRules
* Purpose: Download the raw rules text as a file for backup.
*/
function exportRules() {
  const text = document.getElementById('rulesBox').value || '';
  const blob = new Blob([text], {type: 'text/plain'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'rules_export.txt';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

/** BEGINNER_INLINE_DOCS:importRulesFromFile
* Purpose: Read .txt rules and load them into the textarea, then re-render
* Extra: Alphabetises immediately so imported rules are tidy.
*/
function importRulesFromFile(file) {
  const reader = new FileReader();
  reader.onload = () => {
    const text = reader.result || '';
    const box = document.getElementById('rulesBox');
    box.value = text;
    try { RULES_CHANGED = true; } catch {}
    // Sort after import as well
    sortRulesBox();
    try { box.dispatchEvent(new Event('input', { bubbles: true })); } catch {}
    applyRulesAndRender();
  };
  reader.readAsText(file);
}

// ============================================================================
// SECTION 13: CATEGORY ASSIGNMENT
// ============================================================================

/** BEGINNER_INLINE_DOCS:deriveKeywordFromTxn
* Purpose: Suggest a starting rule keyword from a txn description
* Strategy: Extract alphanumeric tokens; prefer 'PAYPAL' or post-'VISA' tokens
* Why: Speeds up adding rules by guessing a sensible keyword.
*/
function deriveKeywordFromTxn(txn) {
  if (!txn) return "";
  const desc = String(txn.description || txn.desc || "").trim();
  if (!desc) return "";
  const tokens = (desc.match(/[A-Za-z0-9&._]+/g) || []).map(s => s.toLowerCase());
  if (!tokens.length) return "";
  function join3(k) { return tokens.slice(k, k + 3).filter(Boolean).map(s => s.toUpperCase()).join(' '); }
  const up = desc.toUpperCase();
  const paypalIdx = tokens.indexOf('paypal');
  if (paypalIdx !== -1) return join3(paypalIdx);
  if (/\bVISA-/.test(up)) {
    const visaTokIdx = tokens.indexOf('visa');
    if (visaTokIdx !== -1) return join3(Math.min(visaTokIdx + 1, Math.max(0, tokens.length - 1)));
  }
  return join3(0);
}

/** BEGINNER_INLINE_DOCS:addOrUpdateRuleLine
* Purpose: Insert or update a rule in the rules textarea
* Behaviour: Rewrites existing line if keyword exists; otherwise appends
* Post-step: Calls sortRulesBox to keep rules alphabetised.
*/
function addOrUpdateRuleLine(keywordUpper, categoryUpper) {
  if (!keywordUpper || !categoryUpper) return false;
  const box = document.getElementById('rulesBox');
  if (!box) return false;
  const lines = String(box.value || '').split(/\r?\n/);
  let updated = false;
  const kwLower = keywordUpper.toLowerCase();
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || line.startsWith('#')) continue;
    const parts = line.split(/=>/i);
    if (parts.length >= 2) {
      const existingKw = parts[0].trim().toLowerCase();
      if (existingKw === kwLower) {
        lines[i] = `${keywordUpper} => ${categoryUpper}`;
        updated = true;
        break;
      }
    }
  }
  if (!updated) lines.push(`${keywordUpper} => ${categoryUpper}`);
  box.value = lines.join("\n");
  // Ensure alphabetical order after any change
  sortRulesBox();
  try { RULES_CHANGED = true; } catch {}
  try { box.dispatchEvent(new Event('input', { bubbles: true })); } catch {}
  try { localStorage.setItem(LS_KEYS.RULES, box.value); } catch {}
  return true;
}

/** BEGINNER_INLINE_DOCS:assignCategory
* Purpose: Open the category picker, then apply user's selection
* Flow: Build a deduplicated category list (from txns + rules) -> open modal -> onChoose()
* After: If chosen, also suggest adding/updating a rule for future auto-categorisation.
*/
function assignCategory(idx) {
  const fromTxns = (Array.isArray(CURRENT_TXNS) ? CURRENT_TXNS : []).map(x => (x.category || '').trim());
  const fromRules = (Array.isArray(CURRENT_RULES) ? CURRENT_RULES : []).map(r => (r.category || '').trim ? r.category : (r.category || ''));
  const merged = Array.from(new Set([...fromTxns, ...fromRules].map(c => (c || '').trim()).filter(Boolean)));
  let base = Array.from(new Set(merged));
  base = base.map(c => (c.toUpperCase() === 'UNCATEGORISED' ? 'Uncategorised' : c));
  if (!base.includes('Uncategorised')) base.unshift('Uncategorised');
  base.unshift('+ Add new category...');
  const specials = new Set(['+ Add new category...', 'Uncategorised']);
  const rest = base.filter(c => !specials.has(c)).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  const categories = ['+ Add new category...', 'Uncategorised', ...rest];
  const current = ((CURRENT_TXNS && CURRENT_TXNS[idx] && CURRENT_TXNS[idx].category) || '').trim() || 'Uncategorised';

  SL_CatPicker.openCategoryPicker({
    categories,
    current,
    onChoose: (chosen) => {
      if (chosen) {
        const ch = String(chosen).trim();
        const lo = ch.toLowerCase();
        const isAdd = ch.startsWith('➕') || ch.startsWith('+') || lo.indexOf('add new category') !== -1;
        if (isAdd) {
          try { document.getElementById('catpickerBackdrop').classList.remove('show'); } catch {}
          return assignCategory_OLD(idx);
        }
      }
      const norm = (chosen === 'Uncategorised') ? '' : String(chosen).trim().toUpperCase();
      if (CURRENT_TXNS && CURRENT_TXNS[idx]) {
        CURRENT_TXNS[idx].category = norm;
      }
      try {
        if (norm) {
          const kw = deriveKeywordFromTxn(CURRENT_TXNS[idx]);
          if (kw) {
            const added = addOrUpdateRuleLine(kw, norm);
            if (added && typeof applyRulesAndRender === 'function') {
              applyRulesAndRender({keepPage: true});
            } else {
              renderMonthTotals(); renderTransactionsTable();
            }
          } else { renderMonthTotals(); renderTransactionsTable(); }
        } else { renderMonthTotals(); renderTransactionsTable(); }
      } catch (e) { try { renderMonthTotals(); renderTransactionsTable(); } catch {} }
    }
  });
}

function assignCategory_OLD(idx) {
  const txn = CURRENT_TXNS[idx];
  if (!txn) return;
  const suggestedKeyword = deriveKeywordFromTxn(txn);
  const keywordInput = prompt("Enter keyword to match:", suggestedKeyword);
  if (!keywordInput) return;
  const keyword = keywordInput.trim().toUpperCase();
  const defaultCat = (txn.category || "UNCATEGORISED").toUpperCase();
  const catInput = prompt("Enter category name:", defaultCat);
  if (!catInput) return;
  const category = catInput.trim().toUpperCase();
  const box = document.getElementById('rulesBox');
  const lines = String(box.value || "").split(/\r?\n/);
  let updated = false;
  for (let i = 0; i < lines.length; i++) {
    const line = (lines[i] || "").trim();
    if (!line || line.startsWith('#')) continue;
    const parts = line.split(/=>/i);
    if (parts.length >= 2) {
      const k = parts[0].trim().toUpperCase();
      if (k === keyword) { lines[i] = `${keyword} => ${category}`; updated = true; break; }
    }
  }
  if (!updated) lines.push(`${keyword} => ${category}`);
  box.value = lines.join("\n");
  sortRulesBox();
  try { RULES_CHANGED = true; } catch {}
  try { box.dispatchEvent(new Event('input', { bubbles: true })); } catch {}
  try { localStorage.setItem(LS_KEYS.RULES, box.value); } catch {}
  if (typeof applyRulesAndRender === 'function') {
    applyRulesAndRender({keepPage: true});
  }
}

// ============================================================================
// SECTION 14: LOCAL STORAGE
// ============================================================================

/** BEGINNER_INLINE_DOCS:saveTxnsToLocalStorage
* Purpose: Persist CURRENT_TXNS for future sessions
* Reliability: Wrap in try/catch; browsers may block or clear storage
* Compatibility: Writes to multiple keys for older versions.
*/
function saveTxnsToLocalStorage() {
  try {
    const data = JSON.stringify(CURRENT_TXNS || []);
    localStorage.setItem(LS_KEYS.TXNS_JSON, data);
    localStorage.setItem('spendlite_txns_json_v7', data);
    localStorage.setItem('spendlite_txns_json', data);
  } catch {}
}

// ============================================================================
// SECTION 15: COLLAPSE TOGGLE
// ============================================================================

function isTxnsCollapsed() {
  try { return localStorage.getItem(LS_KEYS.TXNS_COLLAPSED) !== 'false'; }
  catch { return true; }
}

function setTxnsCollapsed(v) {
  try { localStorage.setItem(LS_KEYS.TXNS_COLLAPSED, v ? 'true' : 'false'); } catch {}
}

function applyTxnsCollapsedUI() {
  const body = document.getElementById('transactionsBody');
  const toggle = document.getElementById('txnsToggleBtn');
  const collapsed = isTxnsCollapsed();
  if (body) body.style.display = collapsed ? 'none' : '';
  if (toggle) toggle.textContent = collapsed ? 'Show transactions' : 'Hide transactions';
}

function toggleTransactions() {
  const collapsed = isTxnsCollapsed();
  setTxnsCollapsed(!collapsed);
  applyTxnsCollapsedUI();
}

// ============================================================================
// SECTION 16: EVENTS
// ============================================================================

document.getElementById('csvFile').addEventListener('change', (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => { loadCsvText(reader.result); };
  reader.readAsText(file);
});

document.getElementById('recalculateBtn').addEventListener('click', applyRulesAndRender);
document.getElementById('exportRulesBtn').addEventListener('click', exportRules);
document.getElementById('exportTotalsBtn').addEventListener('click', exportTotals);

document.getElementById('importRulesBtn').addEventListener('click', () => 
  document.getElementById('importRulesInput').click()
);
document.getElementById('importRulesInput').addEventListener('change', (e) => {
  const f = e.target.files && e.target.files[0];
  if (f) importRulesFromFile(f);
});

document.getElementById('clearFilterBtn').addEventListener('click', () => {
  CURRENT_FILTER = null;
  try { localStorage.removeItem(LS_KEYS.FILTER); } catch {}
  updateFilterUI();
  CURRENT_PAGE = 1;
  renderTransactionsTable();
  renderMonthTotals(monthFilteredTxns());
});

document.getElementById('clearMonthBtn').addEventListener('click', () => {
  MONTH_FILTER = "";
  try { localStorage.removeItem(LS_KEYS.MONTH); } catch {}
  document.getElementById('monthFilter').value = "";
  updateMonthBanner();
  CURRENT_PAGE = 1;
  applyRulesAndRender();
});

document.getElementById('monthFilter').addEventListener('change', (e) => {
  MONTH_FILTER = e.target.value || "";
  try { localStorage.setItem(LS_KEYS.MONTH, MONTH_FILTER); } catch {}
  updateMonthBanner();
  CURRENT_PAGE = 1;
  applyRulesAndRender();
});

// ============================================================================
// SECTION 17: INIT
// ============================================================================

let INITIAL_RULES = '';
let RULES_CHANGED = false;

window.addEventListener('DOMContentLoaded', async () => {
  let restored = false;
  const box = document.getElementById('rulesBox');

  try {
    const saved = localStorage.getItem(LS_KEYS.RULES);
    if (saved && saved.trim()) {
      box.value = saved;
      restored = true;
    }
  } catch {}

  if (!restored) {
    try {
      const res = await fetch('rules.txt');
      const text = await res.text();
      box.value = text;
      restored = true;
    } catch {}
  }

  if (!restored) {
    box.value = SAMPLE_RULES;
  }

  // NEW: Sort rules once on startup, if needed
  sortRulesBox({silent: true});

  // Track initial snapshot
  INITIAL_RULES = box.value;

  try {
    const savedFilter = localStorage.getItem(LS_KEYS.FILTER);
    CURRENT_FILTER = savedFilter && savedFilter.trim() ? savedFilter.toUpperCase() : null;
  } catch {}
  try {
    const savedMonth = localStorage.getItem(LS_KEYS.MONTH);
    MONTH_FILTER = savedMonth || "";
  } catch {}

  updateFilterUI();
  CURRENT_PAGE = 1;
  updateMonthBanner();
});

document.addEventListener('DOMContentLoaded', () => {
  applyTxnsCollapsedUI();
  try { updateMonthBanner(); } catch {}
});

window.addEventListener('beforeunload', () => {
  try { localStorage.setItem(LS_KEYS.TXNS_JSON, JSON.stringify(CURRENT_TXNS || [])); } catch {}
});

// ============================================================================
// SECTION 23: CLOSE APP WITH AUTO-SAVE
// ============================================================================

window.addEventListener('load', () => {
  const rulesBox = document.getElementById('rulesBox');
  if (rulesBox) {
    rulesBox.addEventListener('input', () => {
      RULES_CHANGED = rulesBox.value !== INITIAL_RULES;
    });
  }
});

function downloadRulesFile(content, filename = 'rules.txt') {
  const blob = new Blob([content], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function showSaveStatus(message, type = 'info') {
  const statusEl = document.getElementById('saveStatus');
  if (!statusEl) return;
  statusEl.textContent = message;
  statusEl.className = `save-status ${type}`;
  statusEl.style.display = 'block';
  setTimeout(() => { statusEl.style.display = 'none'; }, 3000);
}

function handleCloseApp() {
  const rulesBox = document.getElementById('rulesBox');
  if (!rulesBox) return;
  const currentRules = rulesBox.value;
  if ((currentRules || '').trim() !== (INITIAL_RULES || '').trim()) {
    downloadRulesFile(currentRules, 'rules.txt');
    showSaveStatus('✓ Rules file updated', 'success');
    INITIAL_RULES = currentRules;
    RULES_CHANGED = false;
  } else {
    showSaveStatus('ℹ No rule changes', 'info');
  }
}

document.addEventListener('DOMContentLoaded', () => {
  const closeBtn = document.getElementById('closeAppBtn');
  if (closeBtn) closeBtn.addEventListener('click', handleCloseApp);
});
