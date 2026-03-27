(function () {
  'use strict';

  const state = { currentFile: null, previousFile: null, issues: {}, comparisons: {} };
  const NA_VALS = new Set(['n/a', 'na', 'N/A', 'NA', '-', '', 'NOT YET SUPPORTED', 'not yet supported']);
  const SERVICE_SHEETS = new Set(['VAT', 'EPR', 'RsP']);

  document.addEventListener('DOMContentLoaded', () => {
    setupTabs();
    setupDragDrop('dropzone-current', 'file-current', 'current');
    setupDragDrop('dropzone-previous', 'file-previous', 'previous');
    document.getElementById('btn-download-excel').addEventListener('click', downloadExcel);
    document.getElementById('btn-download-analysis').addEventListener('click', downloadAnalysisExcel);
    document.getElementById('sanity-sheet-select').addEventListener('change', renderSanityTable);
  });

  function setupTabs() {
    document.querySelectorAll('.nav-links a').forEach(link => {
      link.addEventListener('click', e => {
        e.preventDefault();
        document.querySelectorAll('.nav-links a').forEach(a => a.classList.remove('active'));
        link.classList.add('active');
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.getElementById('panel-' + link.dataset.tab).classList.add('active');
      });
    });
  }

  function setupDragDrop(zoneId, inputId, type) {
    const zone = document.getElementById(zoneId);
    const input = document.getElementById(inputId);
    ['dragenter', 'dragover'].forEach(e => zone.addEventListener(e, ev => { ev.preventDefault(); zone.classList.add('dragover'); }));
    ['dragleave', 'drop'].forEach(e => zone.addEventListener(e, ev => { ev.preventDefault(); zone.classList.remove('dragover'); }));
    zone.addEventListener('drop', e => { if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0], type); });
    input.addEventListener('change', e => { if (e.target.files[0]) handleFile(e.target.files[0], type); });
  }

  function handleFile(file, type) {
    const ext = file.name.split('.').pop().toLowerCase();
    const isExcel = ['xlsx', 'xls', 'xlsm', 'xlsb'].includes(ext);
    const reader = new FileReader();
    reader.onload = e => {
      let parsed;
      if (isExcel) parsed = parseExcel(e.target.result, file.name);
      else if (ext === 'json') parsed = parseJSON(e.target.result, file.name);
      else if (ext === 'csv') parsed = parseCSV(e.target.result, file.name);
      else { alert('Unsupported file type.'); return; }
      if (type === 'current') {
        state.currentFile = parsed;
        markLoaded('dropzone-current', 'info-current', file.name, parsed);
      } else {
        state.previousFile = parsed;
        markLoaded('dropzone-previous', 'info-previous', file.name, parsed);
      }
      runFullAnalysis();
    };
    isExcel ? reader.readAsArrayBuffer(file) : reader.readAsText(file);
  }

  function markLoaded(zoneId, infoId, name, parsed) {
    document.getElementById(zoneId).classList.add('loaded');
    const sheets = Object.keys(parsed.sheets);
    const rows = sheets.reduce((s, k) => s + parsed.sheets[k].length, 0);
    document.getElementById(infoId).textContent = `✓ ${name} — ${sheets.length} sheet(s), ${rows} rows`;
  }

  // ===== PARSERS =====
  function parseExcel(buf, filename) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    const sheets = {};
    const provider = {};
    const skip = ['instruction', 'faq', 'rsp package list', 'package list'];
    const knownH = ['service_name', 'compliance_jurisdiction_country', 'registration_type',
      'customer_support_type', 'epr_category', 'compliance_region', 'group_type',
      'supported_customer_languages', 'service_description',
      'attributes', 'seller\'s place of establishment', 'place of establishment'];

    for (const sn of wb.SheetNames) {
      if (skip.some(s => sn.toLowerCase().includes(s))) continue;
      const snLo = sn.toLowerCase();
      if (snLo.includes('pan') || snLo.includes('bundle')) continue;

      const ws = wb.Sheets[sn];
      if (!ws || !ws['!ref']) continue;

      // Detect if this is a data sheet (has known data headers) or metadata sheet
      const range = XLSX.utils.decode_range(ws['!ref']);
      let isDataSheet = false;
      for (let r = range.s.r; r <= Math.min(range.s.r + 12, range.e.r); r++) {
        let matchCount = 0;
        for (let c = range.s.c; c <= Math.min(range.e.c, range.s.c + 25); c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          const v = cell ? (cell.v || '').toString().trim().toLowerCase() : '';
          if (v && knownH.some(kh => v.includes(kh))) matchCount++;
        }
        if (matchCount >= 2) { isDataSheet = true; break; }
      }

      // Skip non-data sheets (provider details, instructions, etc.)
      if (!isDataSheet) {
        // Try to extract provider name from non-data sheets
        if (!provider.name) {
          for (let r = range.s.r; r <= Math.min(range.e.r, 20); r++) {
            const labelCell = ws[XLSX.utils.encode_cell({ r, c: range.s.c })];
            const valCell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + 1 })];
            const label = labelCell ? (labelCell.v || '').toString().trim() : '';
            const val = valCell ? (valCell.v || '').toString().trim() : '';
            if (/provider.*name|company.*name|legal.*name/i.test(label) && val) {
              provider.name = val; break;
            }
          }
        }
        continue;
      }

      // === DATA SHEET PARSING ===
      // Find header row
      let hRow = -1, headers = [];
      for (let r = range.s.r; r <= Math.min(range.s.r + 12, range.e.r); r++) {
        const vals = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          vals.push(cell ? (cell.v || '').toString().trim().toLowerCase() : '');
        }
        if (vals.filter(v => knownH.some(kh => v.includes(kh))).length >= 2) {
          hRow = r;
          for (let c = range.s.c; c <= range.e.c; c++) {
            const cell = ws[XLSX.utils.encode_cell({ r, c })];
            headers.push(cell ? (cell.v || '').toString().trim() : '');
          }
          break;
        }
      }
      if (hRow < 0) {
        hRow = range.s.r;
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
          headers.push(cell ? (cell.v || '').toString().trim() : '');
        }
      }

      // Resolve blank headers: check sub-rows for fee-like labels, otherwise use ANY
      // non-empty sub-row value that looks like a label (not a data value)
      const feePattern = /filing|registration|fiscal|authorized|eori|representation|fee|charge|one.?time|recurrent|first|following|year/i;
      const protectedHeaders = /^(attributes|service_name|compliance_jurisdiction_country|epr_category|representative_required|registration_type|customer_support_type|supported_customer_languages|service_provider_locations|service_description|service_fulfillment_sla|fulfillment_documents_required|compliance_region|group_type|parent_asin_count)/i;
      headers = headers.map((h, i) => {
        if (h) {
          // Don't replace known column names — they're correct as-is
          if (protectedHeaders.test(h)) return h;
          if (/place.*establishment|seller/i.test(h)) return h;
          // For other non-blank headers, check if sub-rows have more specific fee labels
          // This handles merged headers like "YEAR 1 charge" that span multiple sub-columns
          // where sub-rows contain "Filing only fee", "Registration only fee", etc.
          for (let dr = 1; dr <= 3; dr++) {
            const cell = ws[XLSX.utils.encode_cell({ r: hRow + dr, c: range.s.c + i })];
            const val = cell ? (cell.v || '').toString().trim() : '';
            if (!val || val.length <= 2 || val.length >= 80) continue;
            if (/^(EUR|\d|http|N\/A|NOT|Yes|No)/i.test(val)) continue;
            // If sub-row has a fee-like label that's DIFFERENT from the parent header,
            // use it as the header (it's more specific)
            if (feePattern.test(val) && val.toLowerCase() !== h.toLowerCase()) {
              return val;
            }
          }
          return h;
        }
        // Blank header: resolve from sub-rows
        for (let dr = 1; dr <= 3; dr++) {
          const cell = ws[XLSX.utils.encode_cell({ r: hRow + dr, c: range.s.c + i })];
          const val = cell ? (cell.v || '').toString().trim() : '';
          if (!val || val.length <= 2 || val.length >= 80) continue;
          if (/^(EUR|\d|http|N\/A|NOT|Yes|No)/i.test(val)) continue;
          if (feePattern.test(val)) return val;
          if (val.length > 4 && /[a-z]/i.test(val)) return val;
        }
        return `Col_${i}`;
      });

      // Determine where data rows start — skip sub-header rows that contain fee labels
      let dataStartRow = hRow + 1;
      for (let dr = 1; dr <= 3; dr++) {
        let hasFeeLabel = false;
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r: hRow + dr, c })];
          const val = cell ? (cell.v || '').toString().trim() : '';
          if (val && feePattern.test(val) && !/^(EUR|\d)/i.test(val) && val.length > 3 && val.length < 80) {
            hasFeeLabel = true; break;
          }
        }
        if (hasFeeLabel) dataStartRow = hRow + dr + 1;
        else break;
      }

      // Parse data rows
      const rows = [];
      for (let r = dataStartRow; r <= range.e.r; r++) {
        const row = {}; let has = false;
        for (let c = 0; c < headers.length; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c: range.s.c + c })];
          const val = cell ? (cell.w || (cell.v != null ? cell.v.toString() : '')).trim() : '';
          row[headers[c]] = val;
          if (val) has = true;
        }
        if (has) rows.push(row);
      }

      // Extract provider name from data if needed
      if (!provider.name) {
        for (const row of rows) {
          for (const k of Object.keys(row)) {
            if (/provider.*name/i.test(k) && row[k] && row[k].length > 2) { provider.name = row[k]; break; }
          }
          if (provider.name) break;
        }
      }

      // Normalize sheet name to VAT/EPR/RsP
      let norm = sn.trim();
      const lo = norm.toLowerCase();
      if (lo.includes('vat') && !lo.includes('epr')) norm = 'VAT';
      else if (lo.includes('epr')) norm = 'EPR';
      else if (lo.includes('rsp') && !lo.includes('package list')) norm = 'RsP';

      if (rows.length > 0) sheets[norm] = rows;
    }
    return { name: filename, sheets, provider };
  }

  function parseCSV(text, filename) {
    const lines = text.split(/\r?\n/).filter(l => l.trim());
    if (lines.length < 2) return { name: filename, sheets: { Sheet1: [] }, provider: {} };
    const headers = csvLine(lines[0]);
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const vals = csvLine(lines[i]);
      const row = {};
      headers.forEach((h, j) => { row[h.trim()] = (vals[j] || '').trim(); });
      rows.push(row);
    }
    const joined = headers.join(' ').toLowerCase();
    let sn = 'Sheet1';
    if (/epr_category/.test(joined)) sn = 'EPR';
    else if (/compliance_region|group_type/.test(joined)) sn = 'RsP';
    else if (/compliance_jurisdiction_country/.test(joined)) sn = 'VAT';
    return { name: filename, sheets: { [sn]: rows }, provider: extractProv(rows) };
  }

  function csvLine(line) {
    const r = []; let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQ) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') inQ = false; else cur += ch; }
      else { if (ch === '"') inQ = true; else if (ch === ',') { r.push(cur); cur = ''; } else cur += ch; }
    }
    r.push(cur); return r;
  }

  function parseJSON(text, filename) {
    try {
      const d = JSON.parse(text);
      let sheets = {}, prov = {};
      if (d.sheets) { sheets = d.sheets; prov = d.provider || {}; }
      else if (Array.isArray(d)) { sheets = { Sheet1: d }; }
      else {
        const ks = Object.keys(d);
        if (ks.some(k => /vat|epr|rsp/i.test(k))) {
          for (const k of ks) if (Array.isArray(d[k])) sheets[k] = d[k];
          prov = d.provider || {};
        } else sheets = { Sheet1: [d] };
      }
      if (!prov.name) prov = extractProv(Object.values(sheets).flat());
      return { name: filename, sheets, provider: prov };
    } catch (e) { alert('Invalid JSON: ' + e.message); return { name: filename, sheets: {}, provider: {} }; }
  }

  function extractProv(rows) {
    for (const r of rows) for (const k of Object.keys(r)) if (/provider.*name/i.test(k) && r[k]) return { name: r[k] };
    return {};
  }

  // ===== HELPERS =====
  function parsePrice(val) {
    if (!val) return null;
    const s = val.toString().trim().toUpperCase();
    if (NA_VALS.has(s)) return null;
    const num = s.replace(/EUR/i, '').replace(/[€\s,]/g, '').trim();
    const p = parseFloat(num);
    return isNaN(p) ? null : p;
  }
  function fmtEUR(n) { return n == null ? '' : 'EUR ' + n.toFixed(2); }
  function esc(s) { return s ? s.toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : ''; }
  function trunc(s, n) { return s && s.length > n ? s.substring(0, n) + '…' : (s || ''); }
  function norm(h) { return h.toLowerCase().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' ').trim(); }
  function classifyEstablishment(val) {
    if (!val) return 'Other';
    const v = val.trim();
    if (/china/i.test(v) || /^cn$/i.test(v)) return 'China';
    if (/^non.?eu/i.test(v) || /out of country/i.test(v)) return 'Non-EU';
    if (/eu/i.test(v) || /same country/i.test(v) || /any/i.test(v)) return 'EU';
    return 'Other';
  }
  function getHeaders(rows) { return rows.length > 0 ? Object.keys(rows[0]) : []; }

  // ===== COLUMN DETECTION =====
  function classifyHeader(h) {
    const lo = norm(h);
    if (/monthly/i.test(lo)) return 'MONTHLY';

    // Breakdown columns (Y1 only — no Y2 breakdown exists)
    if (/filing only/i.test(lo)) return 'BREAKDOWN';
    if (/registration only/i.test(lo)) return 'BREAKDOWN';
    if (/fiscal.*only/i.test(lo)) return 'BREAKDOWN';
    if (/authoris.*only|authorized.*only/i.test(lo)) return 'BREAKDOWN';
    if (/eori only/i.test(lo)) return 'EORI_BREAKDOWN';

    // Real Excel: "YEAR 1 charge" / "YEAR 2 charge"
    if (/year\s*1/i.test(lo) && /charge/i.test(lo) && !/filing|registration|fiscal|eori|authorized/i.test(lo)) return 'Y1_TOTAL';
    if (/year\s*2/i.test(lo) && /charge/i.test(lo) && !/filing|registration|fiscal|eori|authorized/i.test(lo)) return 'Y2_TOTAL';
    if (/^year\s*1$/i.test(lo)) return 'Y1_TOTAL';
    if (/^year\s*2$/i.test(lo)) return 'Y2_TOTAL';

    // JSON format: "One-time charge for YEARLY payment frequency (total year 1)"
    if (/one.?time.*charge.*yearly/i.test(lo)) return 'Y1_TOTAL';
    if (/recurrent.*charge.*yearly/i.test(lo)) return 'Y2_TOTAL';

    // RsP: "First YEAR" = Y1, "Following year" = Y2
    if (/first\s*year/i.test(lo)) return 'Y1_TOTAL';
    if (/following\s*year/i.test(lo)) return 'Y2_TOTAL';

    // RsP fallback: "one-time charge" (not monthly) = Y1
    if (/one.?time.*charge/i.test(lo) && !/monthly/i.test(lo)) return 'Y1_TOTAL';
    // "Recurrent charge" without yearly/monthly = Y2 for RsP
    if (/recurrent.*charge/i.test(lo) && !/monthly/i.test(lo) && !/yearly/i.test(lo)) return 'Y2_TOTAL';

    // RsP: "Offer Attributes" column often contains the Y1 one-time price
    if (/offer\s*attributes/i.test(lo)) return 'Y1_TOTAL';

    return null;
  }

  function detectColumns(headers) {
    let y1Col = null, y2Col = null;
    const breakdownCols = [];
    for (const h of headers) {
      const cat = classifyHeader(h);
      if (cat === 'Y1_TOTAL' && !y1Col) y1Col = h;
      else if (cat === 'Y2_TOTAL' && !y2Col) y2Col = h;
      else if (cat === 'BREAKDOWN') breakdownCols.push(h);
      // EORI_BREAKDOWN intentionally excluded from breakdownCols
    }
    return { y1Col, y2Col, breakdownCols };
  }

  // Fallback: if detectColumns didn't find Y1 for a sheet, try to find it
  // by looking for the first column with EUR prices that appears before Y2
  function detectColumnsWithFallback(headers, rows) {
    let result = detectColumns(headers);
    if (!result.y1Col && rows.length > 0) {
      // Find first column with parseable EUR prices
      for (const h of headers) {
        if (h === result.y2Col) break; // stop before Y2
        if (/^Col_\d+$/.test(h)) continue; // skip generic cols
        if (classifyHeader(h) === 'MONTHLY') continue;
        // Check if this column has EUR prices in at least 3 rows
        let priceCount = 0;
        for (const row of rows) {
          if (parsePrice(row[h]) !== null) priceCount++;
          if (priceCount >= 3) break;
        }
        if (priceCount >= 3) {
          result = { ...result, y1Col: h };
          break;
        }
      }
    }
    if (!result.y2Col && result.y1Col && rows.length > 0) {
      // Find first column with EUR prices AFTER Y1
      let pastY1 = false;
      for (const h of headers) {
        if (h === result.y1Col) { pastY1 = true; continue; }
        if (!pastY1) continue;
        if (/^Col_\d+$/.test(h)) continue;
        if (classifyHeader(h) === 'MONTHLY') continue;
        if (classifyHeader(h) === 'BREAKDOWN' || classifyHeader(h) === 'EORI_BREAKDOWN') continue;
        let priceCount = 0;
        for (const row of rows) {
          if (parsePrice(row[h]) !== null) priceCount++;
          if (priceCount >= 3) break;
        }
        if (priceCount >= 3) {
          result = { ...result, y2Col: h };
          break;
        }
      }
    }
    return result;
  }

  function getEstCol(headers) {
    return headers.find(h => /place.?of.?establishment/i.test(h))
      || headers.find(h => /seller.*place/i.test(h))
      || headers.find(h => /establishment/i.test(h));
  }
  function getCountryCol(headers) {
    return headers.find(h => /compliance_jurisdiction_country/i.test(h));
  }
  function getServiceNameCol(headers) {
    return headers.find(h => /^service_name$/i.test(h))
      || headers.find(h => /service_name/i.test(h));
  }

  // ===== ROW & COLUMN CLEANUP =====
  // A row is "bogus" if it contains no real pricing data — examples, links,
  // sub-headers, blank rows, metadata. These appear in the first few rows
  // after the header in real Excel files.
  function isBogusRow(row, headers, rowPosition, y1Col, y2Col) {
    let filledCount = 0, linkCount = 0;
    for (const h of headers) {
      const v = (row[h] || '').trim();
      if (!v) continue;
      filledCount++;
      if (/^https?:\/\//i.test(v)) linkCount++;
    }
    // Completely empty
    if (filledCount === 0) return true;
    // All filled values are links
    if (linkCount > 0 && linkCount === filledCount) return true;
    // For first 3 rows: check if Y1/Y2 have text descriptions instead of prices
    if (rowPosition < 3) {
      const y1Val = y1Col ? (row[y1Col] || '').trim() : '';
      const y2Val = y2Col ? (row[y2Col] || '').trim() : '';
      const y1IsPrice = parsePrice(y1Val) !== null;
      const y2IsPrice = parsePrice(y2Val) !== null;
      const y1IsEmpty = !y1Val || NA_VALS.has(y1Val.toUpperCase());
      const y2IsEmpty = !y2Val || NA_VALS.has(y2Val.toUpperCase());
      // If Y1 has long text (>20 chars) that's NOT a price, it's a description/sub-header row
      if (y1Val && !y1IsPrice && !y1IsEmpty && y1Val.length > 20) return true;
      if (y2Val && !y2IsPrice && !y2IsEmpty && y2Val.length > 20) return true;
      // If mostly links
      if (linkCount > 0 && linkCount >= filledCount / 2) return true;
      // Very sparse row with no prices
      if (!y1IsPrice && !y2IsPrice && filledCount / headers.length < 0.15) return true;
    }
    return false;
  }

  function getUsefulColumns(headers, rows, y1Col, y2Col) {
    return headers.filter(h => {
      // ALWAYS keep Y1 and Y2 columns
      if (h === y1Col || h === y2Col) return true;
      // Drop Col_N columns that have no real data
      if (/^Col_\d+$/.test(h)) {
        return rows.some(row => {
          const v = (row[h] || '').trim();
          return v && !NA_VALS.has(v.toUpperCase()) && !/^https?:\/\//i.test(v);
        });
      }
      // Drop monthly columns
      if (classifyHeader(h) === 'MONTHLY') return false;
      // Keep everything else that has at least some data
      return rows.some(row => {
        const v = (row[h] || '').trim();
        return v && v.length > 0;
      });
    });
  }

  // ===== FULL ANALYSIS =====
  function runFullAnalysis() {
    if (!state.currentFile) return;
    runSanityChecks();
    if (state.previousFile) runComparisons();
    renderAnalysisTab();
    renderSanityTab();
    document.getElementById('btn-download-excel').disabled = false;
    document.getElementById('btn-download-analysis').disabled = false;
  }

  // ===== SANITY CHECKS =====
  // 1. Y1 exists but no Y2 → issue
  // 2. Y1 exists but no breakdown (VAT/EPR only, not RsP) → issue
  // 3. Both Y1 and Y2 missing → NOT an issue (not quoted)
  // 4. Y2 > Y1 → issue
  // 5. Breakdown sum (excl EORI) ≠ Y1 total (VAT/EPR only) → issue
  function runSanityChecks() {
    state.issues = {};
    const file = state.currentFile;
    for (const [sheet, rows] of Object.entries(file.sheets)) {
      const issues = [];
      const headers = getHeaders(rows);
      const { y1Col, y2Col, breakdownCols } = detectColumnsWithFallback(headers, rows);
      const isRsP = sheet === 'RsP';

      rows.forEach((row, idx) => {
        const msgs = [];
        const y1 = y1Col ? parsePrice(row[y1Col]) : null;
        const y2 = y2Col ? parsePrice(row[y2Col]) : null;

        if (y1 === null && y2 === null) { /* not quoted — no issue */ }
        else {
          if (y1 !== null && y2 === null) msgs.push('Year 1 price quoted but Year 2 is missing');
          if (y1 !== null && y2 !== null && y2 > y1) msgs.push(`Year 2 (${fmtEUR(y2)}) > Year 1 (${fmtEUR(y1)})`);

          if (!isRsP && y1 !== null && breakdownCols.length > 0) {
            let bdSum = 0, bdCount = 0;
            for (const col of breakdownCols) {
              const p = parsePrice(row[col]);
              if (p !== null) { bdSum += p; bdCount++; }
            }
            if (bdCount === 0) msgs.push('Year 1 price quoted but no breakdown fees found');
            else if (Math.abs(y1 - bdSum) > 0.01) msgs.push(`Breakdown sum (${fmtEUR(bdSum)}) ≠ Year 1 total (${fmtEUR(y1)}), delta: ${fmtEUR(y1 - bdSum)}`);
          }
        }
        if (msgs.length > 0) issues.push({ rowIdx: idx, messages: msgs });
      });
      state.issues[sheet] = issues;
    }
  }

  // ===== COMPARISONS =====
  // Composite key: raw establishment + service_name (strict match, no normalization)
  function makeCompKey(row, estCol, snCol) {
    const est = estCol ? (row[estCol] || '').trim().toLowerCase() : '';
    const sn = snCol ? (row[snCol] || '').trim() : '';
    return { composite: `${est}|||${sn}`, simple: sn };
  }

  function runComparisons() {
    state.comparisons = {};
    if (!state.previousFile) return;
    for (const [sheet, currRows] of Object.entries(state.currentFile.sheets)) {
      const prevRows = findPrevSheet(sheet);
      if (!prevRows) continue;
      const headers = getHeaders(currRows);
      const snCol = getServiceNameCol(headers);
      if (!snCol) continue;
      const estCol = getEstCol(headers);
      const curr = detectColumnsWithFallback(headers, currRows);

      const prevH = getHeaders(prevRows);
      const prev = detectColumnsWithFallback(prevH, prevRows);
      const prevSn = getServiceNameCol(prevH);
      if (!prevSn) continue;
      const prevEst = getEstCol(prevH);

      // Build previous index with both composite and simple keys
      const prevByComposite = {};
      const prevBySimple = {};
      prevRows.forEach(r => {
        const keys = makeCompKey(r, prevEst, prevSn);
        if (keys.composite !== 'Other|||') prevByComposite[keys.composite] = r;
        if (keys.simple) prevBySimple[keys.simple] = r;
      });

      // First pass: try composite key matching
      let comps = [];
      currRows.forEach((row, idx) => {
        const keys = makeCompKey(row, estCol, snCol);
        if (!keys.simple) return;
        const prevRow = prevByComposite[keys.composite];
        if (!prevRow) return;
        const delta = computeDelta(row, prevRow, curr, prev);
        if (delta) comps.push({ rowIdx: idx, ...delta });
      });

      // If composite matching found nothing, fall back to simple (service_name only)
      if (comps.length === 0) {
        currRows.forEach((row, idx) => {
          const keys = makeCompKey(row, estCol, snCol);
          if (!keys.simple) return;
          const prevRow = prevBySimple[keys.simple];
          if (!prevRow) return;
          const delta = computeDelta(row, prevRow, curr, prev);
          if (delta) comps.push({ rowIdx: idx, ...delta });
        });
      }

      state.comparisons[sheet] = comps;
    }
  }

  function computeDelta(row, prevRow, curr, prev) {
    let y1Delta = null, y2Delta = null;
    if (curr.y1Col && prev.y1Col) {
      const c = parsePrice(row[curr.y1Col]);
      const p = parsePrice(prevRow[prev.y1Col]);
      if (c !== null && p !== null && c !== p) {
        const pct = p !== 0 ? (((c - p) / p) * 100).toFixed(1) : 'N/A';
        y1Delta = { oldPrice: p, newPrice: c, pct, direction: c > p ? 'up' : 'down' };
      }
    }
    if (curr.y2Col && prev.y2Col) {
      const c = parsePrice(row[curr.y2Col]);
      const p = parsePrice(prevRow[prev.y2Col]);
      if (c !== null && p !== null && c !== p) {
        const pct = p !== 0 ? (((c - p) / p) * 100).toFixed(1) : 'N/A';
        y2Delta = { oldPrice: p, newPrice: c, pct, direction: c > p ? 'up' : 'down' };
      }
    }
    if (y1Delta || y2Delta) return { y1Delta, y2Delta };
    return null;
  }

  function findPrevSheet(name) {
    if (!state.previousFile) return null;
    const ps = state.previousFile.sheets;
    if (ps[name]) return ps[name];
    const n = name.toLowerCase().replace(/[^a-z0-9]/g, '');
    for (const k of Object.keys(ps)) {
      if (k.toLowerCase().replace(/[^a-z0-9]/g, '') === n) return ps[k];
    }
    return null;
  }

  // ===== RENDER TAB 1: UPLOAD & ANALYSIS =====
  // ONLY VAT, EPR, RsP — no provider details, no other sheets
  function renderAnalysisTab() {
    const file = state.currentFile;
    const pBar = document.getElementById('provider-bar');
    if (file.provider && file.provider.name) {
      pBar.style.display = 'flex';
      document.getElementById('provider-name-display').textContent = file.provider.name;
      document.getElementById('provider-file-display').textContent = file.name;
    } else { pBar.style.display = 'none'; }

    const out = document.getElementById('analysis-output');
    let html = '';
    for (const svc of ['VAT', 'EPR', 'RsP']) {
      const rows = file.sheets[svc];
      if (!rows || rows.length === 0) {
        html += `<div class="service-analysis-card">
          <div class="sac-header"><h3>${esc(svc)}</h3><span class="issues-flag no">No data</span></div>
          <div class="sac-body"><p style="color:var(--text-muted);">No ${svc} sheet found in the uploaded file.</p></div>
        </div>`;
        continue;
      }
      html += renderServiceCard(svc, rows);
    }
    out.innerHTML = html;
  }

  function renderServiceCard(sheet, rows) {
    const headers = getHeaders(rows);
    const { y1Col, y2Col } = detectColumnsWithFallback(headers, rows);
    const countryCol = getCountryCol(headers);
    const estCol = getEstCol(headers);

    // Build price data: country × cohort → { y1: [], y2: [] }
    const priceData = {};
    rows.forEach(row => {
      const country = countryCol ? (row[countryCol] || '').trim() : 'N/A';
      const cohort = estCol ? classifyEstablishment(row[estCol]) : 'EU';
      const key = `${country}|||${cohort}`;
      if (!priceData[key]) priceData[key] = { country, cohort, y1: [], y2: [] };
      if (y1Col) { const p = parsePrice(row[y1Col]); if (p !== null) priceData[key].y1.push(p); }
      if (y2Col) { const p = parsePrice(row[y2Col]); if (p !== null) priceData[key].y2.push(p); }
    });

    let tableHtml = `<table class="price-summary-table">
      <thead><tr><th>Country</th><th>Cohort</th><th>Entries</th><th>Year 1 Avg</th><th>Year 2 Avg</th></tr></thead><tbody>`;

    const sortedKeys = Object.keys(priceData).sort();
    for (const key of sortedKeys) {
      const d = priceData[key];
      const entries = Math.max(d.y1.length, d.y2.length);
      if (entries === 0) continue;
      const y1Avg = d.y1.length ? (d.y1.reduce((a, b) => a + b, 0) / d.y1.length) : null;
      const y2Avg = d.y2.length ? (d.y2.reduce((a, b) => a + b, 0) / d.y2.length) : null;
      tableHtml += `<tr>
        <td>${esc(d.country)}</td>
        <td><span class="segment-badge seg-${d.cohort.toLowerCase().replace(/[^a-z]/g,'')}">${esc(d.cohort)}</span></td>
        <td>${entries}</td>
        <td class="price-cell">${y1Avg !== null ? fmtEUR(y1Avg) : '<span style="color:var(--text-muted);">Not quoted</span>'}</td>
        <td class="price-cell">${y2Avg !== null ? fmtEUR(y2Avg) : '<span style="color:var(--text-muted);">Not quoted</span>'}</td>
      </tr>`;
    }
    tableHtml += '</tbody></table>';

    const issues = state.issues[sheet] || [];
    const noData = sortedKeys.length === 0;

    // Diagnostic: show all headers and their detected roles
    let diagHtml = '<details style="margin-top:0.75rem;font-size:0.75rem;color:var(--text-secondary);">';
    diagHtml += '<summary style="cursor:pointer;font-weight:600;">🔍 Column Detection Debug (click to expand)</summary>';
    diagHtml += '<div style="margin-top:0.4rem;background:#f7f8f8;padding:0.5rem;border-radius:4px;overflow-x:auto;">';
    diagHtml += `<div>Y1 col: <code style="color:green;">${esc(y1Col || 'NOT FOUND')}</code> | Y2 col: <code style="color:green;">${esc(y2Col || 'NOT FOUND')}</code></div>`;
    diagHtml += '<table style="font-size:0.7rem;border-collapse:collapse;margin-top:0.3rem;width:100%;">';
    diagHtml += '<tr style="background:#e0e0e0;"><th style="padding:2px 6px;text-align:left;">Index</th><th style="padding:2px 6px;text-align:left;">Header</th><th style="padding:2px 6px;text-align:left;">Classification</th><th style="padding:2px 6px;text-align:left;">Sample Value</th></tr>';
    headers.forEach((h, i) => {
      const cat = classifyHeader(h) || '—';
      const sample = rows.length > 0 ? (rows[0][h] || '').substring(0, 40) : '';
      const isY1 = h === y1Col ? ' ✅ Y1' : '';
      const isY2 = h === y2Col ? ' ✅ Y2' : '';
      diagHtml += `<tr style="border-bottom:1px solid #ddd;"><td style="padding:2px 6px;">${i}</td><td style="padding:2px 6px;max-width:300px;overflow:hidden;text-overflow:ellipsis;">${esc(h)}</td><td style="padding:2px 6px;">${esc(cat)}${isY1}${isY2}</td><td style="padding:2px 6px;">${esc(sample)}</td></tr>`;
    });
    diagHtml += '</table></div></details>';

    return `<div class="service-analysis-card">
      <div class="sac-header">
        <h3>${esc(sheet)}</h3>
        <span class="issues-flag ${issues.length > 0 ? 'yes' : 'no'}">${issues.length > 0 ? issues.length + ' hygiene issue(s)' : 'All OK'}</span>
      </div>
      <div class="sac-body">
        ${noData ? '<p style="color:var(--text-muted);">No pricing data found.</p>' : tableHtml}
        ${diagHtml}
      </div>
    </div>`;
  }

  // ===== COLUMN DISPLAY NAMES =====
  // Map raw column headers to cleaner display names per sheet
  function getDisplayName(header, sheet, y1Col, y2Col) {
    const lo = norm(header);
    // Don't rename the actual Y1/Y2 columns — except to give them cleaner names
    if (header === y1Col) {
      if (sheet === 'RsP') return 'One-time charge (First YEAR)';
      return header;
    }
    if (header === y2Col) {
      if (sheet === 'RsP') return 'Recurrent charge (Following years)';
      return header;
    }
    // VAT: extra "One-time charge for YEARLY..." (not the Y1 col) → EORI
    if (sheet === 'VAT' && /one.?time.*charge.*yearly/i.test(lo)) return 'EORI';
    // EPR: extra "One-time charge for YEARLY..." (not the Y1 col) → Authorized Rep Fee
    if (sheet === 'EPR' && /one.?time.*charge.*yearly/i.test(lo)) return 'Authorized Rep Fee';
    // Clean up generic headers
    if (/^year\s*1\s*charge$/i.test(lo)) return 'Year 1 Charge';
    if (/^year\s*2\s*charge$/i.test(lo)) return 'Year 2 Charge';
    return header;
  }

  // ===== RENDER TAB 2: PRICE COMPARISON & HYGIENE CHECKS =====
  function renderSanityTab() {
    const sel = document.getElementById('sanity-sheet-select');
    const sheets = Object.keys(state.currentFile.sheets).filter(s => SERVICE_SHEETS.has(s));
    sel.innerHTML = sheets.map(s => `<option value="${s}">${s}</option>`).join('');
    renderSanityTable();
  }

  function renderSanityTable() {
    const area = document.getElementById('sanity-table-area');
    const sheet = document.getElementById('sanity-sheet-select').value;
    const allRows = state.currentFile.sheets[sheet];
    if (!allRows || allRows.length === 0) { area.innerHTML = '<div class="empty-state"><p>No data in this sheet</p></div>'; return; }

    const allHeaders = getHeaders(allRows);
    const { y1Col, y2Col } = detectColumnsWithFallback(allHeaders, allRows);
    const issues = state.issues[sheet] || [];
    const comps = state.comparisons[sheet] || [];

    // Filter bogus rows (examples, links, blank rows in first few positions)
    const filteredIndices = [];
    const rows = [];
    allRows.forEach((row, idx) => {
      if (!isBogusRow(row, allHeaders, idx, y1Col, y2Col)) {
        filteredIndices.push(idx);
        rows.push(row);
      }
    });

    // Debug: show comparison diagnostics if previous file loaded but no matches
    let debugHtml = '';
    if (state.previousFile && comps.length === 0) {
      const prevRows = findPrevSheet(sheet);
      const prevH = prevRows ? getHeaders(prevRows) : [];
      const prevSn = getServiceNameCol(prevH);
      const prevEst = getEstCol(prevH);
      const currSn = getServiceNameCol(allHeaders);
      const currEst = getEstCol(allHeaders);
      const prev = prevRows ? detectColumnsWithFallback(prevH, prevRows) : {};

      // Sample keys for debugging
      let currSample = [], prevSample = [];
      if (currSn && rows.length > 0) {
        for (let i = 0; i < Math.min(3, rows.length); i++) {
          const keys = makeCompKey(rows[i], currEst, currSn);
          currSample.push(keys.simple);
        }
      }
      if (prevSn && prevRows && prevRows.length > 0) {
        for (let i = 0; i < Math.min(3, prevRows.length); i++) {
          const keys = makeCompKey(prevRows[i], prevEst, prevSn);
          prevSample.push(keys.simple);
        }
      }

      debugHtml = `<div style="background:#fff8e1;border:1px solid #f0d88a;border-radius:4px;padding:0.6rem 1rem;margin-bottom:0.75rem;font-size:0.78rem;">
        <strong>⚠ No price changes detected for ${esc(sheet)}</strong><br>
        Current: snCol=<code>${esc(currSn || 'NOT FOUND')}</code>, estCol=<code>${esc(currEst || 'NONE')}</code>, y1=<code>${esc(y1Col || 'NONE')}</code>, y2=<code>${esc(y2Col || 'NONE')}</code>, rows=${rows.length}<br>
        Previous: snCol=<code>${esc(prevSn || 'NOT FOUND')}</code>, estCol=<code>${esc(prevEst || 'NONE')}</code>, y1=<code>${esc(prev.y1Col || 'NONE')}</code>, y2=<code>${esc(prev.y2Col || 'NONE')}</code>, rows=${prevRows ? prevRows.length : 0}<br>
        Sample current keys: ${currSample.map(s => '<code>' + esc(trunc(s, 50)) + '</code>').join(', ') || 'none'}<br>
        Sample previous keys: ${prevSample.map(s => '<code>' + esc(trunc(s, 50)) + '</code>').join(', ') || 'none'}
      </div>`;
    }

    // Filter useless columns
    const visibleHeaders = getUsefulColumns(allHeaders, rows, y1Col, y2Col);

    const issueMap = {};
    issues.forEach(i => { issueMap[i.rowIdx] = i.messages; });
    const compMap = {};
    comps.forEach(c => { compMap[c.rowIdx] = c; });
    // Show delta columns if previous file is loaded and this sheet exists in it
    const hasComps = state.previousFile && findPrevSheet(sheet) ? true : false;

    let html = debugHtml;
    html += '<div class="sanity-table-wrap"><table class="sanity-table"><thead><tr>';
    html += '<th>#</th>';
    for (const h of visibleHeaders) {
      const displayName = getDisplayName(h, sheet, y1Col, y2Col);
      html += `<th title="${esc(h)}">${esc(trunc(displayName, 32))}</th>`;
      if (h === y1Col && hasComps) html += '<th class="delta-header">Y1 Δ vs Prev</th>';
      if (h === y2Col && hasComps) html += '<th class="delta-header">Y2 Δ vs Prev</th>';
    }
    html += '<th class="issue-header">Issues</th></tr></thead><tbody>';

    rows.forEach((row, i) => {
      const origIdx = filteredIndices[i];
      const rowIssues = issueMap[origIdx];
      const rowComp = compMap[origIdx];
      let rowClass = '';
      if (rowIssues && rowComp) rowClass = 'row-both';
      else if (rowIssues) rowClass = 'row-issue';
      else if (rowComp) rowClass = 'row-compare';

      html += `<tr class="${rowClass}"><td>${origIdx + 2}</td>`;
      for (const h of visibleHeaders) {
        const val = row[h] || '';
        html += `<td title="${esc(val)}">${esc(trunc(val, 35))}</td>`;
        if (h === y1Col && hasComps) {
          if (rowComp && rowComp.y1Delta) {
            const d = rowComp.y1Delta;
            html += `<td class="col-delta ${d.direction}">${d.direction === 'up' ? '▲' : '▼'} ${Math.abs(parseFloat(d.pct))}% (prev: ${fmtEUR(d.oldPrice)})</td>`;
          } else html += '<td></td>';
        }
        if (h === y2Col && hasComps) {
          if (rowComp && rowComp.y2Delta) {
            const d = rowComp.y2Delta;
            html += `<td class="col-delta ${d.direction}">${d.direction === 'up' ? '▲' : '▼'} ${Math.abs(parseFloat(d.pct))}% (prev: ${fmtEUR(d.oldPrice)})</td>`;
          } else html += '<td></td>';
        }
      }
      html += rowIssues ? `<td class="col-issue">${rowIssues.map(m => esc(m)).join('; ')}</td>` : '<td></td>';
      html += '</tr>';
    });

    html += '</tbody></table></div>';
    area.innerHTML = html;
  }

  // ===== ANALYSIS EXCEL DOWNLOAD =====
  function downloadAnalysisExcel() {
    if (!state.currentFile) return;
    const wb = XLSX.utils.book_new();
    const file = state.currentFile;

    for (const svc of ['VAT', 'EPR', 'RsP']) {
      const rows = file.sheets[svc];
      if (!rows || rows.length === 0) continue;
      const headers = getHeaders(rows);
      const { y1Col, y2Col } = detectColumnsWithFallback(headers, rows);
      const countryCol = getCountryCol(headers);
      const estCol = getEstCol(headers);

      const priceData = {};
      rows.forEach(row => {
        const country = countryCol ? (row[countryCol] || '').trim() : 'N/A';
        const cohort = estCol ? classifyEstablishment(row[estCol]) : 'EU';
        const key = `${country}|||${cohort}`;
        if (!priceData[key]) priceData[key] = { country, cohort, y1: [], y2: [] };
        if (y1Col) { const p = parsePrice(row[y1Col]); if (p !== null) priceData[key].y1.push(p); }
        if (y2Col) { const p = parsePrice(row[y2Col]); if (p !== null) priceData[key].y2.push(p); }
      });

      const hdr = ['Country', 'Cohort', 'Entries', 'Year 1 Avg', 'Year 2 Avg'];
      const dataRows = [];
      for (const key of Object.keys(priceData).sort()) {
        const d = priceData[key];
        const entries = Math.max(d.y1.length, d.y2.length);
        if (entries === 0) continue;
        dataRows.push([
          d.country, d.cohort, entries,
          d.y1.length ? +(d.y1.reduce((a, b) => a + b, 0) / d.y1.length).toFixed(2) : 'Not quoted',
          d.y2.length ? +(d.y2.reduce((a, b) => a + b, 0) / d.y2.length).toFixed(2) : 'Not quoted'
        ]);
      }
      if (dataRows.length === 0) continue;
      const ws = XLSX.utils.aoa_to_sheet([hdr, ...dataRows]);
      ws['!cols'] = [{ wch: 18 }, { wch: 12 }, { wch: 10 }, { wch: 16 }, { wch: 16 }];
      XLSX.utils.book_append_sheet(wb, ws, svc + ' Analysis');
    }

    if (wb.SheetNames.length === 0) { alert('No pricing data found.'); return; }
    const provName = (file.provider?.name || 'provider').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    XLSX.writeFile(wb, `Price_Analysis_${provName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  // ===== COMPARISON EXCEL DOWNLOAD =====
  function downloadExcel() {
    if (!state.currentFile) return;
    const wb = XLSX.utils.book_new();

    for (const [sheet, allRows] of Object.entries(state.currentFile.sheets)) {
      if (!SERVICE_SHEETS.has(sheet) || allRows.length === 0) continue;
      const allHeaders = getHeaders(allRows);
      const { y1Col, y2Col } = detectColumnsWithFallback(allHeaders, allRows);
      const issues = state.issues[sheet] || [];
      const comps = state.comparisons[sheet] || [];

      const filteredIndices = [];
      const rows = [];
      allRows.forEach((row, idx) => {
        if (!isBogusRow(row, allHeaders, idx, y1Col, y2Col)) { filteredIndices.push(idx); rows.push(row); }
      });
      const headers = getUsefulColumns(allHeaders, rows, y1Col, y2Col);

      const issueMap = {};
      issues.forEach(i => { issueMap[i.rowIdx] = i.messages; });
      const compMap = {};
      comps.forEach(c => { compMap[c.rowIdx] = c; });

      const hasPrev = state.previousFile && findPrevSheet(sheet);

      const colOrder = [];
      for (const h of headers) {
        colOrder.push(h);
        if (h === y1Col && hasPrev) colOrder.push('__Y1_DELTA');
        if (h === y2Col && hasPrev) colOrder.push('__Y2_DELTA');
      }
      colOrder.push('__ISSUES');

      const headerRow = colOrder.map(c => {
        if (c === '__Y1_DELTA') return 'Y1 Δ vs Previous';
        if (c === '__Y2_DELTA') return 'Y2 Δ vs Previous';
        if (c === '__ISSUES') return 'Issues';
        return getDisplayName(c, sheet, y1Col, y2Col);
      });

      const dataRows = rows.map((row, i) => {
        const origIdx = filteredIndices[i];
        return colOrder.map(c => {
          if (c === '__Y1_DELTA') {
            const comp = compMap[origIdx];
            if (comp && comp.y1Delta) { const d = comp.y1Delta; return `${d.direction === 'up' ? '▲' : '▼'} ${Math.abs(parseFloat(d.pct))}% (prev: ${fmtEUR(d.oldPrice)})`; }
            return '';
          }
          if (c === '__Y2_DELTA') {
            const comp = compMap[origIdx];
            if (comp && comp.y2Delta) { const d = comp.y2Delta; return `${d.direction === 'up' ? '▲' : '▼'} ${Math.abs(parseFloat(d.pct))}% (prev: ${fmtEUR(d.oldPrice)})`; }
            return '';
          }
          if (c === '__ISSUES') { const iss = issueMap[origIdx]; return iss ? iss.join('; ') : ''; }
          return row[c] || '';
        });
      });

      const ws = XLSX.utils.aoa_to_sheet([headerRow, ...dataRows]);
      ws['!cols'] = colOrder.map((c, i) => {
        if (c === '__ISSUES') return { wch: 50 };
        if (c.startsWith('__')) return { wch: 30 };
        if (/description/i.test(c)) return { wch: 40 };
        return { wch: Math.min(25, Math.max(10, (headerRow[i] || '').length + 2)) };
      });
      XLSX.utils.book_append_sheet(wb, ws, sheet.substring(0, 31));
    }

    // Summary
    const summaryData = [['USSP Price Check Report'], ['Generated', new Date().toISOString()], ['File', state.currentFile.name]];
    if (state.previousFile) summaryData.push(['Compared with', state.previousFile.name]);
    summaryData.push([]);
    for (const [sheet, issues] of Object.entries(state.issues)) summaryData.push([`${sheet}: ${issues.length} rows with issues`]);
    if (state.comparisons) { summaryData.push([]); for (const [sheet, comps] of Object.entries(state.comparisons)) summaryData.push([`${sheet}: ${comps.length} price changes`]); }
    const summWs = XLSX.utils.aoa_to_sheet(summaryData);
    summWs['!cols'] = [{ wch: 30 }, { wch: 50 }];
    XLSX.utils.book_append_sheet(wb, summWs, 'Summary');

    const provName = (state.currentFile.provider?.name || 'provider').replace(/[^a-zA-Z0-9]/g, '_').substring(0, 30);
    XLSX.writeFile(wb, `Price_Check_${provName}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

})();
