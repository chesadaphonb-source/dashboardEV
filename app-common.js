/* =========================================================================
   app-common.js — โค้ดส่วนกลางที่ใช้ร่วมกันทุกหน้า
   หากต้องแก้ไขในอนาคต:
     - แก้ชื่อคอลัมน์ / หมวดงบประมาณ           -> ดูส่วน "CONFIG" ด้านล่าง
     - แก้วิธีอ่านไฟล์ Excel / จัดกลุ่มข้อมูล    -> ดูส่วน "PARSING & AGGREGATION"
     - แก้การจัดเก็บข้อมูลระหว่างหน้า (sessionStorage) -> ดูส่วน "SESSION STORE"
     - แก้แถบเมนูด้านบนของทุกหน้า              -> ดูส่วน "NAV BAR"
   ========================================================================= */

window.BudgetApp = (function () {

  // ------------------------------------------------------------------
  // CONFIG — ปรับตรงนี้ถ้าโครงสร้างไฟล์ Excel เปลี่ยน
  // ------------------------------------------------------------------
  const NUM_COLS = [
    'งบประมาณตั้งต้น', 'ปรับปรุงงบประมาณ', 'โอนย้ายงบประมาณ',
    'ขอจองเงิน', 'คืนเงิน', 'มูลค่าใบสั่งซื้อ', 'มูลค่าใบขอซื้อ',
    'ยอดเงินจริง', 'ยอดเงินที่เหลือ'
  ];
  const REQUIRED_COLS = ['โครงการ', 'งบประมาณตั้งต้น'];       // ใช้เช็คว่าชีทไหน "ใช่" ข้อมูลงบประมาณ
  const CATEGORY_NAMES = ['งบบุคลากร', 'งบดำเนินงาน', 'งบลงทุน', 'งบอุดหนุน', 'งบรายจ่ายอื่น', 'งบกลาง'];
  const CATEGORY_COLOR = {
    'งบบุคลากร': 'var(--teal)',
    'งบดำเนินงาน': 'var(--moss)',
    'งบลงทุน': 'var(--slate)',
    'งบอุดหนุน': 'var(--amber)',
    'งบรายจ่ายอื่น': 'var(--oxblood)',
    'งบกลาง': 'var(--plum)',
    'ไม่ระบุหมวด': 'var(--neutral)'
  };
  const SUMMARY_ROW_NAMES = ['รวม', 'total', 'grand total']; // แถวรวมท้ายชีท ต้องตัดออกไม่ให้นับซ้ำ
  const LARGE_FILE_WARN_MB = 15;
  const SESSION_KEY = 'budgetAppData_v1';

  // ------------------------------------------------------------------
  // FORMATTING HELPERS
  // ------------------------------------------------------------------
  function fmt(n, decimals) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    const d = decimals === undefined ? 0 : decimals;
    const abs = Math.abs(n);
    const s = abs.toLocaleString('th-TH', { minimumFractionDigits: d, maximumFractionDigits: 2 });
    return n < 0 ? '-' + s : s;
  }

  function numCell(value, decimals) {
    const cls = 'num' + (value < 0 ? ' neg' : '');
    return '<td class="' + cls + '">' + fmt(value, decimals) + '</td>';
  }

  function statusFor(pct, budget) {
    if (!budget || budget <= 0) return { label: 'ไม่มีงบ', cls: 'pill--none' };
    if (pct >= 90) return { label: 'ใกล้หมด', cls: 'pill--high' };
    if (pct >= 60) return { label: 'ปานกลาง', cls: 'pill--mid' };
    return { label: 'เหลือมาก', cls: 'pill--low' };
  }

  // ------------------------------------------------------------------
  // PARSING & AGGREGATION
  // ------------------------------------------------------------------
  function isSummaryRow(proj) {
    return SUMMARY_ROW_NAMES.includes(String(proj).trim().toLowerCase());
  }

  // สร้างคีย์เฉพาะของแถว ใช้เช็คว่าแถวนี้ "ซ้ำ" กับแถวที่เจอไปแล้วในชีทก่อนหน้าหรือไม่
  // แก้ไข: เดิมใช้แค่ (โครงการ + รหัสงบประมาณ) ซึ่งไม่ unique จริง เพราะโครงการเดียวกัน
  // มักมีหลายรายการย่อย (เอกสาร/ใบสั่งซื้อคนละใบ) ภายใต้รหัสงบประมาณเดียวกันได้ตามปกติ
  // ถ้าใช้แค่ 2 คอลัมน์นี้ รายการที่ถูกต้องหลายรายการจะถูกเข้าใจผิดว่าเป็นข้อมูลซ้ำและถูกตัดทิ้ง
  // จึงรวมค่าตัวเลขทุกคอลัมน์ + คำอธิบายเข้าไปในคีย์ด้วย เพื่อให้ "ซ้ำจริง" เท่านั้นที่ถูกตัด
  function rowKey(r, proj) {
    const numsPart = NUM_COLS.map(c => r[c]).join(',');
    const descPart = String(r['คำอธิบาย2'] || r['คำอธิบาย'] || '').trim();
    return proj + '|' + String(r['รหัสงบประมาณ'] || '').trim() + '|' + numsPart + '|' + descPart;
  }

  // อ่าน workbook -> คืนแถวดิบทั้งหมด (rawRows) + รายชื่อชีทที่ข้าม (skippedSheets)
  // สำคัญ: บางไฟล์มีชีทที่ export ซ้ำโครงการเดิม (เช่น Sheet2 คัดลอกบางโครงการจาก Sheet1) จึงต้องตัดรายการซ้ำ
  // ดูฟังก์ชัน rowKey() ด้านบนสำหรับนิยามว่าอะไรคือ "ซ้ำ"
  function extractRawRows(wb) {
    const rawRows = [];
    const skippedSheets = [];
    const seenKeys = new Set();
    let duplicateCount = 0;
    wb.SheetNames.forEach(sheetName => {
      const ws = wb.Sheets[sheetName];
      if (!ws) { skippedSheets.push(sheetName); return; } // ชีทอ่านไม่ขึ้น (มักเกิดจากไฟล์บวม)
      // defval: '' (ไม่ใช่ 0) — สำคัญ: ถ้าใช้ 0 แถวที่ช่อง "โครงการ" ว่างจะถูกเติมเป็นเลข 0
      // แล้ว String(0) = "0" ซึ่งเป็นค่า truthy จะไม่ถูกกรองทิ้ง กลายเป็นมี "โครงการ" ปลอมชื่อ "0" ปนเข้ามา
      const json = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
      if (json.length === 0) { skippedSheets.push(sheetName); return; }
      const headers = Object.keys(json[0]).map(h => h.trim());
      const hasAllRequired = REQUIRED_COLS.every(c => headers.includes(c));
      if (!hasAllRequired) { skippedSheets.push(sheetName); return; }
      json.forEach(r => {
        const proj = String(r['โครงการ'] || '').trim();
        if (!proj || isSummaryRow(proj)) return; // ตัดแถว "รวม" ทิ้ง
        const key = rowKey(r, proj);
        if (seenKeys.has(key)) { duplicateCount++; return; } // แถวนี้ซ้ำกับชีทก่อนหน้า ข้ามทิ้ง
        seenKeys.add(key);
        r.__sheet = sheetName;
        rawRows.push(r);
      });
    });
    return { rawRows, skippedSheets, duplicateCount };
  }

  // หาหมวดงบประมาณ + ชื่อโครงการ จากคำอธิบายของแถวตั้งต้นของแต่ละโครงการ
  function buildCategoryAndNameMaps(rawRows) {
    const categoryMap = {};
    const nameMap = {};
    rawRows.forEach(r => {
      const proj = String(r['โครงการ'] || '').trim();
      if (!nameMap[proj]) nameMap[proj] = String(r['คำอธิบาย'] || proj).trim();
      if (!categoryMap[proj]) {
        const desc = String(r['คำอธิบาย'] || '') + ' ' + String(r['คำอธิบาย2'] || '');
        const match = CATEGORY_NAMES.find(c => desc.indexOf(c) !== -1);
        if (match) categoryMap[proj] = match;
      }
    });
    return { categoryMap, nameMap };
  }

  // รวมยอดทุกคอลัมน์ตัวเลขต่อโครงการ (การรวม = การบวกทุกแถวย่อย ใช้ได้เพราะแต่ละแถวเป็นเอกสาร/รายการที่แยกกัน)
  function aggregateByProject(rawRows, categoryMap, nameMap) {
    const agg = {};
    rawRows.forEach(r => {
      const proj = String(r['โครงการ'] || '').trim();
      if (!agg[proj]) {
        agg[proj] = { code: proj, name: nameMap[proj] || proj, category: categoryMap[proj] || 'ไม่ระบุหมวด' };
        NUM_COLS.forEach(c => agg[proj][c] = 0);
      }
      NUM_COLS.forEach(c => { agg[proj][c] += Number(r[c]) || 0; });
    });
    return Object.values(agg).map(p => {
      const budget = p['งบประมาณตั้งต้น'] + p['ปรับปรุงงบประมาณ'] + p['โอนย้ายงบประมาณ'];
      const pct = budget > 0 ? (p['ยอดเงินจริง'] / budget * 100) : (p['ยอดเงินจริง'] > 0 ? 100 : 0);
      return Object.assign({}, p, { budget: budget, pct: pct });
    });
  }

  function computeTotals(projectRows) {
    return projectRows.reduce((t, p) => {
      t.budgetStart += p['งบประมาณตั้งต้น'];
      t.adjust += p['ปรับปรุงงบประมาณ'];
      t.transfer += p['โอนย้ายงบประมาณ'];
      t.reserve += p['ขอจองเงิน'];
      t.refund += p['คืนเงิน'];
      t.po += p['มูลค่าใบสั่งซื้อ'];
      t.pr += p['มูลค่าใบขอซื้อ'];
      t.actual += p['ยอดเงินจริง'];
      t.remaining += p['ยอดเงินที่เหลือ'];
      t.budget += p.budget;
      return t;
    }, { budgetStart:0, adjust:0, transfer:0, reserve:0, refund:0, po:0, pr:0, actual:0, remaining:0, budget:0 });
  }

  // ทำงานทั้งหมดตั้งแต่ workbook -> ผลลัพธ์พร้อมใช้
  function processWorkbook(wb) {
    const { rawRows, skippedSheets, duplicateCount } = extractRawRows(wb);
    const { categoryMap, nameMap } = buildCategoryAndNameMaps(rawRows);
    const projectRows = aggregateByProject(rawRows, categoryMap, nameMap);
    return { rawRows, skippedSheets, duplicateCount, projectRows };
  }

  // ------------------------------------------------------------------
  // SESSION STORE — ส่งข้อมูลระหว่างหน้า index / projects / transactions
  // (ใช้ sessionStorage: อยู่แค่ในแท็บนี้ หายเมื่อปิดแท็บ ไม่กระทบความเป็นส่วนตัวข้ามอุปกรณ์)
  // ------------------------------------------------------------------
  function saveSession(payload) {
    try {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify(payload));
      return true;
    } catch (e) {
      console.error('บันทึกข้อมูลระหว่างหน้าไม่สำเร็จ (ไฟล์อาจใหญ่เกินไป):', e);
      return false;
    }
  }
  function loadSession() {
    try {
      const raw = sessionStorage.getItem(SESSION_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }
  function clearSession() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
  }

  // ------------------------------------------------------------------
  // NAV BAR — เมนูด้านบน เหมือนกันทุกหน้า
  // ------------------------------------------------------------------
  function renderNav(activePage, sessionData) {
    const el = document.getElementById('navPlaceholder');
    if (!el) return;
    const fileInfo = sessionData
      ? '<span class="file-tag">' + sessionData.fileName + ' · ' + sessionData.projectRows.length + ' โครงการ</span>'
      : '';
    const resetBtn = sessionData ? '<button class="reset-link" id="navResetBtn">อัปโหลดไฟล์ใหม่</button>' : '';
    const tab = (href, label, key) =>
      '<a href="' + href + '" class="' + (activePage === key ? 'active' : '') + '">' + label + '</a>';
    el.innerHTML =
      '<div class="topbar">' +
        '<div class="brand"><p class="eyebrow">คณะสิ่งแวดล้อม มก. · ระบบติดตามงบประมาณ</p><h1>สมุดบัญชีงบประมาณ</h1></div>' +
        '<div class="tabs">' +
          tab('index.html', 'ภาพรวม', 'overview') +
          tab('charts.html', 'กราฟเปรียบเทียบ', 'charts') +
          tab('budget-codes.html', 'สรุปตามรหัสงบประมาณ', 'budgetcodes') +
          tab('projects.html', 'รายการโครงการ', 'projects') +
          tab('transactions.html', 'รายการธุรกรรมย่อย', 'transactions') +
        '</div>' +
        '<div style="display:flex; align-items:center; gap:10px;">' + fileInfo + resetBtn + '</div>' +
      '</div>';
    const rb = document.getElementById('navResetBtn');
    if (rb) rb.addEventListener('click', () => { clearSession(); location.href = 'index.html'; });
  }

  // แสดงข้อความ "ยังไม่มีข้อมูล กรุณาอัปโหลดที่หน้าแรก" — ใช้ในหน้า projects/transactions
  function renderEmptyState(container) {
    container.innerHTML =
      '<div class="empty-state">ยังไม่มีข้อมูล — กรุณาอัปโหลดไฟล์ Excel ที่ ' +
      '<a href="index.html">หน้าแรก (ภาพรวม)</a> ก่อน</div>';
  }

  // ------------------------------------------------------------------
  // BUDGET CODE ROLLUP HELPERS — สำหรับหน้า "สรุปตามรหัสงบประมาณ"
  // รหัสงบประมาณ เช่น "5120106-692B15PG00010" -> prefix คือ "5120106"
  // โค้ดเปล่าๆ ไม่มี "-ตามด้วยตัวอักษร" (เช่น "5120135" เดี่ยวๆ) ถือว่าเป็นแถวตั้งต้น/placeholder ไม่นับรวม
  // ------------------------------------------------------------------
  function extractCodePrefix(rawCode) {
    const code = String(rawCode || '').trim();
    const idx = code.indexOf('-');
    if (idx <= 0) return null;               // ไม่มีขีด หรือขีดอยู่ตำแหน่งแรก -> ไม่ใช่รูปแบบที่ต้องการ
    const after = code.slice(idx + 1).trim();
    if (!after) return null;                 // มีขีดแต่ไม่มีอะไรตามหลัง -> ไม่นับ
    return code.slice(0, idx).trim();
  }

  const CODE_NAME_KEY = 'budgetCodeNames_v1'; // localStorage: อยู่ถาวรข้ามการอัปโหลดไฟล์ใหม่ (ต่างจาก sessionStorage ของข้อมูลไฟล์)
  function loadCodeNames() {
    try { return JSON.parse(localStorage.getItem(CODE_NAME_KEY) || '{}'); } catch (e) { return {}; }
  }
  function saveCodeName(prefix, name) {
    try {
      const map = loadCodeNames();
      if (name && name.trim()) map[prefix] = name.trim(); else delete map[prefix];
      localStorage.setItem(CODE_NAME_KEY, JSON.stringify(map));
      return true;
    } catch (e) { return false; }
  }

  return {
    NUM_COLS, CATEGORY_NAMES, CATEGORY_COLOR, LARGE_FILE_WARN_MB,
    fmt, numCell, statusFor,
    processWorkbook, computeTotals,
    saveSession, loadSession, clearSession,
    renderNav, renderEmptyState,
    extractCodePrefix, loadCodeNames, saveCodeName
  };
})();
