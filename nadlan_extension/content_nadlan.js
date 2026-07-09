// content_nadlan.js — רץ על nadlan.taxes.gov.il
// שלב 1: ממלא הטופס אוטומטית
// שלב 2: לאחר CAPTCHA + תוצאות — שולף את כל העסקאות ומוריד Excel

(async () => {
'use strict';

const url = window.location.href;

// קרא בקשה מה-storage
const { nadlan_pending } = await chrome.storage.local.get('nadlan_pending');
if (!nadlan_pending?.gush) return;

const { gush, caseName, saveToDrive, propTypeName, requestId } = nadlan_pending;

// ─── עמוד חיפוש — מלא טופס ───────────────────────────────────────────
if (url.includes('startpageNadlanNewDesign') || url.includes('svinfonadlan2010/') && !url.includes('Perut')) {

  // המתן לטעינת הדף ולנוכחות rbMegush
  await new Promise(r => {
    const iv = setInterval(() => {
      if (document.getElementById('rbMegush') && document.readyState === 'complete') {
        clearInterval(iv); r();
      }
    }, 200);
    setTimeout(r, 5000);
  });
  await new Promise(r => setTimeout(r, 500));

  const rb = document.getElementById('rbMegush');
  if (!rb) return;

  // לחיצה על "לפי גוש/חלקה" — פעולה client-side בלבד, ללא PostBack
  rb.click();

  // המתן עד שתיבת גוש תהפוך לנראית (client-side DOM show/hide, מהיר)
  await new Promise(r => {
    let t = 0;
    const iv = setInterval(() => {
      const gf = document.getElementById('txtmegusha');
      if ((gf && gf.offsetParent !== null) || (t += 100) >= 2000) { clearInterval(iv); r(); }
    }, 100);
  });
  await new Promise(r => setTimeout(r, 300));

  // 1. מלא שדה גוש
  const gushInput = document.getElementById('txtmegusha');
  if (gushInput) {
    gushInput.value = gush;
    gushInput.dispatchEvent(new Event('input', { bubbles: true }));
    gushInput.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // 2. לחץ "העתקת גוש/חלקה" — זהו תג <a>, לא button!
  //    לחיצה מעתיקה את הגוש לשדות txtadGush ו-txtadHelka
  const copyBtn = document.getElementById('ContentUsersPage_copyGushHelka');
  if (copyBtn) {
    copyBtn.click();
    await new Promise(r => setTimeout(r, 400));
  }

  // 3. בחר סוג נכס — dispatch 'change' כדי לעדכן את DDLMahutIska
  const typeEl = document.getElementById('ContentUsersPage_DDLTypeNehes');
  if (typeEl) {
    let targetValue = '1'; // ברירת מחדל: דירת מגורים
    if (propTypeName) {
      const opt = Array.from(typeEl.options).find(o => o.text.trim().includes(propTypeName));
      if (opt) targetValue = opt.value;
    }
    typeEl.value = targetValue;
    // חשוב: dispatch change כדי שה-jQuery listener יפעיל ויעדכן DDLMahutIska
    typeEl.dispatchEvent(new Event('change', { bubbles: true }));
    await new Promise(r => setTimeout(r, 400));
  }

  // 4. בחר מהות עסקה = הכל (999)
  //    DDLMahutIska מתאפשר לאחר בחירת סוג נכס תקין
  const mahutEl = document.getElementById('ContentUsersPage_DDLMahutIska');
  if (mahutEl && !mahutEl.disabled) {
    mahutEl.value = '999';
  }

  // 5. טווח תאריכים = 36 חודשים
  const dateEl = document.getElementById('ContentUsersPage_DDLDateType');
  if (dateEl) dateEl.value = '5';

  await new Promise(r => setTimeout(r, 200));
  showBanner(`✅ גוש ${gush} מולא אוטומטית — ניתן לשנות ערכים, ואז לחץ חיפוש ופתור CAPTCHA`);
  return;
}

// ─── עמוד תוצאות — שלוף עסקאות ──────────────────────────────────────
if (!url.includes('InfoNadlanPerutWithMap')) return;

await new Promise(r => setTimeout(r, 600));

const ui = document.createElement('div');
ui.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;z-index:2147483647;background:rgba(0,0,0,.72);display:flex;align-items:center;justify-content:center;font-family:Arial;direction:rtl';
ui.innerHTML = `<div style="background:#fff;padding:28px 36px;border-radius:14px;min-width:380px;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.3)">
  <div style="font-size:22px;font-weight:bold;color:#154360;margin-bottom:10px">📊 ייצוא עסקאות</div>
  <div style="color:#555;font-size:13px;margin-bottom:8px">${caseName || ''}</div>
  <div id="_m" style="color:#333;font-size:14px;margin-bottom:12px">מאתחל...</div>
  <div style="background:#e9ecef;border-radius:6px;height:10px;overflow:hidden;margin-bottom:6px">
    <div id="_b" style="background:#27ae60;height:100%;width:0;transition:width .3s"></div>
  </div>
  <div id="_c" style="color:#888;font-size:12px">0 / ?</div>
</div>`;
document.body.appendChild(ui);

const upd = (t, d, tot) => {
  ui.querySelector('#_m').textContent = t;
  if (tot) {
    ui.querySelector('#_b').style.width = Math.round(d / tot * 100) + '%';
    ui.querySelector('#_c').textContent = d + ' / ' + tot;
  }
};

try {
  if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS לא זמין — בדוק שexceljs.min.js נמצא בתיקיית ה-Extension');

  // ── שדות ──
  const F = [
    ['lblEzor','אזור'],['lblGush','גוש-חלקה'],['lblTarIska','תאריך עסקה'],
    ['lblYeshuv','ישוב'],['lblRechov','רחוב'],['lblBayit','בית'],
    ['lblKnisa','כניסה'],['lblDira','דירה'],
    ['lblMcirMozhar','מחיר מוצהר (₪)'],['lblMcirMozharDlr','מחיר מוצהר ($)'],
    ['lblMcirMorach','מחיר מוערך (₪)'],['lblMcirMorachDlr','מחיר מוערך ($)'],
    ['lblNameShetachBruto','סוג שטח'],['lblShetachBruto','שטח ברוטו'],
    ['lblShetachNeto','שטח נטו'],['lblShnatBniya','שנת בנייה'],
    ['lblMisHadarim','חדרים'],['lblKoma','קומה'],
    ['lblMisKomot','קומות בבניין'],['lblDirotBnyn','דירות בבניין'],
    ['lblHanaya','חניה'],['lblGag','גג'],['lblMachsan','מחסן'],
    ['lblHzer','חצר'],['lblMigrash','מגרש'],['lblGlrya','גלריה'],
    ['lblMechirLmr','מחיר למ"ר'],['lblMechirCheder','מחיר לחדר'],
    ['lblMalit','מעלית'],
    ['lblSugIska','סוג עסקה'],
    ['lblTifkudBnyn','תפקוד בנין'],
    ['lblTifkudYchida','תפקוד יחידה'],
    ['lblShumaHalakim','שומה חלקים'],
    ['lblMahutZchut','מהות הזכות'],
    ['lblTava','לפי תבע'],
    ['lblMofaGush','מופעי גו"ח']
  ];

  // ── כמה עסקאות? ──
  const reshEl = document.querySelector('span[id*="resh"], div[id*="resh"]');
  const totalMatch = reshEl?.textContent.match(/\d+/);
  const TOTAL = totalMatch ? parseInt(totalMatch[0]) : 0;
  if (!TOTAL) throw new Error('לא נמצאו עסקאות או הדף לא נטען כראוי');

  const PER_PAGE = 12;
  const PAGES = Math.ceil(TOTAL / PER_PAGE);

  const formAction = document.forms[0].action;
  const getFormBase = () => {
    const p = {};
    for (const [k, v] of new FormData(document.forms[0]).entries()) p[k] = v;
    p['__EVENTARGUMENT'] = '';
    return p;
  };

  const fetchRow = async (ctl) => {
    const p = getFormBase();
    p['__EVENTTARGET'] = 'ctl00$ContentUsersPage$GridMultiD1$' + ctl + '$LogShow';
    const r = await fetch(formAction, { method: 'POST', body: new URLSearchParams(p), credentials: 'include', redirect: 'follow' });
    const html = await r.text();
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const row = {};
    for (const [id, lbl] of F) {
      const el = doc.getElementById('ContentUsersPage_' + id);
      row[lbl] = el ? el.textContent.trim() : '';
    }
    return row;
  };

  const goPage = async (n) => {
    const p = getFormBase();
    p['__EVENTTARGET'] = 'ctl00$ContentUsersPage$GridMultiD1';
    p['__EVENTARGUMENT'] = 'Page$' + n;
    const r = await fetch(formAction, { method: 'POST', body: new URLSearchParams(p), credentials: 'include', redirect: 'follow' });
    const html = await r.text();
    const d = new DOMParser().parseFromString(html, 'text/html');
    const vs = d.getElementById('__VIEWSTATE')?.value;
    if (vs) document.getElementById('__VIEWSTATE').value = vs;
    const ev = d.getElementById('__EVENTVALIDATION')?.value;
    if (ev) document.getElementById('__EVENTVALIDATION').value = ev;
  };

  // ── לולאה ראשית ──
  const rows = []; let idx = 1;
  for (let page = 1; page <= PAGES; page++) {
    if (page > 1) {
      upd('עובר לדף ' + page + '...', rows.length, TOTAL);
      await goPage(page);
      await new Promise(r => setTimeout(r, 500));
    }
    const perPage = page < PAGES ? PER_PAGE : TOTAL - (PER_PAGE * (PAGES - 1));
    for (let i = 0; i < perPage; i++) {
      const ctlN = i + 2;
      const ctl = 'ctl' + (ctlN < 10 ? '0' : '') + ctlN;
      upd('דף ' + page + ' — שורה ' + (i + 1) + ' / ' + perPage, idx - 1, TOTAL);
      try { const row = await fetchRow(ctl); row['#'] = idx; rows.push(row); }
      catch (e) { rows.push({ '#': idx, שגיאה: e.message }); }
      idx++;
      await new Promise(r => setTimeout(r, 200));
    }
  }

  // ── עיבוד ──
  upd('מעבד ומעצב נתונים...', TOTAL, TOTAL);
  await new Promise(r => setTimeout(r, 100));

  // סינון שורות ריקות
  const fLabels = F.map(([, l]) => l);
  const validRows = rows.filter(row => fLabels.some(l => row[l] && String(row[l]).trim() !== ''));

  const OUTPUT_COLS = [
    'תאריך עסקה','כתובת','ישוב',                         // 1-3
    'גוש','חלקה','תת חלקה',                              // 4-6
    'חדרים','קומה','שטח ברוטו','שטח נטו',                // 7-10
    'גג','חצר',                                           // 11-12
    'שנת בנייה',                                          // 13
    'שווי מכירה בש"ח','חלק נמכר','מחיר למ"ר ברוטו','מחיר למ"ר נטו', // 14-17
    'כניסה','דירה',                                       // 18-19
    'קומות בבניין','דירות בבניין',                        // 20-21
    'מעלית',                                               // 22
    'מגרש','מחסן','גלריה','חניה',                         // 23-26
    'סוג עסקה','תפקוד בנין','תפקוד יחידה',               // 27-29
    'מהות הזכות','לפי תבע','מופעי גו"ח'                  // 30-32
  ];
  const PRICE_COLS = new Set(['שווי מכירה בש"ח','מחיר למ"ר ברוטו','מחיר למ"ר נטו']);

  const toNum = (v, fn = parseFloat) => { const n = fn(String(v || '').replace(/,/g, '')); return isNaN(n) ? 0 : n; };
  // מספר אם ניתן לפרסר, אחרת מחרוזת (לשדות כמו דירה שיכולים להיות גם אותיות)
  const toNumOrStr = v => { const n = parseInt(String(v || '')); return isNaN(n) ? (v || '') : n; };

  const transformed = validRows.map(r => {
    const parts = (r['גוש-חלקה'] || '').split('-');
    const stripNum = s => { const n = parseInt((s || '').replace(/^0+/, '') || '0'); return isNaN(n) ? 0 : n; };
    const price = toNum(r['מחיר מוצהר (₪)'], parseInt);
    const bruto = toNum(r['שטח ברוטו']);
    const neto  = toNum(r['שטח נטו']);
    const addr  = [r['רחוב'], r['בית']].filter(v => v && v !== '--').join(' ');
    // חלק נמכר: "2 / 1 בלתי מסוימים" → 1/2 = 0.5
    const shumaNums = (r['שומה חלקים'] || '').match(/\d+/g);
    const halakVal  = shumaNums && shumaNums.length >= 2 && parseInt(shumaNums[0]) > 0
      ? parseInt(shumaNums[1]) / parseInt(shumaNums[0]) : '';
    return {
      'תאריך עסקה'       : r['תאריך עסקה'] || '',
      'כתובת'            : addr,
      'כניסה'            : toNumOrStr(r['כניסה']),
      'דירה'             : toNumOrStr(r['דירה']),
      'ישוב'             : r['ישוב']   || '',
      'גוש'              : stripNum(parts[0]),
      'חלקה'             : stripNum(parts[1]),
      'תת חלקה'          : stripNum(parts[2]),
      'חדרים'            : Math.round(toNum(r['חדרים'])),
      'קומה'             : Math.round(toNum(r['קומה'])),
      'שטח ברוטו'        : bruto,
      'שטח נטו'          : neto,
      'גג'               : toNum(r['גג']),
      'חצר'              : toNum(r['חצר']),
      'שנת בנייה'        : toNum(r['שנת בנייה'], parseInt) || '',
      'שווי מכירה בש"ח'  : price,
      'חלק נמכר'         : halakVal,
      // מחיר למ"ר ברוטו/נטו — יוגדרו כנוסחה בגיליון
      'מחיר למ"ר ברוטו'  : null,
      'מחיר למ"ר נטו'    : null,
      'קומות בבניין'     : toNum(r['קומות בבניין'], parseInt),
      'דירות בבניין'     : toNum(r['דירות בבניין'], parseInt),
      'מעלית'            : r['מעלית'] || '',
      'מגרש'             : toNum(r['מגרש']),
      'מחסן'             : toNum(r['מחסן']),
      'גלריה'            : toNum(r['גלריה']),
      'חניה'             : toNum((r['חניה'] || '').replace(/\s*רכבים\s*/g, ''), parseInt),
      'סוג עסקה'         : r['סוג עסקה'] || '',
      'תפקוד בנין'       : r['תפקוד בנין'] || '',
      'תפקוד יחידה'      : r['תפקוד יחידה'] || '',
      'מהות הזכות'       : r['מהות הזכות'] || '',
      'לפי תבע'          : toNum(r['לפי תבע']),
      'מופעי גו"ח'       : toNum(r['מופעי גו"ח'], parseInt)
    };
  });

  // ── מיון לפי תאריך עסקה מהחדש לישן ──
  const parseDate = s => {
    const [d, m, y] = (s || '').split('/');
    return y && m && d ? new Date(+y, +m - 1, +d) : new Date(0);
  };
  transformed.sort((a, b) => parseDate(b['תאריך עסקה']) - parseDate(a['תאריך עסקה']));

  // ── בנה workbook ──
  upd('בונה קובץ Excel...', TOTAL, TOTAL);
  await new Promise(r => setTimeout(r, 100));

  const wb = new ExcelJS.Workbook();

  // ── גיליון 1: גולמי (RTL, ללא עיצוב) ──
  const wsRaw = wb.addWorksheet('עסקאות גולמי', { views: [{ rightToLeft: true }] });
  const rawHeaders = F.map(([, l]) => l).concat(['#']);
  wsRaw.columns = rawHeaders.map(h => ({ header: h, key: h, width: Math.max(h.length + 4, 14) }));
  validRows.forEach(row => wsRaw.addRow(rawHeaders.map(h => row[h] ?? '')));

  // ── גיליון 2: מעוצב ──
  const toColLetter = n => { let s = ''; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; };
  const COL_BRUTO  = OUTPUT_COLS.indexOf('שטח ברוטו')       + 1; // 11 = K
  const COL_NETO   = OUTPUT_COLS.indexOf('שטח נטו')         + 1; // 12 = L
  const COL_SHOVI  = OUTPUT_COLS.indexOf('שווי מכירה בש"ח') + 1; // 16 = P
  const COL_HALAK  = OUTPUT_COLS.indexOf('חלק נמכר')        + 1; // 17 = Q
  const COL_LMR_B  = OUTPUT_COLS.indexOf('מחיר למ"ר ברוטו') + 1; // 18 = R
  const COL_LMR_N  = OUTPUT_COLS.indexOf('מחיר למ"ר נטו')   + 1; // 19 = S
  const K = toColLetter(COL_BRUTO), L = toColLetter(COL_NETO);
  const P = toColLetter(COL_SHOVI), Q = toColLetter(COL_HALAK);

  const wsStyled = wb.addWorksheet('עסקאות נדל"ן', { views: [{ rightToLeft: true }] });
  wsStyled.columns = OUTPUT_COLS.map(col => ({ header: col, key: col, width: Math.max(col.length + 4, 14) }));

  // ── המרת תאריך טקסט (dd/mm/yyyy) ל-Date אמיתי עבור Excel ──
  const parseDateToObj = s => {
    if (!s) return '';
    const [d, m, y] = String(s).split('/');
    if (!y || !m || !d) return s;
    return new Date(Date.UTC(+y, +m - 1, +d));
  };

  const COL_DATE = OUTPUT_COLS.indexOf('תאריך עסקה') + 1;

  // הוסף שורות עם נוסחאות לעמודות מחיר למ"ר
  const styledRows = [];
  transformed.forEach(tRow => {
    const values = OUTPUT_COLS.map(col => {
      if (col === 'מחיר למ"ר ברוטו' || col === 'מחיר למ"ר נטו') return null;
      if (col === 'תאריך עסקה') return parseDateToObj(tRow[col]);
      return tRow[col] ?? '';
    });
    styledRows.push(wsStyled.addRow(values));
  });
  styledRows.forEach((exRow, ri) => {
    const rn = ri + 2; // שורה 1 = כותרות
    exRow.getCell(COL_LMR_B).value = { formula: `${P}${rn}/${Q}${rn}/${K}${rn}` };
    exRow.getCell(COL_LMR_N).value = { formula: `${P}${rn}/${Q}${rn}/${L}${rn}` };
  });

  // סגנונות
  const THIN      = { style: 'thin', color: { argb: 'FF000000' } };
  const ALL_BRD   = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  const FONT_BASE = { name: 'David', size: 12 };
  const FONT_BOLD = { name: 'David', size: 12, bold: true };
  const ALIGN_CTR = { horizontal: 'center', vertical: 'middle', readingOrder: 2 };

  wsStyled.eachRow((row, rn) => {
    row.eachCell({ includeEmpty: true }, cell => {
      cell.font      = rn === 1 ? FONT_BOLD : FONT_BASE;
      cell.alignment = ALIGN_CTR;
      cell.border    = ALL_BRD;
    });
  });

  // פורמט תאריך ומטבע
  const priceColNums = [COL_SHOVI, COL_LMR_B, COL_LMR_N];
  wsStyled.eachRow((row, rn) => {
    if (rn === 1) return;
    row.getCell(COL_DATE).numFmt = 'dd/mm/yyyy';
    priceColNums.forEach(c => { row.getCell(c).numFmt = '"₪ "#,##0'; });
  });

  // ── ייצוא ──
  const buffer = await wb.xlsx.writeBuffer();

  const tikNum = (caseName || '').match(/(\d{4,6})\s*$/)?.[1] || gush;
  const filename = `${tikNum}.xlsx`;

  if (saveToDrive) {
    const uint8 = new Uint8Array(buffer);
    let binary = ''; const CS = 8192;
    for (let i = 0; i < uint8.length; i += CS) binary += String.fromCharCode(...uint8.subarray(i, i + CS));
    const base64 = btoa(binary);
    chrome.storage.local.remove('nadlan_pending');
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: validRows.length, filename, data: base64, saveToDrive: true, requestId });
  } else {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const burl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = burl; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(burl);
    chrome.storage.local.remove('nadlan_pending');
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: validRows.length, filename, requestId });
  }

  upd('✅ הסתיים! ' + validRows.length + ' עסקאות — ' + filename, TOTAL, TOTAL);
  setTimeout(() => ui.remove(), 5000);

} catch (e) {
  alert('❌ שגיאה: ' + e.message);
  ui.remove();
}

function showBanner(text) {
  const b = document.createElement('div');
  b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:9999999;background:#154360;color:#fff;text-align:center;padding:12px 20px;font-size:15px;font-family:Arial;direction:rtl;box-shadow:0 2px 8px rgba(0,0,0,.3)';
  b.textContent = text;
  document.body.prepend(b);
}

})();
