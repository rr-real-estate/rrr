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

  await new Promise(r => {
    const iv = setInterval(() => {
      if (document.getElementById('rbMegush') && document.readyState === 'complete') {
        clearInterval(iv); r();
      }
    }, 200);
    setTimeout(r, 5000);
  });
  await new Promise(r => setTimeout(r, 800));

  const rb = document.getElementById('rbMegush');
  if (!rb) return;
  const vsBeforeRb = document.getElementById('__VIEWSTATE')?.value;
  rb.click();
  await new Promise(r => {
    let t = 0;
    const iv = setInterval(() => {
      if (document.getElementById('__VIEWSTATE')?.value !== vsBeforeRb || (t += 200) >= 5000) {
        clearInterval(iv); r();
      }
    }, 200);
  });
  await new Promise(r => setTimeout(r, 600));

  const gushInputs = Array.from(document.querySelectorAll('input[id*="gusha"],input[name*="gusha"]'))
    .filter(inp => inp.type !== 'checkbox' && inp.type !== 'radio');
  if (gushInputs[0]) gushInputs[0].value = gush;

  const copyBtn = Array.from(document.querySelectorAll('input[type=button],button'))
    .find(b => (b.value || b.textContent || '').includes('העתקת'));
  if (copyBtn) {
    copyBtn.click();
    await new Promise(r => setTimeout(r, 400));
  } else if (gushInputs[1]) {
    gushInputs[1].value = gush;
  }

  const typeEl = document.getElementById('ContentUsersPage_DDLTypeNehes');
  if (typeEl) {
    let targetValue = '1';
    if (propTypeName) {
      const opt = Array.from(typeEl.options).find(o => o.text.trim().includes(propTypeName));
      if (opt) targetValue = opt.value;
    }
    typeEl.value = targetValue;
  }

  const subTypeEl = document.getElementById('ContentUsersPage_DDLSubTypeNehes');
  if (subTypeEl) {
    await new Promise(r => {
      let t = 0;
      const iv = setInterval(() => {
        if (subTypeEl.options.length > 1 || (t += 200) >= 5000) { clearInterval(iv); r(); }
      }, 200);
    });
    const allOpt = Array.from(subTypeEl.options).find(o =>
      o.text.trim() === 'הכל' || o.value === '' || o.value === '0'
    );
    if (allOpt) subTypeEl.value = allOpt.value;
  }

  const dateEl = document.getElementById('ContentUsersPage_DDLDateType');
  if (dateEl) dateEl.value = '5';

  await new Promise(r => setTimeout(r, 300));
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

  const OUTPUT_COLS = [
    'תאריך עסקה','כתובת','כניסה','דירה','ישוב',
    'גוש','חלקה','תת חלקה',
    'חדרים','קומה','שטח ברוטו','שטח נטו',
    'מגרש','גג','מחסן','חצר','גלריה','חניה',
    'קומות בבניין','דירות בבניין','שנת בנייה',
    'שווי מכירה בש"ח','מחיר למ"ר ברוטו','מחיר למ"ר נטו',
    'מעלית','סוג עסקה','תפקוד בנין','תפקוד יחידה',
    'שומה חלקים','מהות הזכות','לפי תבע','מופעי גו"ח'
  ];
  const PRICE_COLS = new Set(['שווי מכירה בש"ח','מחיר למ"ר ברוטו','מחיר למ"ר נטו']);

  const transformed = rows.map(r => {
    const parts = (r['גוש-חלקה'] || '').split('-');
    const strip = s => (s || '').replace(/^0+/, '') || '0';
    const price = parseInt((r['מחיר מוצהר (₪)'] || '0').replace(/,/g, '')) || 0;
    const bruto = parseFloat(r['שטח ברוטו']) || 0;
    const neto  = parseFloat(r['שטח נטו'])   || 0;
    const addr  = [r['רחוב'], r['בית']].filter(v => v && v !== '--').join(' ');
    return {
      'תאריך עסקה'       : r['תאריך עסקה'] || '',
      'כתובת'            : addr,
      'כניסה'            : r['כניסה'] || '',
      'דירה'             : r['דירה']   || '',
      'ישוב'             : r['ישוב']   || '',
      'גוש'              : strip(parts[0]),
      'חלקה'             : strip(parts[1]),
      'תת חלקה'          : strip(parts[2]),
      'חדרים'            : Math.round(parseFloat(r['חדרים']) || 0),
      'קומה'             : Math.round(parseFloat(r['קומה'])  || 0),
      'שטח ברוטו'        : bruto,
      'שטח נטו'          : neto,
      'מגרש'             : r['מגרש']  || '',
      'גג'               : r['גג']    || '',
      'מחסן'             : r['מחסן'] || '',
      'חצר'              : r['חצר']  || '',
      'גלריה'            : r['גלריה'] || '',
      'חניה'             : (r['חניה'] || '').replace(/\s*רכבים\s*/g, '').trim(),
      'קומות בבניין'     : r['קומות בבניין'] || '',
      'דירות בבניין'     : r['דירות בבניין'] || '',
      'שנת בנייה'        : r['שנת בנייה'] || '',
      'שווי מכירה בש"ח'  : price,
      'מחיר למ"ר ברוטו'  : bruto > 0 ? Math.round(price / bruto) : 0,
      'מחיר למ"ר נטו'    : neto  > 0 ? Math.round(price / neto)  : 0,
      'מעלית'            : r['מעלית'] || '',
      'סוג עסקה'         : r['סוג עסקה'] || '',
      'תפקוד בנין'       : r['תפקוד בנין'] || '',
      'תפקוד יחידה'      : r['תפקוד יחידה'] || '',
      'שומה חלקים'       : r['שומה חלקים'] || '',
      'מהות הזכות'       : r['מהות הזכות'] || '',
      'לפי תבע'          : r['לפי תבע'] || '',
      'מופעי גו"ח'       : r['מופעי גו"ח'] || ''
    };
  });

  // ── בנה workbook ──
  upd('בונה קובץ Excel...', TOTAL, TOTAL);
  await new Promise(r => setTimeout(r, 100));

  const wb = new ExcelJS.Workbook();

  // ── גיליון 1: גולמי (RTL, ללא עיצוב) ──
  const wsRaw = wb.addWorksheet('עסקאות גולמי', { views: [{ rightToLeft: true }] });
  const rawHeaders = F.map(([, l]) => l).concat(['#']);
  wsRaw.columns = rawHeaders.map(h => ({ header: h, key: h, width: Math.max(h.length + 4, 14) }));
  rows.forEach(row => wsRaw.addRow(rawHeaders.map(h => row[h] ?? '')));

  // ── גיליון 2: מעוצב ──
  const wsStyled = wb.addWorksheet('עסקאות נדל"ן', { views: [{ rightToLeft: true }] });
  wsStyled.columns = OUTPUT_COLS.map(col => ({ header: col, key: col, width: Math.max(col.length + 4, 14) }));
  transformed.forEach(row => wsStyled.addRow(OUTPUT_COLS.map(col => row[col] ?? '')));

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

  // פורמט מטבע
  const priceColNums = OUTPUT_COLS.reduce((a, c, i) => { if (PRICE_COLS.has(c)) a.push(i + 1); return a; }, []);
  wsStyled.eachRow((row, rn) => {
    if (rn === 1) return;
    priceColNums.forEach(c => { row.getCell(c).numFmt = '#,##0'; });
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
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: rows.length, filename, data: base64, saveToDrive: true, requestId });
  } else {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const burl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = burl; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(burl);
    chrome.storage.local.remove('nadlan_pending');
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: rows.length, filename, requestId });
  }

  upd('✅ הסתיים! ' + rows.length + ' עסקאות — ' + filename, TOTAL, TOTAL);
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
