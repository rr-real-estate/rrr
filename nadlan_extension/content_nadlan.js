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

  // המתן לטעינת הדף המלאה — poll עד שהטופס קיים
  await new Promise(r => {
    const iv = setInterval(() => {
      if (document.getElementById('rbMegush') && document.readyState === 'complete') {
        clearInterval(iv); r();
      }
    }, 200);
    setTimeout(r, 5000); // fallback max 5s
  });
  await new Promise(r => setTimeout(r, 800));

  // לחץ "לפי גוש/חלקה" והמתן לרינדור השדות (VIEWSTATE מתחלף = postback הסתיים)
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

  // מלא גוש בשדה הראשון
  const gushInputs = Array.from(document.querySelectorAll('input[id*="gusha"],input[name*="gusha"]'))
    .filter(inp => inp.type !== 'checkbox' && inp.type !== 'radio');
  if (gushInputs[0]) gushInputs[0].value = gush;

  // לחץ "העתקת גוש חלקה" (מעתיק לשדה השני)
  const copyBtn = Array.from(document.querySelectorAll('input[type=button],button'))
    .find(b => (b.value || b.textContent || '').includes('העתקת'));
  if (copyBtn) {
    copyBtn.click();
    await new Promise(r => setTimeout(r, 400));
  } else if (gushInputs[1]) {
    gushInputs[1].value = gush; // fallback — מלא ידנית
  }

  // סוג נכס — קבע ללא postback
  const typeEl = document.getElementById('ContentUsersPage_DDLTypeNehes');
  if (typeEl) {
    let targetValue = '1';
    if (propTypeName) {
      const opt = Array.from(typeEl.options).find(o => o.text.trim().includes(propTypeName));
      if (opt) targetValue = opt.value;
    }
    typeEl.value = targetValue;
  }

  // המתן עד שDDLSubTypeNehes טעון (יש בו יותר מאפשרות אחת)
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

  // טווח זמן: 5 = 36 חודשים
  const dateEl = document.getElementById('ContentUsersPage_DDLDateType');
  if (dateEl) dateEl.value = '5';

  await new Promise(r => setTimeout(r, 300));

  // הצג הודעה — המשתמש לוחץ חיפוש בעצמו
  showBanner(`✅ גוש ${gush} מולא אוטומטית — ניתן לשנות ערכים, ואז לחץ חיפוש ופתור CAPTCHA`);

  return;
}

// ─── עמוד תוצאות — שלוף עסקאות ──────────────────────────────────────
if (!url.includes('InfoNadlanPerutWithMap')) return;

await new Promise(r => setTimeout(r, 600));

// ── UI פרוגרס ──
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
  // ── SheetJS נטען כ-content script לפני קובץ זה (manifest.json) ──
  if (typeof XLSX === 'undefined') throw new Error('XLSX לא זמין — בדוק שxlsx.full.min.js נמצא בתיקיית ה-Extension');

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
    ['lblMechirLmr','מחיר למ"ר'],['lblMechirCheder','מחיר לחדר']
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

  // ── שליפת שורה בודדת ──
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

  // ── מעבר דף ──
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

  // ── בנה Excel ──
  upd('בונה קובץ Excel...', TOTAL, TOTAL);
  await new Promise(r => setTimeout(r, 100));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 4 }, ...F.map(([, l]) => ({ wch: Math.max(l.length + 4, 14) }))];
  XLSX.utils.book_append_sheet(wb, ws, 'עסקאות');

  // שם קובץ — מספר התיק (5 ספרות מסוף caseName), או גוש אם אין
  const tikNum = (caseName || '').match(/(\d{4,6})\s*$/)?.[1] || gush;
  const filename = `${tikNum}.xlsx`;

  if (saveToDrive) {
    // שלח base64 בחזרה לאפליקציה — היא תשמור ל-Drive
    const base64 = XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
    chrome.storage.local.remove('nadlan_pending');
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: rows.length, filename, data: base64, saveToDrive: true, requestId });
  } else {
    XLSX.writeFile(wb, filename);
    chrome.storage.local.remove('nadlan_pending');
    chrome.runtime.sendMessage({ type: 'NADLAN_DONE', gush, count: rows.length, filename, requestId });
  }

  upd('✅ הסתיים! ' + rows.length + ' עסקאות — ' + filename, TOTAL, TOTAL);
  setTimeout(() => ui.remove(), 5000);

} catch (e) {
  alert('❌ שגיאה: ' + e.message);
  ui.remove();
}

// ── עזר: banner ──
function showBanner(text) {
  const b = document.createElement('div');
  b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:9999999;background:#154360;color:#fff;text-align:center;padding:12px 20px;font-size:15px;font-family:Arial;direction:rtl;box-shadow:0 2px 8px rgba(0,0,0,.3)';
  b.textContent = text;
  document.body.prepend(b);
}

})();
