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
  await new Promise(r => setTimeout(r, 900));

  // לחץ על "לפי גוש/חלקה" והמתן לרינדור השדות
  const rb = document.getElementById('rbMegush');
  if (rb) { rb.click(); await new Promise(r => setTimeout(r, 600)); }

  // מלא גוש — כל שדות הגוש (מגוש + עד גוש)
  document.querySelectorAll('input[id*="gusha"],input[name*="gusha"]').forEach(inp => {
    if (inp.type !== 'checkbox' && inp.type !== 'radio') inp.value = gush;
  });

  // סוג נכס + מהות עסקה — trigger postback ע"י החלפת סוג זמנית
  const typeEl = document.getElementById('ContentUsersPage_DDLTypeNehes');
  if (typeEl && typeEl.options.length > 1) {
    // קבע ערך יעד
    let targetValue = '1';
    if (propTypeName) {
      const opt = Array.from(typeEl.options).find(o => o.text.trim().includes(propTypeName));
      if (opt) targetValue = opt.value;
    }
    // עזר: המתן לרענון VIEWSTATE (postback)
    const waitPostback = async (ms = 2500) => {
      const before = document.getElementById('__VIEWSTATE')?.value;
      await new Promise(r => {
        const t = Date.now();
        const iv = setInterval(() => {
          if (document.getElementById('__VIEWSTATE')?.value !== before || Date.now()-t > ms) {
            clearInterval(iv); r();
          }
        }, 150);
      });
      await new Promise(r => setTimeout(r, 300));
    };
    // בחר ערך שונה מהיעד כדי "לשחרר" את DDLSubTypeNehes
    const tempOpt = Array.from(typeEl.options).find(o =>
      o.value && o.value !== '' && o.value !== '0' && o.value !== targetValue
    );
    if (tempOpt) {
      typeEl.value = tempOpt.value;
      typeEl.dispatchEvent(new Event('change', { bubbles: true }));
      await waitPostback();
    }
    // עכשיו החזר לסוג הנכון
    typeEl.value = targetValue;
    typeEl.dispatchEvent(new Event('change', { bubbles: true }));
    await waitPostback();
  }

  // מהות עסקה: הכל — עכשיו DDLSubTypeNehes מעודכן מהשרת
  const subTypeEl = document.getElementById('ContentUsersPage_DDLSubTypeNehes');
  if (subTypeEl) {
    const allOpt = Array.from(subTypeEl.options).find(o => o.text.trim() === 'הכל' || o.value === '' || o.value === '0');
    if (allOpt) subTypeEl.value = allOpt.value;
  }

  // טווח זמן: 5 = 36 חודשים
  const dateEl = document.getElementById('ContentUsersPage_DDLDateType');
  if (dateEl) dateEl.value = '5';

  await new Promise(r => setTimeout(r, 400));

  // הצג הודעה למשתמש
  showBanner(`✅ גוש ${gush} מולא אוטומטית — לחץ חיפוש ופתור את ה-CAPTCHA`);

  // לחץ חיפוש
  const btn = document.getElementById('ContentUsersPage_btnHipus');
  if (btn) btn.click();

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

  const d = new Date();
  const stamp = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
  const filename = `עסקאות_גוש${gush}_${rows.length}רשומות_${stamp}.xlsx`;

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
