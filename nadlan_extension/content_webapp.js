// content_webapp.js — רץ על rr-real-estate.github.io
// מאפשר לאפליקציה לתקשר עם ה-Extension

// סמן שה-Extension מותקן (האפליקציה בודקת זאת)
document.documentElement.setAttribute('data-rrr-ext', '1');

// קבל בקשת הורדה מהאפליקציה
window.addEventListener('RRR_NADLAN_REQUEST', (e) => {
  const { gush, caseFolder, caseName, saveToDrive, folderName, propTypeName, requestId } = e.detail;
  chrome.storage.local.set({ nadlan_pending: { gush, caseFolder, caseName, saveToDrive, folderName, propTypeName, requestId } });
  chrome.runtime.sendMessage({ type: 'OPEN_NADLAN' });
});

// קבל תוצאה מחזרה מ-content_nadlan.js ושלח לאפליקציה
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'NADLAN_RESULT') {
    window.dispatchEvent(new CustomEvent('RRR_NADLAN_RESULT', { detail: msg }));
  }
});
