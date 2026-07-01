// background.js — Service Worker

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {

  // פתח טאב חדש לאתר הנדלן
  if (msg.type === 'OPEN_NADLAN') {
    chrome.tabs.create({
      url: 'https://nadlan.taxes.gov.il/svinfonadlan2010/startpageNadlanNewDesign.aspx'
    });
  }

  // העבר תוצאות מהנדלן חזרה לאפליקציה — שלח לכל לשוניות האתר
  if (msg.type === 'NADLAN_DONE') {
    chrome.tabs.query({ url: 'https://rr-real-estate.github.io/*' }, (tabs) => {
      console.log('[RRR] NADLAN_DONE received, found', tabs.length, 'tabs, reqId:', msg.requestId);
      tabs.forEach(tab => {
        console.log('[RRR] sending NADLAN_RESULT to tab', tab.id);
        chrome.tabs.sendMessage(tab.id, { ...msg, type: 'NADLAN_RESULT' });
      });
    });
  }

});
