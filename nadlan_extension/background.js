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
      tabs.forEach(tab => {
        chrome.tabs.sendMessage(tab.id, { type: 'NADLAN_RESULT', ...msg });
      });
    });
  }

});
