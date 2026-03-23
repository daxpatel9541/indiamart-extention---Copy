// ============================================
// IndiaMART Lead Extractor — Background Service Worker
// ============================================

// Extension state
let extractionState = {
  isRunning: false,
  currentIndex: 0,
  totalMessages: 0
};

// Listen for messages from popup and content script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {

    case 'getState':
      sendResponse({ state: extractionState });
      return true;

    case 'setState':
      extractionState = { ...extractionState, ...message.payload };
      sendResponse({ success: true });
      return true;

    case 'getLeads':
      chrome.storage.local.get(['leads', 'processedIds'], (result) => {
        sendResponse({
          leads: result.leads || [],
          processedIds: result.processedIds || [],
          count: (result.leads || []).length
        });
      });
      return true; // async sendResponse

    case 'saveLead':
      chrome.storage.local.get(['leads', 'processedIds'], (result) => {
        const leads = result.leads || [];
        const processedIds = result.processedIds || [];
        const lead = message.payload;

        if (!processedIds.includes(lead.uniqueId)) {
          leads.push(lead);
          processedIds.push(lead.uniqueId);
          chrome.storage.local.set({ leads, processedIds }, () => {
            sendResponse({ success: true, newCount: leads.length });
          });
        } else {
          sendResponse({ success: false, reason: 'duplicate', count: leads.length });
        }
      });
      return true; // async sendResponse

    case 'clearLeads':
      chrome.storage.local.set({ leads: [], processedIds: [] }, () => {
        extractionState = { isRunning: false, currentIndex: 0, totalMessages: 0 };
        sendResponse({ success: true });
      });
      return true;

    case 'checkDuplicate':
      chrome.storage.local.get(['processedIds'], (result) => {
        const processedIds = result.processedIds || [];
        sendResponse({ isDuplicate: processedIds.includes(message.uniqueId) });
      });
      return true;

    default:
      sendResponse({ error: 'Unknown action' });
      return false;
  }
});

// On install, initialize storage
chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.get(['leads', 'processedIds'], (result) => {
    if (!result.leads) {
      chrome.storage.local.set({ leads: [], processedIds: [] });
    }
  });
  console.log('[IndiaMART Extractor] Extension installed and storage initialized.');
});
