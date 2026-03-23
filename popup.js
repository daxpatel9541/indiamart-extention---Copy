// ============================================
// IndiaMART Lead Extractor — Popup Script v2
// With Google Sheets integration
// ============================================

const $ = (id) => document.getElementById(id);

const logArea = $('logArea');
const btnStart = $('btnStart');
const btnStop = $('btnStop');
const btnExport = $('btnExport');
const btnClear = $('btnClear');
const totalLeads = $('totalLeads');
const currentProgress = $('currentProgress');
const skippedCount = $('skippedCount');
const progressWrapper = $('progressWrapper');
const progressFill = $('progressFill');
const progressText = $('progressText');

let skipped = 0;

function addLog(msg, type = 'info') {
  const entry = document.createElement('div');
  entry.className = `log-entry ${type}`;
  const time = new Date().toLocaleTimeString('en-IN', { hour12: false });
  entry.textContent = `[${time}] ${msg}`;
  logArea.appendChild(entry);
  logArea.scrollTop = logArea.scrollHeight;
}

// --- Init ---
document.addEventListener('DOMContentLoaded', () => {
  chrome.storage.local.get(['leads', 'isExtracting'], (result) => {
    const count = (result.leads || []).length;
    totalLeads.textContent = count;

    if (result.isExtracting) {
      setUIRunning();
      addLog('Resuming extraction status...', 'info');
    } else if (count > 0) {
      addLog(`Found ${count} previously extracted leads.`, 'info');
    }
  });
});

function setUIRunning() {
  btnStart.style.display = 'none';
  btnStop.style.display = 'flex';
  progressWrapper.style.display = 'flex';
  document.body.classList.add('extracting');
}

// --- Start Extraction ---
btnStart.addEventListener('click', () => {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    const tab = tabs[0];
    if (!tab) { addLog('No active tab.', 'error'); return; }

    if (!tab.url || !tab.url.includes('seller.indiamart.com')) {
      addLog('Not on IndiaMART. Navigating to Lead Manager...', 'warn');
      chrome.tabs.update(tab.id, { url: 'https://seller.indiamart.com/messagecentre/' }, () => {
        addLog('Navigated! Wait for page to load, then click Start again.', 'info');
      });
      return;
    }

    addLog('🚀 Starting extraction...', 'success');
    btnStart.style.display = 'none';
    btnStop.style.display = 'flex';
    progressWrapper.style.display = 'flex';
    document.body.classList.add('extracting');
    skipped = 0;
    skippedCount.textContent = '0';

    chrome.tabs.sendMessage(tab.id, { action: 'startExtraction' }, (response) => {
      if (chrome.runtime.lastError) {
        addLog('Content script not ready. Refresh the page and try again.', 'error');
        resetUI();
        return;
      }
      if (response && response.ack) {
        addLog('Content script acknowledged. Working...', 'success');
      }
    });
  });
});

// --- Stop ---
btnStop.addEventListener('click', () => {
  chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
    if (tabs[0]) chrome.tabs.sendMessage(tabs[0].id, { action: 'stopExtraction' });
  });
  addLog('Stopping...', 'warn');
  resetUI();
});

// --- Progress messages from content script ---
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === 'progress') {
    if (msg.total > 0) {
      currentProgress.textContent = `${msg.current}/${msg.total}`;
      const pct = Math.round((msg.current / msg.total) * 100);
      progressFill.style.width = `${pct}%`;
      progressText.textContent = `${pct}%`;
    } else {
      currentProgress.textContent = `${msg.current}`;
      progressFill.style.width = `100%`;
      progressText.textContent = `...`;
    }

    if (msg.status === 'extracted') {
      addLog(`✅ ${msg.companyName || 'Unknown'}`, 'success');
    } else if (msg.status === 'skipped') {
      skipped++;
      skippedCount.textContent = skipped;
      addLog(`⏭️ Skipped: ${msg.companyName || 'Unknown'}`, 'warn');
    } else if (msg.status === 'error') {
      addLog(`❌ Error: ${msg.companyName || 'See console'}`, 'error');
    }
  }
  if (msg.type === 'leadCount') totalLeads.textContent = msg.count;
  if (msg.type === 'log') addLog(msg.text, msg.level || 'info');
  if (msg.type === 'done') {
    addLog(`🎉 Done! ${msg.total} total leads.`, 'success');
    totalLeads.textContent = msg.total;
    resetUI();
  }
  if (msg.type === 'extractionError') {
    addLog(`❌ ${msg.text}`, 'error');
    resetUI();
  }
});

function resetUI() {
  btnStart.style.display = 'flex';
  btnStop.style.display = 'none';
  document.body.classList.remove('extracting');
  chrome.storage.local.set({ isExtracting: false });
}

// --- Export Excel ---
btnExport.addEventListener('click', () => {
  chrome.runtime.sendMessage({ action: 'getLeads' }, (response) => {
    if (!response || !response.leads || response.leads.length === 0) {
      addLog('No leads to export.', 'warn');
      return;
    }

    addLog(`Exporting ${response.leads.length} leads to Excel...`, 'extract');

    try {
      const data = response.leads.map((lead, idx) => ({
        'S.No': idx + 1,
        'Company Name': lead.companyName || '',
        'Mobile Number': lead.mobileNumber || '',
        'Contact Person': lead.contactPerson || '',
        'Address': lead.address || '',
        'Rating': lead.rating || '',
        'Review Count': lead.reviewCount || '',
        'Review 1': lead.review1 || '',
        'Review 2': lead.review2 || '',
        'Response Rate': lead.responseRate || '',
        'Quality Rate': lead.qualityRate || '',
        'Delivery Rate': lead.deliveryRate || '',
        'Business Type': lead.businessType || '',
        'Company Owner': lead.ownerName || '',
        'Total Employees': lead.totalEmployees || '',
        'Year of Establishment': lead.yearEstablished || '',
        'IndiaMART Member Since': lead.memberSince || '',
        'Annual Turnover': lead.annualTurnover || '',
        'GST No.': lead.gstNumber || '',
        'PAN': lead.panNumber || '',
        'Sells / Products': lead.products || ''
      }));

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(data);

      // Auto-width
      const colWidths = Object.keys(data[0]).map(key => {
        const maxLen = Math.max(key.length, ...data.map(r => String(r[key] || '').length));
        return { wch: Math.min(maxLen + 2, 50) };
      });
      ws['!cols'] = colWidths;

      XLSX.utils.book_append_sheet(wb, ws, 'Leads');
      XLSX.writeFile(wb, 'IndiaMart_Leads.xlsx');
      addLog('✅ Excel downloaded: IndiaMart_Leads.xlsx', 'success');
    } catch (err) {
      addLog(`Export error: ${err.message}`, 'error');
    }
  });
});

// --- Clear ---
btnClear.addEventListener('click', () => {
  if (!confirm('Clear all extracted leads? Cannot be undone.')) return;
  chrome.runtime.sendMessage({ action: 'clearLeads' }, (resp) => {
    if (resp && resp.success) {
      totalLeads.textContent = '0';
      currentProgress.textContent = '—';
      skippedCount.textContent = '0';
      progressFill.style.width = '0%';
      progressText.textContent = '0%';
      addLog('🗑️ All data cleared.', 'info');
    }
  });
});
