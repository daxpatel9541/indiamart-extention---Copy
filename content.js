// ============================================
// IndiaMART Lead Extractor — Content Script v2
// Based on actual IndiaMART DOM structure
// ============================================

(function () {
  'use strict';

  let isRunning = false;
  let shouldStop = false;

  const CONFIG = {
    LEAD_MANAGER_URL: 'https://seller.indiamart.com/messagecentre/',
    INITIAL_LOAD_WAIT: 4000,
    CLICK_DELAY: 2000,
    DETAIL_LOAD_DELAY: 3000,
    BETWEEN_MESSAGES: 2000,
    SCROLL_DELAY: 1000,
    MOBILE_EXTRACT_DELAY: 1500,
    MAX_RETRIES: 5
  };

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function waitForDomStable(timeout = 2000) {
    return new Promise(resolve => {
      let timer;
      const observer = new MutationObserver(() => {
        clearTimeout(timer);
        timer = setTimeout(() => { observer.disconnect(); resolve(); }, 500);
      });
      observer.observe(document.body, { childList: true, subtree: true });
      timer = setTimeout(() => { observer.disconnect(); resolve(); }, timeout);
    });
  }

  function generateUniqueId(companyName, address, mobile = '') {
    // Priority 1: Mobile + Address (Most unique)
    // Priority 2: Company Name + Address
    const base = mobile ? `${mobile.replace(/\D/g, '')}__${(address || '').trim().toLowerCase()}` : 
                          `${(companyName || '').trim().toLowerCase()}__${(address || '').trim().toLowerCase()}`;
    
    let hash = 0;
    for (let i = 0; i < base.length; i++) {
      hash = ((hash << 5) - hash) + base.charCodeAt(i);
      hash |= 0;
    }
    return 'lead_' + Math.abs(hash).toString(36);
  }

  function sendLog(text, level = 'info') {
    console.log(`[IM-Extractor][${level}] ${text}`);
    try { chrome.runtime.sendMessage({ type: 'log', text, level }); } catch (e) {}
  }

  function sendProgress(current, total, companyName, status) {
    try { chrome.runtime.sendMessage({ type: 'progress', current, total, companyName, status }); } catch (e) {}
  }

  // ============================================
  // STEP 1: Find all message items in left panel
  // ============================================
  function getMessageItems() {
    // From screenshot: message items are in a list under the "All / Unread" tabs.
    // Each item contains: profile image, company name, location, preview text, date.
    // We'll use multiple strategies to identify them.

    let items = [];

    // --- Strategy A: Common IndiaMART class patterns ---
    const selectors = [
      // IndiaMART message centre known patterns
      '.lm_lst_item', '.mc_lt_li', '.msg-item', '.chat-list-item',
      '[class*="mc_lt"] > div', '[class*="mc_lt"] > li',
      '#leftMsg > div', '#leftMsg > li', '#leftMsg li',
      '[class*="msgList"] > div', '[class*="msgList"] > li',
      '[class*="msg_list"] > div', '[class*="msg_list"] > li',
      '[class*="leftPanel"] li', '[class*="left_panel"] li',
      '[class*="chatList"] > div', '[class*="chat_list"] > div',
      // Data attribute based
      '[data-enqid]', '[data-enquiry]', '[data-msgid]', '[data-glid]',
      '[data-uid]', '[data-contact]',
    ];

    for (const sel of selectors) {
      try {
        items = Array.from(document.querySelectorAll(sel));
        items = items.filter(el => el.offsetHeight > 30 && el.offsetWidth > 100);
        if (items.length >= 2) {
          sendLog(`Found ${items.length} messages with selector: "${sel}"`, 'success');
          return items;
        }
      } catch (e) {}
    }    // --- Strategy B: Find "Messages" header and get sibling list ---
    try {
      const allElements = document.querySelectorAll('h1, h2, h3, h4, div, span, p');
      for (const el of allElements) {
        const text = el.textContent.trim();
        if (text === 'Messages' || text === 'All Messages') {
          let container = el.parentElement;
          for (let depth = 0; depth < 8; depth++) {
            if (!container) break;
            const kids = Array.from(container.children);
            const messageItems = kids.filter(c => {
               const h = c.offsetHeight;
               const w = c.offsetWidth;
               return h >= 50 && h <= 180 && w >= 200 && c.textContent.trim().length > 10;
            });
            if (messageItems.length >= 2) {
              sendLog(`Found ${messageItems.length} messages via "Messages" header strategy`, 'success');
              return messageItems;
            }
            container = container.parentElement;
          }
        }
      }
    } catch (e) {}

    // --- Strategy C: Find bold company names and walk up ---
    try {
      const bolds = document.querySelectorAll('b, strong, h3, h4');
      const items = [];
      for (const b of bolds) {
        const bt = b.textContent.trim();
        if (bt.length < 3 || bt.length > 100) continue;
        
        let parent = b.parentElement;
        for (let d = 0; d < 6; d++) {
          if (!parent) break;
          const h = parent.offsetHeight;
          const w = parent.offsetWidth;
          // Typical message item dimensions
          if (h >= 50 && h <= 180 && w >= 200 && w <= 500) {
             if (!items.includes(parent) && parent.textContent.includes(bt)) {
               items.push(parent);
             }
             break;
          }
          parent = parent.parentElement;
        }
      }
      if (items.length >= 2) {
        sendLog(`Found ${items.length} messages via "Bold-Text" strategy`, 'success');
        return items;
      }
    } catch (e) {}

    // --- Strategy D: Identify the main scrollable list on the left ---
    try {
      const scrollables = Array.from(document.querySelectorAll('div, ul, section'))
        .filter(el => {
           const style = window.getComputedStyle(el);
           return (style.overflowY === 'auto' || style.overflowY === 'scroll') && el.offsetWidth < 500;
        });
      
      for (const s of scrollables) {
        const kids = Array.from(s.children);
        const candidates = kids.filter(c => c.offsetHeight >= 50 && c.offsetHeight <= 180 && c.offsetWidth >= 100);
        if (candidates.length >= 2) {
          sendLog(`Found ${candidates.length} messages via "Scrollable-List" strategy`, 'success');
          return candidates;
        }
      }
    } catch (e) {}

    // --- Strategy C (Original, now E): Find items with profile images + location text ---
    try {
      const allImgs = document.querySelectorAll('img');
      const messageItems = [];

      for (const img of allImgs) {
        // Profile images are typically 40-80px
        if (img.offsetWidth < 25 || img.offsetWidth > 100) continue;
        if (img.offsetHeight < 25 || img.offsetHeight > 100) continue;

        // Walk up to find the message item container
        let parent = img.parentElement;
        for (let d = 0; d < 6; d++) {
          if (!parent) break;
          if (parent.offsetHeight >= 50 && parent.offsetHeight <= 160 &&
              parent.offsetWidth >= 200 && parent.offsetWidth <= 500 &&
              parent.textContent.trim().length >= 15) {
            if (!messageItems.includes(parent)) {
              messageItems.push(parent);
            }
            break;
          }
          parent = parent.parentElement;
        }
      }

      // Group by common parent to find the actual list
      if (messageItems.length >= 2) {
        const parentMap = new Map();
        messageItems.forEach(item => {
          const p = item.parentElement;
          if (!parentMap.has(p)) parentMap.set(p, []);
          parentMap.get(p).push(item);
        });

        let bestParent = null, bestCount = 0;
        parentMap.forEach((children, parent) => {
          if (children.length > bestCount) {
            bestCount = children.length;
            bestParent = parent;
          }
        });

        if (bestParent && bestCount >= 2) {
          // Return ALL visible children of this parent, not just ones we found
          const allChildren = Array.from(bestParent.children).filter(c =>
            c.offsetHeight >= 40 && c.offsetWidth >= 150 && c.textContent.trim().length > 10
          );
          if (allChildren.length >= 2) {
            sendLog(`Found ${allChildren.length} messages via profile-image + parent strategy`, 'success');
            return allChildren;
          }
          sendLog(`Found ${bestCount} messages via profile-image strategy`, 'success');
          return parentMap.get(bestParent);
        }
      }
    } catch (e) {
      sendLog('Profile-image strategy error: ' + e.message, 'warn');
    }

    // --- Strategy D: Find scrollable containers with uniform children ---
    try {
      const allContainers = document.querySelectorAll('div, ul, section');
      let best = null, bestCount = 0;

      for (const container of allContainers) {
        if (container.offsetWidth < 200 || container.offsetWidth > 500) continue;
        if (container.offsetHeight < 200) continue;
        if (container === document.body) continue;

        const children = Array.from(container.children).filter(c =>
          c.offsetHeight >= 50 && c.offsetHeight <= 160 && c.offsetWidth >= 180 &&
          c.textContent.trim().length > 15
        );

        if (children.length >= 3 && children.length > bestCount) {
          const heights = children.map(c => c.offsetHeight);
          const avg = heights.reduce((a, b) => a + b) / heights.length;
          const variance = heights.reduce((a, h) => a + Math.abs(h - avg), 0) / heights.length;
          if (variance < avg * 0.4) {
            best = container;
            bestCount = children.length;
          }
        }
      }

      if (best) {
        items = Array.from(best.children).filter(c =>
          c.offsetHeight >= 50 && c.offsetHeight <= 160 && c.offsetWidth >= 180 &&
          c.textContent.trim().length > 15
        );
        sendLog(`Found ${items.length} messages via container-scan (${best.className || best.id})`, 'success');
        return items;
      }
    } catch (e) {
      sendLog('Container-scan error: ' + e.message, 'warn');
    }

    // --- Strategy E: Last resort - dump DOM debug info ---
    sendLog('All strategies failed! Dumping DOM info to console...', 'error');
    const allDivClasses = new Set();
    document.querySelectorAll('div[class]').forEach(d =>
      d.classList.forEach(c => allDivClasses.add(c))
    );
    console.log('[IM-Extractor] ALL CSS classes:', Array.from(allDivClasses).sort());
    console.log('[IM-Extractor] Body children:', document.body.children.length);
    console.log('[IM-Extractor] Total DIVs:', document.querySelectorAll('div').length);
    console.log('[IM-Extractor] Total IMGs:', document.querySelectorAll('img').length);

    // Try absolutely any clickable-looking item
    const anyItems = Array.from(document.querySelectorAll('div, li, a')).filter(el =>
      el.offsetHeight >= 50 && el.offsetHeight <= 160 &&
      el.offsetWidth >= 200 && el.offsetWidth <= 500 &&
      el.querySelector('img') &&
      el.textContent.trim().length > 20 &&
      el.textContent.trim().length < 400
    );
    const deduped = anyItems.filter(c => !anyItems.some(o => o !== c && o.contains(c)));
    if (deduped.length >= 2) {
      sendLog(`Found ${deduped.length} messages via nuclear fallback`, 'success');
      return deduped;
    }

    return [];
  }

  // Helper: Get a stable signature for a message item (ignores time, read status, etc.)
  function getMessageSignature(msgEl) {
    try {
      // 1. Try Immutable Data Attributes first
      const dataId = msgEl.getAttribute('data-enqid') || 
                     msgEl.getAttribute('data-msgid') || 
                     msgEl.getAttribute('data-uid') ||
                     msgEl.getAttribute('data-contact');
      if (dataId) return `id_${dataId}`.toLowerCase();

      // 2. Extract Company Name (Most stable text element)
      const nameEl = msgEl.querySelector('b, strong, [class*="name"], [class*="company"]');
      const name = nameEl ? nameEl.textContent.trim().toLowerCase().replace(/\s+/g, ' ') : '';
      
      // 3. Extract Location
      const locationEl = msgEl.querySelector('[class*="location"], [class*="city"]');
      const loc = locationEl ? locationEl.textContent.trim().toLowerCase().replace(/\s+/g, ' ') : '';
      
      // 4. Extract first few words of preview, but strip any trailing time-like strings
      const previewEl = msgEl.querySelector('[class*="preview"], [class*="msg-text"], [class*="desc"]');
      let preview = previewEl ? previewEl.textContent.trim().toLowerCase() : '';
      
      // STRIP DYNAMIC CONTENT: Remove relative times like "2 min ago", "12:30 PM", etc.
      preview = preview.replace(/\d+[:.]\d+\s*(am|pm)?/gi, '') // 12:30 pm
                      .replace(/\d+\s*(min|hr|h|d|day|week|month|year)s?\s*ago/gi, '') // 2 mins ago
                      .replace(/just\s*now/gi, '')
                      .replace(/(yesterday|today|monday|tuesday|wednesday|thursday|friday|saturday|sunday)/gi, '')
                      .replace(/\d+\s*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/gi, '')
                      .replace(/\s+/g, ' ').trim();

      if (!name && !loc && !preview) {
        // Absolute fallback, but clean it heavily
        return msgEl.innerText.toLowerCase().replace(/\d/g, '').replace(/\s+/g, ' ').substring(0, 50).trim();
      }

      // Combine Name + Location + a slice of preview for a robust signature
      return `${name}|${loc}|${preview.substring(0, 30)}`.trim();
    } catch (e) {
      return msgEl.innerText.trim().substring(0, 60).toLowerCase();
    }
  }

  // ============================================
  // STEP 2: Find "View Details «" button
  // ============================================
  function findViewDetailsButton() {
    // From screenshot: it's a green/teal outlined button with text "View Details «"
    // Located in the top-right area of the chat panel

    // Strategy 1: Text content match (most reliable)
    const allClickable = document.querySelectorAll('a, button, span, div, p, label');
    for (const el of allClickable) {
      if (el.offsetWidth === 0 || el.offsetHeight === 0) continue;
      const text = el.textContent.trim().toLowerCase();

      // Match various forms of "View Details"
      if (text.match(/^view\s*details?\s*[«»›>]?\s*$/i) ||
          text === 'view details' || text === 'view details «' ||
          text === 'view details »' || text === 'view details ›') {
        sendLog(`Found "View Details" button: <${el.tagName}> "${el.textContent.trim()}"`, 'success');
        return el;
      }
    }

    // Strategy 2: Broader match
    for (const el of allClickable) {
      if (el.offsetWidth === 0 || el.offsetHeight === 0) continue;
      const text = el.textContent.trim().toLowerCase();
      if (text.length < 30 && text.includes('view') && text.includes('detail')) {
        sendLog(`Found View Details (broad): <${el.tagName}> "${el.textContent.trim()}"`, 'success');
        return el;
      }
    }

    // Strategy 3: CSS class match
    const detailSelectors = [
      '[class*="view_detail"]', '[class*="viewDetail"]', '[class*="ViewDetail"]',
      '[class*="view-detail"]', '[class*="vw_dtl"]', '[class*="vwDtl"]',
      'a[href*="viewdetail"]', 'a[href*="view-detail"]', 'a[href*="viewcompany"]',
    ];
    for (const sel of detailSelectors) {
      try {
        const el = document.querySelector(sel);
        if (el && el.offsetWidth > 0) {
          sendLog(`Found View Details via selector: ${sel}`, 'success');
          return el;
        }
      } catch (e) {}
    }

    return null;
  }

  // ============================================
  // STEP 2.5: Extract mobile from chat header
  // ============================================
  function extractMobileFromHeader() {
    // Target: Chat header (top-right section near company name and icons)
    sendLog('Scanning chat header for mobile number...', 'info');
    
    const phoneRegex = /(\+?\d[\d\-\s]{8,15}\d)/g;
    let headerContainer = null;

    // Strategy 1: Look for tel: links anywhere in the top area (Very reliable)
    const telLinks = document.querySelectorAll('a[href^="tel:"]');
    for (const link of telLinks) {
       if (link.offsetWidth > 0) {
         const num = link.href.replace('tel:', '').trim();
         if (num.replace(/\D/g, '').length >= 10) {
           sendLog(`Found mobile via tel link: ${num}`, 'success');
           return num;
         }
       }
    }

    // Strategy 2: Look for "last seen" or phone icons to find header
    const iconSelectors = [
      '[class*="phone-icon"]', '[class*="contact-icon"]', '[class*="call-icon"]',
      'i[class*="phone"]', 'i[class*="call"]', '.mc_rt_top', '.chat-header'
    ];
    
    for (const sel of iconSelectors) {
      const el = document.querySelector(sel);
      if (el && el.offsetWidth > 0) {
        headerContainer = el.closest('div[class*="header"], div[class*="top"], div[class*="panel"]');
        if (headerContainer) break;
      }
    }

    if (!headerContainer) {
      // Strategy 3: Find "View Details" and look at its siblings/parent
      const vdBtn = findViewDetailsButton();
      if (vdBtn) {
        headerContainer = vdBtn.parentElement;
        // Walk up to find a larger header container
        for (let d = 0; d < 3; d++) {
          if (headerContainer && headerContainer.offsetWidth > 300) break;
          if (headerContainer) headerContainer = headerContainer.parentElement;
        }
      }
    }

    if (!headerContainer) {
      // Fallback Strategy: Search for "last seen"
      const allEls = document.querySelectorAll('div, span, p');
      for (const el of allEls) {
        if (el.offsetWidth > 0 && el.textContent.toLowerCase().includes('last seen')) {
           headerContainer = el.parentElement;
           for (let d = 0; d < 3; d++) {
             if (headerContainer && headerContainer.offsetWidth > 300) break;
             if (headerContainer) headerContainer = headerContainer.parentElement;
           }
           if (headerContainer) break;
        }
      }
    }

    // Final Fallback: The main right-side top panel
    if (!headerContainer) {
      headerContainer = document.querySelector('[class*="chat-panel"] [class*="top"], #chatPanel [class*="top"], .mc_rt_top');
    }

    if (headerContainer) {
      // Clean up text and find potential numbers
      const text = headerContainer.innerText;
      const matches = text.match(phoneRegex);
      if (matches) {
        for (const match of matches) {
          const digits = match.replace(/\D/g, '');
          // IndiaMART numbers are usually 10-13 digits. 
          // Avoid matching ratings like "3.9 (47)" or memberships like "3 yrs"
          if (digits.length >= 10 && digits.length <= 13) {
            // Check if it's near a label or has a specific format
            sendLog(`Found mobile in header text: ${match}`, 'success');
            return match.trim();
          }
        }
      }
    }

    sendLog('Mobile number not found in header area.', 'warn');
    return '';
  }

  // ============================================
  // STEP 3: Wait for and find the detail panel
  // ============================================
  async function waitForDetailPanel() {
    // From screenshots: the detail panel is a modal/side panel with sections:
    // "Contact Details", "Reviews and Ratings", "Fact Sheet"
    // It has an "x" close button at the top

    for (let attempt = 0; attempt < CONFIG.MAX_RETRIES; attempt++) {
      await sleep(1000);

      // Look for "Contact Details" or "Fact Sheet" text appearing
      const allElements = document.querySelectorAll('*');
      for (const el of allElements) {
        const text = el.textContent.trim();
        if (el.offsetWidth > 150 && el.offsetHeight > 300 &&
            (text.includes('Contact Details') || text.includes('Fact Sheet')) &&
            (text.includes('Business Type') || text.includes('Year of Establishment') ||
             text.includes('Company Owner') || text.includes('GST'))) {
          // This is likely the detail panel or one of its parents
          // Find the most specific container that has all sections
          let container = el;
          while (container.parentElement &&
                 container.parentElement !== document.body &&
                 container.parentElement.offsetWidth < window.innerWidth * 0.6) {
            container = container.parentElement;
          }
          sendLog('Found detail panel with Fact Sheet content', 'success');
          return container;
        }
      }

      // Also check for modals / dialogs
      const modalSelectors = [
        '[role="dialog"]', '.modal.show', '[class*="modal"][style*="display: block"]',
        '[class*="detail-panel"]', '[class*="detailPanel"]',
        '[class*="company-detail"]', '[class*="companyDetail"]',
      ];
      for (const sel of modalSelectors) {
        try {
          const modal = document.querySelector(sel);
          if (modal && modal.offsetWidth > 150 && modal.offsetHeight > 200 &&
              modal.textContent.includes('Contact Details')) {
            sendLog(`Found detail panel via modal selector: ${sel}`, 'success');
            return modal;
          }
        } catch (e) {}
      }

      sendLog(`Waiting for detail panel (attempt ${attempt + 1})...`, 'info');
    }

    // Fallback: return null
    sendLog('Detail panel not found after retries', 'warn');
    return null;
  }

  // ============================================
  // STEP 4: Extract data from the detail panel
  // ============================================
  function extractLeadData(container, listCompanyName = '', headerMobile = '') {
    if (!container && !listCompanyName) return {};
    
    const data = {
      companyName: listCompanyName || '',
      mobileNumber: headerMobile || '',
      contactPerson: '',
      address: '',
      rating: '',
      reviewCount: '',
      review1: '',
      review2: '',
      responseRate: '',
      qualityRate: '',
      deliveryRate: '',
      businessType: '',
      ownerName: '',
      totalEmployees: '',
      yearEstablished: '',
      memberSince: '',
      annualTurnover: '',
      gstNumber: '',
      panNumber: '',
      products: ''
    };

    // --- Helper: Find value for a specific label ---
    function findValueByLabel(labelText) {
      const allEls = container.querySelectorAll('div, span, td, p, b, strong, label');
      for (const el of allEls) {
        if (el.children.length > 8) continue; // Relaxed child count
        const text = el.textContent.trim();
        
        // Fuzzy match: label can be part of the text
        if (text.toLowerCase().includes(labelText.toLowerCase())) {
          // Case 1: Value is in the same element after a colon or just after the label
          if (text.includes(':')) {
            const parts = text.split(':');
            if (parts.length > 1 && parts[1].trim().length > 0) return parts[1].trim();
          }
          
          // Case 1b: If text starts with label and is longer, the rest might be the value
          if (text.toLowerCase().startsWith(labelText.toLowerCase()) && text.length > labelText.length + 2) {
            const val = text.substring(labelText.length).trim().replace(/^[:\s-]+/, '');
            if (val.length > 0 && val.length < 100) return val;
          }

          // Case 2: Value is the NEXT sibling element
          let next = el.nextElementSibling;
          if (next && next.textContent.trim().length > 0 && next.textContent.trim().length < 500) {
            return next.textContent.trim();
          }

          // Case 3: Parent-based check (Label and value are siblings)
          const parent = el.parentElement;
          if (parent) {
            const pText = parent.textContent.trim();
            if (pText.toLowerCase().includes(labelText.toLowerCase()) && pText.length > labelText.length + 3) {
              const val = pText.replace(new RegExp(labelText, 'i'), '').trim().replace(/^[:\s-]+/, '');
              if (val.length > 0 && val.length < 500) return val;
            }
          }
        }
      }
      return '';
    }

    // --- 1. Company Name (Top of modal) ---
    // If the modal has a better one than the list, we'll use it.
    const topHeading = container ? container.querySelector('h1, h2, h3, h4, [class*="title"], [class*="name"]') : null;
    if (topHeading) {
      // Clean the header: Remove common close button patterns ('X' at the end)
      let headerText = topHeading.textContent.trim();
      // If it ends with a lone 'X', it's likely the close button
      headerText = headerText.replace(/\s+X$/i, '').replace(/X$/, '').trim();
      
      if (headerText.length > 3 && !headerText.includes('Details') && !headerText.includes('Reviews')) {
        data.companyName = headerText;
      }
    }

    // Fallback if still empty
    if (!data.companyName && container) {
      const bolds = container.querySelectorAll('b, strong');
      for (const b of bolds) {
        const bt = b.textContent.trim();
        if (bt.length > 5 && bt.length < 60 && !bt.includes('Details') && !bt.includes('Reviews')) {
          data.companyName = bt;
          break;
        }
      }
    }

    // --- 2. Contact Details Section ---
    const contactSection = findSectionByHeading(container, 'Contact Details');
    if (contactSection) {
      const texts = Array.from(contactSection.querySelectorAll('div, p, span'))
        .map(i => i.textContent.trim())
        .filter(t => t.length > 2 && t.length < 300 && !t.includes('Contact Details'));
      
      if (texts.length > 0) {
        data.contactPerson = texts[0];
        data.address = texts.slice(1).join(', ').substring(0, 300);
      }

      // Try icon-based picking if above is messy
      const personIcon = contactSection.querySelector('[class*="user"], [class*="person"], [class*="contact"]');
      if (personIcon) {
        const p = personIcon.parentElement;
        if (p && p.textContent.trim().length < 60) data.contactPerson = p.textContent.trim();
      }
    }

    // --- 3. Reviews and Ratings ---
    const reviewSection = findSectionByHeading(container, 'Reviews and Ratings');
    if (reviewSection) {
      const fullText = reviewSection.textContent;
      const m = fullText.match(/([\d.]+)\s*\(\d+\)/);
      if (m) data.rating = m[1];
      const countMatch = fullText.match(/\((\d+)\)/);
      if (countMatch) data.reviewCount = countMatch[1];
    }

    // --- 4. User Satisfaction ---
    const satSection = findSectionByHeading(container, 'User Satisfaction');
    if (satSection) {
      const t = satSection.textContent;
      const r = t.match(/(\d+%)\s*Response/i);
      const q = t.match(/(\d+%)\s*Quality/i);
      const d = t.match(/(\d+%)\s*Delivery/i);
      if (r) data.responseRate = r[1];
      if (q) data.qualityRate = q[1];
      if (d) data.deliveryRate = d[1];
    }

    // --- 5. Customer Reviews ---
    const custReviewSection = findSectionByHeading(container, 'Customer Reviews');
    if (custReviewSection) {
      const reviews = Array.from(custReviewSection.querySelectorAll('div'))
        .filter(d => d.textContent.trim().length > 20 && d.children.length < 5)
        .map(d => d.textContent.trim().replace(/\s+/g, ' '));
      data.review1 = reviews[0] || '';
      data.review2 = reviews[1] || '';
    }

    // --- 6. Fact Sheet ---
    data.businessType = findValueByLabel('Business Type');
    data.ownerName = findValueByLabel('Company Owner');
    data.totalEmployees = findValueByLabel('Total Number of Employees');
    data.yearEstablished = findValueByLabel('Year of Establishment');
    data.memberSince = findValueByLabel('IndiaMART Member Since') || findValueByLabel('Member Since');
    data.annualTurnover = findValueByLabel('Annual Turnover');
    data.gstNumber = findValueByLabel('GST No.');
    data.panNumber = findValueByLabel('PAN');
    data.products = findValueByLabel('Sells');

    // Extra fallback for Products/Sells from "Sells" section
    if (!data.products) {
      const sellsSection = findSectionByHeading(container, 'Sells');
      if (sellsSection) data.products = sellsSection.textContent.replace('Sells', '').trim().substring(0, 500);
    }

    // Generate unique ID based on Mobile and Address (More robust than Company Name)
    data.uniqueId = generateUniqueId(data.companyName, data.address, data.mobileNumber);
    
    sendLog(`📊 Extracted: ${data.companyName || 'Unknown'} (ID: ${data.uniqueId})`, 'extract');
    return data;
  }

  // Helper: Find a section container by its heading text
  function findSectionByHeading(container, headingText) {
    const allEls = container.querySelectorAll('h1, h2, h3, h4, h5, h6, b, strong, div, span, p');
    for (const el of allEls) {
      const directText = Array.from(el.childNodes)
        .filter(n => n.nodeType === 3)
        .map(n => n.textContent.trim())
        .join(' ').trim();

      if (directText === headingText || el.textContent.trim() === headingText) {
        // Return the parent container that holds the entire section
        let sectionContainer = el.parentElement;
        // Walk up until we find a container with reasonable size
        for (let d = 0; d < 5; d++) {
          if (!sectionContainer) break;
          if (sectionContainer.offsetHeight > 100) {
            return sectionContainer;
          }
          sectionContainer = sectionContainer.parentElement;
        }
        return el.parentElement;
      }
    }
    return null;
  }

  // ============================================
  // STEP 5: Close the detail panel
  // ============================================
  async function closeDetailPanel() {
    // From screenshot: there's an "x" close button at top of the panel
    const closeSelectors = [
      // Various close button patterns
      '[class*="close"]', '[class*="Close"]', '[class*="cls"]',
      'button[aria-label="Close"]', '.modal .close',
    ];

    for (const sel of closeSelectors) {
      try {
        const btns = document.querySelectorAll(sel);
        for (const btn of btns) {
          if (btn.offsetWidth > 0 && btn.offsetHeight > 0 && btn.offsetHeight < 50) {
            const text = btn.textContent.trim();
            if (text === '×' || text === 'x' || text === 'X' || text === '✕' ||
                text === '' || text.toLowerCase() === 'close') {
              btn.click();
              await sleep(500);
              return;
            }
          }
        }
      } catch (e) {}
    }

    // Text-based close
    const allEls = document.querySelectorAll('button, a, span, div, i');
    for (const el of allEls) {
      const t = el.textContent.trim();
      if ((t === '×' || t === 'x' || t === 'X' || t === '✕') &&
          el.offsetWidth > 0 && el.offsetHeight > 0 && el.offsetHeight <= 40 && el.offsetWidth <= 40) {
        el.click();
        await sleep(500);
        return;
      }
    }

    // Escape key
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', keyCode: 27, bubbles: true }));
    await sleep(500);

    // Click backdrop
    const overlays = document.querySelectorAll('[class*="overlay"], [class*="backdrop"], [class*="mask"]');
    for (const ov of overlays) {
      if (ov.offsetWidth > window.innerWidth * 0.5 && ov.offsetHeight > window.innerHeight * 0.5) {
        ov.click();
        await sleep(300);
        return;
      }
    }
  }

  // ============================================
  // MAIN EXTRACTION FLOW
  // ============================================
  async function runExtraction() {
    if (isRunning) {
      sendLog('Already running!', 'warn');
      return;
    }
    isRunning = true;
    shouldStop = false;
    chrome.storage.local.set({ isExtracting: true });

    try {
      // Check URL
      const url = window.location.href.toLowerCase();
      const isLeadManager = url.includes('messagecentre') || url.includes('message-centre') || url.includes('leadmanager');

      if (!isLeadManager) {
        sendLog('Not on Lead Manager. Setting auto-start and redirecting...', 'warn');
        chrome.storage.local.set({ extractionAutoStart: true }, () => {
          window.location.href = CONFIG.LEAD_MANAGER_URL;
        });
        return;
      }

      sendLog('Waiting for page to fully load...', 'info');
      await sleep(CONFIG.INITIAL_LOAD_WAIT);
      await waitForDomStable(3000);

      const processedSignatures = new Set();
      let lastProcessedSignature = null;
      let totalExtracted = 0;
      let consecutiveEmptyScrolls = 0;
      const MAX_EMPTY_SCROLLS = 5; 
      const MAX_TOTAL_LEADS = 500; 
      
      let listParent = null;

      while (!shouldStop && totalExtracted < MAX_TOTAL_LEADS) {
        const messages = getMessageItems();
        if (messages.length > 0 && !listParent) listParent = messages[0].parentElement;

        // Filter for un-processed messages
        const newMessages = messages.filter(msg => {
          const sig = getMessageSignature(msg);
          // STOP CONDITION: If session signature exists, don't re-process
          if (processedSignatures.has(sig)) return false;
          // IMPORTANT: Double check against the very last processed message
          if (sig === lastProcessedSignature) return false;
          return true;
        });

        if (newMessages.length === 0) {
          // STRICT STOP CONDITION: 
          // 1. If we are at the bottom
          // 2. OR if the last message in the current list has already been processed
          const isAtBottom = listParent && (listParent.scrollTop + listParent.clientHeight >= listParent.scrollHeight - 20);
          
          let lastMessageProcessed = false;
          if (messages.length > 0) {
            const lastSig = getMessageSignature(messages[messages.length - 1]);
            if (processedSignatures.has(lastSig)) lastMessageProcessed = true;
          }

          if (isAtBottom && lastMessageProcessed) {
            sendLog('Reached end of list and last message already processed. Stopping.', 'success');
            break;
          }

          consecutiveEmptyScrolls++;
          if (consecutiveEmptyScrolls >= MAX_EMPTY_SCROLLS) {
            sendLog('No more new messages found after multiple scrolls. Finishing...', 'success');
            break;
          }
          
          if (listParent) {
            sendLog(`Scanning... (Attempt ${consecutiveEmptyScrolls})`, 'info');
            listParent.scrollTop += 600;
            await sleep(2000);
            await waitForDomStable(1500);
            continue;
          } else {
            sendLog('Could not find scrollable container. Stopping.', 'error');
            break;
          }
        }

        consecutiveEmptyScrolls = 0;
        sendLog(`Found ${newMessages.length} new messages in current view.`, 'info');

        for (const msgEl of newMessages) {
          if (shouldStop) break;

          const sig = getMessageSignature(msgEl);
          processedSignatures.add(sig);

          // Get Company Name from the list item
          let listCompanyName = '';
          const nameEl = msgEl.querySelector('b, strong, [class*="name"], [class*="company"]');
          if (nameEl) listCompanyName = nameEl.textContent.trim();

          totalExtracted++;
          sendLog(`📋 Processing lead #${totalExtracted}: ${listCompanyName || 'Unknown'}`, 'info');
          sendProgress(totalExtracted, 0, listCompanyName, 'processing');

          // Click message
          msgEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
          await sleep(600);
          msgEl.click();
          await sleep(CONFIG.CLICK_DELAY);
          await waitForDomStable(2000);

          // EXTRA STEP: Extract mobile from header
          const headerMobile = extractMobileFromHeader();

          // Find and click "View Details"
          let vdBtn = null;
          for (let r = 0; r < CONFIG.MAX_RETRIES; r++) {
            vdBtn = findViewDetailsButton();
            if (vdBtn) break;
            sendLog(`Waiting for View Details button... (retry ${r + 1})`, 'info');
            await sleep(1500);
          }

          if (!vdBtn) {
            sendLog(`⚠️ "View Details" not found for ${listCompanyName || 'this lead'}. Skipping.`, 'warn');
            sendProgress(totalExtracted, 0, listCompanyName, 'error');
            continue;
          }

          sendLog('Clicking "View Details"...', 'info');
          vdBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
          await sleep(300);
          vdBtn.click();
          await sleep(CONFIG.DETAIL_LOAD_DELAY);
          await waitForDomStable(3000);

          // Wait for detail panel
          const detailPanel = await waitForDetailPanel();

          // SCROLL DETAIL PANEL
          if (detailPanel) {
            for (let s = 0; s < 2; s++) {
              detailPanel.scrollTop = (s + 1) * 500;
              await sleep(500);
            }
            detailPanel.scrollTop = 0;
            await sleep(500);
          }

          // Extract data
          const leadData = extractLeadData(detailPanel, listCompanyName, headerMobile);

          // Save
          const saveResp = await new Promise(resolve => {
            chrome.runtime.sendMessage({ action: 'saveLead', payload: leadData }, resolve);
          });

          if (saveResp && saveResp.success) {
            lastProcessedSignature = sig; // Update last processed tracker
            sendProgress(totalExtracted, 0, leadData.companyName, 'extracted');
            chrome.runtime.sendMessage({ type: 'leadCount', count: saveResp.newCount });
          } else if (saveResp && saveResp.reason === 'duplicate') {
            sendProgress(totalExtracted, 0, leadData.companyName, 'skipped');
          }

          // Close panel
          await closeDetailPanel();
          await sleep(CONFIG.BETWEEN_MESSAGES);
        }

        // After processing a batch, scroll the list a bit to reveal more
        if (listParent && !shouldStop) {
          const prevScrollTop = listParent.scrollTop;
          sendLog('Scrolling to next batch...', 'info');
          listParent.scrollTop += 800;
          await sleep(2000);
          await waitForDomStable(1500);
          
          // Check if we actually scrolled anything
          if (listParent.scrollTop === prevScrollTop) {
             consecutiveEmptyScrolls++;
             sendLog(`At bottom of list (Attempt ${consecutiveEmptyScrolls})`, 'info');
          } else {
             consecutiveEmptyScrolls = 0;
          }
        }
      }

      sendLog(`🎉 Extraction Finished! Total Leads Processed: ${totalExtracted}`, 'success');
      chrome.runtime.sendMessage({ type: 'done', total: totalExtracted });

    } catch (error) {
      sendLog(`❌ Fatal: ${error.message}`, 'error');
      chrome.runtime.sendMessage({ type: 'extractionError', text: error.message });
      console.error('[IM-Extractor]', error);
    } finally {
      isRunning = false;
      shouldStop = false;
      chrome.storage.local.set({ isExtracting: false });
    }
  }

  // ---- Message listener ----
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'startExtraction') {
      sendResponse({ ack: true });
      runExtraction();
      return true;
    }
    if (message.action === 'stopExtraction') {
      shouldStop = true;
      sendResponse({ ack: true });
      return true;
    }
    return false;
  });

  // --- Auto-start logic on load ---
  chrome.storage.local.get(['extractionAutoStart'], (result) => {
    if (result.extractionAutoStart) {
      console.log('[IM-Extractor] Auto-starting extraction after redirect...');
      // Clear flag so it doesn't loop
      chrome.storage.local.remove('extractionAutoStart');
      runExtraction();
    }
  });

  console.log('[IM-Extractor] Content script loaded:', window.location.href);
})();
