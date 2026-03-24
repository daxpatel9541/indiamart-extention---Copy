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

  function normalizeText(text) {
    if (!text) return '';
    return text
      .replace(/\s+/g, ' ')
      .replace(/[×xX✕]\s*$/g, '') // Remove close button 'X' from end of strings
      .replace(/view details\s*[«»›>]?\s*$/gi, '') // Remove "View Details" noise
      .trim();
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
  async function extractMobileFromHeader(expectedName = '') {
    // Target: Chat header (top-right section near company name and icons)
    sendLog('Scanning chat header for mobile number...', 'info');
    
    const phoneRegex = /(\+?\d[\d\-\s]{8,15}\d)/g;
    let headerContainer = null;

    // High priority: Specific IndiaMART header classes
    const knownHeaderSelectors = [
      '.mc_rt_top', '.chat-header', '[class*="chat-panel"] [class*="top"]',
      '.topHdr', '[class*="Header"]', '[class*="top-section"]', '.mc_rt_hd'
    ];
    
    for (let attempt = 0; attempt < 12; attempt++) {
      for (const sel of knownHeaderSelectors) {
        try {
          const el = document.querySelector(sel);
          if (el && el.offsetWidth > 0) {
            const text = el.innerText.toLowerCase();
            if (expectedName) {
               if (text.includes(expectedName.toLowerCase())) {
                 headerContainer = el;
                 break;
               }
            } else {
              headerContainer = el;
              break;
            }
          }
        } catch(e) {}
      }
      if (headerContainer) break;
      
      // Fallback within attempt: search for any div containing the name that's at the top
      if (expectedName && attempt > 3) {
        const potentialHeaders = Array.from(document.querySelectorAll('div, header, section')).filter(el => {
          return el.offsetHeight > 40 && el.offsetHeight < 150 && 
                 el.getBoundingClientRect().top < 200 &&
                 el.innerText.toLowerCase().includes(expectedName.toLowerCase());
        });
        if (potentialHeaders.length > 0) {
          headerContainer = potentialHeaders[0];
          break;
        }
      }
      
      await sleep(400); // Wait for header to update
    }

    if (!headerContainer) {
       sendLog(`Header mismatch or not found for "${expectedName}". Scanning all top elements...`, 'warn');
    }

    if (headerContainer) {
      const text = headerContainer.innerText;
      const matches = text.match(phoneRegex);
      if (matches) {
        for (const match of matches) {
          const digits = match.replace(/\D/g, '');
          if (digits.length >= 10 && digits.length <= 13) {
            sendLog(`Found mobile in header text: ${match}`, 'success');
            return match.trim();
          }
        }
      }
    }

    // Strategy 2: Look for tel: links anywhere in the top area (Very reliable)
    const telLinks = document.querySelectorAll('a[href^="tel:"]');
    for (const link of telLinks) {
       const rect = link.getBoundingClientRect();
       if (rect.top < 300 && rect.width > 0) {
         const num = link.href.replace('tel:', '').trim();
         if (num.replace(/\D/g, '').length >= 10) {
           sendLog(`Found mobile via tel link in top area: ${num}`, 'success');
           return num;
         }
       }
    }

    // Strategy 3: Aggressive Scan of Top-Right Quadrant
    const topElements = Array.from(document.querySelectorAll('span, div, b, strong, a, p')).filter(el => {
      const rect = el.getBoundingClientRect();
      return rect.top < 200 && rect.left > window.innerWidth * 0.4 && rect.width > 0;
    });

    for (const el of topElements) {
      if (el.children.length > 0) continue; // Only check leaf nodes
      const text = el.textContent.trim();
      if (phoneRegex.test(text)) {
        const match = text.match(phoneRegex)[0];
        const digits = match.replace(/\D/g, '');
        if (digits.length >= 10 && digits.length <= 13) {
          sendLog(`Found mobile via top-right scan: ${match}`, 'success');
          return match;
        }
      }
    }

    sendLog('Mobile number not found in header area.', 'warn');
    return '';
  }

  // ============================================
  // STEP 3: Wait for and find the detail panel
  // ============================================
  async function waitForDetailPanel(expectedName = '') {
    // From screenshots: the detail panel is a modal/side panel with sections:
    // "Contact Details", "Reviews and Ratings", "Fact Sheet"
    
    for (let attempt = 0; attempt < CONFIG.MAX_RETRIES + 4; attempt++) {
      await sleep(1000);

      // Look for any modal/dialog that contains "Contact Details"
      const candidates = Array.from(document.querySelectorAll('div, [role="dialog"]')).filter(el => {
        if (el.offsetWidth < 200 || el.offsetHeight < 200) return false;
        const text = el.textContent;
        // Basic signature of the IndiaMART fact sheet modal
        return text.includes('Contact Details') && (text.includes('Fact Sheet') || text.includes('Business Type') || text.includes('Establishment'));
      });

      for (const modal of candidates) {
        // Find the specific container (walk up to find the root of the modal if needed)
        let container = modal;
        while (container.parentElement && 
               container.parentElement !== document.body && 
               container.parentElement.offsetWidth < window.innerWidth * 0.8) {
          container = container.parentElement;
        }

        // VERIFICATION: Does the title match our expected company?
        if (expectedName) {
           // Find all potential title elements
           const headings = Array.from(container.querySelectorAll('h1, h2, h3, h4, [class*="title"], [class*="name"], b, strong'));
           let foundMatch = false;
           
           for (const el of headings) {
             const titleText = normalizeText(el.textContent);
             // Skip known section headers
             if (!titleText || /contact detail|review|fact sheet|satisfaction|business type/i.test(titleText)) continue;
             
             if (titleText.toLowerCase().includes(expectedName.toLowerCase()) || 
                 expectedName.toLowerCase().includes(titleText.toLowerCase())) {
               sendLog(`Confirmed detail panel for: ${titleText}`, 'success');
               foundMatch = true;
               break;
             }
           }
           
           if (foundMatch) return container;
           
           // If we found headers but none matched the name, log the first non-empty one for debug
           const firstHeader = headings.map(h => normalizeText(h.textContent)).filter(t => t.length > 2)[0];
           sendLog(`Modal headers found, but no match for "${expectedName}". (Top header: "${firstHeader}"). Still waiting...`, 'info');
        } else {
          return container;
        }
      }

      sendLog(`Waiting for detail panel update (attempt ${attempt + 1})...`, 'info');
    }

    sendLog(`Detail panel not found or title mismatch for "${expectedName}" after retries`, 'warn');
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
      if (!container) return '';
      // Prioritize small elements that specifically contain the label
      const allEls = Array.from(container.querySelectorAll('div, span, td, p, b, strong, label, th, dt, dd'));
      
      for (const el of allEls) {
        // Skip hidden elements or huge containers
        if (el.offsetWidth === 0 || el.children.length > 3) continue;
        
        const rawText = el.textContent.trim();
        if (!rawText) continue;

        // Clean label for matching
        const cleanLabel = labelText.toLowerCase().replace(/[:\s]+$/, '');
        const cleanText = rawText.toLowerCase();

        // Check if this element matches the label
        // We want it to be either EXACTLY the label, or START with the label (+ colon/space)
        const isExactMatch = (cleanText === cleanLabel);
        
        // Starts with match: "Label: " or "Label "
        const startsWithMatch = cleanText.startsWith(cleanLabel) && (
           cleanText === cleanLabel || 
           /^[:\s\-]/.test(rawText.substring(labelText.length))
        );

        if (isExactMatch || startsWithMatch) {
          // Case A: Value is in the same element (e.g. "Label: Value")
          if (rawText.includes(':')) {
            const parts = rawText.split(':');
            if (parts.length > 1) {
              const val = normalizeText(parts.slice(1).join(':'));
              if (val) {
                console.log(`[IM-Extractor] Found ${labelText} (Same-El): "${val}"`);
                return val;
              }
            }
          }
          
          if (startsWithMatch) {
             const val = normalizeText(rawText.substring(labelText.length).replace(/^[:\s\-]+/, ''));
             if (val) {
                console.log(`[IM-Extractor] Found ${labelText} (Starts-With): "${val}"`);
                return val;
             }
          }

          // Case B: Value is in the next sibling
          let next = el.nextElementSibling;
          if (next) {
            const val = normalizeText(next.textContent);
            if (val && val.length < 300) {
              console.log(`[IM-Extractor] Found ${labelText} (Sibling): "${val}"`);
              return val;
            }
          }

          // Case C: Value is in parent's next sibling
          const parent = el.parentElement;
          if (parent && parent.children.length <= 2) {
             const pNext = parent.nextElementSibling;
             if (pNext) {
                const val = normalizeText(pNext.textContent);
                if (val && val.length < 300) {
                  console.log(`[IM-Extractor] Found ${labelText} (Parent-Sibling): "${val}"`);
                  return val;
                }
             }
          }
        }
      }
      return '';
    }

    // --- 1. Company Name (Top of modal) ---
    const topHeading = container ? container.querySelector('h1, h2, h3, h4, [class*="title"], [class*="name"]') : null;
    if (topHeading) {
      const headerText = normalizeText(topHeading.textContent);
      if (headerText.length > 3 && !headerText.includes('Details') && !headerText.includes('Reviews')) {
        data.companyName = headerText;
      }
    }

    // --- 2. Contact Details Section ---
    const contactSection = findSectionByHeading(container, 'Contact Details');
    if (contactSection) {
      // Use clean innerText instead of querySelectorAll tags to avoid recursive duplication
      const fullText = contactSection.innerText;
      const lines = fullText.split('\n')
        .map(l => normalizeText(l))
        .filter(l => l.length > 2 && !l.toLowerCase().includes('contact details'));
      
      // Usually: Line 1 = Name, Lines 2+ = Address
      if (lines.length > 0) {
        data.contactPerson = lines[0];
        data.address = lines.slice(1).join(', ').substring(0, 400);
      }

      // Refinement: if there's a bold element, it's likely the person
      const boldEl = contactSection.querySelector('b, strong');
      if (boldEl) {
        const boldText = normalizeText(boldEl.textContent);
        if (boldText && boldText.length < 60) {
           data.contactPerson = boldText;
           // Address is lines that aren't the person
           data.address = lines.filter(l => l !== boldText).join(', ').substring(0, 400);
        }
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
      const reviews = Array.from(custReviewSection.querySelectorAll('div, p'))
        .filter(d => d.textContent.trim().length > 20 && d.children.length < 3)
        .map(d => normalizeText(d.textContent));
      data.review1 = reviews[0] || '';
      data.review2 = reviews[1] || '';
    }

    // --- 6. Fact Sheet (Robust Extraction) ---
    const factSheet = findSectionByHeading(container, 'Fact Sheet') || container;
    
    const factLabels = {
      'Business Type': 'businessType',
      'Company Owner': 'ownerName',
      'Total Number of Employees': 'totalEmployees',
      'Year of Establishment': 'yearEstablished',
      'IndiaMART Member Since': 'memberSince',
      'Member Since': 'memberSince',
      'Annual Turnover': 'annualTurnover',
      'GST No.': 'gstNumber',
      'GST': 'gstNumber',
      'PAN': 'panNumber',
      'Sells': 'products'
    };

    for (const [label, key] of Object.entries(factLabels)) {
       if (!data[key]) {
         data[key] = findValueByLabel(label);
       }
    }

    if (!data.products) {
      const sellsSection = findSectionByHeading(container, 'Sells');
      if (sellsSection) data.products = normalizeText(sellsSection.textContent.replace('Sells', ''));
    }

    data.uniqueId = generateUniqueId(data.companyName, data.address, data.mobileNumber);
    
    console.log('[IM-Extractor] FINAL LEAD DATA:', data);
    sendLog(`📊 Extracted: ${data.companyName}`, 'extract');
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
          const headerMobile = await extractMobileFromHeader(listCompanyName);

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

          // Wait for detail panel (ensure it matches the current lead!)
          const detailPanel = await waitForDetailPanel(listCompanyName);

          // SCROLL DETAIL PANEL
          if (detailPanel) {
            sendLog('Scrolling detail panel to trigger lazy loading...', 'info');
            // Try to find the actual scrollable element inside the modal
            const scrollable = [detailPanel, ...detailPanel.querySelectorAll('*')].find(el => {
              const style = window.getComputedStyle(el);
              return (style.overflowY === 'auto' || style.overflowY === 'scroll') && el.scrollHeight > el.clientHeight;
            }) || detailPanel;

            for (let s = 1; s <= 3; s++) {
              scrollable.scrollTop = s * 600;
              await sleep(800);
            }
            scrollable.scrollTop = 0;
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
