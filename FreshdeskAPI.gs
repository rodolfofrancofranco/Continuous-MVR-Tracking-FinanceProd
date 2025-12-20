/**
 * FRESHDESK API WRAPPER
 * MVR TICKET TRACKER - API Integration
 * 
 * Purpose: Fetch tickets matching MVR email pattern from Freshdesk
 * Pattern: "One Off Continuous MVR - (Suspension Check|Recheck) - {partner} - {turnId}"
 */

// ═══════════════════════════════════════════════════════════════════════════════
// API RETRY LOGIC
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch with retries and exponential backoff
 * @param {string} url - URL to fetch
 * @param {Object} options - Fetch options
 * @param {number} maxRetries - Maximum retry attempts
 * @param {number} baseDelayMs - Base delay for exponential backoff
 * @return {Object} HTTP response
 */
function fetchWithRetries(url, options, maxRetries = 4, baseDelayMs = 1000) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      Logger.log(`🌐 HTTP ${code} for ${url} (attempt ${attempt + 1})`);

      // Success
      if (code === 200) return resp;

      // Rate limited or server error -> retry with backoff
      if (code === 429 || code >= 500) {
        const waitMs = baseDelayMs * Math.pow(2, attempt);
        Logger.log(`⏳ Received ${code}; retrying after ${waitMs}ms`);
        Utilities.sleep(waitMs);
        continue;
      }

      // Client error (4xx other than 429) — treat as fatal
      throw new Error(`HTTP ${code} for ${url}: ${resp.getContentText().slice(0, 300)}`);
    } catch (err) {
      // If this was the last attempt, rethrow
      if (attempt === maxRetries - 1) {
        throw err;
      }
      const waitMs = baseDelayMs * Math.pow(2, attempt);
      Logger.log(`⚠️ Fetch attempt ${attempt + 1} failed for ${url}: ${err} — retrying in ${waitMs}ms`);
      Utilities.sleep(waitMs);
    }
  }
  throw new Error("Exhausted retries for " + url);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MVR TICKET FETCHING (Using updated_since for server-side filtering)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch MVR tickets from Freshdesk API using server-side date filtering
 * Uses updated_since parameter for efficient server-side filtering
 * Then filters by subject pattern client-side
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {number} hours - Number of hours to look back (default: 720 = 30 days)
 * @return {Array} Array of MVR tickets within time window
 */
function fetchMVRTickets(apiKey, domain, hours = TIME_CONFIG.LOOKBACK_HOURS) {
  const options = {
    method: "get",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(apiKey + ":X")
    },
    muteHttpExceptions: true
  };
  
  // Calculate cutoff date - use updated_since for server-side filtering
  const updatedSince = new Date(Date.now() - hours * 60 * 60 * 1000);
  const updatedSinceISO = updatedSince.toISOString();
  
  Logger.log(`\n🔍 === FETCHING MVR TICKETS ===`);
  Logger.log(`📅 Updated since: ${updatedSinceISO}`);
  Logger.log(`📅 Local: ${updatedSince.toLocaleString('en-US', { timeZone: TIME_CONFIG.TIMEZONE })} CST`);
  Logger.log(`🔍 Using server-side updated_since filter (efficient)`);
  
  let allMVRTickets = [];
  let page = 1;
  let totalTicketsScanned = 0;
  const MAX_PAGES_SAFETY = 500; // Safety limit to prevent infinite loops
  
  while (page <= MAX_PAGES_SAFETY) {
    Logger.log(`🔎 Fetching page ${page}...`);
    
    // Use updated_since for SERVER-SIDE filtering - Freshdesk only returns matching tickets
    // include=description is REQUIRED to get email body (State Vendor, DL State, Tier, etc.)
    const url = `https://${domain}.freshdesk.com/api/v2/tickets`
      + `?updated_since=${encodeURIComponent(updatedSinceISO)}`
      + `&include=description`
      + `&page=${page}`
      + `&per_page=${API_CONFIG.PER_PAGE}`;
    
    let response;
    try {
      response = fetchWithRetries(url, options, API_CONFIG.RETRY_ATTEMPTS, API_CONFIG.RETRY_DELAY_BASE);
    } catch (e) {
      Logger.log(`⚠️ API Error on page ${page}: ${e.message}`);
      if (allMVRTickets.length > 0) {
        Logger.log(`⚠️ Returning ${allMVRTickets.length} tickets collected before error`);
        break;
      }
      throw e;
    }
    
    let tickets;
    try {
      tickets = JSON.parse(response.getContentText());
    } catch (e) {
      Logger.log(`⚠️ JSON parse error on page ${page}: ${e.message}`);
      break;
    }
    
    // Natural end - no more tickets matching updated_since
    if (!Array.isArray(tickets) || tickets.length === 0) {
      Logger.log(`✅ Reached end of tickets at page ${page}`);
      break;
    }
    
    totalTicketsScanned += tickets.length;
    Logger.log(`📥 Retrieved ${tickets.length} tickets (${totalTicketsScanned} total scanned)`);
    
    // Filter by MVR subject patterns (client-side)
    let mvrOnPage = 0;
    for (const ticket of tickets) {
      if (isMVRSubject(ticket.subject)) {
        allMVRTickets.push(ticket);
        mvrOnPage++;
      }
    }
    
    Logger.log(`📄 Page ${page}: ${mvrOnPage} MVR tickets (${allMVRTickets.length} total MVR)`);
    
    page++;
    
    // Rate limiting between requests
    Utilities.sleep(API_CONFIG.RATE_LIMIT_DELAY);
    
    // Progress logging every 10 pages
    if (page % 10 === 0) {
      Logger.log(`📊 Progress: Page ${page}, ${allMVRTickets.length} MVR tickets, ${totalTicketsScanned} total scanned`);
    }
  }
  
  // Remove duplicates
  const uniqueTickets = Array.from(new Map(allMVRTickets.map(t => [t.id, t])).values());
  
  // Summary statistics
  if (uniqueTickets.length > 0) {
    const statusCounts = {
      open: uniqueTickets.filter(t => t.status === TICKET_STATUS.OPEN).length,
      pending: uniqueTickets.filter(t => t.status === TICKET_STATUS.PENDING).length,
      resolved: uniqueTickets.filter(t => t.status === TICKET_STATUS.RESOLVED).length,
      closed: uniqueTickets.filter(t => t.status === TICKET_STATUS.CLOSED).length
    };
    
    Logger.log(`\n📊 === FETCH SUMMARY ===`);
    Logger.log(`   Total MVR tickets: ${uniqueTickets.length}`);
    Logger.log(`   Pages fetched: ${page - 1}`);
    Logger.log(`   Total tickets scanned: ${totalTicketsScanned}`);
    Logger.log(`   Status: Open=${statusCounts.open}, Pending=${statusCounts.pending}, Resolved=${statusCounts.resolved}, Closed=${statusCounts.closed}`);
  } else {
    Logger.log(`⚠️ No MVR tickets found in the specified time window`);
  }
  
  return uniqueTickets;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TICKET DETAILS FETCHING (For Full Information)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch complete ticket details for a single ticket
 * Use this when you need full ticket information including custom fields
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {number} ticketId - Ticket ID to fetch
 * @param {boolean} includeCustomFields - Whether to include custom fields (default: true)
 * @return {Object|null} Complete ticket object or null if not found
 */
function fetchTicketDetails(apiKey, domain, ticketId, includeStats = true) {
  const options = {
    method: "get",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(apiKey + ":X")
    },
    muteHttpExceptions: true
  };
  
  // Description IS included by default for single-ticket endpoint (View a Ticket API)
  // Valid include values: conversations, requester, company, stats, sla_policy
  const include = includeStats ? "?include=stats" : "";
  const url = `https://${domain}.freshdesk.com/api/v2/tickets/${ticketId}${include}`;
  
  try {
    Logger.log(`🔍 Fetching details for ticket ${ticketId}...`);
    const response = fetchWithRetries(url, options, API_CONFIG.RETRY_ATTEMPTS, API_CONFIG.RETRY_DELAY_BASE);
    const ticket = JSON.parse(response.getContentText());
    Logger.log(`✅ Retrieved details for ticket ${ticketId}`);
    return ticket;
  } catch (e) {
    Logger.log(`⚠️ Failed to fetch ticket ${ticketId}: ${e.message}`);
    return null;
  }
}

/**
 * Fetch conversations for a single ticket
 * Conversations include agent replies, notes, and customer responses
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {number} ticketId - Ticket ID to fetch conversations for
 * @return {Array} Array of conversation objects or empty array on error
 */
function fetchTicketConversations(apiKey, domain, ticketId) {
  const options = {
    method: "get",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(apiKey + ":X")
    },
    muteHttpExceptions: true
  };
  
  const url = `https://${domain}.freshdesk.com/api/v2/tickets/${ticketId}/conversations`;
  
  try {
    const response = fetchWithRetries(url, options, API_CONFIG.RETRY_ATTEMPTS, API_CONFIG.RETRY_DELAY_BASE);
    const conversations = JSON.parse(response.getContentText());
    Logger.log(`💬 Retrieved ${conversations.length} conversations for ticket ${ticketId}`);
    return conversations;
  } catch (e) {
    Logger.log(`⚠️ Failed to fetch conversations for ticket ${ticketId}: ${e.message}`);
    return [];
  }
}

/**
 * Batch fetch complete ticket details for multiple tickets
 * Includes rate limiting to respect API limits
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {Array} ticketIds - Array of ticket IDs to fetch
 * @return {Array} Array of complete ticket objects
 */
function batchFetchTicketDetails(apiKey, domain, ticketIds) {
  Logger.log(`🔍 Batch fetching ${ticketIds.length} ticket details...`);
  
  const tickets = [];
  let successCount = 0;
  let failCount = 0;
  
  for (let i = 0; i < ticketIds.length; i++) {
    const ticketId = ticketIds[i];
    
    try {
      const ticket = fetchTicketDetails(apiKey, domain, ticketId);
      if (ticket) {
        tickets.push(ticket);
        successCount++;
      } else {
        failCount++;
      }
    } catch (e) {
      Logger.log(`⚠️ Error fetching ticket ${ticketId}: ${e.message}`);
      failCount++;
    }
    
    // Rate limiting between requests
    if (i < ticketIds.length - 1) {
      Utilities.sleep(API_CONFIG.RATE_LIMIT_DELAY);
    }
    
    // Progress logging every 10 tickets
    if ((i + 1) % 10 === 0) {
      Logger.log(`📊 Progress: ${i + 1}/${ticketIds.length} tickets (${successCount} success, ${failCount} failed)`);
    }
  }
  
  Logger.log(`✅ Batch fetch complete: ${successCount} success, ${failCount} failed`);
  return tickets;
}

// ═══════════════════════════════════════════════════════════════════════════════
// API CONNECTION TESTING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test API connection
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @return {boolean} True if connection successful
 */
function testApiConnection(apiKey, domain) {
  const headers = {
    "Authorization": "Basic " + Utilities.base64Encode(apiKey + ":X")
  };
  
  try {
    const url = `https://${domain}.freshdesk.com/api/v2/tickets?page=1&per_page=1`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: headers,
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    Logger.log(`🔌 API Test: HTTP ${statusCode}`);
    
    if (statusCode === 200) {
      Logger.log("✅ Freshdesk API connection successful");
      return true;
    } else {
      Logger.log("❌ Freshdesk API connection failed");
      Logger.log(response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log(`❌ API Connection Error: ${e}`);
    return false;
  }
}

/**
 * Test MVR ticket fetching
 * Quick test to verify MVR ticket pattern matching works
 */
function testMVRTicketFetch() {
  try {
    const creds = getFreshdeskCredentials();
    Logger.log('🧪 Testing MVR ticket fetch...');
    
    // Test API connection first
    if (!testApiConnection(creds.apiKey, creds.domain)) {
      throw new Error('API connection test failed');
    }
    
    // Fetch last 24 hours for quick test
    const tickets = fetchMVRTickets(creds.apiKey, creds.domain, 24);
    
    Logger.log(`✅ Test complete: Found ${tickets.length} MVR tickets in last 24 hours`);
    
    if (tickets.length > 0) {
      Logger.log('\n📋 Sample ticket:');
      Logger.log(`   ID: ${tickets[0].id}`);
      Logger.log(`   Subject: ${tickets[0].subject}`);
      Logger.log(`   Status: ${STATUS_NAMES[tickets[0].status]}`);
      Logger.log(`   Created: ${tickets[0].created_at}`);
    }
    
    return tickets.length;
  } catch (e) {
    Logger.log(`❌ Test failed: ${e.message}`);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// BATCH FETCH WITH CONVERSATIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch conversations for existing tickets (OPTIMIZED - recommended)
 * Takes tickets that already have data from list endpoint, adds conversations only
 * This is faster than batchFetchWithConversations (1 call/ticket vs 2 calls/ticket)
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {Array} tickets - Array of ticket objects from list endpoint
 * @param {Object} options - Fetch options
 * @param {number} options.batchSize - Tickets per batch (default: 50)
 * @param {number} options.batchDelayMs - Delay between batches (default: 1000ms)
 * @param {number} options.rateLimitMs - Delay between individual calls (default: 500ms)
 * @return {Array} Same tickets array with conversations property added
 */
function fetchConversationsForTickets(apiKey, domain, tickets, options = {}) {
  const batchSize = options.batchSize || BATCH_CONFIG.BATCH_SIZE;
  const batchDelayMs = options.batchDelayMs || BATCH_CONFIG.BATCH_DELAY_MS;
  const rateLimitMs = options.rateLimitMs || API_CONFIG.RATE_LIMIT_DELAY;
  
  Logger.log(`\n💬 Fetching conversations for ${tickets.length} tickets...`);
  Logger.log(`   Batch size: ${batchSize}, Rate limit: ${rateLimitMs}ms, Batch delay: ${batchDelayMs}ms`);
  Logger.log(`   Estimated time: ${((tickets.length * rateLimitMs + Math.ceil(tickets.length / batchSize) * batchDelayMs) / 60000).toFixed(1)} minutes`);
  
  const totalBatches = Math.ceil(tickets.length / batchSize);
  let successCount = 0;
  let failCount = 0;
  const startTime = new Date();
  
  for (let batchNum = 0; batchNum < totalBatches; batchNum++) {
    const startIdx = batchNum * batchSize;
    const endIdx = Math.min(startIdx + batchSize, tickets.length);
    const batchTickets = tickets.slice(startIdx, endIdx);
    
    Logger.log(`\n📦 Batch ${batchNum + 1}/${totalBatches} (tickets ${startIdx + 1}-${endIdx})...`);
    
    for (const ticket of batchTickets) {
      try {
        ticket.conversations = fetchTicketConversations(apiKey, domain, ticket.id);
        successCount++;
      } catch (e) {
        Logger.log(`⚠️ Ticket ${ticket.id}: ${e.message}`);
        ticket.conversations = [];
        failCount++;
      }
      
      // Rate limiting between calls
      Utilities.sleep(rateLimitMs);
    }
    
    // Progress update
    const elapsed = (new Date() - startTime) / 1000;
    const rate = successCount / elapsed;
    Logger.log(`   ✅ Batch ${batchNum + 1}: ${successCount} success, ${failCount} failed (${rate.toFixed(1)} tickets/sec)`);
    
    // Delay between batches (not after last batch)
    if (batchNum < totalBatches - 1) {
      Utilities.sleep(batchDelayMs);
    }
  }
  
  const totalDuration = (new Date() - startTime) / 1000;
  Logger.log(`\n✅ Conversation fetch complete: ${successCount} success, ${failCount} failed in ${totalDuration.toFixed(1)}s`);
  return tickets;
}

/**
 * Batch fetch tickets with their conversations
 * Processes in batches with delays to respect rate limits
 * 
 * @param {string} apiKey - Freshdesk API key
 * @param {string} domain - Freshdesk domain
 * @param {Array} ticketIds - Array of ticket IDs to fetch
 * @param {Object} options - Optional configuration
 * @param {number} options.batchSize - Tickets per batch (default: BATCH_CONFIG.BATCH_SIZE)
 * @param {number} options.batchDelayMs - Delay between batches (default: BATCH_CONFIG.BATCH_DELAY_MS)
 * @param {boolean} options.includeConversations - Fetch conversations (default: true)
 * @return {Array} Array of ticket objects with conversations property
 */
function batchFetchWithConversations(apiKey, domain, ticketIds, options = {}) {
  const batchSize = options.batchSize || BATCH_CONFIG.BATCH_SIZE;
  const batchDelayMs = options.batchDelayMs || BATCH_CONFIG.BATCH_DELAY_MS;
  const includeConversations = options.includeConversations !== false;
  
  Logger.log(`\n🔄 Batch fetching ${ticketIds.length} tickets with conversations...`);
  Logger.log(`   Batch size: ${batchSize}, Delay: ${batchDelayMs}ms`);
  
  const results = [];
  const totalBatches = Math.ceil(ticketIds.length / batchSize);
  let successCount = 0;
  let failCount = 0;
  
  for (let batchNum = 0; batchNum < totalBatches; batchNum++) {
    const startIdx = batchNum * batchSize;
    const endIdx = Math.min(startIdx + batchSize, ticketIds.length);
    const batchIds = ticketIds.slice(startIdx, endIdx);
    
    Logger.log(`\n📦 Processing batch ${batchNum + 1}/${totalBatches} (tickets ${startIdx + 1}-${endIdx})...`);
    
    for (const ticketId of batchIds) {
      try {
        // Fetch ticket details with custom fields
        const ticket = fetchTicketDetails(apiKey, domain, ticketId, true);
        
        if (ticket) {
          // Fetch conversations if requested
          if (includeConversations) {
            Utilities.sleep(API_CONFIG.RATE_LIMIT_DELAY); // Rate limit between calls
            ticket.conversations = fetchTicketConversations(apiKey, domain, ticketId);
          } else {
            ticket.conversations = [];
          }
          
          results.push(ticket);
          successCount++;
        } else {
          failCount++;
        }
      } catch (e) {
        Logger.log(`⚠️ Error processing ticket ${ticketId}: ${e.message}`);
        failCount++;
      }
      
      // Rate limiting between individual tickets
      Utilities.sleep(API_CONFIG.RATE_LIMIT_DELAY);
    }
    
    // Delay between batches (not after last batch)
    if (batchNum < totalBatches - 1) {
      Logger.log(`⏳ Batch ${batchNum + 1} complete. Waiting ${batchDelayMs}ms before next batch...`);
      Utilities.sleep(batchDelayMs);
    }
  }
  
  Logger.log(`\n✅ Batch fetch complete: ${successCount} success, ${failCount} failed`);
  return results;
}

/**
 * Fetch tickets with conversations for backfill operation
 * Specifically designed for updating historical data with new fields
 * 
 * @param {Array} ticketIds - Array of ticket IDs to fetch
 * @return {Array} Array of enriched ticket objects with conversations
 */
function fetchTicketsForBackfill(ticketIds) {
  Logger.log(`\n🔄 Starting backfill fetch for ${ticketIds.length} tickets...`);
  
  const creds = getFreshdeskCredentials();
  return batchFetchWithConversations(creds.apiKey, creds.domain, ticketIds, {
    batchSize: BATCH_CONFIG.BATCH_SIZE,
    batchDelayMs: BATCH_CONFIG.BATCH_DELAY_MS,
    includeConversations: true
  });
}
