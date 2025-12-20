/**
 * DATA PROCESSING & HISTORICAL APPEND SYSTEM
 * MVR TICKET TRACKER - Deduplication & Append Logic
 * 
 * Purpose: Maintain historical append-only record of MVR tickets
 * - Fetch new tickets from Freshdesk (10-day lookback)
 * - Deduplicate against existing history
 * - Append only new tickets to history sheet
 */

// ═══════════════════════════════════════════════════════════════════════════════
// DEDUPLICATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Load existing ticket IDs from history sheet for deduplication
 * @param {Sheet} historySheet - MVR_Ticket_History sheet
 * @return {Set} Set of existing ticket IDs
 */
function loadExistingTicketIds(historySheet) {
  const data = historySheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    // Only headers or empty sheet
    return new Set();
  }
  
  // Get ticket IDs from first column (skip header row)
  const ticketIds = new Set();
  for (let i = 1; i < data.length; i++) {
    const ticketId = data[i][0]; // First column is Ticket_ID
    if (ticketId) {
      ticketIds.add(ticketId);
    }
  }
  
  Logger.log(`📋 Loaded ${ticketIds.size} existing ticket IDs from history`);
  return ticketIds;
}

/**
 * Filter out tickets that already exist in history
 * @param {Array} tickets - Array of enriched ticket objects
 * @param {Set} existingIds - Set of existing ticket IDs
 * @return {Array} Array of new tickets only
 */
function filterNewTickets(tickets, existingIds) {
  const newTickets = tickets.filter(ticket => !existingIds.has(ticket.id));
  
  Logger.log(`🔍 Filtered ${tickets.length} tickets → ${newTickets.length} new tickets`);
  return newTickets;
}

// ═══════════════════════════════════════════════════════════════════════════════
// HISTORICAL APPEND
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Append new tickets to history sheet
 * @param {Array} tickets - Array of enriched ticket objects
 * @param {Sheet} historySheet - MVR_Ticket_History sheet
 * @return {number} Number of tickets appended
 */
function appendToHistory(tickets, historySheet) {
  if (tickets.length === 0) {
    Logger.log('⚠️ No new tickets to append');
    return 0;
  }
  
  // Convert tickets to rows
  const rows = ticketsToRows(tickets);
  
  // Get next available row
  const lastRow = historySheet.getLastRow();
  const startRow = lastRow + 1;
  
  // Append rows
  historySheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  
  Logger.log(`✅ Appended ${rows.length} tickets to history (rows ${startRow}-${startRow + rows.length - 1})`);
  return rows.length;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN PROCESSING FUNCTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fetch and append new MVR tickets to history
 * Main entry point for data processing
 * 
 * @return {Object} Summary object with counts
 */
function fetchAndAppendNewTickets() {
  Logger.log('\n🚀 Starting MVR ticket fetch and append process...\n');
  
  const startTime = new Date();
  
  try {
    // 1. Get Freshdesk credentials
    Logger.log('1️⃣ Loading Freshdesk credentials...');
    const creds = getFreshdeskCredentials();
    Logger.log('✅ Credentials loaded\n');
    
    // 2. Get or create history sheet
    Logger.log('2️⃣ Initializing history sheet...');
    const historySheet = getOrCreateSheet(SHEET_NAMES.MVR_TICKET_HISTORY, HISTORY_HEADERS);
    Logger.log('✅ History sheet ready\n');
    
    // 3. Load existing ticket IDs for deduplication
    Logger.log('3️⃣ Loading existing ticket IDs...');
    const existingIds = loadExistingTicketIds(historySheet);
    Logger.log(`✅ Loaded ${existingIds.size} existing tickets\n`);
    
    // 4. Fetch tickets from Freshdesk (10-day lookback)
    Logger.log('4️⃣ Fetching MVR tickets from Freshdesk...');
    const rawTickets = fetchMVRTickets(creds.apiKey, creds.domain, TIME_CONFIG.LOOKBACK_HOURS);
    Logger.log(`✅ Fetched ${rawTickets.length} MVR tickets\n`);
    
    if (rawTickets.length === 0) {
      Logger.log('⚠️ No MVR tickets found in lookback window');
      return {
        success: true,
        fetched: 0,
        existing: existingIds.size,
        new: 0,
        appended: 0,
        duration: (new Date() - startTime) / 1000
      };
    }
    
    // 5. Enrich tickets
    Logger.log('5️⃣ Enriching tickets...');
    const enrichedTickets = enrichTickets(rawTickets);
    Logger.log('✅ Enrichment complete\n');
    
    // 6. Filter new tickets
    Logger.log('6️⃣ Filtering new tickets...');
    const newTickets = filterNewTickets(enrichedTickets, existingIds);
    Logger.log(`✅ Found ${newTickets.length} new tickets\n`);
    
    if (newTickets.length === 0) {
      Logger.log('✅ No new tickets to append - history is up to date');
      return {
        success: true,
        fetched: rawTickets.length,
        existing: existingIds.size,
        new: 0,
        appended: 0,
        duration: (new Date() - startTime) / 1000
      };
    }
    
    // 7. Append new tickets to history
    Logger.log('7️⃣ Appending new tickets to history...');
    const appended = appendToHistory(newTickets, historySheet);
    Logger.log(`✅ Append complete\n`);
    
    // 8. Summary
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n📊 === PROCESS SUMMARY ===');
    Logger.log(`   Total fetched: ${rawTickets.length}`);
    Logger.log(`   Existing in history: ${existingIds.size}`);
    Logger.log(`   New tickets: ${newTickets.length}`);
    Logger.log(`   Appended: ${appended}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    Logger.log('\n✅ Process complete!\n');
    
    return {
      success: true,
      fetched: rawTickets.length,
      existing: existingIds.size,
      new: newTickets.length,
      appended: appended,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Error in fetch and append process: ${e.message}`);
    Logger.log(e.stack);
    
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SYNC TICKETS (FETCH + APPEND NEW + UPDATE CHANGED)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Sync MVR tickets - fetches, appends new, and updates changed tickets
 * This is the unified function for both periodic and manual syncs
 * 
 * @param {Object} options - Sync options
 * @param {number} options.lookbackHours - Hours to look back (default: 30 days)
 * @param {boolean} options.includeConversations - Fetch conversations (default: true)
 * @return {Object} Summary with counts
 */
function syncTickets(options = {}) {
  const lookbackHours = options.lookbackHours || TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS;
  const includeConversations = options.includeConversations !== false;
  
  Logger.log('\n🔄 === STARTING TICKET SYNC ===');
  Logger.log(`📅 Lookback: ${lookbackHours / 24} days (${lookbackHours} hours)`);
  Logger.log(`💬 Include conversations: ${includeConversations}\n`);
  
  const startTime = new Date();
  
  try {
    // 1. Get credentials and sheet
    const creds = getFreshdeskCredentials();
    const historySheet = getOrCreateSheet(SHEET_NAMES.MVR_TICKET_HISTORY, HISTORY_HEADERS);
    
    // 2. Load existing data for comparison
    const existingData = loadExistingTicketData(historySheet);
    Logger.log(`📋 Loaded ${existingData.size} existing tickets from history`);
    
    // 3. Fetch tickets from Freshdesk
    Logger.log(`\n🔍 Fetching tickets from Freshdesk...`);
    const rawTickets = fetchMVRTickets(creds.apiKey, creds.domain, lookbackHours);
    Logger.log(`📥 Fetched ${rawTickets.length} MVR tickets from API`);
    
    if (rawTickets.length === 0) {
      return {
        success: true,
        fetched: 0,
        newTickets: 0,
        updatedTickets: 0,
        unchangedTickets: 0,
        duration: (new Date() - startTime) / 1000
      };
    }
    
    // 4. Categorize tickets: new vs existing (changed or unchanged)
    const newTickets = [];
    const changedTickets = [];
    const unchangedTickets = [];
    
    for (const ticket of rawTickets) {
      const existing = existingData.get(ticket.id);
      
      if (!existing) {
        newTickets.push(ticket);
      } else if (hasTicketChanged(ticket, existing)) {
        changedTickets.push({ ticket, rowNum: existing.rowNum });
      } else {
        unchangedTickets.push(ticket);
      }
    }
    
    Logger.log(`\n📊 Categorization:`);
    Logger.log(`   New: ${newTickets.length}`);
    Logger.log(`   Changed: ${changedTickets.length}`);
    Logger.log(`   Unchanged: ${unchangedTickets.length}`);
    
    // 5. Fetch conversations for new and changed tickets
    const ticketsNeedingProcessing = [...newTickets, ...changedTickets.map(c => c.ticket)];
    
    if (includeConversations && ticketsNeedingProcessing.length > 0) {
      Logger.log(`\n💬 Fetching conversations for ${ticketsNeedingProcessing.length} tickets...`);
      const ticketIds = ticketsNeedingProcessing.map(t => t.id);
      const ticketsWithConversations = fetchTicketsForBackfill(ticketIds);
      
      // Merge conversations back
      const convMap = new Map(ticketsWithConversations.map(t => [t.id, t.conversations || []]));
      
      for (const ticket of newTickets) {
        ticket.conversations = convMap.get(ticket.id) || [];
      }
      for (const item of changedTickets) {
        item.ticket.conversations = convMap.get(item.ticket.id) || [];
      }
    }
    
    // 6. Load existing overrides to preserve
    const existingOverridesMap = loadExistingOverrides(historySheet);
    
    // 7. Enrich and append new tickets
    let appendedCount = 0;
    if (newTickets.length > 0) {
      Logger.log(`\n📝 Processing ${newTickets.length} new tickets...`);
      const enrichedNew = enrichTickets(newTickets, existingOverridesMap);
      appendedCount = appendToHistory(enrichedNew, historySheet);
    }
    
    // 8. Enrich and update changed tickets
    let updatedCount = 0;
    if (changedTickets.length > 0) {
      Logger.log(`\n🔄 Updating ${changedTickets.length} changed tickets...`);
      
      for (const { ticket, rowNum } of changedTickets) {
        // Preserve overrides
        const existingOverride = existingOverridesMap.get(ticket.id);
        const enriched = enrichTicket(ticket, ticket.conversations || [], existingOverride || {});
        const newRow = ticketToRow(enriched);
        
        historySheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
        updatedCount++;
        
        if (updatedCount % 25 === 0) {
          Logger.log(`   Progress: ${updatedCount}/${changedTickets.length} updated`);
        }
      }
    }
    
    // 9. Apply highlighting
    applyOverrideHighlighting(historySheet);
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === SYNC COMPLETE ===');
    Logger.log(`   Fetched: ${rawTickets.length}`);
    Logger.log(`   New (appended): ${appendedCount}`);
    Logger.log(`   Changed (updated): ${updatedCount}`);
    Logger.log(`   Unchanged: ${unchangedTickets.length}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      fetched: rawTickets.length,
      newTickets: appendedCount,
      updatedTickets: updatedCount,
      unchangedTickets: unchangedTickets.length,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Sync error: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

/**
 * Load existing ticket data for change comparison
 * Returns Map of ticketId -> { rowNum, status, tags, resolvedAt, ... }
 */
function loadExistingTicketData(historySheet) {
  const data = historySheet.getDataRange().getValues();
  const result = new Map();
  
  if (data.length <= 1) return result;
  
  const headers = data[0];
  const colMap = createColumnMapper(headers);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ticketId = row[colMap['Ticket ID']];
    
    if (ticketId) {
      result.set(ticketId, {
        rowNum: i + 1,  // 1-indexed
        status: row[colMap['Status']] || '',
        tags: row[colMap['Tags']] || '',
        priority: row[colMap['Priority']] || '',
        agentName: row[colMap['Agent Name']] || '',
        groupName: row[colMap['Group']] || '',
        resolvedAt: row[colMap['Date Resolved']] || '',
        closedAt: row[colMap['Date Closed']] || '',
        mvrOutcome: row[colMap['MVR Outcome']] || ''
      });
    }
  }
  
  return result;
}

/**
 * Check if a ticket has changed compared to stored data
 * @param {Object} freshTicket - Ticket from API
 * @param {Object} storedData - Data from history sheet
 * @return {boolean} True if ticket has changed
 */
function hasTicketChanged(freshTicket, storedData) {
  // Status change
  const freshStatus = STATUS_NAMES[freshTicket.status] || String(freshTicket.status);
  if (freshStatus !== storedData.status) {
    Logger.log(`   🔄 Ticket ${freshTicket.id}: Status changed (${storedData.status} → ${freshStatus})`);
    return true;
  }
  
  // Tags change (most important for outcomes)
  const freshTags = (freshTicket.tags || []).sort().join(', ');
  const storedTags = storedData.tags;
  if (freshTags !== storedTags) {
    Logger.log(`   🔄 Ticket ${freshTicket.id}: Tags changed`);
    return true;
  }
  
  // Resolution timestamp (ticket was resolved)
  if (freshTicket.stats?.resolved_at && !storedData.resolvedAt) {
    Logger.log(`   🔄 Ticket ${freshTicket.id}: Now resolved`);
    return true;
  }
  
  // Closure timestamp (ticket was closed)
  if (freshTicket.stats?.closed_at && !storedData.closedAt) {
    Logger.log(`   🔄 Ticket ${freshTicket.id}: Now closed`);
    return true;
  }
  
  return false;
}

/**
 * Periodic sync - 10 day lookback (for automated triggers)
 */
function syncTicketsPeriodic() {
  return syncTickets({
    lookbackHours: TIME_CONFIG.PERIODIC_LOOKBACK_HOURS,
    includeConversations: true
  });
}

/**
 * Full sync - 30 day lookback (for manual runs)
 */
function syncTicketsFull() {
  return syncTickets({
    lookbackHours: TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS,
    includeConversations: true
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// RAW TICKET PULL (No Transformation - Subject Filtering Only)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Headers for raw ticket data - all native Freshdesk fields
 */
const RAW_TICKET_HEADERS = [
  'Ticket ID',
  'Subject',
  'Description',
  'Status',
  'Priority',
  'Type',
  'Source',
  'Tags',
  'Requester ID',
  'Requester Email',
  'Requester Name',
  'Responder ID',
  'Group ID',
  'Company ID',
  'Created At',
  'Updated At',
  'Due By',
  'First Response Due By',
  'First Responded At',
  'Resolved At',
  'Closed At',
  'Is Escalated',
  'Spam',
  'Custom Fields (JSON)',
  'Raw Ticket JSON',
  'Change Log',  // Inline changelog: MM/DD HH:MM action | action | ...
  // Complete data columns (added for full ticket capture)
  'All Notes Text',           // Description + all conversation body_text joined by ---
  'Conversation Count',       // Number of conversations
  'Attachment Names',         // Comma-separated attachment filenames
  'Attachment URLs',          // Newline-separated attachment URLs
  'Last Conversation Date',   // Timestamp of most recent conversation
  'Last Responder Name'       // Name of last person who responded
];

/**
 * Pull raw MVR tickets based on subject patterns only
 * No transformation, no analysis - just raw Freshdesk data
 * 
 * @param {Object} options - Pull options
 * @param {number} options.lookbackHours - Hours to look back (default: 720 = 30 days)
 * @return {Object} Summary with counts
 */
function pullRawMVRTickets(options = {}) {
  const lookbackHours = options.lookbackHours || TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS;
  
  Logger.log('\n📦 === STARTING RAW MVR TICKET PULL ===');
  Logger.log(`📅 Lookback: ${lookbackHours / 24} days (${lookbackHours} hours)`);
  Logger.log(`🔍 Filter: Subject patterns only (no ticket type filter)\n`);
  
  const startTime = new Date();
  
  try {
    // 1. Get credentials
    const creds = getFreshdeskCredentials();
    
    // 2. Get or create raw tickets sheet
    const rawSheet = getOrCreateSheet(SHEET_NAMES.MVR_RAW_TICKETS, RAW_TICKET_HEADERS);
    
    // Clear existing data (keep headers)
    if (rawSheet.getLastRow() > 1) {
      rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, rawSheet.getLastColumn()).clear();
    }
    
    // 3. Fetch tickets using expanded MVR subject patterns
    Logger.log('🔍 Fetching MVR tickets from Freshdesk...');
    const rawTickets = fetchMVRTickets(creds.apiKey, creds.domain, lookbackHours);
    Logger.log(`📥 Fetched ${rawTickets.length} MVR tickets\n`);
    
    if (rawTickets.length === 0) {
      Logger.log('⚠️ No MVR tickets found in lookback window');
      return {
        success: true,
        totalTickets: 0,
        patternsMatched: 0,
        duration: (new Date() - startTime) / 1000
      };
    }
    
    // 4. Log pattern distribution
    const patternCounts = {};
    MVR_SUBJECT_PATTERNS.forEach((pattern, idx) => {
      const patternStr = pattern.toString();
      patternCounts[patternStr] = rawTickets.filter(t => pattern.test(t.subject)).length;
    });
    
    Logger.log('📊 Pattern Distribution:');
    Object.entries(patternCounts).forEach(([pattern, count]) => {
      if (count > 0) {
        Logger.log(`   ${pattern}: ${count} tickets`);
      }
    });
    
    // 5. Convert tickets to raw rows
    Logger.log(`\n📝 Converting ${rawTickets.length} tickets to raw format...`);
    const rows = rawTickets.map(ticket => ticketToRawRow(ticket));
    
    // 6. Write to sheet
    if (rows.length > 0) {
      rawSheet.getRange(2, 1, rows.length, RAW_TICKET_HEADERS.length).setValues(rows);
    }
    
    // 7. Auto-resize columns
    rawSheet.autoResizeColumns(1, RAW_TICKET_HEADERS.length);
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === RAW PULL COMPLETE ===');
    Logger.log(`   Total tickets: ${rows.length}`);
    Logger.log(`   Sheet: ${SHEET_NAMES.MVR_RAW_TICKETS}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      totalTickets: rows.length,
      patternsMatched: Object.values(patternCounts).reduce((a, b) => a + b, 0),
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Raw pull error: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

/**
 * Convert a single ticket to a raw row (no transformation)
 * Includes conversations, attachments, and complete notes
 * @param {Object} ticket - Raw Freshdesk ticket object (may include conversations)
 * @return {Array} Row data matching RAW_TICKET_HEADERS
 */
function ticketToRawRow(ticket) {
  // Process conversations and attachments
  const conversationData = processTicketConversations(ticket);
  
  return [
    ticket.id,
    ticket.subject || '',
    (ticket.description_text || ticket.description || '').substring(0, 5000),
    STATUS_NAMES[ticket.status] || ticket.status,
    PRIORITY_NAMES[ticket.priority] || ticket.priority,
    ticket.type || '',
    ticket.source || '',
    (ticket.tags || []).join(', '),
    ticket.requester_id || '',
    ticket.requester?.email || '',
    ticket.requester?.name || '',
    ticket.responder_id || '',
    ticket.group_id || '',
    ticket.company_id || '',
    ticket.created_at || '',
    ticket.updated_at || '',
    ticket.due_by || '',
    ticket.fr_due_by || '',
    ticket.stats?.first_responded_at || ticket.fr_escalated_at || '',
    ticket.stats?.resolved_at || '',
    ticket.stats?.closed_at || '',
    ticket.is_escalated || false,
    ticket.spam || false,
    JSON.stringify(ticket.custom_fields || {}),
    truncateJsonForCell(ticket, ['conversations', 'description', 'description_text']),  // Exclude large redundant fields
    ticket._changeLog || '',
    // Complete data fields
    conversationData.allNotesText,
    conversationData.conversationCount,
    conversationData.attachmentNames,
    conversationData.attachmentUrls,
    conversationData.lastConversationDate,
    conversationData.lastResponderName
  ];
}

/**
 * Process ticket conversations and attachments
 * Extracts all notes text, attachment info, and conversation metadata
 * @param {Object} ticket - Ticket object (may have conversations array)
 * @return {Object} Processed conversation data
 */
function processTicketConversations(ticket) {
  const conversations = ticket.conversations || [];
  const descriptionText = ticket.description_text || ticket.description || '';
  
  // Build combined notes text
  const allNotes = [descriptionText.trim()];
  const attachmentNames = [];
  const attachmentUrls = [];
  let lastConversationDate = '';
  let lastResponderName = '';
  
  if (conversations.length > 0) {
    // Sort by created_at to get chronological order
    const sortedConvos = conversations.sort((a, b) => 
      new Date(a.created_at) - new Date(b.created_at)
    );
    
    sortedConvos.forEach(convo => {
      // Add body text
      if (convo.body_text) {
        allNotes.push(convo.body_text.trim());
      }
      
      // Collect attachments
      if (convo.attachments && convo.attachments.length > 0) {
        convo.attachments.forEach(att => {
          if (att.name) attachmentNames.push(att.name);
          if (att.attachment_url) attachmentUrls.push(att.attachment_url);
        });
      }
      
      // Track last conversation metadata
      lastConversationDate = convo.created_at || lastConversationDate;
      if (convo.user_id) {
        lastResponderName = getAgentName(convo.user_id) || convo.user_id;
      }
    });
  }
  
  // Truncate text fields to avoid Google Sheets 50k char cell limit
  const rawNotesText = allNotes.join('\n---\n');
  const rawUrls = attachmentUrls.join('\n');
  
  return {
    allNotesText: truncateForCell(rawNotesText, CELL_LIMITS.TEXT_FIELD_LIMIT),
    conversationCount: conversations.length,
    attachmentNames: attachmentNames.join(', '),
    attachmentUrls: truncateForCell(rawUrls, CELL_LIMITS.URL_FIELD_LIMIT),
    lastConversationDate: lastConversationDate,
    lastResponderName: lastResponderName
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// RAW PULL WITH CHANGE TRACKING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Fields to track for changes in raw tickets
 */
const RAW_CHANGE_FIELDS = ['Status', 'Priority', 'Tags', 'Responder ID', 'Group ID', 'Resolved At', 'Closed At'];

/**
 * Pull raw MVR tickets with change detection and inline changelog
 * Compares against existing data before overwriting
 * 
 * @param {Object} options - Pull options
 * @param {number} options.lookbackHours - Hours to look back
 * @return {Object} Summary with counts
 */
function pullRawWithChangeTracking(options = {}) {
  const lookbackHours = options.lookbackHours || TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS;
  
  Logger.log('\n📦 === RAW PULL WITH CHANGE TRACKING ===');
  Logger.log(`📅 Lookback: ${lookbackHours} hours (${(lookbackHours / 24).toFixed(1)} days)`);
  
  const startTime = new Date();
  const now = new Date();
  const timestamp = `${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getDate().toString().padStart(2, '0')} ${now.getHours().toString().padStart(2, '0')}:${now.getMinutes().toString().padStart(2, '0')}`;
  
  try {
    const creds = getFreshdeskCredentials();
    const rawSheet = getOrCreateSheet(SHEET_NAMES.MVR_RAW_TICKETS, RAW_TICKET_HEADERS);
    
    // 1. Load existing data for comparison
    const existingData = loadExistingRawData(rawSheet);
    Logger.log(`📋 Loaded ${existingData.size} existing tickets from raw sheet`);
    
    // 2. Fetch fresh tickets from API
    Logger.log('🔍 Fetching MVR tickets from Freshdesk...');
    const freshTickets = fetchMVRTickets(creds.apiKey, creds.domain, lookbackHours);
    
    if (!freshTickets || freshTickets.length === 0) {
      return { success: true, totalTickets: 0, newTickets: 0, changedTickets: 0, duration: (new Date() - startTime) / 1000 };
    }
    
    Logger.log(`📥 Fetched ${freshTickets.length} tickets from list endpoint`);
    
    // 3. Fetch conversations for all tickets (1 API call per ticket)
    Logger.log(`💬 Fetching conversations for ${freshTickets.length} tickets...`);
    fetchConversationsForTickets(creds.apiKey, creds.domain, freshTickets, {
      batchSize: BATCH_CONFIG.BATCH_SIZE,
      batchDelayMs: BATCH_CONFIG.BATCH_DELAY_MS,
      rateLimitMs: API_CONFIG.RATE_LIMIT_DELAY
    });
    
    Logger.log(`✅ Complete data ready for ${freshTickets.length} tickets`);
    
    // 4. Compare and build change logs
    let newCount = 0;
    let changedCount = 0;
    
    for (const ticket of freshTickets) {
      const existing = existingData.get(ticket.id);
      
      if (!existing) {
        // New ticket
        ticket._changeLog = `${timestamp} Created`;
        newCount++;
      } else {
        // Existing ticket - check for changes
        const changes = detectRawChanges(ticket, existing);
        
        if (changes.length > 0) {
          const changeNotes = changes.map(c => `${c.field}:${c.old}→${c.new}`).join(', ');
          const newEntry = `${timestamp} ${changeNotes}`;
          ticket._changeLog = existing.changeLog ? `${existing.changeLog} | ${newEntry}` : newEntry;
          changedCount++;
          Logger.log(`   🔄 Ticket ${ticket.id}: ${changeNotes}`);
        } else {
          // No changes - preserve existing changelog
          ticket._changeLog = existing.changeLog || '';
        }
      }
    }
    
    Logger.log(`\n📊 Change Summary: ${newCount} new, ${changedCount} changed, ${freshTickets.length - newCount - changedCount} unchanged`);
    
    // 4. Clear and overwrite sheet
    if (rawSheet.getLastRow() > 1) {
      rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, rawSheet.getLastColumn()).clear();
    }
    
    // 5. Convert and write
    const rows = freshTickets.map(ticket => ticketToRawRow(ticket));
    if (rows.length > 0) {
      rawSheet.getRange(2, 1, rows.length, RAW_TICKET_HEADERS.length).setValues(rows);
    }
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === RAW PULL WITH TRACKING COMPLETE ===');
    Logger.log(`   Total: ${rows.length}, New: ${newCount}, Changed: ${changedCount}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      totalTickets: rows.length,
      newTickets: newCount,
      changedTickets: changedCount,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Error: ${e.message}`);
    Logger.log(e.stack);
    return { success: false, error: e.message, duration: (new Date() - startTime) / 1000 };
  }
}

/**
 * Load existing raw ticket data for comparison
 * @param {Sheet} rawSheet - MVR_Raw_Tickets sheet
 * @return {Map} Map of ticketId -> { status, priority, tags, ..., changeLog }
 */
function loadExistingRawData(rawSheet) {
  const data = rawSheet.getDataRange().getValues();
  const result = new Map();
  
  if (data.length <= 1) return result;
  
  const headers = data[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ticketId = row[colMap['Ticket ID']];
    
    if (ticketId) {
      result.set(ticketId, {
        status: row[colMap['Status']] || '',
        priority: row[colMap['Priority']] || '',
        tags: row[colMap['Tags']] || '',
        responderId: row[colMap['Responder ID']] || '',
        groupId: row[colMap['Group ID']] || '',
        resolvedAt: row[colMap['Resolved At']] || '',
        closedAt: row[colMap['Closed At']] || '',
        changeLog: row[colMap['Change Log']] || ''
      });
    }
  }
  
  return result;
}

/**
 * Detect changes between fresh ticket and existing data
 * @param {Object} fresh - Fresh ticket from API
 * @param {Object} existing - Existing data from sheet
 * @return {Array} Array of { field, old, new } objects
 */
function detectRawChanges(fresh, existing) {
  const changes = [];
  
  // Status
  const freshStatus = STATUS_NAMES[fresh.status] || String(fresh.status);
  if (freshStatus !== existing.status && existing.status) {
    changes.push({ field: 'status', old: existing.status, new: freshStatus });
  }
  
  // Priority  
  const freshPriority = PRIORITY_NAMES[fresh.priority] || String(fresh.priority);
  if (freshPriority !== existing.priority && existing.priority) {
    changes.push({ field: 'priority', old: existing.priority, new: freshPriority });
  }
  
  // Tags
  const freshTags = (fresh.tags || []).sort().join(', ');
  if (freshTags !== existing.tags) {
    const oldTags = existing.tags.split(', ').filter(t => t);
    const newTags = (fresh.tags || []);
    const added = newTags.filter(t => !oldTags.includes(t));
    const removed = oldTags.filter(t => !newTags.includes(t));
    
    if (added.length > 0) changes.push({ field: 'tags', old: '', new: '+' + added.join(',+') });
    if (removed.length > 0) changes.push({ field: 'tags', old: '-' + removed.join(',-'), new: '' });
  }
  
  // Agent assignment
  const freshResponder = fresh.responder_id ? String(fresh.responder_id) : '';
  const existingResponder = existing.responderId ? String(existing.responderId) : '';
  if (freshResponder !== existingResponder) {
    changes.push({ field: 'agent', old: existingResponder || 'none', new: freshResponder || 'none' });
  }
  
  // Resolved
  if (fresh.stats?.resolved_at && !existing.resolvedAt) {
    changes.push({ field: 'resolved', old: '', new: 'yes' });
  }
  
  // Closed
  if (fresh.stats?.closed_at && !existing.closedAt) {
    changes.push({ field: 'closed', old: '', new: 'yes' });
  }
  
  return changes;
}

// ═══════════════════════════════════════════════════════════════════════════════
// POST-PULL TAG DISCOVERY
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Run after every pull to discover and track tags
 * Notifies user of unmapped tags
 */
function afterPullComplete() {
  Logger.log('\n📊 === POST-PULL TAG DISCOVERY ===');
  
  try {
    // 1. Discover tags from raw data
    const discovery = discoverTagsFromRaw();
    
    if (!discovery.success) {
      Logger.log(`   ⚠️ Tag discovery failed: ${discovery.error}`);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Tag discovery failed. Check logs for details.',
        'Warning',
        5
      );
      return;
    }
    
    Logger.log(`   ✅ Found ${discovery.totalTags} unique tags (${discovery.newTags} new)`);
    
    // 2. Identify unmapped tags
    const unmapped = identifyUnmappedTags();
    Logger.log(`   ⚠️ Unmapped tags: ${unmapped.length}`);
    
    // 3. Calculate data quality: % of tickets with unknown outcome
    const rawSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
    const totalTickets = rawSheet ? rawSheet.getLastRow() - 1 : 0;
    const unmappedRate = totalTickets > 0 ? (unmapped.reduce((sum, t) => sum + t.count, 0) / totalTickets) : 0;
    
    Logger.log(`   📊 Data Quality: ${(unmappedRate * 100).toFixed(1)}% tickets may have Unknown outcome`);
    
    // 4. Notify user based on severity
    if (unmapped.length > 0) {
      const topTags = unmapped.slice(0, 3).map(t => `${t.tag} (${t.count})`).join(', ');
      
      // Critical alert if >5% tickets affected
      if (unmappedRate > 0.05) {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `⚠️ DATA QUALITY WARNING: ${unmapped.length} unmapped tags affecting ${(unmappedRate * 100).toFixed(1)}% of tickets (${topTags}...). Finance reports may be inaccurate! Go to 🏷️ Tag Management → View Unmapped Tags`,
          '🚨 Action Required',
          15
        );
      } else if (discovery.newTags > 0) {
        // Info alert for new tags
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `Found ${discovery.newTags} new unmapped tags (${topTags}...). Go to 🏷️ Tag Management → View Unmapped Tags`,
          'New Tags Discovered',
          10
        );
      }
    }
    
    Logger.log('✅ Post-pull discovery complete');
  } catch (e) {
    Logger.log(`⚠️ Tag discovery error: ${e.message}`);
    // Don't fail the pull if discovery has issues
  }
}

/**
 * Full raw refresh - 30 day lookback
 */
function pullRawFull() {
  const result = pullRawWithChangeTracking({
    lookbackHours: TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS
  });
  
  if (result.success) {
    afterPullComplete();
  }
  
  return result;
}

/**
 * Initial pull - 30 day lookback WITHOUT change tracking
 * Used for first-time setup or complete reset
 * Writes fresh data without comparing to existing
 */
function pullRawInitial() {
  Logger.log('\n🆕 === INITIAL RAW PULL (No Change Tracking) ===');
  
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get API credentials (matching existing pattern in codebase)
    const creds = getFreshdeskCredentials();
    if (!creds.apiKey || !creds.domain) {
      throw new Error('API credentials not configured. Go to Settings → Configure Freshdesk API Key');
    }
    
    const lookbackHours = TIME_CONFIG.FULL_SYNC_LOOKBACK_HOURS;
    Logger.log(`📅 Lookback: ${lookbackHours} hours (30 days)`);
    
    // Fetch tickets from API (list endpoint includes all needed ticket data)
    const tickets = fetchMVRTickets(creds.apiKey, creds.domain, lookbackHours);
    
    if (!tickets || tickets.length === 0) {
      Logger.log('⚠️ No tickets found');
      return { success: true, totalTickets: 0, duration: 0 };
    }
    
    Logger.log(`📥 Fetched ${tickets.length} tickets from list endpoint`);
    
    // Fetch conversations for all tickets (1 API call per ticket)
    Logger.log(`💬 Fetching conversations for ${tickets.length} tickets (batched)...`);
    fetchConversationsForTickets(creds.apiKey, creds.domain, tickets, {
      batchSize: BATCH_CONFIG.BATCH_SIZE,
      batchDelayMs: BATCH_CONFIG.BATCH_DELAY_MS,
      rateLimitMs: API_CONFIG.RATE_LIMIT_DELAY
    });
    
    Logger.log(`✅ Complete data ready for ${tickets.length} tickets with conversations`);
    
    // Get/create raw sheet
    const rawSheet = getOrCreateSheet(SHEET_NAMES.MVR_RAW_TICKETS, RAW_TICKET_HEADERS);
    
    // Clear existing data (keep headers)
    if (rawSheet.getLastRow() > 1) {
      rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, rawSheet.getLastColumn()).clear();
    }
    
    // Build rows using ticketToRawRow (includes conversations/attachments)
    const rows = tickets.map(ticket => ticketToRawRow(ticket));
    
    // Write all rows
    if (rows.length > 0) {
      rawSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === INITIAL PULL COMPLETE ===');
    Logger.log(`   Total tickets: ${rows.length}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      totalTickets: rows.length,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Initial pull error: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PROCESS RAW TO HISTORY (Bridge Function)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Process tickets from MVR_Raw_Tickets to MVR_Ticket_History
 * Reads raw data, enriches each ticket, and writes to history
 * - New tickets are appended
 * - Changed tickets are updated in place
 * - Manual overrides are preserved
 * - Change logs are carried over from raw
 * 
 * @param {Object} options - Processing options
 * @param {boolean} options.forceReprocess - If true, reprocess all tickets even if unchanged
 * @return {Object} Summary with counts
 */
function processRawToHistory(options = {}) {
  const forceReprocess = options.forceReprocess || false;
  
  Logger.log('\n⚙️ === PROCESSING RAW TO HISTORY ===');
  Logger.log(`🔧 Force reprocess: ${forceReprocess}`);
  
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Load raw tickets
    const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      Logger.log('⚠️ No raw tickets to process');
      return { success: true, processed: 0, newTickets: 0, updatedTickets: 0, duration: 0 };
    }
    
    const rawData = rawSheet.getDataRange().getValues();
    const rawHeaders = rawData[0];
    const rawColMap = {};
    rawHeaders.forEach((h, i) => rawColMap[h] = i);
    
    Logger.log(`📋 Loaded ${rawData.length - 1} raw tickets`);
    
    // 2. Load existing history for comparison and override preservation
    const historySheet = getOrCreateSheet(SHEET_NAMES.MVR_TICKET_HISTORY, HISTORY_HEADERS);
    const existingHistory = loadExistingHistoryData(historySheet);
    const existingOverrides = loadExistingOverrides(historySheet);
    Logger.log(`📋 Loaded ${existingHistory.size} existing history tickets`);
    
    // 3. Process each raw ticket
    const newTickets = [];
    const updatedTickets = [];
    let unchangedCount = 0;
    
    for (let i = 1; i < rawData.length; i++) {
      const rawRow = rawData[i];
      const ticketId = rawRow[rawColMap['Ticket ID']];
      
      if (!ticketId) continue;
      
      // Convert raw row to ticket object
      const rawTicket = rawRowToTicketObject(rawRow, rawColMap);
      
      // Check if exists in history
      const existingEntry = existingHistory.get(ticketId);
      
      if (!existingEntry) {
        // New ticket - enrich and add
        const enriched = enrichTicketFromRaw(rawTicket, existingOverrides.get(ticketId));
        newTickets.push(enriched);
      } else if (forceReprocess || hasRawTicketChanged(rawTicket, existingEntry)) {
        // Changed ticket - enrich and update
        const enriched = enrichTicketFromRaw(rawTicket, existingOverrides.get(ticketId));
        updatedTickets.push({ ticket: enriched, rowNum: existingEntry.rowNum });
      } else {
        unchangedCount++;
      }
      
      // Progress logging
      if (i % 200 === 0) {
        Logger.log(`   Progress: ${i}/${rawData.length - 1} processed`);
      }
    }
    
    Logger.log(`\n📊 Processing Summary:`);
    Logger.log(`   New: ${newTickets.length}`);
    Logger.log(`   Updated: ${updatedTickets.length}`);
    Logger.log(`   Unchanged: ${unchangedCount}`);
    
    // 4. Append new tickets
    if (newTickets.length > 0) {
      Logger.log(`\n📝 Appending ${newTickets.length} new tickets...`);
      const newRows = newTickets.map(t => ticketToRow(t));
      const startRow = historySheet.getLastRow() + 1;
      historySheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    }
    
    // 5. Update changed tickets
    if (updatedTickets.length > 0) {
      Logger.log(`\n🔄 Updating ${updatedTickets.length} changed tickets...`);
      for (const { ticket, rowNum } of updatedTickets) {
        const row = ticketToRow(ticket);
        historySheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      }
    }
    
    // 6. Apply override highlighting
    if (newTickets.length > 0 || updatedTickets.length > 0) {
      applyOverrideHighlighting(historySheet);
    }
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === PROCESSING COMPLETE ===');
    Logger.log(`   Processed: ${newTickets.length + updatedTickets.length}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      processed: rawData.length - 1,
      newTickets: newTickets.length,
      updatedTickets: updatedTickets.length,
      unchangedTickets: unchangedCount,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Processing error: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

/**
 * Convert raw sheet row to ticket object
 */
function rawRowToTicketObject(row, colMap) {
  // Parse the Raw Ticket JSON if available for complete data
  let rawJson = {};
  try {
    const jsonStr = row[colMap['Raw Ticket JSON']];
    if (jsonStr) {
      rawJson = JSON.parse(jsonStr);
    }
  } catch (e) {
    // Ignore parse errors, use row data
  }
  
  return {
    id: row[colMap['Ticket ID']],
    subject: row[colMap['Subject']] || '',
    description: row[colMap['Description']] || '',
    status: rawJson.status || row[colMap['Status']],
    priority: rawJson.priority || row[colMap['Priority']],
    type: row[colMap['Type']] || '',
    source: rawJson.source || row[colMap['Source']],
    tags: row[colMap['Tags']] ? row[colMap['Tags']].split(', ').filter(t => t) : [],
    requester_id: row[colMap['Requester ID']] || '',
    responder_id: row[colMap['Responder ID']] || '',
    group_id: row[colMap['Group ID']] || '',
    company_id: row[colMap['Company ID']] || '',
    created_at: row[colMap['Created At']] || '',
    updated_at: row[colMap['Updated At']] || '',
    resolved_at: row[colMap['Resolved At']] || '',
    closed_at: row[colMap['Closed At']] || '',
    custom_fields: rawJson.custom_fields || {},
    change_log: row[colMap['Change Log']] || '',
    // Stats from raw JSON
    stats: rawJson.stats || {}
  };
}

/**
 * Enrich a ticket from raw data (similar to enrichTicket but from raw row)
 * Uses extractVendorGroup from OutcomeTracking.gs for consistent vendor detection
 */
function enrichTicketFromRaw(rawTicket, existingOverride) {
  // Parse subject line
  const parsed = parseSubjectLine(rawTicket.subject);
  
  // Get agent and group names
  const agentName = getAgentName(rawTicket.responder_id);
  const groupName = getGroupName(rawTicket.group_id);
  
  // Parse description for workflow data
  const workflowData = parseTicketDescription(rawTicket.description);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 1: Extract DL State with fallbacks (needed for vendor determination)
  // ═══════════════════════════════════════════════════════════════════════════
  let dlState = workflowData.dl_state || '';
  
  // If no state from description, try subject line
  if (!dlState && parsed.dl_state) {
    dlState = parsed.dl_state;
  }
  
  // If still no state, try aggressive extraction from description + notes
  if (!dlState) {
    const allText = (rawTicket.description || '') + ' ' + (rawTicket.all_notes_text || '');
    dlState = extractStateFromText(allText) || '';
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STEP 2: Extract vendor group using subject AND state
  // ═══════════════════════════════════════════════════════════════════════════
  const vendorGroup = extractVendorGroup(rawTicket.subject, dlState);
  
  // Calculate times
  const createdDate = new Date(rawTicket.created_at);
  const now = new Date();
  const ageHours = (now - createdDate) / (1000 * 60 * 60);
  
  let resolutionHours = null;
  let resolvedDate = null;
  if (rawTicket.resolved_at) {
    resolvedDate = rawTicket.resolved_at;
    resolutionHours = (new Date(rawTicket.resolved_at) - createdDate) / (1000 * 60 * 60);
  }
  
  // Classify outcome from tags (no conversations in raw data)
  const outcome = classifyOutcome(rawTicket.tags, []);
  
  // Build enriched ticket
  return {
    id: rawTicket.id,
    subject: rawTicket.subject,
    type: parsed.type,
    partner: parsed.partner || extractPartnerName(rawTicket),
    turnId: parsed.turnId,
    status: rawTicket.status,
    status_name: STATUS_NAMES[rawTicket.status] || String(rawTicket.status),
    priority: rawTicket.priority,
    priority_name: PRIORITY_NAMES[rawTicket.priority] || String(rawTicket.priority),
    created_at: rawTicket.created_at,
    updated_at: rawTicket.updated_at,
    resolved_date: resolvedDate,
    age_hours: ageHours,
    resolution_hours: resolutionHours,
    responder_id: rawTicket.responder_id,
    agent_name: agentName,
    group_id: rawTicket.group_id,
    group_name: groupName,
    requester_id: rawTicket.requester_id,
    company_id: rawTicket.company_id,
    tags: rawTicket.tags,
    source: rawTicket.source,
    source_name: getSourceName(rawTicket.source),
    sla: '',
    last_pulled: new Date().toISOString(),
    // Workflow data - vendor_group now uses unified extractVendorGroup from OutcomeTracking.gs
    vendor_group: vendorGroup,
    dl_state: dlState,
    dl_number: workflowData.dl_number || '',
    days_since_last: workflowData.days_since_last || '',
    last_check_date: workflowData.last_check_date || '',
    cadence_days: workflowData.cadence_days || '',
    tier: workflowData.tier || '',
    enrollment_type: workflowData.enrollment_type || '',
    ta_url: workflowData.ta_url || '',
    // Outcome
    mvr_outcome: outcome.type,
    outcome_source: outcome.source,
    outcome_date: outcome.date || '',
    resolution_notes: '',
    conversation_count: 0,
    last_responder: agentName,
    // Override (preserve existing)
    outcome_override: existingOverride?.override || '',
    override_reason: existingOverride?.reason || '',
    // Change log (from raw)
    change_log: rawTicket.change_log || ''
  };
}

/**
 * Load existing history data for comparison
 */
function loadExistingHistoryData(historySheet) {
  const data = historySheet.getDataRange().getValues();
  const result = new Map();
  
  if (data.length <= 1) return result;
  
  const headers = data[0];
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const ticketId = row[colMap['Ticket ID']];
    
    if (ticketId) {
      result.set(ticketId, {
        rowNum: i + 1,
        status: row[colMap['Status']] || '',
        tags: row[colMap['Tags']] || '',
        updatedAt: row[colMap['Last Updated']] || '',
        mvrOutcome: row[colMap['MVR Outcome']] || ''
      });
    }
  }
  
  return result;
}

/**
 * Check if raw ticket has changed compared to history
 */
function hasRawTicketChanged(rawTicket, historyEntry) {
  // Status change
  const rawStatus = STATUS_NAMES[rawTicket.status] || String(rawTicket.status);
  if (rawStatus !== historyEntry.status) return true;
  
  // Tags change
  const rawTags = (rawTicket.tags || []).sort().join(', ');
  if (rawTags !== historyEntry.tags) return true;
  
  // Updated timestamp change
  if (rawTicket.updated_at !== historyEntry.updatedAt) return true;
  
  return false;
}

// Note: extractVendorGroup is defined in OutcomeTracking.gs
// It handles InformData, Sentinel, Certn, PennDOT, and state-based fallback

// ═══════════════════════════════════════════════════════════════════════════════
// UPDATE EXISTING TICKETS (Optional - for status changes)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Update existing tickets in history with latest status
 * Now includes conversations and new outcome fields
 * Use this to refresh status, resolution times for tickets that were open
 * 
 * @param {Array} ticketIds - Optional array of specific ticket IDs to update
 * @return {number} Number of tickets updated
 */
function updateExistingTickets(ticketIds = null) {
  Logger.log('\n🔄 Starting ticket update process...\n');
  
  try {
    // Get credentials and history sheet
    const creds = getFreshdeskCredentials();
    const historySheet = getOrCreateSheet(SHEET_NAMES.MVR_TICKET_HISTORY, HISTORY_HEADERS);
    
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('⚠️ No tickets in history to update');
      return 0;
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Determine which tickets to update
    let ticketsToUpdate = [];
    const rowMap = new Map(); // ticketId -> row index
    
    if (ticketIds && ticketIds.length > 0) {
      // Update specific tickets
      Logger.log(`Updating ${ticketIds.length} specific tickets...`);
      ticketsToUpdate = ticketIds;
      
      // Build row map for specified tickets
      for (let i = 1; i < data.length; i++) {
        const ticketId = data[i][colMap['Ticket ID']];
        if (ticketIds.includes(ticketId)) {
          rowMap.set(ticketId, i + 1);
        }
      }
    } else {
      // Update all open/pending tickets
      Logger.log('Updating all open/pending tickets...');
      for (let i = 1; i < data.length; i++) {
        const status = data[i][colMap['Status']];
        const ticketId = data[i][colMap['Ticket ID']];
        
        if ((status === 'Open' || status === 'Pending') && ticketId) {
          ticketsToUpdate.push(ticketId);
          rowMap.set(ticketId, i + 1);
        }
      }
    }
    
    if (ticketsToUpdate.length === 0) {
      Logger.log('✅ No tickets need updating');
      return 0;
    }
    
    Logger.log(`📋 Found ${ticketsToUpdate.length} tickets to update`);
    
    // Load existing overrides to preserve them
    const existingOverridesMap = loadExistingOverrides(historySheet);
    
    // Fetch fresh ticket data WITH conversations using batch processing
    const freshTickets = batchFetchWithConversations(creds.apiKey, creds.domain, ticketsToUpdate, {
      batchSize: BATCH_CONFIG.BATCH_SIZE,
      batchDelayMs: BATCH_CONFIG.BATCH_DELAY_MS,
      includeConversations: true
    });
    
    if (freshTickets.length === 0) {
      Logger.log('⚠️ No fresh ticket data retrieved');
      return 0;
    }
    
    // Enrich fresh tickets with preserved overrides
    const enrichedFresh = enrichTickets(freshTickets, existingOverridesMap);
    
    // Create map for fast lookup
    const freshMap = new Map(enrichedFresh.map(t => [t.id, t]));
    
    // Track override changes for audit logging
    let updateCount = 0;
    for (let i = 1; i < data.length; i++) {
      const ticketId = data[i][colMap['Ticket ID']];
      
      if (freshMap.has(ticketId)) {
        const freshTicket = freshMap.get(ticketId);
        const oldRow = data[i];
        const newRow = ticketToRow(freshTicket);
        
        // Check for override changes
        const overrideChange = detectOverrideChange(oldRow, newRow, colMap);
        if (overrideChange) {
          logOverrideChange(
            overrideChange.ticketId,
            overrideChange.turnId,
            overrideChange.partner,
            overrideChange.previousOutcome,
            overrideChange.newOverride,
            overrideChange.reason
          );
        }
        
        historySheet.getRange(i + 1, 1, 1, newRow.length).setValues([newRow]);
        updateCount++;
      }
    }
    
    // Apply override highlighting
    applyOverrideHighlighting(historySheet);
    
    Logger.log(`✅ Updated ${updateCount} tickets in history`);
    return updateCount;
    
  } catch (e) {
    Logger.log(`❌ Error updating tickets: ${e.message}`);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// BACKFILL HISTORICAL DATA
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Backfill all existing tickets with new fields (conversations, outcomes, workflow data)
 * Processes in batches to respect API rate limits
 * 
 * @return {Object} Summary of backfill operation
 */
function backfillAllTickets() {
  Logger.log('\n🔄 === STARTING HISTORICAL BACKFILL ===\n');
  
  const startTime = new Date();
  
  try {
    // Get history sheet
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found');
    }
    
    const data = historySheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('⚠️ No tickets in history to backfill');
      return { success: true, processed: 0, duration: 0 };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Check if we need to add new columns
    const currentColCount = headers.length;
    const expectedColCount = HISTORY_HEADERS.length;
    
    if (currentColCount < expectedColCount) {
      Logger.log(`📝 Adding ${expectedColCount - currentColCount} new columns to history sheet...`);
      
      // Add new headers
      const newHeaders = HISTORY_HEADERS.slice(currentColCount);
      historySheet.getRange(1, currentColCount + 1, 1, newHeaders.length).setValues([newHeaders]);
      
      // Format new header cells
      const newHeaderRange = historySheet.getRange(1, currentColCount + 1, 1, newHeaders.length);
      newHeaderRange.setBackground('#4a86e8');
      newHeaderRange.setFontColor('#ffffff');
      newHeaderRange.setFontWeight('bold');
      newHeaderRange.setHorizontalAlignment('center');
      
      Logger.log(`✅ Added columns: ${newHeaders.join(', ')}`);
    }
    
    // Collect all ticket IDs
    const ticketIds = [];
    const rowMap = new Map(); // ticketId -> row index
    const existingOverridesMap = new Map(); // ticketId -> {outcomeOverride, overrideReason}
    
    for (let i = 1; i < data.length; i++) {
      const ticketId = data[i][colMap['Ticket ID']];
      if (ticketId) {
        ticketIds.push(ticketId);
        rowMap.set(ticketId, i + 1); // 1-indexed row number
        
        // Preserve existing overrides
        const outcomeOverrideIdx = colMap['Outcome Override'];
        const overrideReasonIdx = colMap['Override Reason'];
        if (outcomeOverrideIdx !== undefined) {
          existingOverridesMap.set(ticketId, {
            outcomeOverride: data[i][outcomeOverrideIdx] || "",
            overrideReason: data[i][overrideReasonIdx] || ""
          });
        }
      }
    }
    
    Logger.log(`📋 Found ${ticketIds.length} tickets to backfill`);
    
    // Fetch tickets with conversations in batches
    const freshTickets = fetchTicketsForBackfill(ticketIds);
    
    Logger.log(`📥 Retrieved ${freshTickets.length} tickets from API`);
    
    // Enrich tickets with preserved overrides
    const enrichedTickets = enrichTickets(freshTickets, existingOverridesMap);
    
    // Update rows in sheet
    let updatedCount = 0;
    for (const ticket of enrichedTickets) {
      const rowNum = rowMap.get(ticket.id);
      if (rowNum) {
        const newRow = ticketToRow(ticket);
        historySheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
        updatedCount++;
        
        // Progress logging every 50 tickets
        if (updatedCount % 50 === 0) {
          Logger.log(`📊 Progress: ${updatedCount}/${enrichedTickets.length} tickets updated`);
        }
      }
    }
    
    // Apply override highlighting
    applyOverrideHighlighting(historySheet);
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n✅ === BACKFILL COMPLETE ===');
    Logger.log(`   Total tickets: ${ticketIds.length}`);
    Logger.log(`   Retrieved: ${freshTickets.length}`);
    Logger.log(`   Updated: ${updatedCount}`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    
    return {
      success: true,
      total: ticketIds.length,
      retrieved: freshTickets.length,
      updated: updatedCount,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`❌ Backfill failed: ${e.message}`);
    Logger.log(e.stack);
    return {
      success: false,
      error: e.message,
      duration: (new Date() - startTime) / 1000
    };
  }
}

/**
 * Load existing override data from history sheet
 * Used to preserve manual overrides during updates
 * 
 * @param {Sheet} historySheet - MVR_Ticket_History sheet
 * @return {Map} Map of ticketId -> {outcomeOverride, overrideReason}
 */
function loadExistingOverrides(historySheet) {
  const overrides = new Map();
  
  const data = historySheet.getDataRange().getValues();
  if (data.length <= 1) return overrides;
  
  const headers = data[0];
  const colMap = createColumnMapper(headers);
  
  const ticketIdIdx = colMap['Ticket ID'];
  const overrideIdx = colMap['Outcome Override'];
  const reasonIdx = colMap['Override Reason'];
  
  if (overrideIdx === undefined) return overrides;
  
  for (let i = 1; i < data.length; i++) {
    const ticketId = data[i][ticketIdIdx];
    const override = data[i][overrideIdx];
    
    if (ticketId && override) {
      overrides.set(ticketId, {
        outcomeOverride: override,
        overrideReason: data[i][reasonIdx] || ""
      });
    }
  }
  
  Logger.log(`📋 Loaded ${overrides.size} existing overrides`);
  return overrides;
}

// ═══════════════════════════════════════════════════════════════════════════════
// OVERRIDE HIGHLIGHTING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Apply yellow highlighting to rows with manual overrides
 * 
 * @param {Sheet} sheet - Sheet to apply highlighting to
 */
function applyOverrideHighlighting(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    const overrideIdx = colMap['Outcome Override'];
    
    if (overrideIdx === undefined) {
      Logger.log('⚠️ Outcome Override column not found');
      return;
    }
    
    // Create conditional formatting rule for override column
    const range = sheet.getRange(2, 1, data.length - 1, headers.length);
    
    // Build conditional format rule
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${columnToLetter(overrideIdx + 1)}2<>""`)
      .setBackground('#FFEB3B') // Yellow
      .setRanges([range])
      .build();
    
    // Get existing rules and add new one
    const rules = sheet.getConditionalFormatRules();
    
    // Remove existing override highlighting rules (to avoid duplicates)
    const filteredRules = rules.filter(r => {
      const bg = r.getBooleanCondition()?.getBackground();
      return bg !== '#FFEB3B' && bg !== '#ffeb3b';
    });
    
    filteredRules.push(rule);
    sheet.setConditionalFormatRules(filteredRules);
    
    Logger.log('🎨 Applied override highlighting');
    
  } catch (e) {
    Logger.log(`⚠️ Error applying highlighting: ${e.message}`);
  }
}

/**
 * Convert column index to letter (A, B, C, ..., AA, AB, etc.)
 * @param {number} column - 1-indexed column number
 * @return {string} Column letter
 */
function columnToLetter(column) {
  let letter = '';
  while (column > 0) {
    const temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

// ═══════════════════════════════════════════════════════════════════════════════
// STATISTICS & VALIDATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get statistics about current history
 * @return {Object} Statistics object
 */
function getHistoryStats() {
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      return { error: 'History sheet not found' };
    }
    
    const data = historySheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { total: 0 };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Count by type, status, partner, vendor group, outcome
    const stats = {
      total: data.length - 1,
      by_type: {},
      by_status: {},
      by_partner: {},
      by_vendor_group: {},
      by_outcome: {},
      overrides_count: 0,
      oldest_created: null,
      newest_created: null
    };
    
    for (let i = 1; i < data.length; i++) {
      const type = data[i][colMap['Request Type (SC/RC)']];
      const status = data[i][colMap['Status']];
      const partner = data[i][colMap['Partner Name']];
      const created = data[i][colMap['Date Created']];
      const vendorGroup = data[i][colMap['Vendor Group']];
      const outcome = data[i][colMap['MVR Outcome']];
      const override = data[i][colMap['Outcome Override']];
      
      // Count by type
      if (type) stats.by_type[type] = (stats.by_type[type] || 0) + 1;
      
      // Count by status
      if (status) stats.by_status[status] = (stats.by_status[status] || 0) + 1;
      
      // Count by partner
      if (partner) stats.by_partner[partner] = (stats.by_partner[partner] || 0) + 1;
      
      // Count by vendor group
      if (vendorGroup) stats.by_vendor_group[vendorGroup] = (stats.by_vendor_group[vendorGroup] || 0) + 1;
      
      // Count by outcome
      if (outcome) stats.by_outcome[outcome] = (stats.by_outcome[outcome] || 0) + 1;
      
      // Count overrides
      if (override) stats.overrides_count++;
      
      // Track date range
      if (created) {
        const createdDate = new Date(created);
        if (!stats.oldest_created || createdDate < stats.oldest_created) {
          stats.oldest_created = createdDate;
        }
        if (!stats.newest_created || createdDate > stats.newest_created) {
          stats.newest_created = createdDate;
        }
      }
    }
    
    Logger.log('\n📊 === HISTORY STATISTICS ===');
    Logger.log(`   Total tickets: ${stats.total}`);
    Logger.log(`   By type: ${JSON.stringify(stats.by_type)}`);
    Logger.log(`   By status: ${JSON.stringify(stats.by_status)}`);
    Logger.log(`   By vendor group: ${JSON.stringify(stats.by_vendor_group)}`);
    Logger.log(`   By outcome: ${JSON.stringify(stats.by_outcome)}`);
    Logger.log(`   Overrides: ${stats.overrides_count}`);
    Logger.log(`   Date range: ${stats.oldest_created?.toLocaleDateString()} to ${stats.newest_created?.toLocaleDateString()}`);
    Logger.log(`   Unique partners: ${Object.keys(stats.by_partner).length}`);
    
    return stats;
    
  } catch (e) {
    Logger.log(`❌ Error getting history stats: ${e.message}`);
    return { error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// STATE EXTRACTION UTILITIES
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Extract state from text using multiple patterns
 * Tries various formats found in ticket data
 * Used as fallback when subject/description parsing doesn't find state
 * @param {string} text - Text to search for state
 * @returns {string|null} - 2-letter state code or null
 */
function extractStateFromText(text) {
  if (!text) return null;
  
  // Patterns to try, in order of specificity
  const patterns = [
    /DL State:\s*\*?([A-Z]{2})\*?/i,                    // DL State: *TN* or DL State: TN
    /State:\s*\*?([A-Z]{2})\*?/i,                        // State: TN
    /Driver[''']?s?\s+License\s+State:\s*([A-Z]{2})/i,   // Driver's License State: TN
    /\bState\b[:\s]+([A-Z]{2})\b/i,                      // State: TN or State TN
    /License\s+State:\s*([A-Z]{2})/i,                    // License State: TN
    /\|\s*([A-Z]{2})\s*\|/,                              // | TN | (table format)
    /State Vendor:\s*[^|]+\|\s*([A-Z]{2})\b/i,           // State Vendor: xxx | TN
    /MVR.*\b([A-Z]{2})\b.*Driver/i,                      // MVR ... TN ... Driver
    /\bDL[:\s]+[A-Z0-9]+[,\s]+([A-Z]{2})\b/i,           // DL: ABC123, TN
    /Driver License[:\s]+[A-Z0-9]+[,\s]+([A-Z]{2})\b/i, // Driver License: ABC123, TN
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const state = match[1].toUpperCase();
      // Validate it's a real US state code
      if (isValidStateCode(state)) {
        return state;
      }
    }
  }
  
  return null;
}

/**
 * Check if a 2-letter code is a valid US state
 * @param {string} code - 2-letter code to check
 * @returns {boolean} - True if valid state code
 */
function isValidStateCode(code) {
  const validStates = [
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA',
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
  ];
  return validStates.includes(code);
}
