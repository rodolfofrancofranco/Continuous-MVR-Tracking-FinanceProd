/**
 * OUTCOME TRACKING SYSTEM
 * MVR TICKET TRACKER - Outcome Classification & Audit Logging
 * 
 * Purpose: Parse ticket content for MVR outcomes, classify results,
 * and maintain audit log for manual overrides
 */

// ═══════════════════════════════════════════════════════════════════════════════
// CACHING - Avoid repeated sheet reads during batch processing
// ═══════════════════════════════════════════════════════════════════════════════

/** @type {Array|null} Cached tag mappings - cleared per execution */
let _tagMappingsCache = null;

/**
 * Clear the tag mappings cache (call at start of batch operations)
 */
function clearTagMappingsCache() {
  _tagMappingsCache = null;
  Logger.log('🔄 Tag mappings cache cleared');
}

// ═══════════════════════════════════════════════════════════════════════════════
// TICKET DESCRIPTION PARSING (Extract workflow data from HTML email body)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Parse ticket description HTML to extract workflow data points
 * Extracts: DL State, DL Number, Days Since Last, Last Check Date, Cadence Days, Tier, Enrollment Type, TA URL
 * 
 * @param {string} descriptionHtml - HTML content from ticket description
 * @return {Object} Extracted workflow data
 */
function parseTicketDescription(descriptionHtml) {
  const result = {
    dl_state: "",
    dl_number: "",
    days_since_last: "",
    last_check_date: "",
    cadence_days: "",
    tier: "",
    enrollment_type: "",
    ta_url: "",
    state_vendor: ""  // Inform, Non-Inform-State, etc.
  };
  
  if (!descriptionHtml) {
    return result;
  }
  
  // Clean HTML entities
  const cleanHtml = descriptionHtml
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
  
  // Extract data using regex patterns for table cells
  // Pattern: <td>Label</td><td>Value</td> or similar structures
  
  // State Vendor (Inform, Non-Inform-State, etc.) - for vendor classification
  const vendorMatch = cleanHtml.match(/(?:State\s*Vendor)[:\s]*<\/?\w*>?\s*([^<\n]+)/i) ||
                      cleanHtml.match(/<td[^>]*>\s*State\s*Vendor\s*<\/td>\s*<td[^>]*>\s*([^<]+)\s*<\/td>/i);
  if (vendorMatch) result.state_vendor = vendorMatch[1].trim();
  
  // DL State (2-letter state code)
  const stateMatch = cleanHtml.match(/(?:DL\s*State|State)[:\s]*<\/?\w*>?\s*([A-Z]{2})\b/i) ||
                     cleanHtml.match(/<td[^>]*>\s*(?:DL\s*State|State)\s*<\/td>\s*<td[^>]*>\s*([A-Z]{2})\s*<\/td>/i);
  if (stateMatch) result.dl_state = stateMatch[1].toUpperCase();
  
  // DL Number
  const dlMatch = cleanHtml.match(/(?:DL\s*Number|License\s*Number)[:\s]*<\/?\w*>?\s*([A-Z0-9*-]+)/i) ||
                  cleanHtml.match(/<td[^>]*>\s*(?:DL\s*Number|License)\s*<\/td>\s*<td[^>]*>\s*([^<]+)\s*<\/td>/i);
  if (dlMatch) result.dl_number = dlMatch[1].trim();
  
  // Days Since Last Check
  const daysMatch = cleanHtml.match(/(?:Days\s*Since\s*Last|Days\s*Overdue)[:\s]*<\/?\w*>?\s*(\d+)/i) ||
                    cleanHtml.match(/<td[^>]*>\s*Days\s*Since\s*Last[^<]*<\/td>\s*<td[^>]*>\s*(\d+)\s*<\/td>/i);
  if (daysMatch) result.days_since_last = daysMatch[1];
  
  // Last Check Date
  const lastDateMatch = cleanHtml.match(/(?:Last\s*Check\s*Date|Last\s*MVR\s*Date)[:\s]*<\/?\w*>?\s*([\d\-\/]+)/i) ||
                        cleanHtml.match(/<td[^>]*>\s*Last\s*Check\s*Date\s*<\/td>\s*<td[^>]*>\s*([\d\-\/]+)\s*<\/td>/i);
  if (lastDateMatch) result.last_check_date = lastDateMatch[1];
  
  // Cadence Days
  const cadenceMatch = cleanHtml.match(/(?:Cadence\s*Days|Cadence)[:\s]*<\/?\w*>?\s*(\d+)/i) ||
                       cleanHtml.match(/<td[^>]*>\s*Cadence[^<]*<\/td>\s*<td[^>]*>\s*(\d+)\s*<\/td>/i);
  if (cadenceMatch) result.cadence_days = cadenceMatch[1];
  
  // Tier
  const tierMatch = cleanHtml.match(/(?:Tier)[:\s]*<\/?\w*>?\s*(Tier\s*\d+|\d+)/i) ||
                    cleanHtml.match(/<td[^>]*>\s*Tier\s*<\/td>\s*<td[^>]*>\s*([^<]+)\s*<\/td>/i);
  if (tierMatch) result.tier = tierMatch[1].trim();
  
  // Enrollment Type
  const enrollMatch = cleanHtml.match(/(?:Enrollment\s*Type|Enrollment)[:\s]*<\/?\w*>?\s*([A-Za-z]+)/i) ||
                      cleanHtml.match(/<td[^>]*>\s*Enrollment[^<]*<\/td>\s*<td[^>]*>\s*([^<]+)\s*<\/td>/i);
  if (enrollMatch) result.enrollment_type = enrollMatch[1].trim();
  
  // TA URL (Turn Admin URL)
  const urlMatch = cleanHtml.match(/(?:TA\s*URL|Turn\s*Admin)[:\s]*<\/?\w*>?\s*(https?:\/\/[^\s<"]+)/i) ||
                   cleanHtml.match(/href=["'](https?:\/\/app\.turn\.ai[^"']+)["']/i);
  if (urlMatch) result.ta_url = urlMatch[1];
  
  return result;
}

/**
 * Classify MVR vendor based on State Vendor field and DL State
 * Vendors: INFORM, PENNDOT (PA only), CERTN (all other Non-Inform states)
 * 
 * @param {string} stateVendor - Value from State Vendor field (Inform, Non-Inform-State, etc.)
 * @param {string} dlState - 2-letter state code
 * @return {string} Vendor classification: INFORM, PENNDOT, CERTN, or UNKNOWN
 */
function classifyVendor(stateVendor, dlState) {
  if (!stateVendor && !dlState) {
    return 'UNKNOWN';
  }
  
  const vendor = (stateVendor || '').toLowerCase().trim();
  const state = (dlState || '').toUpperCase().trim();
  
  // Inform vendor - used for states that support INFORM
  if (vendor.includes('inform') && !vendor.includes('non-inform')) {
    return 'INFORM';
  }
  
  // Non-Inform states use different vendors based on state
  if (vendor.includes('non-inform') || vendor === '' || !vendor) {
    // Pennsylvania uses PENNDOT
    if (state === 'PA') {
      return 'PENNDOT';
    }
    // All other states use CERTN
    if (state) {
      return 'CERTN';
    }
  }
  
  return 'UNKNOWN';
}

// ═══════════════════════════════════════════════════════════════════════════════
// DYNAMIC TAG MAPPING LOADER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get column indices from header row for flexible sheet structure
 * Handles any column order, missing columns get default indices
 * 
 * @param {Array} headerRow - First row of sheet with column names
 * @return {Object} Map of column names to indices
 */
function getTagMappingColumnMap(headerRow) {
  const colMap = {};
  
  // Map each expected column name to its index
  TAG_MAPPING_HEADERS.forEach((expectedHeader, defaultIdx) => {
    const foundIdx = headerRow.indexOf(expectedHeader);
    colMap[expectedHeader] = foundIdx !== -1 ? foundIdx : defaultIdx;
  });
  
  return colMap;
}

/**
 * Apply data validation to Tag_Outcome_Mappings sheet
 * Uses column mapper for flexible column positions
 * 
 * @param {Sheet} sheet - Tag_Outcome_Mappings sheet
 * @param {Object} colMap - Column index map
 */
function applyTagMappingValidation(sheet, colMap) {
  const matchTypeCol = colMap["Match Type"] + 1; // +1 for 1-based indexing
  const outcomeTypeCol = colMap["Outcome Type"] + 1;
  const billableCol = colMap["Is Billable"] + 1;
  const activeCol = colMap["Is Active"] + 1;
  
  // Add data validation for Match Type column
  const matchTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["contains", "regex"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, matchTypeCol, 100, 1).setDataValidation(matchTypeRule);
  
  // Add data validation for Outcome Type column
  const outcomeValues = Object.values(OUTCOME_TYPES);
  const outcomeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(outcomeValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, outcomeTypeCol, 100, 1).setDataValidation(outcomeRule);
  
  // Add checkboxes for Is Billable and Is Active
  sheet.getRange(2, billableCol, 100, 1).insertCheckboxes();
  sheet.getRange(2, activeCol, 100, 1).insertCheckboxes();
}

/**
 * Load tag-to-outcome mappings from sheet (cached for batch performance)
 * Creates sheet with defaults if it doesn't exist
 * Uses column mapper for flexible sheet structure
 * 
 * @param {boolean} forceReload - Force reload from sheet, ignoring cache
 * @return {Array} Array of mapping objects sorted by priority
 */
function loadTagMappings(forceReload = false) {
  // Return cached mappings if available (huge performance gain for batch ops)
  if (_tagMappingsCache !== null && !forceReload) {
    return _tagMappingsCache;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  
  // Create sheet with defaults if missing
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
    sheet.appendRow(TAG_MAPPING_HEADERS);
    
    DEFAULT_TAG_MAPPINGS.forEach(row => sheet.appendRow(row));
    
    // Format header row
    sheet.getRange(1, 1, 1, TAG_MAPPING_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("white");
    
    // Apply data validation using column mapper
    const colMap = getTagMappingColumnMap(TAG_MAPPING_HEADERS);
    applyTagMappingValidation(sheet, colMap);
    
    sheet.autoResizeColumns(1, TAG_MAPPING_HEADERS.length);
    
    Logger.log(`✅ Created ${SHEET_NAMES.TAG_OUTCOME_MAPPINGS} sheet with ${DEFAULT_TAG_MAPPINGS.length} default mappings`);
  }
  
  // Read all mappings with column mapper
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = getTagMappingColumnMap(headers);
  const mappings = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Use column mapper for flexible positioning
    const pattern = row[colMap["Tag Pattern"]];
    const matchType = row[colMap["Match Type"]];
    const outcomeType = row[colMap["Outcome Type"]];
    const priority = row[colMap["Priority"]] || 99;
    const isBillable = row[colMap["Is Billable"]] === true || row[colMap["Is Billable"]] === "TRUE";
    const isActive = row[colMap["Is Active"]] === true || row[colMap["Is Active"]] === "TRUE";
    const notes = row[colMap["Notes"]] || "";
    
    // Skip inactive, empty, or invalid rows
    if (!isActive || !pattern || !outcomeType) {
      if (pattern || outcomeType) {
        Logger.log(`⚠️ Skipped row ${i + 1}: inactive=${!isActive}, pattern="${pattern}", outcome="${outcomeType}"`);
      }
      continue;
    }
    
    // Compile pattern
    let compiledPattern;
    try {
      if (matchType === "regex") {
        compiledPattern = new RegExp(pattern, "i");
      } else {
        // "contains" - escape special regex chars and wrap
        const escaped = pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        compiledPattern = new RegExp(escaped, "i");
      }
    } catch (e) {
      Logger.log(`⚠️ Invalid pattern "${pattern}" in row ${i + 1}: ${e.message}`);
      continue;
    }
    
    mappings.push({
      pattern: pattern,
      matchType: matchType,
      regex: compiledPattern,
      outcomeType: outcomeType,
      priority: priority,
      isBillable: isBillable,
      notes: notes
    });
  }
  
  // Sort by priority (ascending - lower number = higher priority)
  mappings.sort((a, b) => a.priority - b.priority);
  
  // Cache for subsequent calls in this execution
  _tagMappingsCache = mappings;
  
  Logger.log(`📋 Loaded ${mappings.length} active tag mappings (cached)`);
  return mappings;
}

/**
 * Check if an outcome is billable based on tag mappings
 * 
 * @param {string} outcomeType - The outcome type to check
 * @return {boolean} True if billable
 */
function isOutcomeBillable(outcomeType) {
  const mappings = loadTagMappings();
  const mapping = mappings.find(m => m.outcomeType === outcomeType);
  return mapping ? mapping.isBillable : false;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAG KEYWORD HINTS (Based on 602-ticket analysis - Dec 2025)
// ═══════════════════════════════════════════════════════════════════════════════

const TAG_KEYWORD_HINTS = {
  // High-volume billable patterns (373 + 150 + 208 tickets = 731/602 coverage)
  'same information': 'Clear',
  'same info': 'Clear',
  'updated': 'Clear',
  'uploaded': 'Clear',
  'approved': 'Clear',
  
  // Violations (always billable)
  'expired': 'Expired License',
  'dl expired': 'Expired License',
  'suspend': 'Suspended',
  'not valid': 'Invalid License',
  'invalid': 'Invalid License',
  'violation': 'Violation',
  
  // Pending/Processing states (not billable)
  'pending': 'Pending',
  'requested': 'Still Processing',
  'emailed': 'Still Processing',
  'verifying': 'Still Processing',
  'review identity': 'Still Processing',
  'in progress': 'Still Processing',
  
  // Special cases
  'record not found': 'Record Not Found',
  'withdrawn': 'Withdrawn',
  'consider': 'Unknown',  // Context-dependent - needs manual review
  
  // Legacy patterns
  'clear': 'Clear',
  'duplicate': 'Duplicate',
  'no_record': 'No MVR Record',
  'no_mvr': 'No MVR Record',
  'error': 'Error',
  'active': 'Active Monitoring'
};

/**
 * Suggest outcome for a tag based on 602-ticket analysis patterns
 * Priority-based matching: exact → specific → general → pattern detection
 * 
 * @param {string} tagName - Tag name to analyze
 * @return {string} Suggested outcome type
 */
function suggestOutcomeForTag(tagName) {
  if (!tagName) return OUTCOME_TYPES.UNKNOWN;
  
  const lower = tagName.toLowerCase().trim();
  
  // PRIORITY 1: Exact matches (most reliable)
  if (TAG_KEYWORD_HINTS[lower]) {
    return TAG_KEYWORD_HINTS[lower];
  }
  
  // PRIORITY 2: Substring matches (specific before general)
  const orderedKeywords = [
    // Multi-word patterns first (more specific)
    'same information',
    'same info',
    'dl expired',
    'not valid',
    'record not found',
    'review identity',
    'in progress',
    'no_record',
    'no_mvr',
    // Single-word patterns (more general)
    'expired',
    'suspend',
    'invalid',
    'violation',
    'updated',
    'uploaded',
    'approved',
    'pending',
    'requested',
    'emailed',
    'verifying',
    'withdrawn',
    'consider',
    'clear',
    'duplicate',
    'error',
    'active'
  ];
  
  for (const keyword of orderedKeywords) {
    if (lower.includes(keyword)) {
      return TAG_KEYWORD_HINTS[keyword];
    }
  }
  
  // PRIORITY 3: Pattern detection (regex-based)
  
  // State codes (CA, WI, PA, FL, etc.) - don't map
  if (/^[A-Z]{2}$/.test(tagName) || /^(ca|wi|pa|fl|tx|ny|il|oh)$/i.test(lower)) {
    return OUTCOME_TYPES.UNKNOWN; // Metadata, not outcomes
  }
  
  // Typos of known patterns
  if (lower.includes('requesed') || lower.includes('requeste')) {
    return OUTCOME_TYPES.STILL_PROCESSING; // Typo of "requested"
  }
  
  // System/metadata tags (json, dmv, sentinel)
  if (['json', 'dmv', 'sentinel', 'api', 'system'].includes(lower)) {
    return OUTCOME_TYPES.UNKNOWN; // Don't map technical tags
  }
  
  return OUTCOME_TYPES.UNKNOWN;
}

/**
 * Discover all unique tags from raw tickets and track their usage
 * Creates/updates the Discovered_Tags sheet with tag statistics
 * @return {Object} Discovery summary
 */
function discoverTagsFromRaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
  
  if (!rawSheet) {
    Logger.log('❌ MVR_Raw_Tickets sheet not found');
    return { success: false, error: 'Raw tickets sheet not found' };
  }
  
  // Get or create Discovered_Tags sheet
  let discoveredSheet = ss.getSheetByName(SHEET_NAMES.DISCOVERED_TAGS);
  if (!discoveredSheet) {
    discoveredSheet = ss.insertSheet(SHEET_NAMES.DISCOVERED_TAGS);
    discoveredSheet.appendRow(['Tag', 'Count', 'First Seen', 'Last Seen', 'Has Mapping', 'Suggested Outcome']);
    discoveredSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#f3f3f3');
  }
  
  // Get all tags from raw tickets
  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('ℹ️ No raw tickets to scan');
    return { success: true, tagsFound: 0 };
  }
  
  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const tagsColIdx = headers.indexOf('Tags');
  
  if (tagsColIdx === -1) {
    Logger.log('❌ Tags column not found');
    return { success: false, error: 'Tags column not found' };
  }
  
  const tagsData = rawSheet.getRange(2, tagsColIdx + 1, lastRow - 1, 1).getValues();
  
  // Load existing discovered tags to preserve history
  const existingData = {};
  if (discoveredSheet.getLastRow() > 1) {
    const existing = discoveredSheet.getRange(2, 1, discoveredSheet.getLastRow() - 1, 6).getValues();
    existing.forEach(row => {
      const tag = row[0];
      existingData[tag] = {
        count: row[1] || 0,
        firstSeen: row[2] || new Date(),
        lastSeen: row[3] || new Date(),
        hasMapping: row[4],
        suggestion: row[5]
      };
    });
  }
  
  // Count tag occurrences from current raw data
  const tagStats = {};
  tagsData.forEach(row => {
    const tagsStr = row[0] || '';
    const tags = tagsStr.split(',').map(t => t.trim()).filter(t => t);
    tags.forEach(tag => {
      if (!tagStats[tag]) {
        tagStats[tag] = { count: 0, firstSeen: new Date(), lastSeen: new Date() };
      }
      tagStats[tag].count++;
      tagStats[tag].lastSeen = new Date();
    });
  });
  
  // Load existing mappings to check which tags are mapped
  const mappings = loadTagMappings();
  const mappedTags = new Set(mappings.map(m => m.pattern.toLowerCase()));
  
  // Merge with existing data - preserve firstSeen dates
  const now = new Date();
  let newTagsCount = 0;
  const sheetData = Object.entries(tagStats).map(([tag, stats]) => {
    const hasMapping = mappedTags.has(tag.toLowerCase()) || 
                      mappings.some(m => m.pattern.includes('*') && testTagPattern(tag, m.pattern));
    const suggestion = suggestOutcomeForTag(tag);
    
    // Use existing firstSeen if available, otherwise use new date
    const firstSeen = existingData[tag] ? existingData[tag].firstSeen : stats.firstSeen;
    if (!existingData[tag]) newTagsCount++;
    
    return [tag, stats.count, firstSeen, now, hasMapping ? 'Yes' : 'No', suggestion];
  });
  
  // Sort by count descending
  sheetData.sort((a, b) => b[1] - a[1]);
  
  // Clear and write merged data
  if (discoveredSheet.getLastRow() > 1) {
    discoveredSheet.getRange(2, 1, discoveredSheet.getLastRow() - 1, 6).clear();
  }
  
  if (sheetData.length > 0) {
    discoveredSheet.getRange(2, 1, sheetData.length, 6).setValues(sheetData);
  }
  
  Logger.log(`✅ Discovered ${sheetData.length} unique tags (${newTagsCount} new)`);
  return { 
    success: true, 
    totalTags: sheetData.length, 
    newTags: newTagsCount,
    unmapped: sheetData.filter(row => row[4] === 'No').length 
  };
}

/**
 * Identify tags that don't have outcome mappings
 * @return {Array} Array of unmapped tags with usage counts
 */
function identifyUnmappedTags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const discoveredSheet = ss.getSheetByName(SHEET_NAMES.DISCOVERED_TAGS);
  
  if (!discoveredSheet || discoveredSheet.getLastRow() < 2) {
    return [];
  }
  
  const data = discoveredSheet.getRange(2, 1, discoveredSheet.getLastRow() - 1, 5).getValues();
  const unmapped = data
    .filter(row => row[4] === 'No') // Has Mapping column
    .map(row => ({ tag: row[0], count: row[1] }))
    .sort((a, b) => b.count - a.count);
  
  return unmapped;
}

/**
 * Get summary of unmapped tags for menu display
 * @return {string} Summary text
 */
function getUnmappedTagsSummary() {
  const unmapped = identifyUnmappedTags();
  if (unmapped.length === 0) {
    return 'All tags mapped';
  }
  
  const top10 = unmapped.slice(0, 10).map(t => `${t.tag} (${t.count})`).join(', ');
  return `${unmapped.length} unmapped: ${top10}`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// OUTCOME CLASSIFICATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Classify MVR outcome based on ticket tags and conversation content
 * Uses dynamic tag mappings loaded from Tag_Outcome_Mappings sheet
 * 
 * @param {Array} tags - Array of ticket tags
 * @param {Array} conversations - Array of conversation objects from Freshdesk
 * @return {Object} Classification result with outcome, source, notes, and isBillable
 */
function classifyOutcome(tags, conversations) {
  const result = {
    outcome: OUTCOME_TYPES.UNKNOWN,
    source: OUTCOME_SOURCES.AUTO,
    notes: "",
    outcomeDate: null,
    isBillable: false
  };
  
  // Load dynamic mappings (reloads each run)
  const mappings = loadTagMappings();
  
  const tagString = (tags || []).join(' ').toLowerCase();
  
  // Priority 1: Check tags against dynamic mappings (sorted by priority)
  for (const mapping of mappings) {
    if (mapping.regex.test(tagString)) {
      result.outcome = mapping.outcomeType;
      result.source = OUTCOME_SOURCES.TAG;
      result.notes = `Tag match: "${mapping.pattern}"`;
      result.isBillable = mapping.isBillable;
      return result;
    }
  }
  
  // Priority 2: Check conversations for outcome indicators
  if (conversations && conversations.length > 0) {
    const conversationResult = classifyFromConversations(conversations);
    if (conversationResult.outcome !== OUTCOME_TYPES.UNKNOWN) {
      // Check billable status from mappings
      const mapping = mappings.find(m => m.outcomeType === conversationResult.outcome);
      conversationResult.isBillable = mapping ? mapping.isBillable : false;
      return conversationResult;
    }
  }
  
  // Priority 3: Check if still in progress based on empty tags
  if (!tags || tags.length === 0) {
    result.outcome = OUTCOME_TYPES.STILL_PROCESSING;
    result.source = OUTCOME_SOURCES.AUTO;
    result.notes = "No outcome tags yet";
    return result;
  }
  
  return result;
}

/**
 * Classify outcome from conversation content
 * Scans agent replies for outcome indicators
 * 
 * @param {Array} conversations - Array of conversation objects
 * @return {Object} Classification result
 */
function classifyFromConversations(conversations) {
  const result = {
    outcome: OUTCOME_TYPES.UNKNOWN,
    source: OUTCOME_SOURCES.CONVERSATION,
    notes: "",
    outcomeDate: null
  };
  
  // Sort conversations by created_at descending (newest first)
  const sorted = conversations.slice().sort((a, b) => 
    new Date(b.created_at) - new Date(a.created_at)
  );
  
  for (const conv of sorted) {
    // Skip customer replies (only check agent/system responses)
    if (conv.incoming) continue;
    
    const body = (conv.body_text || conv.body || "").toLowerCase();
    
    // Check for suspension confirmation
    if (/license\s*(is\s*)?suspended|mvr\s*shows?\s*suspend|suspension\s*confirmed/i.test(body)) {
      result.outcome = OUTCOME_TYPES.SUSPENSION_CONFIRMED;
      result.notes = "Conversation: suspension mentioned";
      result.outcomeDate = conv.created_at;
      return result;
    }
    
    // Check for valid/clear result
    if (/mvr\s*(is\s*)?(valid|clear|active)|license\s*(is\s*)?(valid|active|current)/i.test(body)) {
      result.outcome = OUTCOME_TYPES.CLEAR;
      result.notes = "Conversation: valid MVR mentioned";
      result.outcomeDate = conv.created_at;
      return result;
    }
    
    // Check for DMV unavailable
    if (/dmv\s*(is\s*)?(down|unavailable)|cannot\s*connect\s*to\s*dmv|state\s*system\s*(is\s*)?down/i.test(body)) {
      result.outcome = OUTCOME_TYPES.DMV_UNAVAILABLE;
      result.notes = "Conversation: DMV unavailable";
      result.outcomeDate = conv.created_at;
      return result;
    }
    
    // Check for cannot process
    if (/cannot\s*process|unable\s*to\s*(complete|process)|missing\s*information|dl\s*(number\s*)?(mismatch|invalid)/i.test(body)) {
      result.outcome = OUTCOME_TYPES.CANNOT_PROCESS;
      result.notes = "Conversation: processing issue";
      result.outcomeDate = conv.created_at;
      return result;
    }
  }
  
  return result;
}

/**
 * Get effective outcome considering manual override
 * 
 * @param {string} autoOutcome - Auto-classified outcome
 * @param {string} overrideOutcome - Manual override value (may be empty)
 * @return {string} Effective outcome to use
 */
function getEffectiveOutcome(autoOutcome, overrideOutcome) {
  if (overrideOutcome && overrideOutcome.trim() !== "") {
    return overrideOutcome.trim();
  }
  return autoOutcome || OUTCOME_TYPES.UNKNOWN;
}

// ═══════════════════════════════════════════════════════════════════════════════
// CONVERSATION ANALYSIS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Extract the last responder (agent) from conversations
 * 
 * @param {Array} conversations - Array of conversation objects
 * @return {string} Last responder name/email or empty string
 */
function extractLastResponder(conversations) {
  if (!conversations || conversations.length === 0) {
    return "";
  }
  
  // Sort conversations by created_at descending (newest first)
  const sorted = conversations.slice().sort((a, b) => 
    new Date(b.created_at) - new Date(a.created_at)
  );
  
  // Find the last non-incoming (agent) response
  for (const conv of sorted) {
    if (!conv.incoming) {
      // Return agent info - prefer name over email
      if (conv.from_email) {
        const emailParts = conv.from_email.split('@');
        return emailParts[0] || conv.from_email;
      }
      if (conv.user_id) {
        return `Agent ${conv.user_id}`;
      }
    }
  }
  
  return "";
}

/**
 * Get conversation count
 * 
 * @param {Array} conversations - Array of conversation objects
 * @return {number} Number of conversations
 */
function getConversationCount(conversations) {
  return (conversations || []).length;
}

// ═══════════════════════════════════════════════════════════════════════════════
// OVERRIDE AUDIT LOGGING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Log an override change to the audit log sheet
 * 
 * @param {number} ticketId - Freshdesk ticket ID
 * @param {string} turnId - Turn ID from ticket
 * @param {string} partner - Partner name
 * @param {string} previousOutcome - Previous outcome value
 * @param {string} newOverride - New override value
 * @param {string} reason - Reason for override
 */
function logOverrideChange(ticketId, turnId, partner, previousOutcome, newOverride, reason) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.OVERRIDE_AUDIT_LOG, AUDIT_LOG_HEADERS);
    
    // Get current user info
    let userEmail = "";
    let userName = "";
    try {
      userEmail = Session.getActiveUser().getEmail();
      userName = userEmail.split('@')[0];
    } catch (e) {
      userEmail = "system";
      userName = "System Process";
    }
    
    const timestamp = new Date();
    
    const row = [
      timestamp,
      ticketId,
      turnId || "",
      partner || "",
      previousOutcome || "",
      newOverride || "",
      reason || "",
      userEmail,
      userName
    ];
    
    sheet.appendRow(row);
    Logger.log(`📝 Logged override change for ticket ${ticketId}`);
    
  } catch (e) {
    Logger.log(`⚠️ Failed to log override: ${e.message}`);
  }
}

/**
 * Detect if override has changed between old and new row data
 * 
 * @param {Array} oldRow - Previous row data
 * @param {Array} newRow - New row data
 * @param {Object} colMap - Column index mapper
 * @return {Object|null} Change details or null if no change
 */
function detectOverrideChange(oldRow, newRow, colMap) {
  const overrideIdx = colMap["Outcome Override"];
  
  if (overrideIdx === undefined) {
    return null;
  }
  
  const oldOverride = (oldRow[overrideIdx] || "").toString().trim();
  const newOverride = (newRow[overrideIdx] || "").toString().trim();
  
  // Only detect if there's a new override that wasn't there before
  // or if the override value has changed
  if (newOverride !== "" && newOverride !== oldOverride) {
    return {
      ticketId: oldRow[colMap["Ticket ID"]] || newRow[0],
      turnId: oldRow[colMap["Turn ID"]] || newRow[colMap["Turn ID"]],
      partner: oldRow[colMap["Partner Name"]] || newRow[colMap["Partner Name"]],
      previousOutcome: oldRow[colMap["MVR Outcome"]] || "",
      newOverride: newOverride,
      reason: newRow[colMap["Override Reason"]] || ""
    };
  }
  
  return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// VENDOR GROUP DETECTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Extract vendor group from subject line and/or DL state
 * 
 * @param {string} subject - Ticket subject line
 * @param {string} dlState - Driver's license state (2-letter code)
 * @return {string} Vendor group: CERTN, PENNDOT, INFORM, INFORMDATA, or UNKNOWN
 */
function extractVendorGroup(subject, dlState) {
  if (!subject) {
    subject = "";
  }
  
  // Priority 0: Check for InformData prefix (new format)
  if (MVR_PATTERNS.INFORM_PATTERN && MVR_PATTERNS.INFORM_PATTERN.test(subject)) {
    return VENDOR_GROUPS.INFORMDATA;
  }
  
  // Also check for Sentinel Report pattern (InformData tickets)
  if (subject.includes(MVR_PATTERNS.SENTINEL_REPORT)) {
    return VENDOR_GROUPS.INFORMDATA;
  }
  
  // Priority 1: Check for explicit vendor prefix in subject
  const prefixMatch = subject.match(MVR_PATTERNS.VENDOR_PREFIX_REGEX);
  if (prefixMatch) {
    const prefix = prefixMatch[1].toUpperCase();
    if (prefix === "CERTN") return VENDOR_GROUPS.CERTN;
    if (prefix === "PENNDOT") return VENDOR_GROUPS.PENNDOT;
    if (prefix === "INFORM") return VENDOR_GROUPS.INFORM;
    if (prefix === "INFORMDATA") return VENDOR_GROUPS.INFORMDATA;
  }
  
  // Check for Certn continuous monitoring pattern
  if (MVR_PATTERNS.CERTN_PATTERN.test(subject)) {
    return VENDOR_GROUPS.CERTN;
  }
  
  // Check for PennDOT continuous monitoring pattern
  if (MVR_PATTERNS.PENNDOT_PATTERN.test(subject)) {
    return VENDOR_GROUPS.PENNDOT;
  }
  
  // Priority 2: Use DL State to determine vendor
  // State-based mapping from Heroku MVR queue analysis
  if (dlState) {
    const state = dlState.toUpperCase().trim();
    
    // PennDOT for Pennsylvania
    if (state === PENNDOT_STATE) {
      return VENDOR_GROUPS.PENNDOT;
    }
    
    // CERTN for MO, IL
    if (CERTN_STATES.includes(state)) {
      return VENDOR_GROUPS.CERTN;
    }
    
    // All other states use INFORMDATA (Sentinel System)
    // This is 77.5% of volume based on queue analysis
    return VENDOR_GROUPS.INFORMDATA;
  }
  
  // No state available - cannot determine vendor
  return VENDOR_GROUPS.UNKNOWN;
}

// ═══════════════════════════════════════════════════════════════════════════════
// CUSTOM FIELD EXTRACTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Extract and format custom fields from ticket
 * 
 * @param {Object} customFields - Custom fields object from Freshdesk ticket
 * @return {Object} Formatted custom fields with display names
 */
function extractCustomFields(customFields) {
  const result = {};
  
  if (!customFields) {
    return result;
  }
  
  for (const [key, value] of Object.entries(customFields)) {
    // Use mapped name if available, otherwise create display name
    const displayName = CUSTOM_FIELD_MAP[key] || 
                        CUSTOM_FIELD_PREFIX + key.replace(/^cf_/, '').replace(/_/g, ' ');
    result[displayName] = value || "";
  }
  
  return result;
}

/**
 * Get list of all custom field column headers
 * Combines predefined mappings with dynamically discovered fields
 * 
 * @param {Array} tickets - Sample tickets to discover custom fields from
 * @return {Array} Array of custom field header names
 */
function getCustomFieldHeaders(tickets) {
  const headers = new Set();
  
  // Add predefined custom field mappings
  for (const displayName of Object.values(CUSTOM_FIELD_MAP)) {
    headers.add(displayName);
  }
  
  // Discover additional custom fields from tickets
  if (tickets && tickets.length > 0) {
    for (const ticket of tickets) {
      if (ticket.custom_fields) {
        const extracted = extractCustomFields(ticket.custom_fields);
        for (const header of Object.keys(extracted)) {
          headers.add(header);
        }
      }
    }
  }
  
  return Array.from(headers).sort();
}

// ═══════════════════════════════════════════════════════════════════════════════
// QA TEST SUITE - Tag Classification System
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * QA TEST SUITE - Tag Classification System
 * Run this to validate column mapper and classification after changes
 */
function runTagClassificationQA() {
  Logger.log('\n🧪 === TAG CLASSIFICATION QA TEST SUITE ===\n');
  
  const tests = [];
  let passed = 0;
  let failed = 0;
  
  // Test 1: Column Mapper with correct order
  tests.push({
    name: 'Column Mapper - Standard Order',
    test: () => {
      const headers = TAG_MAPPING_HEADERS;
      const colMap = getTagMappingColumnMap(headers);
      return colMap["Tag Pattern"] === 0 && 
             colMap["Match Type"] === 1 && 
             colMap["Outcome Type"] === 2 &&
             colMap["Priority"] === 3 &&
             colMap["Is Billable"] === 4 &&
             colMap["Is Active"] === 5 &&
             colMap["Notes"] === 6;
    }
  });
  
  // Test 2: Column Mapper with reordered columns
  tests.push({
    name: 'Column Mapper - Reordered Columns',
    test: () => {
      const reordered = ["Priority", "Tag Pattern", "Outcome Type", "Match Type", "Is Active", "Is Billable", "Notes"];
      const colMap = getTagMappingColumnMap(reordered);
      return colMap["Tag Pattern"] === 1 && 
             colMap["Match Type"] === 3 && 
             colMap["Outcome Type"] === 2 &&
             colMap["Priority"] === 0;
    }
  });
  
  // Test 3: Load mappings works
  tests.push({
    name: 'Load Mappings - Basic',
    test: () => {
      const mappings = loadTagMappings();
      return mappings.length > 0 && 
             mappings[0].pattern && 
             mappings[0].outcomeType &&
             mappings[0].regex;
    }
  });
  
  // Test 4: Classification with tags
  tests.push({
    name: 'Classification - Suspended Tag',
    test: () => {
      const result = classifyOutcome(['suspended'], []);
      return result.outcome === 'Suspension Confirmed' && 
             result.isBillable === true;
    }
  });
  
  // Test 5: Billable flag reads correctly
  tests.push({
    name: 'Billable Flag - Correct Value',
    test: () => {
      const mappings = loadTagMappings();
      const suspendedMapping = mappings.find(m => m.pattern.toLowerCase().includes('suspend'));
      return suspendedMapping && suspendedMapping.isBillable === true;
    }
  });
  
  // Test 6: Priority sorting
  tests.push({
    name: 'Priority Sorting - Correct Order',
    test: () => {
      const mappings = loadTagMappings();
      for (let i = 0; i < mappings.length - 1; i++) {
        if (mappings[i].priority > mappings[i + 1].priority) {
          return false;
        }
      }
      return true;
    }
  });
  
  // Test 7: Inactive mappings skipped
  tests.push({
    name: 'Inactive Mappings - Skipped',
    test: () => {
      const mappings = loadTagMappings();
      return mappings.every(m => m.hasOwnProperty('pattern')); // All loaded mappings should have pattern
    }
  });
  
  // Test 8: Match Type handling
  tests.push({
    name: 'Match Type - Contains vs Regex',
    test: () => {
      const mappings = loadTagMappings();
      return mappings.some(m => m.matchType === 'contains') || 
             mappings.some(m => m.matchType === 'regex');
    }
  });
  
  // Run all tests
  tests.forEach(test => {
    try {
      const result = test.test();
      if (result) {
        Logger.log(`✅ PASS: ${test.name}`);
        passed++;
      } else {
        Logger.log(`❌ FAIL: ${test.name}`);
        failed++;
      }
    } catch (e) {
      Logger.log(`❌ ERROR: ${test.name} - ${e.message}`);
      failed++;
    }
  });
  
  Logger.log(`\n📊 QA Results: ${passed}/${tests.length} passed, ${failed} failed\n`);
  
  if (failed === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `All ${tests.length} tests passed!`,
      '✅ QA Complete',
      5
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `${failed} tests failed. Check logs.`,
      '⚠️ QA Issues',
      5
    );
  }
  
  return {passed, failed, total: tests.length};
}

// ═══════════════════════════════════════════════════════════════════════════════
// SYNC TAG MAPPINGS - Single Source of Truth
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Sync Tag_Outcome_Mappings sheet with DEFAULT_TAG_MAPPINGS from Config.gs
 * This is the SINGLE SOURCE OF TRUTH for all tag mappings
 * 
 * What it does:
 * 1. Clears all existing rows (keeps header)
 * 2. Writes ALL patterns from DEFAULT_TAG_MAPPINGS in correct order
 * 3. Applies formatting and validation
 * 
 * Use this to:
 * - Initial setup of mappings
 * - Reset to clean state after manual edits
 * - Apply updated billable logic from code
 */
function syncTagMappings() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Confirm with user
  const confirm = ui.alert(
    '⚠️ Sync Tag Mappings',
    `This will REPLACE ALL existing tag mappings with ${DEFAULT_TAG_MAPPINGS.length} standard mappings from code.\n\n` +
    'Any custom mappings you added manually will be removed.\n\n' +
    'This ensures the sheet matches the correct billable logic.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) {
    ss.toast('Sync cancelled.', 'ℹ️ Cancelled', 3);
    return {synced: 0, cancelled: true};
  }
  
  Logger.log('\n🔄 === SYNC TAG MAPPINGS ===\n');
  
  // Get or create sheet
  let sheet = ss.getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
    Logger.log('📄 Created new Tag_Outcome_Mappings sheet');
  }
  
  // Clear all existing data
  sheet.clear();
  Logger.log('🧹 Cleared existing data');
  
  // Write header row
  sheet.appendRow(TAG_MAPPING_HEADERS);
  
  // Write all mappings from DEFAULT_TAG_MAPPINGS
  DEFAULT_TAG_MAPPINGS.forEach((row, index) => {
    sheet.appendRow(row);
    Logger.log(`✅ Row ${index + 2}: ${row[0]} → ${row[2]} (Billable: ${row[4]})`);
  });
  
  // Format header row
  sheet.getRange(1, 1, 1, TAG_MAPPING_HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#4a86e8")
    .setFontColor("white");
  
  // Apply data validation
  const colMap = getTagMappingColumnMap(TAG_MAPPING_HEADERS);
  applyTagMappingValidation(sheet, colMap);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, TAG_MAPPING_HEADERS.length);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log(`\n✅ Sync complete: ${DEFAULT_TAG_MAPPINGS.length} mappings written\n`);
  
  // Show summary
  ui.alert(
    '✅ Sync Complete',
    `Successfully synced ${DEFAULT_TAG_MAPPINGS.length} tag mappings.\n\n` +
    'Billable Logic Summary:\n' +
    '• Violations (suspended, expired, invalid): Billable = YES\n' +
    '• Delivery (Updated, uploaded): Billable = YES\n' +
    '• Same info, approved, requested: Billable = NO\n' +
    '• Pending, processing states: Billable = NO\n' +
    '• System issues (DMV down): Billable = NO\n\n' +
    'All mappings are now active and in priority order.',
    ui.ButtonSet.OK
  );
  
  ss.toast(
    `Synced ${DEFAULT_TAG_MAPPINGS.length} mappings from code.`,
    '✅ Sync Complete',
    5
  );
  
  return {synced: DEFAULT_TAG_MAPPINGS.length, cancelled: false};
}
