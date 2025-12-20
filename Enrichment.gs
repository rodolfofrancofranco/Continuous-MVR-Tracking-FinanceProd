/**
 * TICKET ENRICHMENT SYSTEM
 * MVR TICKET TRACKER - Data Enrichment
 * 
 * Purpose: Enrich raw Freshdesk tickets with complete information
 * - Parse subject line to extract Type (SC/RC), Partner, Turn_ID
 * - Add human-readable names (status, priority, agent, group)
 * - Calculate resolution times and SLA status
 * - Extract partner from custom fields
 * - Parse ticket description for workflow data points
 * - Classify MVR outcomes from tags and conversations
 * - Extract vendor group and custom fields
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SUBJECT LINE PARSING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Parse MVR ticket subject line to extract components
 * Pattern: "One Off Continuous MVR - (Suspension Check|Recheck) - {partner} - {turnId}"
 * 
 * @param {string} subject - Ticket subject line
 * @return {Object} Object with type, partner, turnId (null if parsing fails)
 */
function parseSubjectLine(subject) {
  if (!subject) {
    return { type: null, partner: null, turnId: null, isInformData: false };
  }
  
  // Check for InformData prefix and strip it
  let cleanSubject = subject;
  let isInformData = false;
  
  if (MVR_PATTERNS.INFORM_PATTERN && MVR_PATTERNS.INFORM_PATTERN.test(subject)) {
    cleanSubject = subject.replace(/^\[InformData\]\s*/i, '');
    isInformData = true;
  }
  
  // Try standard subject regex first
  const match = cleanSubject.match(MVR_PATTERNS.SUBJECT_REGEX);
  
  if (match) {
    // Extract components from regex groups
    const actionType = match[1]; // "Suspension Check" or "Recheck"
    const partner = match[2].trim();
    const turnId = match[3].trim();
    
    // Convert to short codes
    let type = null;
    if (actionType.toLowerCase().includes('suspension check')) {
      type = 'SC';
    } else if (actionType.toLowerCase().includes('recheck')) {
      type = 'RC';
    }
    
    return { type, partner, turnId, isInformData };
  }
  
  // Try InformData-specific regex (Sentinel Report format)
  if (MVR_PATTERNS.INFORM_SUBJECT_REGEX) {
    const informMatch = cleanSubject.match(MVR_PATTERNS.INFORM_SUBJECT_REGEX);
    
    if (informMatch) {
      const actionType = informMatch[1]; // "Sentinel Report"
      const turnId = informMatch[2].trim();
      
      // Map Sentinel Report to SC
      let type = 'SC';
      if (actionType.toLowerCase().includes('sentinel')) {
        type = 'SC';
      }
      
      // Partner will be extracted from notes
      return { type, partner: null, turnId, isInformData: true };
    }
  }
  
  Logger.log(`⚠️ Failed to parse subject: ${subject}`);
  return { type: null, partner: null, turnId: null, isInformData };
}

/**
 * Test subject line parsing with sample subjects
 */
function testSubjectParsing() {
  const testSubjects = [
    "One Off Continuous MVR - Suspension Check - ABC Transport - T12345",
    "One Off Continuous MVR - Recheck - XYZ Logistics - T67890",
    "One Off Continuous MVR - Suspension Check - Test Partner Name - T99999"
  ];
  
  Logger.log('🧪 Testing subject line parsing...\n');
  
  testSubjects.forEach(subject => {
    const parsed = parseSubjectLine(subject);
    Logger.log(`Subject: ${subject}`);
    Logger.log(`  Type: ${parsed.type}`);
    Logger.log(`  Partner: ${parsed.partner}`);
    Logger.log(`  Turn ID: ${parsed.turnId}\n`);
  });
}

// ═══════════════════════════════════════════════════════════════════════════════
// AGENT & GROUP NAME LOOKUPS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get agent name from Agent-Mappings sheet or return ID if not found
 * @param {number} agentId - Freshdesk agent ID
 * @return {string} Agent name or "Agent {id}"
 */
function getAgentName(agentId) {
  if (!agentId) return "";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.AGENT_MAPPINGS);
    
    if (!sheet) {
      return `Agent ${agentId}`;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Skip header row, find matching agent ID
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] == agentId) { // Column A: Agent ID
        return row[1] || `Agent ${agentId}`; // Column B: Agent Name
      }
    }
    
    return `Agent ${agentId}`;
  } catch (e) {
    Logger.log(`⚠️ Error getting agent name for ${agentId}: ${e.message}`);
    return `Agent ${agentId}`;
  }
}

/**
 * Get group name from Freshdesk-Mappings sheet or return ID if not found
 * @param {number} groupId - Freshdesk group ID
 * @return {string} Group name or "Group {id}"
 */
function getGroupName(groupId) {
  if (!groupId) return "";
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.FRESHDESK_MAPPINGS);
    
    if (!sheet) {
      return `Group ${groupId}`;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find "Group" section and matching ID
    let inGroupSection = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const fieldName = row[0]; // Column A: Field Name
      
      if (fieldName === "Group") {
        inGroupSection = true;
        continue;
      }
      
      // Stop if we hit another section
      if (inGroupSection && fieldName && fieldName !== "Group") {
        break;
      }
      
      if (inGroupSection && row[1] == groupId) { // Column B: Code
        return row[2] || `Group ${groupId}`; // Column C: Label
      }
    }
    
    return `Group ${groupId}`;
  } catch (e) {
    Logger.log(`⚠️ Error getting group name for ${groupId}: ${e.message}`);
    return `Group ${groupId}`;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PARTNER EXTRACTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Extract partner name from ticket
 * Priority: 1) Subject line, 2) Custom field cf_partner_name, 3) Company field
 * @param {Object} ticket - Freshdesk ticket object
 * @param {Object} parsedSubject - Parsed subject line components
 * @return {string} Partner name
 */
function extractPartnerName(ticket, parsedSubject) {
  // Priority 1: Partner from subject line
  if (parsedSubject && parsedSubject.partner) {
    return parsedSubject.partner;
  }
  
  // Priority 2: Custom field cf_partner_name
  if (ticket.custom_fields && ticket.custom_fields.cf_partner_name) {
    return ticket.custom_fields.cf_partner_name;
  }
  
  // Priority 3: Company name (if available - requires additional API call)
  // For now, return empty if not in subject
  return "";
}

// ═══════════════════════════════════════════════════════════════════════════════
// ALL NOTES TEXT PARSING (Fallback for InformData/Sentinel tickets)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Parse All Notes Text to extract InformData/Sentinel ticket fields
 * Used as fallback when subject/description parsing fails
 * 
 * InformData notes contain asterisk-wrapped values:
 * Partner: *Papa*
 * DL State: *TN*
 * Turn ID: *C8078635470*
 * 
 * @param {string} notesText - All Notes Text content
 * @return {Object} Extracted fields from notes
 */
function parseAllNotesText(notesText) {
  const result = {
    turn_id: "",
    partner: "",
    dl_state: "",
    dl_number: "",
    order_id: "",
    tier: "",
    cadence_days: "",
    days_since_last: "",
    last_check_date: "",
    ta_url: "",
    process_type: "",
    vendor: "",
    source: ""
  };
  
  if (!notesText) {
    return result;
  }
  
  // Extract each field using NOTES_FIELD_PATTERNS
  const turnIdMatch = notesText.match(NOTES_FIELD_PATTERNS.TURN_ID);
  if (turnIdMatch) result.turn_id = turnIdMatch[1].trim();
  
  const partnerMatch = notesText.match(NOTES_FIELD_PATTERNS.PARTNER);
  if (partnerMatch) result.partner = partnerMatch[1].trim();
  
  const dlStateMatch = notesText.match(NOTES_FIELD_PATTERNS.DL_STATE);
  if (dlStateMatch) result.dl_state = dlStateMatch[1].toUpperCase().trim();
  
  const dlNumberMatch = notesText.match(NOTES_FIELD_PATTERNS.DL_NUMBER);
  if (dlNumberMatch) result.dl_number = dlNumberMatch[1].trim();
  
  const orderIdMatch = notesText.match(NOTES_FIELD_PATTERNS.ORDER_ID);
  if (orderIdMatch) result.order_id = orderIdMatch[1].trim();
  
  const tierMatch = notesText.match(NOTES_FIELD_PATTERNS.TIER);
  if (tierMatch) result.tier = tierMatch[1].trim();
  
  const cadenceMatch = notesText.match(NOTES_FIELD_PATTERNS.CADENCE_DAYS);
  if (cadenceMatch) result.cadence_days = cadenceMatch[1].trim();
  
  const daysSinceMatch = notesText.match(NOTES_FIELD_PATTERNS.DAYS_SINCE_LAST);
  if (daysSinceMatch) result.days_since_last = daysSinceMatch[1].trim();
  
  const lastCheckMatch = notesText.match(NOTES_FIELD_PATTERNS.LAST_CHECK_DATE);
  if (lastCheckMatch) result.last_check_date = lastCheckMatch[1].trim();
  
  const taUrlMatch = notesText.match(NOTES_FIELD_PATTERNS.TA_URL);
  if (taUrlMatch) result.ta_url = taUrlMatch[1].trim();
  
  const processTypeMatch = notesText.match(NOTES_FIELD_PATTERNS.PROCESS_TYPE);
  if (processTypeMatch) result.process_type = processTypeMatch[1].trim();
  
  const vendorMatch = notesText.match(NOTES_FIELD_PATTERNS.VENDOR);
  if (vendorMatch) result.vendor = vendorMatch[1].trim();
  
  const sourceMatch = notesText.match(NOTES_FIELD_PATTERNS.SOURCE);
  if (sourceMatch) result.source = sourceMatch[1].trim();
  
  return result;
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMPLETE TICKET ENRICHMENT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Enrich a single ticket with all additional information
 * @param {Object} ticket - Raw Freshdesk ticket object
 * @param {Array} conversations - Optional array of conversation objects
 * @param {Object} existingOverrides - Optional existing override data to preserve
 * @return {Object} Enriched ticket object
 */
function enrichTicket(ticket, conversations = null, existingOverrides = null) {
  // Parse subject line
  const parsed = parseSubjectLine(ticket.subject);
  
  // Calculate time metrics
  const createdAt = new Date(ticket.created_at);
  const ageHours = calculateAgeHours(createdAt);
  const resolutionHours = calculateResolutionHours(ticket);
  
  // Get human-readable names
  const statusName = STATUS_NAMES[ticket.status] || "Unknown";
  const priorityName = PRIORITY_NAMES[ticket.priority] || "Unknown";
  const agentName = getAgentName(ticket.responder_id);
  const groupName = getGroupName(ticket.group_id);
  
  // Classify SLA
  const sla = classifySLA(ticket, ageHours, resolutionHours);
  
  // Get resolved date
  const resolvedDate = getResolvedDate(ticket);
  
  // Get source name
  const sourceName = getSourceName(ticket.source);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Parse workflow data from ticket description
  // ═══════════════════════════════════════════════════════════════════════════
  const description = ticket.description || ticket.description_text || "";
  let workflowData = parseTicketDescription(description);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // UNIVERSAL FALLBACK: Always parse All Notes Text for any missing fields
  // This applies to ALL ticket types, not just InformData
  // ═══════════════════════════════════════════════════════════════════════════
  const allNotesText = ticket.all_notes_text || ticket.allNotesText || "";
  let notesData = null;
  
  // Always parse notes if we have notes text - fill in any missing fields
  if (allNotesText) {
    notesData = parseAllNotesText(allNotesText);
    
    // Merge notes data into workflow data (notes fills in missing fields only)
    if (!workflowData.dl_state && notesData.dl_state) workflowData.dl_state = notesData.dl_state;
    if (!workflowData.dl_number && notesData.dl_number) workflowData.dl_number = notesData.dl_number;
    if (!workflowData.tier && notesData.tier) workflowData.tier = notesData.tier;
    if (!workflowData.cadence_days && notesData.cadence_days) workflowData.cadence_days = notesData.cadence_days;
    if (!workflowData.days_since_last && notesData.days_since_last) workflowData.days_since_last = notesData.days_since_last;
    if (!workflowData.last_check_date && notesData.last_check_date) workflowData.last_check_date = notesData.last_check_date;
    if (!workflowData.ta_url && notesData.ta_url) workflowData.ta_url = notesData.ta_url;
    if (!workflowData.order_id && notesData.order_id) workflowData.order_id = notesData.order_id;
    if (!workflowData.process_type && notesData.process_type) workflowData.process_type = notesData.process_type;
    if (!workflowData.vendor && notesData.vendor) workflowData.vendor = notesData.vendor;
    if (!workflowData.source && notesData.source) workflowData.source = notesData.source;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Extract partner with notes fallback
  // ═══════════════════════════════════════════════════════════════════════════
  let partnerName = extractPartnerName(ticket, parsed);
  
  // If still no partner, try notes data
  if (!partnerName && notesData && notesData.partner) {
    partnerName = notesData.partner;
  }
  // Graceful default if still missing
  partnerName = partnerName || "";
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Extract Turn ID with notes fallback  
  // ═══════════════════════════════════════════════════════════════════════════
  let turnId = parsed.turnId;
  if (!turnId && notesData && notesData.turn_id) {
    turnId = notesData.turn_id;
  }
  // Graceful default if still missing
  turnId = turnId || "";
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Extract Type (SC/RC) with notes fallback
  // ═══════════════════════════════════════════════════════════════════════════
  let requestType = parsed.type;
  if (!requestType && notesData && notesData.process_type) {
    // Try to infer from notes process_type field
    const processType = notesData.process_type.toLowerCase();
    if (processType.includes('suspension') || processType.includes('sentinel')) {
      requestType = 'SC';
    } else if (processType.includes('recheck')) {
      requestType = 'RC';
    }
  }
  // Graceful default if still missing
  requestType = requestType || "";
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Determine vendor group (with InformData detection)
  // ═══════════════════════════════════════════════════════════════════════════
  const vendorGroup = extractVendorGroup(ticket.subject, workflowData.dl_state);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Classify outcome from tags and conversations
  // ═══════════════════════════════════════════════════════════════════════════
  const ticketConversations = conversations || ticket.conversations || [];
  const outcomeClassification = classifyOutcome(ticket.tags, ticketConversations);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Extract conversation metadata
  // ═══════════════════════════════════════════════════════════════════════════
  const conversationCount = getConversationCount(ticketConversations);
  const lastResponder = extractLastResponder(ticketConversations);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Extract custom fields
  // ═══════════════════════════════════════════════════════════════════════════
  const customFieldsExtracted = extractCustomFields(ticket.custom_fields);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Preserve existing overrides if provided
  // ═══════════════════════════════════════════════════════════════════════════
  const outcomeOverride = existingOverrides?.outcomeOverride || "";
  const overrideReason = existingOverrides?.overrideReason || "";
  
  // Return enriched ticket
  return {
    // Original ticket data
    ...ticket,
    
    // Parsed subject components (with notes fallback and graceful defaults)
    type: requestType,
    partner: partnerName,
    turnId: turnId,
    
    // Human-readable names
    status_name: statusName,
    priority_name: priorityName,
    agent_name: agentName,
    group_name: groupName,
    source_name: sourceName,
    
    // Time calculations
    age_hours: ageHours,
    resolution_hours: resolutionHours,
    resolved_date: resolvedDate,
    
    // SLA classification
    sla: sla,
    
    // Pull timestamp
    last_pulled: new Date(),
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Workflow origin data (from email body or notes fallback)
    // ═══════════════════════════════════════════════════════════════════════════
    vendor_group: vendorGroup,
    dl_state: workflowData.dl_state,
    dl_number: workflowData.dl_number,
    days_since_last: workflowData.days_since_last,
    last_check_date: workflowData.last_check_date,
    cadence_days: workflowData.cadence_days,
    tier: workflowData.tier,
    enrollment_type: workflowData.enrollment_type,
    ta_url: workflowData.ta_url,
    
    // InformData-specific field
    order_id: notesData?.order_id || "",
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Outcome classification
    // ═══════════════════════════════════════════════════════════════════════════
    mvr_outcome: outcomeClassification.outcome,
    outcome_source: outcomeClassification.source,
    outcome_date: outcomeClassification.outcomeDate || "",
    resolution_notes: outcomeClassification.notes,
    conversation_count: conversationCount,
    last_responder: lastResponder,
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Manual override fields (preserved from existing data)
    // ═══════════════════════════════════════════════════════════════════════════
    outcome_override: outcomeOverride,
    override_reason: overrideReason,
    
    // ═══════════════════════════════════════════════════════════════════════════
    // Custom fields (dynamic)
    // ═══════════════════════════════════════════════════════════════════════════
    custom_fields_extracted: customFieldsExtracted
  };
}

/**
 * Batch enrich multiple tickets
 * @param {Array} tickets - Array of raw Freshdesk ticket objects (may include conversations)
 * @param {Map} existingOverridesMap - Optional map of ticketId -> override data
 * @return {Array} Array of enriched ticket objects
 */
function enrichTickets(tickets, existingOverridesMap = null) {
  Logger.log(`🔄 Enriching ${tickets.length} tickets...`);
  
  const enriched = tickets.map(ticket => {
    const conversations = ticket.conversations || null;
    const overrides = existingOverridesMap?.get(ticket.id) || null;
    return enrichTicket(ticket, conversations, overrides);
  });
  
  Logger.log(`✅ Enrichment complete`);
  return enriched;
}

// ═══════════════════════════════════════════════════════════════════════════════
// SOURCE NAME MAPPING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get human-readable source name from source code
 * @param {number} source - Freshdesk source code
 * @return {string} Source name
 */
function getSourceName(source) {
  const sourceMap = {
    1: "Email",
    2: "Portal",
    3: "Phone",
    4: "Chat",
    5: "Mobihelp",
    6: "Feedback Widget",
    7: "Outbound Email"
  };
  
  return sourceMap[source] || `Source ${source}`;
}

// ═══════════════════════════════════════════════════════════════════════════════
// TICKET TO ROW CONVERSION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Convert enriched ticket to sheet row format
 * Includes all new workflow, outcome, and override columns
 * @param {Object} ticket - Enriched ticket object
 * @return {Array} Row data in HISTORY_HEADERS order (42 columns + custom fields)
 */
function ticketToRow(ticket) {
  return [
    // Original columns (1-25)
    ticket.id,
    ticket.subject || "",
    ticket.type || "",
    ticket.partner || "",
    ticket.turnId || "",
    ticket.status,
    ticket.status_name || "",
    ticket.priority,
    ticket.priority_name || "",
    ticket.created_at,
    ticket.updated_at,
    ticket.resolved_date || "",
    ticket.age_hours ? ticket.age_hours.toFixed(1) : "",
    ticket.resolution_hours ? ticket.resolution_hours.toFixed(1) : "",
    ticket.responder_id || "",
    ticket.agent_name || "",
    ticket.group_id || "",
    ticket.group_name || "",
    ticket.requester_id || "",
    ticket.company_id || "",
    (ticket.tags || []).join(", "),
    ticket.source || "",
    ticket.source_name || "",
    ticket.sla || "",
    ticket.last_pulled,
    // Workflow Origin columns (26-34)
    ticket.vendor_group || "",
    ticket.dl_state || "",
    ticket.dl_number || "",
    ticket.days_since_last || "",
    ticket.last_check_date || "",
    ticket.cadence_days || "",
    ticket.tier || "",
    ticket.enrollment_type || "",
    ticket.ta_url || "",
    // Outcome columns (35-40)
    ticket.mvr_outcome || "",
    ticket.outcome_source || "",
    ticket.outcome_date || "",
    ticket.resolution_notes || "",
    ticket.conversation_count || 0,
    ticket.last_responder || "",
    // Override columns (41-42)
    ticket.outcome_override || "",
    ticket.override_reason || "",
    // Change tracking (43)
    ticket.change_log || ""
    // Note: Custom fields are handled separately via getCustomFieldValues()
  ];
}

/**
 * Convert multiple enriched tickets to sheet rows
 * @param {Array} tickets - Array of enriched ticket objects
 * @return {Array} 2D array of row data
 */
function ticketsToRows(tickets) {
  return tickets.map(ticket => ticketToRow(ticket));
}

// ═══════════════════════════════════════════════════════════════════════════════
// VALIDATION & TESTING
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test enrichment with sample ticket data
 */
function testEnrichment() {
  Logger.log('🧪 Testing ticket enrichment...\n');
  
  // Sample ticket (minimal structure)
  const sampleTicket = {
    id: 12345,
    subject: "One Off Continuous MVR - Suspension Check - ABC Transport - T67890",
    status: 2, // Open
    priority: 3, // High
    created_at: new Date(Date.now() - 12 * 3600000).toISOString(), // 12 hours ago
    updated_at: new Date().toISOString(),
    responder_id: 101,
    group_id: 5,
    requester_id: 202,
    company_id: 303,
    tags: ["mvr", "urgent"],
    source: 1, // Email
    custom_fields: {
      cf_partner_name: "ABC Transport",
      cf_turn_id: "T67890"
    }
  };
  
  Logger.log('Sample ticket:');
  Logger.log(JSON.stringify(sampleTicket, null, 2));
  
  const enriched = enrichTicket(sampleTicket);
  
  Logger.log('\nEnriched ticket:');
  Logger.log(`  Type: ${enriched.type}`);
  Logger.log(`  Partner: ${enriched.partner}`);
  Logger.log(`  Turn ID: ${enriched.turnId}`);
  Logger.log(`  Status: ${enriched.status_name}`);
  Logger.log(`  Priority: ${enriched.priority_name}`);
  Logger.log(`  Age: ${enriched.age_hours.toFixed(1)} hours`);
  Logger.log(`  SLA: ${enriched.sla}`);
  Logger.log(`  Agent: ${enriched.agent_name}`);
  Logger.log(`  Group: ${enriched.group_name}`);
  Logger.log(`  Source: ${enriched.source_name}`);
  
  Logger.log('\n✅ Enrichment test complete');
}
