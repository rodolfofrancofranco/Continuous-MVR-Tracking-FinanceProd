/**
 * MVR TICKET TRACKER - REPORTING SYSTEM
 * Configuration Constants and Setup
 * 
 * Purpose: Track Freshdesk tickets created by MVR automation emails
 * Source: Tickets with subject pattern "One Off Continuous MVR"
 * Data Retention: 10-day lookback + historical append-only record
 */

// ═══════════════════════════════════════════════════════════════════════════════
// SHEET NAMES
// ═══════════════════════════════════════════════════════════════════════════════

const SHEET_NAMES = {
  MVR_TICKET_HISTORY: "MVR_Ticket_History",
  MVR_RAW_TICKETS: "MVR_Raw_Tickets",
  BY_PARTNER_SUMMARY: "By_Partner_Summary",
  MONTHLY_REPORT: "Monthly_MVR_Report",
  AGENT_MAPPINGS: "Agent-Mappings",
  FRESHDESK_MAPPINGS: "Freshdesk-Mappings",
  TAG_OUTCOME_MAPPINGS: "Tag_Outcome_Mappings",
  OVERRIDE_AUDIT_LOG: "Override_Audit_Log",
  EXECUTIVE_DASHBOARD: "Executive_Dashboard",
  FINANCE_AUDIT_TRAIL: "Finance_Audit_Trail",
  OPS_PERFORMANCE: "Ops_Performance",
  // New reporting suite sheets
  ASSUMPTIONS_LOG: "Assumptions_Log",
  ANNUAL_PROJECTION: "Annual_Projection",
  OPERATIVE_PLAN: "Operative_Plan",
  GROWTH_PROJECTION: "Growth_Projection",
  // Dashboard views
  DASHBOARD_PARTNERS: "Dashboard_Partners",
  DASHBOARD_VENDORS: "Dashboard_Vendors",
  DASHBOARD_STATES: "Dashboard_States",
  DASHBOARD_GEOGRAPHIC: "Dashboard_Geographic",
  RECONCILIATION_PREFIX: "Reconciliation_"  // + YYYY for year-specific sheets
};

// ═══════════════════════════════════════════════════════════════════════════════
// TIME WINDOWS & CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const TIME_CONFIG = {
  // Hourly updates (automated triggers - every hour)
  HOURLY_LOOKBACK_HOURS: 2,       // 2 hours - catches recent changes
  
  // Periodic updates (automated triggers)
  PERIODIC_LOOKBACK_DAYS: 10,
  PERIODIC_LOOKBACK_HOURS: 240,   // 10 days in hours
  
  // Manual full sync (on-demand)
  FULL_SYNC_LOOKBACK_DAYS: 30,
  FULL_SYNC_LOOKBACK_HOURS: 720,  // 30 days in hours
  
  // Default (used by existing code)
  LOOKBACK_DAYS: 30,
  LOOKBACK_HOURS: 720,
  
  TIMEZONE: "America/Chicago"     // CST/CDT
};

// Fields to check for changes when syncing existing tickets
const CHANGE_DETECTION_FIELDS = [
  'status',           // Ticket status (Open, Pending, Resolved, Closed)
  'tags',             // Tags array - affects outcome classification
  'priority',         // Priority changes
  'agent_id',         // Reassignment
  'group_id',         // Group changes
  'resolved_at',      // Resolution timestamp
  'closed_at'         // Closure timestamp
];

// ═══════════════════════════════════════════════════════════════════════════════
// BATCH PROCESSING CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const BATCH_CONFIG = {
  BATCH_SIZE: 100,                // Tickets per batch for conversation fetching (increased for speed)
  BATCH_DELAY_MS: 500,            // 500ms delay between batches (reduced from 1000ms)
  BACKFILL_DAYS: 30               // Days to look back for backfill operation
};

// ═══════════════════════════════════════════════════════════════════════════════
// CELL CHARACTER LIMITS (Google Sheets max = 50,000)
// ═══════════════════════════════════════════════════════════════════════════════

const CELL_LIMITS = {
  MAX_CHARS: 50000,               // Google Sheets hard limit
  SAFE_LIMIT: 45000,              // Safe limit with buffer
  TEXT_FIELD_LIMIT: 40000,        // For text fields (notes, descriptions)
  JSON_FIELD_LIMIT: 45000,        // For JSON fields
  URL_FIELD_LIMIT: 30000          // For URL lists
};

// ═══════════════════════════════════════════════════════════════════════════════
// TIER CONFIGURATION (Check Cadences)
// ═══════════════════════════════════════════════════════════════════════════════

const TIER_CONFIG = {
  1: { cadence_days: 60,  checks_per_year: 6, description: 'High-risk states - 60 day cadence' },
  2: { cadence_days: 90,  checks_per_year: 4, description: 'Medium-risk states - 90 day cadence' },
  3: { cadence_days: 180, checks_per_year: 2, description: 'Standard states - 180 day cadence' }
};

// ═══════════════════════════════════════════════════════════════════════════════
// STATE CONFIGURATION (All 50 States)
// tier: 1/2/3, is_free: true = DMV site check (no vendor cost)
// ═══════════════════════════════════════════════════════════════════════════════

const STATE_CONFIG = {
  // Tier 1 States (60-day cadence) - All PAID
  'DE': { tier: 1, is_free: false, name: 'Delaware' },
  'MT': { tier: 1, is_free: false, name: 'Montana' },
  'MS': { tier: 1, is_free: false, name: 'Mississippi' },
  
  // Tier 2 States (90-day cadence) - Mixed FREE/PAID
  'GA': { tier: 2, is_free: true,  name: 'Georgia' },
  'LA': { tier: 2, is_free: true,  name: 'Louisiana' },
  'TN': { tier: 2, is_free: true,  name: 'Tennessee' },
  'MI': { tier: 2, is_free: false, name: 'Michigan' },
  
  // Tier 3 States (180-day cadence) - Mixed FREE/PAID
  // FREE States (DMV site check)
  'AL': { tier: 3, is_free: true,  name: 'Alabama' },
  'AK': { tier: 3, is_free: true,  name: 'Alaska' },
  'CT': { tier: 3, is_free: true,  name: 'Connecticut' },
  'FL': { tier: 3, is_free: true,  name: 'Florida' },
  'ID': { tier: 3, is_free: true,  name: 'Idaho' },
  'KS': { tier: 3, is_free: true,  name: 'Kansas' },
  'MD': { tier: 3, is_free: true,  name: 'Maryland' },
  'MA': { tier: 3, is_free: true,  name: 'Massachusetts' },
  'MN': { tier: 3, is_free: true,  name: 'Minnesota' },
  'NE': { tier: 3, is_free: true,  name: 'Nebraska' },
  'ND': { tier: 3, is_free: true,  name: 'North Dakota' },
  'OR': { tier: 3, is_free: true,  name: 'Oregon' },
  'RI': { tier: 3, is_free: true,  name: 'Rhode Island' },
  'TX': { tier: 3, is_free: true,  name: 'Texas' },
  'WA': { tier: 3, is_free: true,  name: 'Washington' },
  'WI': { tier: 3, is_free: true,  name: 'Wisconsin' },
  
  // PAID States (Vendor check)
  'AZ': { tier: 3, is_free: false, name: 'Arizona' },
  'AR': { tier: 3, is_free: false, name: 'Arkansas' },
  'CA': { tier: 3, is_free: false, name: 'California' },
  'CO': { tier: 3, is_free: false, name: 'Colorado' },
  'DC': { tier: 3, is_free: false, name: 'District of Columbia' },
  'HI': { tier: 3, is_free: false, name: 'Hawaii' },
  'IL': { tier: 3, is_free: false, name: 'Illinois' },
  'IN': { tier: 3, is_free: false, name: 'Indiana' },
  'IA': { tier: 3, is_free: false, name: 'Iowa' },
  'KY': { tier: 3, is_free: false, name: 'Kentucky' },
  'ME': { tier: 3, is_free: false, name: 'Maine' },
  'MO': { tier: 3, is_free: false, name: 'Missouri' },
  'NV': { tier: 3, is_free: false, name: 'Nevada' },
  'NH': { tier: 3, is_free: false, name: 'New Hampshire' },
  'NJ': { tier: 3, is_free: false, name: 'New Jersey' },
  'NM': { tier: 3, is_free: false, name: 'New Mexico' },
  'NY': { tier: 3, is_free: false, name: 'New York' },
  'NC': { tier: 3, is_free: false, name: 'North Carolina' },
  'OH': { tier: 3, is_free: false, name: 'Ohio' },
  'OK': { tier: 3, is_free: false, name: 'Oklahoma' },
  'PA': { tier: 3, is_free: false, name: 'Pennsylvania' },  // PENNDOT vendor
  'SC': { tier: 3, is_free: false, name: 'South Carolina' },
  'SD': { tier: 3, is_free: false, name: 'South Dakota' },
  'UT': { tier: 3, is_free: false, name: 'Utah' },
  'VT': { tier: 3, is_free: false, name: 'Vermont' },
  'VA': { tier: 3, is_free: false, name: 'Virginia' },
  'WV': { tier: 3, is_free: false, name: 'West Virginia' },
  'WY': { tier: 3, is_free: false, name: 'Wyoming' }
};

// ═══════════════════════════════════════════════════════════════════════════════
// CAPACITY CONFIGURATION (Agent Workload)
// ═══════════════════════════════════════════════════════════════════════════════

const CAPACITY_CONFIG = {
  MAX_CHECKS_PER_DAY: 70,         // Max tickets an agent can process if 100% MVR
  MVR_TIME_ALLOCATION_PCT: 10,    // % of agent time dedicated to MVR
  EFFECTIVE_CHECKS_PER_DAY: 7,    // 70 * 10% = 7 effective MVR checks/day/agent
  WORKING_DAYS_PER_MONTH: 21,     // Average working days
  WORKING_DAYS_PER_YEAR: 252      // 21 * 12
};

// ═══════════════════════════════════════════════════════════════════════════════
// REPORTING ASSUMPTIONS (Documented Uncertainty)
// Used in projections with Low/Mid/High scenarios
// ═══════════════════════════════════════════════════════════════════════════════

const ASSUMPTIONS = {
  // What % of monthly checks are from enrolled (continuous) vs unenrolled (one-off) drivers?
  ENROLLED_PCT: {
    low: 40, mid: 60, high: 80,
    description: 'Percentage of monthly checks that are enrolled (continuous monitoring) vs one-off',
    source: 'Unknown - requires operational data analysis'
  },
  
  // What % of enrolled profiles are actively requiring checks (not stale)?
  ACTIVE_PROFILE_PCT: {
    low: 30, mid: 50, high: 70,
    description: 'Percentage of enrolled profiles that actively require monitoring (not stale/inactive)',
    source: 'Unknown - profiles can go stale if driver leaves company'
  },
  
  // FREE state driver distribution
  FREE_STATE_PCT: {
    low: 40, mid: 45, high: 50,
    description: 'Percentage of drivers in FREE (DMV site check) states',
    source: 'Based on 19 FREE states out of 50, weighted by population'
  },
  
  // Problem tag thresholds for OPS reporting
  PROBLEM_THRESHOLDS: {
    open_hours_warning: 24,       // Flag tickets open > 24 hours
    open_hours_critical: 48,      // Critical if open > 48 hours
    pending_hours_warning: 48,    // Flag pending > 48 hours
    pending_hours_critical: 72    // Critical if pending > 72 hours
  },
  
  // Growth scenarios for projections
  GROWTH_SCENARIOS: {
    conservative: 0.10,           // 10% annual growth
    moderate: 0.25,               // 25% annual growth  
    aggressive: 0.50              // 50% annual growth
  },
  
  // Projection horizon
  PROJECTION_YEARS: 5
};

// Problem tags to flag in OPS lifecycle report
const PROBLEM_TAGS = [
  'DMV Down',
  'Cannot Process', 
  'Missing Data',
  'Discrepancy',
  'Record Not Found',
  'System Error',
  'Pending Verification'
];

/**
 * Truncate text to fit within Google Sheets cell character limit
 * Preserves beginning of text and adds truncation marker
 * @param {string} text - Text to truncate
 * @param {number} maxLength - Maximum length (default: CELL_LIMITS.SAFE_LIMIT)
 * @return {string} Truncated text with marker if needed
 */
function truncateForCell(text, maxLength = CELL_LIMITS.SAFE_LIMIT) {
  if (!text || typeof text !== 'string') return text || '';
  if (text.length <= maxLength) return text;
  
  const truncationMarker = '\n\n[...TRUNCATED - ' + (text.length - maxLength + 50) + ' chars removed...]';
  return text.substring(0, maxLength - 50) + truncationMarker;
}

/**
 * Truncate JSON string, optionally removing large nested arrays first
 * @param {Object} obj - Object to stringify and truncate
 * @param {Array<string>} excludeKeys - Keys to remove before stringifying (e.g., ['conversations'])
 * @param {number} maxLength - Maximum length
 * @return {string} Truncated JSON string
 */
function truncateJsonForCell(obj, excludeKeys = [], maxLength = CELL_LIMITS.JSON_FIELD_LIMIT) {
  if (!obj) return '';
  
  // Create shallow copy without excluded keys
  const cleanObj = { ...obj };
  excludeKeys.forEach(key => delete cleanObj[key]);
  
  let jsonStr = JSON.stringify(cleanObj);
  
  if (jsonStr.length > maxLength) {
    const marker = '..."_truncated":true}';
    jsonStr = jsonStr.substring(0, maxLength - marker.length) + marker;
  }
  
  return jsonStr;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MVR EMAIL PATTERNS
// ═══════════════════════════════════════════════════════════════════════════════

const MVR_PATTERNS = {
  SUBJECT_PREFIX: "One Off Continuous MVR",
  SUSPENSION_CHECK: "Suspension Check",
  RECHECK: "Recheck",
  SENTINEL_REPORT: "Sentinel Report",
  
  // Standard format: "One Off Continuous MVR - (Suspension Check|Recheck) - {partner} - {turnId}"
  SUBJECT_REGEX: /One Off Continuous MVR - (Suspension Check|Recheck) - (.+?) - (.+?)$/i,
  
  // InformData/Sentinel format: "[InformData] One Off Continuous MVR - Sentinel Report - {turnId}"
  // No partner in subject - must extract from notes
  INFORM_SUBJECT_REGEX: /One Off Continuous MVR - (Sentinel Report) - ([A-Z]\d+)$/i,
  
  // Vendor group prefix patterns (e.g., [Certn], [PennDOT], [InformData])
  VENDOR_PREFIX_REGEX: /^\[(\w+)\]\s*/i,
  
  // Alternative continuous monitoring patterns
  CERTN_PATTERN: /\[Certn\]\s*Continuous Monitoring Due/i,
  PENNDOT_PATTERN: /\[PennDOT\]\s*Continuous Monitoring Due/i,
  INFORM_PATTERN: /\[InformData\]\s*One Off Continuous MVR/i
};

// ═══════════════════════════════════════════════════════════════════════════════
// NOTES FIELD EXTRACTION PATTERNS (Fallback for InformData/Sentinel tickets)
// Used when subject/description parsing fails - extracts from All Notes Text
// ═══════════════════════════════════════════════════════════════════════════════

const NOTES_FIELD_PATTERNS = {
  // InformData uses asterisk-wrapped values: "Turn ID: *C8078635470*"
  TURN_ID: /Turn ID:\s*\*?([A-Z]\d+)\*?/i,
  PARTNER: /Partner:\s*\*?([^*\n]+)\*?/i,
  DL_STATE: /DL State:\s*\*?([A-Z]{2})\*?/i,
  DL_NUMBER: /DL Number:\s*([^\n*]+)/i,
  ORDER_ID: /Order ID:\s*\*?(\d+)\*?/i,
  TIER: /Tier:\s*\*?(Tier[^\n*]+)\*?/i,
  CADENCE_DAYS: /Cadence Days:\s*\*?(\d+)\s*days?\*?/i,
  PROCESS_TYPE: /Process Type:\s*\*?([^*\n]+)\*?/i,
  VENDOR: /Vendor:\s*\*?(\w+)\*?/i,
  SOURCE: /Source:\s*\*?([^|*\n]+)\*?/i,
  DAYS_SINCE_LAST: /Days Since Last Check:\s*\*?(\d+)\s*days?\*?/i,
  LAST_CHECK_DATE: /Last Check Date:\s*\*?([^\n*]+)\*?/i,
  TA_URL: /TA URL:\s*(https:\/\/[^\s\n]+)/i
};

// ═══════════════════════════════════════════════════════════════════════════════
// MVR SUBJECT PATTERNS (For Raw Pull - All Continuous MVR Ticket Types)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * All subject patterns that identify MVR/Continuous Monitoring tickets
 * Based on analysis of actual Freshdesk ticket subjects
 */
const MVR_SUBJECT_PATTERNS = [
  /\[Certn\]\s*Continuous Monitoring Due/i,           // [Certn] Continuous Monitoring Due
  /\[PennDOT\]\s*Continuous Monitoring Due/i,         // [PennDOT] Continuous Monitoring Due
  /\[InformData\]\s*One Off Continuous MVR/i,         // [InformData] One Off Continuous MVR - Sentinel Report
  /One Off Continuous MVR\s*-\s*Suspension Check/i,   // One Off Continuous MVR - Suspension Check
  /One Off Continuous MVR\s*-\s*Recheck/i,            // One Off Continuous MVR - Recheck
  /One Off Continuous MVR\s*-\s*Sentinel Report/i,    // One Off Continuous MVR - Sentinel Report
  /MVR Suspension Check/i,                            // MVR Suspension Check
  /MVR Recheck/i,                                     // MVR Recheck
  /Continuous MVR/i                                   // Continuous MVR (catch-all)
];

/**
 * Check if a subject line matches any MVR pattern
 * @param {string} subject - Ticket subject line
 * @return {boolean} True if this is an MVR ticket
 */
function isMVRSubject(subject) {
  if (!subject) return false;
  return MVR_SUBJECT_PATTERNS.some(pattern => pattern.test(subject));
}

// ═══════════════════════════════════════════════════════════════════════════════
// VENDOR GROUP CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const VENDOR_GROUPS = {
  CERTN: "CERTN",
  PENNDOT: "PENNDOT",
  INFORM: "INFORM",
  INFORMDATA: "INFORMDATA",  // Sentinel System
  UNKNOWN: "UNKNOWN"          // Could not determine from ticket content
};

// ═══════════════════════════════════════════════════════════════════════════════
// STATE-BASED VENDOR MAPPING
// SC cadence is defined by STATE, not vendor. Vendor is determined by state.
// Based on Heroku MVR queue analysis (Dec 2025)
// ═══════════════════════════════════════════════════════════════════════════════

// PennDOT applies to Pennsylvania only
const PENNDOT_STATE = "PA";

// States that use CERTN (from Heroku MVR data analysis)
const CERTN_STATES = ["MO", "IL"];

// All other states use INFORMDATA (Sentinel System) - 77.5% of volume
// This is the default when a state is known but not PA, MO, or IL

// Legacy placeholder - no longer used
const INFORM_STATES = [];

// ═══════════════════════════════════════════════════════════════════════════════
// OUTCOME CLASSIFICATION
// ═══════════════════════════════════════════════════════════════════════════════

const OUTCOME_TYPES = {
  SUSPENSION_CONFIRMED: "Suspension Confirmed",
  CLEAR: "Clear",
  DMV_UNAVAILABLE: "DMV Unavailable",
  CANNOT_PROCESS: "Cannot Process",
  STILL_PROCESSING: "Still Processing",
  UNKNOWN: "Unknown"
};

const OUTCOME_SOURCES = {
  TAG: "Tag",
  CONVERSATION: "Conversation",
  OVERRIDE: "Manual Override",
  AUTO: "Auto-Classified"
};

// ═══════════════════════════════════════════════════════════════════════════════
// TAG OUTCOME MAPPING CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

// Tag Outcome Mapping sheet headers
const TAG_MAPPING_HEADERS = [
  "Tag Pattern",      // The pattern to match (text or regex)
  "Match Type",       // "contains" or "regex"
  "Outcome Type",     // Maps to OUTCOME_TYPES value
  "Priority",         // Lower number = higher priority (1-99)
  "Is Billable",      // TRUE/FALSE - affects finance reporting
  "Is Active",        // TRUE/FALSE - inactive mappings are skipped
  "Notes"             // Optional description
];

// Default tag mappings (used to seed the sheet)
// SINGLE SOURCE OF TRUTH - All patterns, billable logic, priorities
// Format: [Pattern, Match Type, Outcome Type, Priority, Is Billable, Is Active, Notes]
const DEFAULT_TAG_MAPPINGS = [
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 1-10: VIOLATIONS (Billable = YES - Customer pays)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["suspended", "contains", "Suspension Confirmed", 1, true, true, "License suspended - billable violation"],
  ["DL Expired", "contains", "Expired License", 2, true, true, "License expired - billable violation"],
  ["expired", "contains", "Expired License", 3, true, true, "Expired license - billable violation"],
  ["DL NOT VALID", "contains", "Invalid License", 4, true, true, "Invalid license - billable violation"],
  ["not valid", "contains", "Invalid License", 5, true, true, "License not valid - billable violation"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 11-20: STATUS TAGS (Billable = NO - Already in system, no work done)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["valid", "contains", "Clear", 11, false, true, "Driver already valid - no work needed"],
  ["active", "contains", "Clear", 12, false, true, "Driver already active - no work needed"],
  ["clear", "contains", "Clear", 13, false, true, "MVR already clear - no work needed"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 21-30: SYSTEM ISSUES (Billable = NO - DMV/vendor problems)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["dmv[\\s_-]?down", "regex", "DMV Unavailable", 21, false, true, "DMV system down - we eat cost"],
  ["state[\\s_-]?down", "regex", "DMV Unavailable", 22, false, true, "State system down - we eat cost"],
  ["[A-Z]{2}\\s+(down|unavailable)", "regex", "DMV Unavailable", 23, false, true, "State code down (e.g., TX down)"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 31-40: CANNOT PROCESS (Billable = NO - Data/processing issues)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["no[\\s_-]?dl", "regex", "Cannot Process", 31, false, true, "No driver license - cannot process"],
  ["discrepancy", "contains", "Cannot Process", 32, false, true, "Data discrepancy - cannot process"],
  ["discrepancies", "contains", "Cannot Process", 33, false, true, "Multiple discrepancies"],
  ["missing", "contains", "Cannot Process", 34, false, true, "Missing data - cannot process"],
  ["cannot[\\s_-]?process", "regex", "Cannot Process", 35, false, true, "Cannot process request"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 41-50: SPECIAL CASES (Billable = NO - Applicant/vendor issues)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["Record not found", "contains", "Record Not Found", 41, false, true, "Record not found - we eat cost"],
  ["Withdrawn", "contains", "Withdrawn", 42, false, true, "Applicant withdrew - no billing"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 51-60: DELIVERY/UPLOAD (Billable = YES - Report delivered to customer)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["Updated", "contains", "Clear", 51, true, true, "Report uploaded to employer - billable"],
  ["uploaded", "contains", "Clear", 52, true, true, "Report uploaded - billable"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 61-70: STATUS INDICATORS (Billable = NO - Operations status only)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["approved", "contains", "Clear", 61, false, true, "Status tag only - final approval, not billable"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 71-80: NO CHANGES FOUND (Billable = NO - We eat cost, wasted check)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["same information", "contains", "Clear", 71, false, true, "No changes found - we eat the cost"],
  ["same info", "contains", "Clear", 72, false, true, "No changes found - we eat the cost"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 81-90: PENDING/PROCESSING (Billable = NO - Still in progress)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["Pending.*", "regex", "Pending", 81, false, true, "Pending state (regex) - not billable"],
  ["pending", "contains", "Pending", 82, false, true, "Awaiting information - not billable"],
  ["emailed", "contains", "Still Processing", 83, false, true, "Waiting for applicant info - not billable"],
  ["Review identity", "contains", "Still Processing", 84, false, true, "Identity verification required - not billable"],
  ["Verifying", "contains", "Still Processing", 85, false, true, "Verification in progress - not billable"],
  ["In Progress", "contains", "Still Processing", 86, false, true, "Processing - not billable"],
  
  // ═══════════════════════════════════════════════════════════════════════════════
  // PRIORITY 91-99: LOWEST PRIORITY (Billable = NO - Just initiated, not delivered)
  // ═══════════════════════════════════════════════════════════════════════════════
  ["requested", "contains", "Still Processing", 91, false, true, "Check requested only - not delivered, not billable"]
];

// ═══════════════════════════════════════════════════════════════════════════════
// TICKET STATUS CODES (Freshdesk Standard)
// ═══════════════════════════════════════════════════════════════════════════════

const TICKET_STATUS = {
  OPEN: 2,
  PENDING: 3,
  RESOLVED: 4,
  CLOSED: 5
};

const STATUS_NAMES = {
  2: "Open",
  3: "Pending",
  4: "Resolved",
  5: "Closed"
};

// ═══════════════════════════════════════════════════════════════════════════════
// PRIORITY CODES (Freshdesk Standard)
// ═══════════════════════════════════════════════════════════════════════════════

const TICKET_PRIORITY = {
  LOW: 1,
  MEDIUM: 2,
  HIGH: 3,
  URGENT: 4
};

const PRIORITY_NAMES = {
  1: "Low",
  2: "Medium",
  3: "High",
  4: "Urgent"
};

// ═══════════════════════════════════════════════════════════════════════════════
// API CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const API_CONFIG = {
  PER_PAGE: 100,                  // Tickets per API request
  MAX_PAGES: 50,                  // Safety limit for pagination
  RATE_LIMIT_DELAY: 300,          // Milliseconds between requests (reduced from 500ms for speed)
  RETRY_ATTEMPTS: 3,              // Max retries on API failure
  RETRY_DELAY_BASE: 1000          // Base delay for exponential backoff
};

// ═══════════════════════════════════════════════════════════════════════════════
// BUSINESS HOURS CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════

const BUSINESS_HOURS = {
  START_HOUR: 9,                  // 9 AM
  END_HOUR: 18,                   // 6 PM
  WORK_DAYS: [1, 2, 3, 4, 5],    // Monday-Friday (0=Sunday, 6=Saturday)
  TIMEZONE: "America/Chicago"     // CST/CDT
};

// ═══════════════════════════════════════════════════════════════════════════════
// PERFORMANCE THRESHOLDS (24-HOUR SLA)
// ═══════════════════════════════════════════════════════════════════════════════

const PERFORMANCE_THRESHOLDS = {
  RESOLUTION_STANDARD: 24,        // Hours - Standard SLA for One-Off tickets
  EXCELLENT: 6,                   // Hours - Excellent performance
  GOOD: 12,                       // Hours - Good performance
  WARNING: 18,                    // Hours - Warning threshold
  CRITICAL: 24,                   // Hours - Critical (at deadline)
  OVERDUE: 24                     // Hours - Overdue (SLA breach)
};

// ═══════════════════════════════════════════════════════════════════════════════
// SCRIPT PROPERTIES SETUP
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Setup Script Properties
 * Run this function once to configure Freshdesk API credentials
 */
function setupScriptProperties() {
  const ui = SpreadsheetApp.getUi();
  
  const apiKeyResponse = ui.prompt(
    'Freshdesk API Key',
    'Enter your Freshdesk API Key:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (apiKeyResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Setup cancelled');
    return;
  }
  
  const domainResponse = ui.prompt(
    'Freshdesk Domain',
    'Enter your Freshdesk domain (e.g., "turnhq" for turnhq.freshdesk.com):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (domainResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Setup cancelled');
    return;
  }
  
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    FRESHDESK_API_KEY: apiKeyResponse.getResponseText().trim(),
    FRESHDESK_DOMAIN: domainResponse.getResponseText().trim()
  });
  
  Logger.log('✅ Freshdesk credentials saved successfully');
  ui.alert('✅ Setup Complete', 'Freshdesk credentials saved successfully!', ui.ButtonSet.OK);
}

/**
 * Get Freshdesk credentials from Script Properties
 * @return {Object} Object with apiKey and domain
 */
function getFreshdeskCredentials() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('FRESHDESK_API_KEY');
  const domain = props.getProperty('FRESHDESK_DOMAIN');
  
  if (!apiKey || !domain) {
    throw new Error('Freshdesk credentials not configured. Run setupScriptProperties() first.');
  }
  
  return { apiKey, domain };
}

// ═══════════════════════════════════════════════════════════════════════════════
// SHEET HEADERS DEFINITION
// ═══════════════════════════════════════════════════════════════════════════════

const HISTORY_HEADERS = [
  // Original columns (1-25)
  "Ticket ID",
  "Email Subject Line",
  "Request Type (SC/RC)",
  "Partner Name",
  "Turn ID",
  "Status Code",
  "Status",
  "Priority Code",
  "Priority Level",
  "Date Created",
  "Last Updated",
  "Date Resolved",
  "Age (Hours)",
  "Resolution Time (Hours)",
  "Agent ID",
  "Assigned Agent",
  "Group ID",
  "Support Group",
  "Requester ID",
  "Company ID",
  "Tags",
  "Source Code",
  "Ticket Source",
  "SLA Status",
  "Last Data Pull",
  // Workflow Origin columns (26-34)
  "Vendor Group",
  "DL State",
  "DL Number",
  "Days Since Last",
  "Last Check Date",
  "Cadence Days",
  "Tier",
  "Enrollment Type",
  "TA URL",
  // Outcome columns (35-40)
  "MVR Outcome",
  "Outcome Source",
  "Outcome Date",
  "Resolution Notes",
  "Conversation Count",
  "Last Responder",
  // Override columns (41-42)
  "Outcome Override",
  "Override Reason",
  // Change tracking (43)
  "Change Log"
  // Custom fields will be added dynamically (44+)
];

// Override Audit Log Headers
const AUDIT_LOG_HEADERS = [
  "Timestamp",
  "Ticket ID",
  "Turn ID",
  "Partner",
  "Previous Outcome",
  "New Override",
  "Override Reason",
  "User Email",
  "User Name"
];

// Custom field prefix for dynamic columns
const CUSTOM_FIELD_PREFIX = "CF: ";

// Known custom field mappings (Freshdesk field key -> display name)
const CUSTOM_FIELD_MAP = {
  cf_partner_name: "CF: Partner Name",
  cf_turn_id: "CF: Turn ID",
  cf_driver_name: "CF: Driver Name",
  cf_dl_state: "CF: DL State",
  cf_dl_number: "CF: DL Number"
};

const PARTNER_SUMMARY_HEADERS = [
  "Partner Name",
  "Total Tickets",
  "Suspension Checks (SC)",
  "Rechecks (RC)",
  "Open Tickets",
  "Pending Tickets",
  "Resolved Tickets",
  "Closed Tickets",
  "Avg Resolution Time (Hours)",
  "Avg Ticket Age (Hours)",
  "SLA Met Count",
  "SLA Breach Count",
  "Oldest Ticket Date",
  "Newest Ticket Date",
  "Report Generated"
];

// ═══════════════════════════════════════════════════════════════════════════════
// UTILITY FUNCTIONS - Date & Time Calculations
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Adjust date to next business day/hour if created after hours
 * Business hours: Monday-Friday, 9 AM - 6 PM CST
 * If ticket created outside business hours, adjust to 9:00 AM next business day
 * @param {Date} date - Original creation date
 * @return {Date} Adjusted date (first minute of next business day if after hours)
 */
function adjustToBusinessHours(date) {
  if (!date || !(date instanceof Date)) return date;
  
  const adjusted = new Date(date);
  const dayOfWeek = adjusted.getDay(); // 0=Sunday, 6=Saturday
  const hour = adjusted.getHours();
  
  // Check if it's a weekend
  if (dayOfWeek === 0) {
    // Sunday -> Move to Monday 9 AM
    adjusted.setDate(adjusted.getDate() + 1);
    adjusted.setHours(BUSINESS_HOURS.START_HOUR, 0, 0, 0);
    return adjusted;
  }
  
  if (dayOfWeek === 6) {
    // Saturday -> Move to Monday 9 AM
    adjusted.setDate(adjusted.getDate() + 2);
    adjusted.setHours(BUSINESS_HOURS.START_HOUR, 0, 0, 0);
    return adjusted;
  }
  
  // It's a weekday - check if before/after business hours
  if (hour < BUSINESS_HOURS.START_HOUR) {
    // Before 9 AM -> Set to 9 AM same day
    adjusted.setHours(BUSINESS_HOURS.START_HOUR, 0, 0, 0);
    return adjusted;
  }
  
  if (hour >= BUSINESS_HOURS.END_HOUR) {
    // After 6 PM -> Move to 9 AM next business day
    let nextDay = new Date(adjusted);
    nextDay.setDate(nextDay.getDate() + 1);
    nextDay.setHours(BUSINESS_HOURS.START_HOUR, 0, 0, 0);
    
    // If next day is Saturday, jump to Monday
    if (nextDay.getDay() === 6) {
      nextDay.setDate(nextDay.getDate() + 2);
    }
    // If next day is Sunday, jump to Monday
    else if (nextDay.getDay() === 0) {
      nextDay.setDate(nextDay.getDate() + 1);
    }
    
    return nextDay;
  }
  
  // Within business hours - no adjustment needed
  return adjusted;
}

/**
 * Get the resolved date for a ticket
 * Returns updated_at for RESOLVED/CLOSED tickets, null otherwise
 * @param {Object} ticket - Freshdesk ticket object
 * @return {Date|null} Resolved date or null
 */
function getResolvedDate(ticket) {
  if (ticket.status === TICKET_STATUS.RESOLVED || ticket.status === TICKET_STATUS.CLOSED) {
    return new Date(ticket.updated_at);
  }
  return null;
}

/**
 * Calculate age in hours from creation to now
 * Uses business hours adjustment for accurate measurement
 * @param {Date} createdAt - Ticket creation date
 * @return {number} Age in hours
 */
function calculateAgeHours(createdAt) {
  const adjustedCreated = adjustToBusinessHours(createdAt);
  const now = new Date();
  return (now - adjustedCreated) / 3600000; // Convert ms to hours
}

/**
 * Calculate resolution time in hours
 * Only for RESOLVED/CLOSED tickets
 * @param {Object} ticket - Freshdesk ticket object
 * @return {number|null} Resolution time in hours or null if not resolved
 */
function calculateResolutionHours(ticket) {
  if (ticket.status !== TICKET_STATUS.RESOLVED && ticket.status !== TICKET_STATUS.CLOSED) {
    return null;
  }
  
  const createdAt = new Date(ticket.created_at);
  const resolvedAt = new Date(ticket.updated_at);
  const adjustedCreated = adjustToBusinessHours(createdAt);
  
  return (resolvedAt - adjustedCreated) / 3600000; // Convert ms to hours
}

/**
 * Classify SLA status based on resolution time or current age
 * @param {Object} ticket - Freshdesk ticket object
 * @param {number} ageHours - Current age in hours
 * @param {number|null} resolutionHours - Resolution time in hours (null if not resolved)
 * @return {string} SLA classification: "MET", "BREACH", "PEND", "Ongoing"
 */
function classifySLA(ticket, ageHours, resolutionHours) {
  if (ticket.status === TICKET_STATUS.RESOLVED || ticket.status === TICKET_STATUS.CLOSED) {
    // For resolved tickets, check resolution time
    return (resolutionHours <= PERFORMANCE_THRESHOLDS.RESOLUTION_STANDARD) ? "MET" : "BREACH";
  } else if (ticket.status === TICKET_STATUS.OPEN) {
    // For open tickets, check current age
    return (ageHours > PERFORMANCE_THRESHOLDS.RESOLUTION_STANDARD) ? "BREACH" : "PEND";
  } else if (ticket.status === TICKET_STATUS.PENDING) {
    // PENDING status - ongoing work with customer
    return "Ongoing";
  }
  return "N/A";
}

// ═══════════════════════════════════════════════════════════════════════════════
// HEADER UTILITY FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create column index mapper from headers
 * @param {Array} headers - Array of header names
 * @return {Object} Map of header name to column index
 */
function createColumnMapper(headers) {
  const mapper = {};
  headers.forEach((header, index) => {
    mapper[header] = index;
  });
  return mapper;
}

/**
 * Get or create sheet with headers
 * @param {string} sheetName - Name of sheet to get/create
 * @param {Array} headers - Array of header names
 * @return {Sheet} Google Sheets Sheet object
 */
function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // Set headers
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row - blue background, white text, bold, centered
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#4a86e8');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      headerRange.setVerticalAlignment('middle');
      headerRange.setWrap(true);
      
      // Set row height for headers
      sheet.setRowHeight(1, 50);
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Auto-resize columns
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
    }
    
    Logger.log(`✅ Created sheet: ${sheetName}`);
  }
  
  return sheet;
}
