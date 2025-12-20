/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * ASSUMPTIONS SHEET - Central Configuration & Parameters
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * Creates and manages the Assumptions sheet that serves as a single source of
 * truth for all configurable parameters. Reports reference these via named ranges.
 * 
 * Categories:
 * - VENDOR COSTS: Per-check pricing for CERTN, INFORMDATA, PENNDOT
 * - CAPACITY: Agent throughput and FTE calculations
 * - TIERS: Cadence days and checks per year
 * - THRESHOLDS: SLA and escalation timing
 * - GROWTH: Projection scenarios
 * - UNCERTAINTY: Low/Mid/High range parameters
 */

// ═══════════════════════════════════════════════════════════════════════════════
// PROFESSIONAL COLOR PALETTE
// ═══════════════════════════════════════════════════════════════════════════════
const COLORS = {
  // Primary
  NAVY: '#1a365d',
  SLATE: '#334155',
  WHITE: '#ffffff',
  
  // Accents
  ROYAL_BLUE: '#2563eb',
  EMERALD: '#059669',
  AMBER: '#d97706',
  RED: '#dc2626',
  
  // Backgrounds
  LIGHT_GRAY: '#f8fafc',
  MEDIUM_GRAY: '#e2e8f0',
  BORDER_GRAY: '#cbd5e1',
  
  // Category Headers
  COST_HEADER: '#7c3aed',      // Purple for costs
  CAPACITY_HEADER: '#0891b2',   // Cyan for capacity
  TIER_HEADER: '#059669',       // Green for tiers
  THRESHOLD_HEADER: '#ea580c',  // Orange for thresholds
  GROWTH_HEADER: '#2563eb',     // Blue for growth
  UNCERTAINTY_HEADER: '#64748b' // Gray for uncertainty
};

// ═══════════════════════════════════════════════════════════════════════════════
// SECTION EXPLANATIONS - Displayed in the sheet
// ═══════════════════════════════════════════════════════════════════════════════
const SECTION_EXPLANATIONS = {
  VENDOR_COSTS: {
    title: 'How Vendor Costs Affect Your Reports',
    lines: [
      '• These costs are used to calculate total spend in Finance Dashboard and CEO Dashboard.',
      '• Changing vendor costs updates all cost projections, monthly totals, and ROI calculations.',
      '• Incident Cost × Prevention Rate = estimated savings from detecting suspended drivers.',
      '• Example: $50,000 incident cost × 10% prevention = $5,000 value per detection.',
    ]
  },
  CAPACITY: {
    title: 'How Capacity Settings Affect Projections',
    lines: [
      '• Max Checks × Time Allocation = Effective daily capacity per agent.',
      '• These values drive FTE requirements in Annual Projection report.',
      '• Increasing Time Allocation reduces FTE needs but increases agent MVR burden.',
      '• Working days affect monthly/yearly capacity calculations for planning.',
    ]
  },
  TIERS: {
    title: 'How Tier Configuration Affects Workload',
    lines: [
      '• Cadence = days between checks. Shorter cadence = more checks per driver.',
      '• Tier 1 (high-risk states): More frequent checks, higher workload.',
      '• Changing cadence directly impacts volume projections and capacity needs.',
      '• Checks per Year = 365 ÷ Cadence. Used in growth projections.',
    ]
  },
  THRESHOLDS: {
    title: 'How SLA Thresholds Affect Dashboards',
    lines: [
      '• SLA Target: Tickets resolved faster than this are "SLA Met" (green).',
      '• Warning/Critical levels trigger color coding in CEO Dashboard insights.',
      '• Lowering thresholds makes metrics stricter (fewer green, more warnings).',
      '• These drive the escalation alerts and operational health indicators.',
    ]
  },
  GROWTH: {
    title: 'How Growth Rates Affect 5-Year Projections',
    lines: [
      '• Base Profiles × Growth Rate = next year\'s projected volume.',
      '• Conservative/Moderate/Aggressive give low/mid/high scenarios.',
      '• Higher growth rates show earlier hiring triggers and capacity gaps.',
      '• Projection Horizon sets how many years forward to project.',
    ]
  },
  UNCERTAINTY: {
    title: 'How Uncertainty Ranges Affect Estimates',
    lines: [
      '• Low/Mid/High provide confidence bands for projections.',
      '• Enrolled % = what portion of drivers are in continuous monitoring.',
      '• Active % = what portion of enrolled actually need checks.',
      '• FREE State % = portion avoiding vendor costs (using DMV direct).',
      '• Tier Distribution affects weighted average check frequency.',
    ]
  },
  LABOR: {
    title: 'How Labor Costs Affect ROI & Capacity',
    lines: [
      '• Agent Monthly Salary is the base compensation for a 5-day, 8-hour shift.',
      '• Hourly Rate = Salary ÷ (Working Days × Hours per Shift).',
      '• Labor Cost per Ticket = Hourly Rate ÷ Effective Checks per Hour.',
      '• Total Cost = Vendor Cost + Labor Cost per ticket.',
      '• Used in ROI calculations and capacity planning projections.',
    ]
  }
};

// ═══════════════════════════════════════════════════════════════════════════════
// ASSUMPTIONS DATA STRUCTURE
// ═══════════════════════════════════════════════════════════════════════════════
const ASSUMPTIONS_DATA = {
  UNIVERSE: {
    header: '🌐 ANALYSIS UNIVERSE',
    color: '#6366f1',
    explanation: {
      title: 'Understanding the Data Scope',
      lines: [
        '• This analysis covers ONLY drivers flagged for recheck in continuous monitoring.',
        '• Total Monitored Profiles = the "N" (denominator) for all rate calculations.',
        '• MVR_Ticket_History = subset of monitored profiles that triggered a recheck.',
        '• NOT the full universe of all workers/profiles in the system.',
        '• Rates and projections are limited to this monitored population scope.',
      ]
    },
    params: [
      { name: 'Total Monitored Profiles', value: 15000, unit: 'profiles', description: 'Total drivers enrolled in continuous monitoring (our N)' },
      { name: 'Recheck Tickets (from data)', value: 0, unit: 'tickets', description: '= COUNT of MVR_Ticket_History rows (auto-calculated)', formula: true },
      { name: 'Recheck Rate (actual)', value: 0, unit: '%', description: '= Tickets ÷ Monitored Profiles (auto-calculated)', formula: true }
    ]
  },
  VENDOR_COSTS: {
    header: '💰 VENDOR COSTS',
    color: COLORS.COST_HEADER,
    explanation: SECTION_EXPLANATIONS.VENDOR_COSTS,
    params: [
      { name: 'Cost per Check - CERTN', value: 3.50, unit: '$ per check', description: 'CERTN MVR check cost (MO, IL)' },
      { name: 'Cost per Check - INFORMDATA', value: 2.75, unit: '$ per check', description: 'InformData/Sentinel MVR check cost' },
      { name: 'Cost per Check - PENNDOT', value: 0.00, unit: '$ per check', description: 'Pennsylvania DMV direct (FREE)' },
      { name: 'Incident Cost Estimate', value: 50000, unit: 'USD', description: 'Average cost per driving incident for ROI' },
      { name: 'Prevention Rate', value: 0.10, unit: '%', description: 'Estimated incidents prevented by detection' }
    ]
  },
  CAPACITY: {
    header: '⚙️ CAPACITY PLANNING',
    color: COLORS.CAPACITY_HEADER,
    explanation: SECTION_EXPLANATIONS.CAPACITY,
    params: [
      { name: 'Max Checks per Day @100%', value: 70, unit: 'tickets/day', description: 'Agent capacity if 100% MVR time' },
      { name: 'MVR Time Allocation', value: 0.10, unit: '%', description: 'Actual % of agent time on MVR' },
      { name: 'Effective Checks per Day', value: 7, unit: 'tickets/day', description: 'Calculated: Max × Allocation', formula: true },
      { name: 'Working Days per Month', value: 21, unit: 'days', description: 'Standard month assumption' },
      { name: 'Working Days per Year', value: 252, unit: 'days', description: '21 × 12 months' },
      { name: 'Current FTE', value: 1, unit: 'FTE', description: 'Current MVR headcount' }
    ]
  },
  LABOR: {
    header: '💵 LABOR COSTS (Derived from 10% Allocation)',
    color: '#be185d',
    explanation: {
      title: 'How Labor Cost is Calculated',
      lines: [
        '• Agents dedicate 10% of their time to MVR recheck processing.',
        '• Monthly MVR Labor = Agent Salary × 10% allocation.',
        '• Cost per Ticket = Monthly MVR Labor ÷ Tickets processed that month.',
        '• This is DERIVED from actual data, not assumed per-ticket.',
        '• Formula: (Salary × 0.10) ÷ (Tickets from MVR_Ticket_History)',
      ]
    },
    params: [
      { name: 'Agent Monthly Salary', value: 1500, unit: 'USD/month', description: 'Full monthly compensation per agent' },
      { name: 'MVR Time Allocation', value: 0.10, unit: '%', description: 'Agents dedicate 10% of time to MVR rechecks' },
      { name: 'Monthly MVR Labor Cost', value: 150, unit: 'USD/month', description: '= Salary × 10% = $150/mo per agent for MVR work' },
      { name: 'Number of MVR Agents', value: 1, unit: 'agents', description: 'Agents handling MVR rechecks' },
      { name: 'Total Monthly Labor', value: 150, unit: 'USD/month', description: '= Monthly MVR Labor × Agents' },
      { name: 'Labor Cost per Ticket', value: 0, unit: 'USD/ticket', description: '= Total Labor ÷ Tickets (auto from analysis)', formula: true }
    ]
  },
  PARTNER_PRICING: {
    header: '🏷️ PRICING MODEL',
    color: '#0d9488',
    explanation: {
      title: 'How Pricing Affects Revenue & ROI',
      lines: [
        '• Driver Subscription: Monthly fee per recheck ticket (ALL tickets pay this).',
        '• Check Price: Additional fee only for billable checks (uploaded/violations).',
        '• Revenue = (All Tickets × Subscription) + (Billable Tickets × Check Price).',
        '• Non-billable tickets: We absorb vendor cost but still collect subscription.',
        '• Each row in MVR_Ticket_History = 1 recheck ticket to process.',
      ]
    },
    params: [
      { name: 'Driver Monthly Subscription', value: 2.00, unit: '$ per ticket', description: 'Fee per recheck ticket (all tickets pay this)' },
      { name: 'Price per Billable Check', value: 5.00, unit: '$ per check', description: 'Additional fee only for billable checks (uploaded)' },
      { name: 'Market Price per Check', value: 15.00, unit: '$ per check', description: 'Industry standard all-in price for comparison' }
    ]
  },
  TIERS: {
    header: '🎯 TIER CONFIGURATION',
    color: COLORS.TIER_HEADER,
    explanation: SECTION_EXPLANATIONS.TIERS,
    params: [
      { name: 'Tier 1 Cadence', value: 60, unit: 'days', description: 'High-risk states check frequency' },
      { name: 'Tier 1 Checks per Year', value: 6, unit: 'checks/year', description: '365 ÷ 60 rounded' },
      { name: 'Tier 2 Cadence', value: 90, unit: 'days', description: 'Medium-risk states check frequency' },
      { name: 'Tier 2 Checks per Year', value: 4, unit: 'checks/year', description: '365 ÷ 90 rounded' },
      { name: 'Tier 3 Cadence', value: 180, unit: 'days', description: 'Standard states check frequency' },
      { name: 'Tier 3 Checks per Year', value: 2, unit: 'checks/year', description: '365 ÷ 180 rounded' }
    ]
  },
  THRESHOLDS: {
    header: '⏱️ SLA & ESCALATION',
    color: COLORS.THRESHOLD_HEADER,
    explanation: SECTION_EXPLANATIONS.THRESHOLDS,
    params: [
      { name: 'SLA Target', value: 24, unit: 'hours', description: 'Resolution deadline' },
      { name: 'Resolution Excellent', value: 6, unit: 'hours', description: 'Performance tier: Excellent' },
      { name: 'Resolution Good', value: 12, unit: 'hours', description: 'Performance tier: Good' },
      { name: 'Resolution Warning', value: 18, unit: 'hours', description: 'Performance tier: Warning' },
      { name: 'Open Ticket Warning', value: 24, unit: 'hours', description: 'Flag open tickets' },
      { name: 'Open Ticket Critical', value: 48, unit: 'hours', description: 'Escalate open tickets' },
      { name: 'Pending Ticket Warning', value: 48, unit: 'hours', description: 'Flag pending tickets' },
      { name: 'Pending Ticket Critical', value: 72, unit: 'hours', description: 'Escalate pending tickets' }
    ]
  },
  GROWTH: {
    header: '📈 GROWTH SCENARIOS',
    color: COLORS.GROWTH_HEADER,
    explanation: SECTION_EXPLANATIONS.GROWTH,
    params: [
      { name: 'Base Total Profiles', value: 10000, unit: 'profiles', description: 'Current enrolled driver base' },
      { name: 'Conservative Growth Rate', value: 0.10, unit: '% annual', description: 'Minimal expansion scenario' },
      { name: 'Moderate Growth Rate', value: 0.25, unit: '% annual', description: 'Target growth scenario' },
      { name: 'Aggressive Growth Rate', value: 0.50, unit: '% annual', description: 'Stretch goal scenario' },
      { name: 'Projection Horizon', value: 5, unit: 'years', description: 'Long-term projection period' }
    ]
  },
  UNCERTAINTY: {
    header: '📊 UNCERTAINTY RANGES',
    color: COLORS.UNCERTAINTY_HEADER,
    explanation: SECTION_EXPLANATIONS.UNCERTAINTY,
    params: [
      { name: 'Enrolled Profile % - Low', value: 0.40, unit: '%', description: 'Low estimate: % in continuous monitoring' },
      { name: 'Enrolled Profile % - Mid', value: 0.60, unit: '%', description: 'Mid estimate: % in continuous monitoring' },
      { name: 'Enrolled Profile % - High', value: 0.80, unit: '%', description: 'High estimate: % in continuous monitoring' },
      { name: 'Active Profile % - Low', value: 0.30, unit: '%', description: 'Low estimate: % requiring checks' },
      { name: 'Active Profile % - Mid', value: 0.50, unit: '%', description: 'Mid estimate: % requiring checks' },
      { name: 'Active Profile % - High', value: 0.70, unit: '%', description: 'High estimate: % requiring checks' },
      { name: 'FREE State % - Low', value: 0.40, unit: '%', description: 'Low estimate: drivers in FREE states' },
      { name: 'FREE State % - Mid', value: 0.45, unit: '%', description: 'Mid estimate: drivers in FREE states' },
      { name: 'FREE State % - High', value: 0.50, unit: '%', description: 'High estimate: drivers in FREE states' },
      { name: 'Tier 1 Distribution', value: 0.10, unit: '%', description: 'Profiles in Tier 1' },
      { name: 'Tier 2 Distribution', value: 0.20, unit: '%', description: 'Profiles in Tier 2' },
      { name: 'Tier 3 Distribution', value: 0.70, unit: '%', description: 'Profiles in Tier 3' }
    ]
  }
};


// ═══════════════════════════════════════════════════════════════════════════════
// CREATE ASSUMPTIONS SHEET
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create or update the Assumptions sheet with all configurable parameters
 * Sets up named ranges for formula references
 */
function createAssumptionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Check if sheet exists
  let sheet = ss.getSheetByName('Assumptions');
  if (sheet) {
    const response = ui.alert(
      '⚠️ Assumptions Sheet Exists',
      'This will reset all values to defaults. Continue?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      ui.alert('Operation cancelled.');
      return;
    }
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Assumptions');
  }
  
  // Move to front
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  
  // Set up columns
  sheet.setColumnWidth(1, 30);   // Spacer
  sheet.setColumnWidth(2, 250);  // Parameter Name
  sheet.setColumnWidth(3, 120);  // Value
  sheet.setColumnWidth(4, 100);  // Unit
  sheet.setColumnWidth(5, 350);  // Description
  sheet.setColumnWidth(6, 30);   // Spacer
  
  // Hide gridlines for professional look
  sheet.setHiddenGridlines(true);
  
  let currentRow = 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // TITLE SECTION
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 2, 1, 4).merge()
    .setValue('📋 ASSUMPTIONS & PARAMETERS')
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(currentRow, 45);
  currentRow++;
  
  // Subtitle
  sheet.getRange(currentRow, 2, 1, 4).merge()
    .setValue('Edit values below to update all reports • Changes take effect on next refresh')
    .setFontColor(COLORS.SLATE)
    .setFontSize(10)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  currentRow++;
  
  // Last updated
  sheet.getRange(currentRow, 2, 1, 4).merge()
    .setValue(`Last Updated: ${new Date().toLocaleString()}`)
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setHorizontalAlignment('center');
  currentRow += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // COLUMN HEADERS
  // ═══════════════════════════════════════════════════════════════════════════
  const headers = ['Parameter', 'Value', 'Unit', 'Description'];
  sheet.getRange(currentRow, 2, 1, 4).setValues([headers])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 28);
  currentRow++;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // WRITE EACH CATEGORY
  // ═══════════════════════════════════════════════════════════════════════════
  const namedRanges = [];
  
  for (const category of Object.values(ASSUMPTIONS_DATA)) {
    // Category header
    sheet.getRange(currentRow, 2, 1, 4).merge()
      .setValue(category.header)
      .setBackground(category.color)
      .setFontColor(COLORS.WHITE)
      .setFontWeight('bold')
      .setFontSize(11);
    sheet.setRowHeight(currentRow, 30);
    currentRow++;
    
    // ═══════════════════════════════════════════════════════════════════════
    // SECTION EXPLANATION BOX
    // ═══════════════════════════════════════════════════════════════════════
    if (category.explanation) {
      // Title row
      sheet.getRange(currentRow, 2, 1, 4).merge()
        .setValue(`📖 ${category.explanation.title}`)
        .setBackground('#f0f9ff')  // Very light blue
        .setFontColor(COLORS.NAVY)
        .setFontWeight('bold')
        .setFontSize(10);
      sheet.setRowHeight(currentRow, 24);
      currentRow++;
      
      // Explanation lines
      for (const line of category.explanation.lines) {
        sheet.getRange(currentRow, 2, 1, 4).merge()
          .setValue(line)
          .setBackground('#f0f9ff')
          .setFontColor(COLORS.SLATE)
          .setFontSize(9)
          .setWrap(true);
        sheet.setRowHeight(currentRow, 20);
        currentRow++;
      }
      
      // Small gap after explanation
      currentRow++;
    }
    
    // Parameters
    let altRow = false;
    for (const param of category.params) {
      const bgColor = altRow ? COLORS.LIGHT_GRAY : COLORS.WHITE;
      
      // Parameter name
      sheet.getRange(currentRow, 2)
        .setValue(param.name)
        .setBackground(bgColor)
        .setFontWeight('bold');
      
      // Value - this is the editable cell
      const valueCell = sheet.getRange(currentRow, 3);
      valueCell.setValue(param.value)
        .setBackground(bgColor)
        .setHorizontalAlignment('right');
      
      // Format based on type
      if (param.unit.includes('%')) {
        valueCell.setNumberFormat('0%');
      } else if (param.unit.includes('$')) {
        valueCell.setNumberFormat('$#,##0.00');
      } else if (param.unit.includes('USD')) {
        valueCell.setNumberFormat('$#,##0');
      } else {
        valueCell.setNumberFormat('#,##0.##');
      }
      
      // Add light blue border to indicate editable
      valueCell.setBorder(true, true, true, true, false, false, COLORS.ROYAL_BLUE, SpreadsheetApp.BorderStyle.SOLID);
      
      // Unit
      sheet.getRange(currentRow, 4)
        .setValue(param.unit)
        .setBackground(bgColor)
        .setFontColor(COLORS.SLATE)
        .setFontSize(9);
      
      // Description
      sheet.getRange(currentRow, 5)
        .setValue(param.description)
        .setBackground(bgColor)
        .setFontColor(COLORS.SLATE)
        .setFontSize(9)
        .setFontStyle('italic');
      
      // Create named range for this parameter
      const rangeName = paramToRangeName(param.name);
      namedRanges.push({ name: rangeName, row: currentRow, col: 3 });
      
      currentRow++;
      altRow = !altRow;
    }
    
    // Spacer between categories
    currentRow++;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CREATE NAMED RANGES
  // ═══════════════════════════════════════════════════════════════════════════
  for (const nr of namedRanges) {
    try {
      // Remove existing named range if it exists
      const existing = ss.getRangeByName(nr.name);
      if (existing) {
        ss.removeNamedRange(nr.name);
      }
      // Create new named range
      ss.setNamedRange(nr.name, sheet.getRange(nr.row, nr.col));
    } catch (e) {
      Logger.log(`Could not create named range ${nr.name}: ${e.message}`);
    }
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FOOTER
  // ═══════════════════════════════════════════════════════════════════════════
  currentRow++;
  sheet.getRange(currentRow, 2, 1, 4).merge()
    .setValue('ℹ️ Values with blue borders are editable. Reports will use these values on next refresh.')
    .setFontColor(COLORS.ROYAL_BLUE)
    .setFontSize(9)
    .setFontStyle('italic');
  
  // Freeze header rows
  sheet.setFrozenRows(5);
  
  // Protect non-value cells
  const protection = sheet.protect().setDescription('Assumptions Structure');
  protection.setUnprotectedRanges(getEditableRanges(sheet, namedRanges));
  protection.setWarningOnly(true);
  
  ui.alert('✅ Assumptions Sheet Created', 
    `Created ${namedRanges.length} configurable parameters with named ranges.\n\n` +
    'You can now edit values (blue-bordered cells) and reports will reference them automatically.',
    ui.ButtonSet.OK);
  
  Logger.log(`✅ Created Assumptions sheet with ${namedRanges.length} named ranges`);
}

/**
 * Convert parameter name to valid named range name
 */
function paramToRangeName(paramName) {
  return 'PARAM_' + paramName
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '');
}

/**
 * Get editable ranges for protection
 */
function getEditableRanges(sheet, namedRanges) {
  return namedRanges.map(nr => sheet.getRange(nr.row, nr.col));
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER: GET ASSUMPTION VALUE
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get a parameter value from the Assumptions sheet
 * Falls back to default if sheet doesn't exist
 * @param {string} paramName - The parameter name (e.g., 'Cost per Check - CERTN')
 * @returns {number|string} - The parameter value
 */
function getAssumption(paramName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rangeName = paramToRangeName(paramName);
  
  try {
    const range = ss.getRangeByName(rangeName);
    if (range) {
      return range.getValue();
    }
  } catch (e) {
    Logger.log(`Could not get assumption ${paramName}: ${e.message}`);
  }
  
  // Fall back to defaults
  for (const category of Object.values(ASSUMPTIONS_DATA)) {
    const param = category.params.find(p => p.name === paramName);
    if (param) return param.value;
  }
  
  return null;
}

/**
 * Get all assumptions as an object
 * @returns {Object} - All assumptions keyed by parameter name
 */
function getAllAssumptions() {
  const assumptions = {};
  
  for (const category of Object.values(ASSUMPTIONS_DATA)) {
    for (const param of category.params) {
      assumptions[param.name] = getAssumption(param.name);
    }
  }
  
  return assumptions;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Menu entry point for creating Assumptions sheet
 */
function runCreateAssumptions() {
  createAssumptionsSheet();
}
