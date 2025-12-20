/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * FINANCE DASHBOARD - True Cost Input & Analytics
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * Professional finance dashboard featuring:
 * - All 51 states with pre-populated vendors based on STATE_CONFIG
 * - Editable True Cost input per state (starts at $0)
 * - Vendor dropdown (CERTN, INFORMDATA, PENNDOT)
 * - Hidden sheet to persist entered costs between sessions
 * - Refresh button that calculates weighted averages and updates Assumptions
 * 
 * WORKFLOW:
 * 1. Finance enters actual invoice costs per state in True Cost column
 * 2. Click REFRESH button (or run from menu)
 * 3. System calculates weighted averages per vendor
 * 4. Updates Assumptions sheet with new averages
 * 5. All dashboards using those named ranges update automatically
 * 
 * COLUMN REFERENCE (MVR_Ticket_History):
 * - Column AA (27) = DL State
 * - Column Z (26) = Vendor Group
 */

// Hidden sheet name for persisting costs
var FINANCE_COSTS_SHEET = '_Finance_Costs';

// All 51 states with their default vendor based on STATE_CONFIG
// PA → PENNDOT, MO/IL → CERTN, all others → INFORMDATA
var ALL_STATES_WITH_VENDORS = [
  { state: 'AL', vendor: 'INFORMDATA', name: 'Alabama' },
  { state: 'AK', vendor: 'INFORMDATA', name: 'Alaska' },
  { state: 'AZ', vendor: 'INFORMDATA', name: 'Arizona' },
  { state: 'AR', vendor: 'INFORMDATA', name: 'Arkansas' },
  { state: 'CA', vendor: 'INFORMDATA', name: 'California' },
  { state: 'CO', vendor: 'INFORMDATA', name: 'Colorado' },
  { state: 'CT', vendor: 'INFORMDATA', name: 'Connecticut' },
  { state: 'DE', vendor: 'INFORMDATA', name: 'Delaware' },
  { state: 'DC', vendor: 'INFORMDATA', name: 'District of Columbia' },
  { state: 'FL', vendor: 'INFORMDATA', name: 'Florida' },
  { state: 'GA', vendor: 'INFORMDATA', name: 'Georgia' },
  { state: 'HI', vendor: 'INFORMDATA', name: 'Hawaii' },
  { state: 'ID', vendor: 'INFORMDATA', name: 'Idaho' },
  { state: 'IL', vendor: 'CERTN', name: 'Illinois' },
  { state: 'IN', vendor: 'INFORMDATA', name: 'Indiana' },
  { state: 'IA', vendor: 'INFORMDATA', name: 'Iowa' },
  { state: 'KS', vendor: 'INFORMDATA', name: 'Kansas' },
  { state: 'KY', vendor: 'INFORMDATA', name: 'Kentucky' },
  { state: 'LA', vendor: 'INFORMDATA', name: 'Louisiana' },
  { state: 'ME', vendor: 'INFORMDATA', name: 'Maine' },
  { state: 'MD', vendor: 'INFORMDATA', name: 'Maryland' },
  { state: 'MA', vendor: 'INFORMDATA', name: 'Massachusetts' },
  { state: 'MI', vendor: 'INFORMDATA', name: 'Michigan' },
  { state: 'MN', vendor: 'INFORMDATA', name: 'Minnesota' },
  { state: 'MS', vendor: 'INFORMDATA', name: 'Mississippi' },
  { state: 'MO', vendor: 'CERTN', name: 'Missouri' },
  { state: 'MT', vendor: 'INFORMDATA', name: 'Montana' },
  { state: 'NE', vendor: 'INFORMDATA', name: 'Nebraska' },
  { state: 'NV', vendor: 'INFORMDATA', name: 'Nevada' },
  { state: 'NH', vendor: 'INFORMDATA', name: 'New Hampshire' },
  { state: 'NJ', vendor: 'INFORMDATA', name: 'New Jersey' },
  { state: 'NM', vendor: 'INFORMDATA', name: 'New Mexico' },
  { state: 'NY', vendor: 'INFORMDATA', name: 'New York' },
  { state: 'NC', vendor: 'INFORMDATA', name: 'North Carolina' },
  { state: 'ND', vendor: 'INFORMDATA', name: 'North Dakota' },
  { state: 'OH', vendor: 'INFORMDATA', name: 'Ohio' },
  { state: 'OK', vendor: 'INFORMDATA', name: 'Oklahoma' },
  { state: 'OR', vendor: 'INFORMDATA', name: 'Oregon' },
  { state: 'PA', vendor: 'PENNDOT', name: 'Pennsylvania' },
  { state: 'RI', vendor: 'INFORMDATA', name: 'Rhode Island' },
  { state: 'SC', vendor: 'INFORMDATA', name: 'South Carolina' },
  { state: 'SD', vendor: 'INFORMDATA', name: 'South Dakota' },
  { state: 'TN', vendor: 'INFORMDATA', name: 'Tennessee' },
  { state: 'TX', vendor: 'INFORMDATA', name: 'Texas' },
  { state: 'UT', vendor: 'INFORMDATA', name: 'Utah' },
  { state: 'VT', vendor: 'INFORMDATA', name: 'Vermont' },
  { state: 'VA', vendor: 'INFORMDATA', name: 'Virginia' },
  { state: 'WA', vendor: 'INFORMDATA', name: 'Washington' },
  { state: 'WV', vendor: 'INFORMDATA', name: 'West Virginia' },
  { state: 'WI', vendor: 'INFORMDATA', name: 'Wisconsin' },
  { state: 'WY', vendor: 'INFORMDATA', name: 'Wyoming' }
];

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER: LOAD OUTCOME ANALYSIS FROM TAG_OUTCOME_MAPPINGS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Reads Tag_Outcome_Mappings and aggregates by Outcome + Billable status.
 * Returns array of objects with outcome info for dynamic row generation.
 * "Clear" is split into billable and non-billable based on Is Billable column.
 */
function loadOutcomeAnalysisFromMappings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mappingSheet = ss.getSheetByName('Tag_Outcome_Mappings');
  
  if (!mappingSheet) {
    Logger.log('Tag_Outcome_Mappings sheet not found - using defaults');
    return getDefaultOutcomeAnalysis();
  }
  
  var data = mappingSheet.getDataRange().getValues();
  if (data.length < 2) return getDefaultOutcomeAnalysis();
  
  // Header row: Tag Pattern, Match Type, Outcome Type, Priority, Is Billable, Is Active, Notes
  var headers = data[0];
  var colIdx = {
    pattern: 0,
    matchType: 1,
    outcome: 2,
    priority: 3,
    billable: 4,
    active: 5,
    notes: 6
  };
  
  // Group by Outcome + Billable status
  var outcomeMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var pattern = String(row[colIdx.pattern] || '').trim();
    var outcome = String(row[colIdx.outcome] || '').trim();
    var isBillable = String(row[colIdx.billable]).toUpperCase() === 'TRUE';
    var isActive = String(row[colIdx.active]).toUpperCase() === 'TRUE';
    var priority = parseInt(row[colIdx.priority]) || 999;
    
    if (!outcome || !isActive || !pattern) continue;
    
    // Create unique key: outcome_billable (e.g., "Clear_true" vs "Clear_false")
    var key = outcome + '_' + isBillable;
    
    if (!outcomeMap[key]) {
      outcomeMap[key] = {
        outcome: outcome,
        isBillable: isBillable,
        patterns: [],
        priority: priority,
        count: 0
      };
    }
    
    outcomeMap[key].patterns.push(pattern);
    if (priority < outcomeMap[key].priority) {
      outcomeMap[key].priority = priority;
    }
  }
  
  // Convert to array and sort by priority
  var result = [];
  for (var k in outcomeMap) {
    result.push(outcomeMap[k]);
  }
  result.sort(function(a, b) { return a.priority - b.priority; });
  
  return result;
}

/**
 * Default outcome analysis if Tag_Outcome_Mappings doesn't exist
 */
function getDefaultOutcomeAnalysis() {
  return [
    { outcome: 'Suspension Confirmed', isBillable: true, patterns: ['Suspension'], priority: 1 },
    { outcome: 'Expired License', isBillable: true, patterns: ['Expired'], priority: 2 },
    { outcome: 'Invalid License', isBillable: true, patterns: ['Invalid'], priority: 3 },
    { outcome: 'Clear', isBillable: true, patterns: ['Updated', 'uploaded'], priority: 4 },
    { outcome: 'Clear', isBillable: false, patterns: ['same info', 'valid'], priority: 5 },
    { outcome: 'DMV Unavailable', isBillable: false, patterns: ['DMV'], priority: 6 },
    { outcome: 'Cannot Process', isBillable: false, patterns: ['Cannot'], priority: 7 },
    { outcome: 'Record Not Found', isBillable: false, patterns: ['Not Found'], priority: 8 },
    { outcome: 'Withdrawn', isBillable: false, patterns: ['Withdrawn'], priority: 9 },
    { outcome: 'Pending', isBillable: false, patterns: ['Pending'], priority: 10 },
    { outcome: 'Still Processing', isBillable: false, patterns: ['processing'], priority: 11 }
  ];
}

/**
 * Builds COUNTIF or SUMPRODUCT formula based on outcome type.
 * For Clear outcomes: uses tag-based matching
 * For other outcomes: uses simple COUNTIF on AI column
 */
function buildOutcomeCountFormula(item) {
  // Clear outcomes need tag-based detection
  if (item.outcome === 'Clear' && item.patterns.length > 0) {
    // Build pattern search: (ISNUMBER(SEARCH("pattern1",U:U))+ISNUMBER(SEARCH("pattern2",U:U))>0)
    var patternParts = item.patterns.map(function(p) {
      return 'ISNUMBER(SEARCH("' + p + '",\'MVR_Ticket_History\'!U:U))';
    });
    return '=IFERROR(SUMPRODUCT((\'MVR_Ticket_History\'!AI:AI="Clear")*(' + patternParts.join('+') + '>0)*1),0)';
  }
  
  // Simple COUNTIF for non-Clear outcomes
  return '=IFERROR(COUNTIF(\'MVR_Ticket_History\'!AI:AI,"' + item.outcome + '"),0)';
}

/**
 * Builds cost formula based on outcome type.
 */
function buildOutcomeCostFormula(item) {
  var vendorCostPart = 'IF(\'MVR_Ticket_History\'!Z:Z="CERTN",PARAM_COST_PER_CHECK_CERTN,' +
    'IF(\'MVR_Ticket_History\'!Z:Z="INFORMDATA",PARAM_COST_PER_CHECK_INFORMDATA,' +
    'IF(\'MVR_Ticket_History\'!Z:Z="PENNDOT",PARAM_COST_PER_CHECK_PENNDOT,0)))';
  
  // Clear outcomes need tag-based detection
  if (item.outcome === 'Clear' && item.patterns.length > 0) {
    var patternParts = item.patterns.map(function(p) {
      return 'ISNUMBER(SEARCH("' + p + '",\'MVR_Ticket_History\'!U:U))';
    });
    return '=IFERROR(SUMPRODUCT((\'MVR_Ticket_History\'!AI:AI="Clear")*(' + patternParts.join('+') + '>0)*' + vendorCostPart + '),0)';
  }
  
  // Simple SUMPRODUCT for non-Clear outcomes
  return '=IFERROR(SUMPRODUCT((\'MVR_Ticket_History\'!AI:AI="' + item.outcome + '")*' + vendorCostPart + '),0)';
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN DASHBOARD CREATION
// ═══════════════════════════════════════════════════════════════════════════════

function createFinanceDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // Check for Assumptions sheet first
  if (!ss.getSheetByName('Assumptions')) {
    ui.alert('⚠️ Setup Required', 
      'Please run "Setup Assumptions Sheet" first from the Reports menu.\n\n' +
      'The Assumptions sheet contains cost parameters needed for calculations.',
      ui.ButtonSet.OK);
    return;
  }
  
  // Ensure hidden costs sheet exists
  ensureFinanceCostsSheet(ss);
  
  // Get or create dashboard sheet
  var sheet = ss.getSheetByName('Finance Dashboard');
  if (sheet) {
    var response = ui.alert(
      '⚠️ Finance Dashboard Exists',
      'This will recreate the dashboard.\n\nYour entered costs are preserved in a hidden sheet and will be restored.',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    
    // Remove existing charts
    var charts = sheet.getCharts();
    charts.forEach(function(c) { sheet.removeChart(c); });
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Finance Dashboard');
  }
  
  // Professional setup
  sheet.setHiddenGridlines(true);
  
  // Column widths: A=spacer, B=State, C=StateName, D=Vendor, E=TrueCost, F=Volume, G=TotalCost, H=spacer, I=RefreshBtn
  var colWidths = [30, 60, 150, 120, 120, 100, 120, 30, 120];
  colWidths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  
  var row = 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // TITLE SECTION WITH REFRESH BUTTON
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💰 FINANCE DASHBOARD - True Cost Input')
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 50);
  
  // REFRESH BUTTON in column I
  sheet.getRange(row, 9, 2, 1).merge()
    .setValue('🔄 REFRESH\n& Update Assumptions')
    .setBackground('#059669')
    .setFontColor(COLORS.WHITE)
    .setFontSize(11)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, false, false, '#047857', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(row, 9, 2, 1).setNote('To refresh:\n1. Menu → MVR Reports → Dashboards\n2. Click "🔄 Refresh Finance Costs → Assumptions"\n\nThis will:\n• Save your True Cost entries\n• Calculate weighted averages\n• Update Assumptions sheet');
  
  row++;
  
  // Subtitle with last refresh time
  sheet.getRange(row, 2, 1, 6).merge()
    .setFormula('="Last Refreshed: "&TEXT(NOW(),"yyyy-mm-dd hh:mm")')
    .setFontColor(COLORS.SLATE)
    .setFontSize(10)
    .setHorizontalAlignment('center');
  row++;
  
  // Instructions
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📝 Enter actual invoice costs in "True Cost" column (blue border). Run Refresh from menu to update Assumptions.')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // KPI SUMMARY CARDS
  // ═══════════════════════════════════════════════════════════════════════════
  var kpiRow = row;
  
  // We'll reference the data table rows after they're created
  var dataStartRow = row + 12; // After KPIs + explanation + headers
  var dataEndRow = dataStartRow + 50; // 51 states
  
  // KPI 1: Total Tickets
  createFinanceKPICard(sheet, row, 2, 
    "=IFERROR(COUNTA('MVR_Ticket_History'!A:A)-1,0)",
    'Total Tickets', '#,##0');
  
  // KPI 2: Total Cost (from True Cost input × Volume)
  createFinanceKPICard(sheet, row, 4,
    "=IFERROR(SUM(G" + dataStartRow + ":G" + dataEndRow + "),0)",
    'Total Vendor Cost', '$#,##0.00');
  
  // KPI 3: Weighted Avg Cost
  createFinanceKPICard(sheet, row, 6,
    "=IFERROR(D" + row + "/B" + row + ",0)",
    'Avg Cost/Ticket', '$#,##0.00');
  
  row += 4;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // EXPLANATION BOX
  // ═══════════════════════════════════════════════════════════════════════════
  addFinanceExplanationBox(sheet, row, 2, 6, [
    '📖 How to Use This Dashboard:',
    '• True Cost (blue border): Enter your actual invoice cost per check for each state',
    '• Vendor: Select from dropdown if the vendor differs from default assignment',
    '• Volume: Ticket count from MVR_Ticket_History (auto-calculated, read-only)',
    '• Total Cost: True Cost × Volume (auto-calculated)',
    '• Click 🔄 REFRESH (from menu) to save costs and update Assumptions averages'
  ]);
  row += 7;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // STATE COST INPUT TABLE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📋 TRUE COST INPUT BY STATE (All 51 States)')
    .setBackground(COLORS.COST_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  // Table headers
  var headers = ['State', 'State Name', 'Vendor', 'True Cost', 'Volume', 'Total Cost'];
  sheet.getRange(row, 2, 1, 6).setValues([headers])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  // Store actual starting row for formulas
  var actualDataStartRow = row;
  
  // Load saved costs if any
  var savedCosts = loadSavedCosts(ss);
  
  // Create vendor dropdown validation
  var vendorRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CERTN', 'INFORMDATA', 'PENNDOT'], true)
    .setAllowInvalid(false)
    .build();
  
  // Populate all 51 states
  ALL_STATES_WITH_VENDORS.forEach(function(stateInfo, idx) {
    var bgColor = idx % 2 === 0 ? COLORS.WHITE : COLORS.LIGHT_GRAY;
    var savedData = savedCosts[stateInfo.state] || { cost: 0, vendor: stateInfo.vendor };
    
    // State code
    sheet.getRange(row, 2).setValue(stateInfo.state)
      .setBackground(bgColor)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // State name
    sheet.getRange(row, 3).setValue(stateInfo.name)
      .setBackground(bgColor);
    
    // Vendor dropdown (pre-populated from saved or default)
    sheet.getRange(row, 4).setValue(savedData.vendor)
      .setBackground(bgColor)
      .setDataValidation(vendorRule);
    
    // True Cost - EDITABLE (blue border, starts at $0 or saved value)
    sheet.getRange(row, 5)
      .setValue(savedData.cost)
      .setBackground(bgColor)
      .setNumberFormat('$#,##0.00')
      .setBorder(true, true, true, true, false, false, COLORS.ROYAL_BLUE, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    
    // Volume - count from history (Column AA = DL State)
    sheet.getRange(row, 6)
      .setFormula("=IFERROR(COUNTIF('MVR_Ticket_History'!AA:AA,\"" + stateInfo.state + "\"),0)")
      .setBackground(bgColor)
      .setNumberFormat('#,##0');
    
    // Total Cost = True Cost × Volume
    sheet.getRange(row, 7)
      .setFormula("=E" + row + "*F" + row)
      .setBackground(bgColor)
      .setNumberFormat('$#,##0.00');
    
    row++;
  });
  
  var actualDataEndRow = row - 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // TOTALS ROW
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 3).merge()
    .setValue('TOTALS')
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('right');
  
  // Empty True Cost column (doesn't sum)
  sheet.getRange(row, 5)
    .setValue('')
    .setBackground(COLORS.NAVY);
  
  // Total Volume
  sheet.getRange(row, 6)
    .setFormula("=SUM(F" + actualDataStartRow + ":F" + actualDataEndRow + ")")
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setNumberFormat('#,##0');
  
  // Total Cost
  sheet.getRange(row, 7)
    .setFormula("=SUM(G" + actualDataStartRow + ":G" + actualDataEndRow + ")")
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setNumberFormat('$#,##0.00');
  
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // VENDOR SUMMARY SECTION
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📊 VENDOR SUMMARY (Weighted Averages)')
    .setBackground(COLORS.CAPACITY_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  // Explanation for this section
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 "Weighted Avg" is calculated from your True Cost entries. "Assumptions Value" shows current named range value.')
    .setBackground('#f0f9ff')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  // Vendor summary headers
  var vendorHeaders = ['Vendor', '', 'Weighted Avg', 'Total Volume', 'Total Cost', 'Assumptions Value'];
  sheet.getRange(row, 2, 1, 6).setValues([vendorHeaders])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  // CERTN row
  sheet.getRange(row, 2).setValue('CERTN').setFontWeight('bold');
  sheet.getRange(row, 3).setValue('').setBackground(COLORS.WHITE);
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"CERTN\")/SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"CERTN\"),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold');
  sheet.getRange(row, 5)
    .setFormula("=SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"CERTN\")")
    .setNumberFormat('#,##0');
  sheet.getRange(row, 6)
    .setFormula("=SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"CERTN\")")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 7)
    .setFormula("=IFERROR(PARAM_COST_PER_CHECK_CERTN,\"Not Set\")")
    .setNumberFormat('$#,##0.00')
    .setFontColor(COLORS.SLATE)
    .setFontStyle('italic');
  row++;
  
  // INFORMDATA row
  sheet.getRange(row, 2).setValue('INFORMDATA').setFontWeight('bold').setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 3, 1, 5).setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"INFORMDATA\")/SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"INFORMDATA\"),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold');
  sheet.getRange(row, 5)
    .setFormula("=SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"INFORMDATA\")")
    .setNumberFormat('#,##0');
  sheet.getRange(row, 6)
    .setFormula("=SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"INFORMDATA\")")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 7)
    .setFormula("=IFERROR(PARAM_COST_PER_CHECK_INFORMDATA,\"Not Set\")")
    .setNumberFormat('$#,##0.00')
    .setFontColor(COLORS.SLATE)
    .setFontStyle('italic');
  row++;
  
  // PENNDOT row
  sheet.getRange(row, 2).setValue('PENNDOT').setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"PENNDOT\")/SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"PENNDOT\"),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold');
  sheet.getRange(row, 5)
    .setFormula("=SUMIFS(F" + actualDataStartRow + ":F" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"PENNDOT\")")
    .setNumberFormat('#,##0');
  sheet.getRange(row, 6)
    .setFormula("=SUMIFS(G" + actualDataStartRow + ":G" + actualDataEndRow + ",D" + actualDataStartRow + ":D" + actualDataEndRow + ",\"PENNDOT\")")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 7)
    .setFormula("=IFERROR(PARAM_COST_PER_CHECK_PENNDOT,\"Not Set\")")
    .setNumberFormat('$#,##0.00')
    .setFontColor(COLORS.SLATE)
    .setFontStyle('italic');
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // BILLABLE COST ANALYSIS SECTION (DYNAMIC FROM TAG_OUTCOME_MAPPINGS)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📊 BILLABLE COST ANALYSIS (Dynamic from Tag_Outcome_Mappings)')
    .setBackground('#059669')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  // Data source explanation
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📅 Data Source: MVR_Ticket_History | Classifications: Tag_Outcome_Mappings | Costs: Vendor Group (Col Z) × Assumptions')
    .setBackground('#ecfdf5')
    .setFontColor('#059669')
    .setFontSize(9)
    .setFontWeight('bold');
  row++;
  
  // Explanation for billable section
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 Each row dynamically generated from Tag_Outcome_Mappings. Billable status determined by "Is Billable" column in mappings.')
    .setBackground('#ecfdf5')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  // Billable analysis headers
  var billableHeaders = ['Outcome Type', 'Tags/Patterns', 'Count', 'Vendor Cost', '% of Total', 'Status'];
  sheet.getRange(row, 2, 1, 6).setValues([billableHeaders])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  var billableStartRow = row;
  
  // Load outcome analysis from Tag_Outcome_Mappings
  var outcomeAnalysis = loadOutcomeAnalysisFromMappings();
  var rowIndex = 0;
  var billableRowsForMetrics = []; // Track billable rows for metrics calculation
  
  outcomeAnalysis.forEach(function(item) {
    var isAltRow = (rowIndex % 2 === 1);
    var bgColor = isAltRow ? COLORS.LIGHT_GRAY : null;
    
    // Determine display label and color
    var icon, color, statusText;
    if (item.outcome === 'Pending' || item.outcome === 'Still Processing' || item.outcome === 'Unknown') {
      icon = '⏳';
      color = '#d97706'; // amber
      statusText = 'TBD';
    } else if (item.isBillable) {
      icon = '✅';
      color = '#059669'; // green
      statusText = 'Revenue';
      billableRowsForMetrics.push(row);
    } else {
      icon = '❌';
      color = '#dc2626'; // red
      statusText = 'Cost Eaten';
    }
    
    // Build label: "✅ Clear (Billable)" or "❌ Clear (Non-billable)" for Clear outcomes
    var displayLabel = icon + ' ' + item.outcome;
    if (item.outcome === 'Clear') {
      displayLabel += item.isBillable ? ' (Uploaded)' : ' (No Change)';
    }
    
    // Set row background if alt row
    if (bgColor) {
      sheet.getRange(row, 2, 1, 6).setBackground(bgColor);
    }
    
    // Column B-C: Outcome label
    sheet.getRange(row, 2).setValue(displayLabel)
      .setFontWeight('bold').setFontColor(color);
    
    // Column C: Tag patterns (truncated)
    var patternsDisplay = item.patterns.slice(0, 3).join(', ');
    if (item.patterns.length > 3) patternsDisplay += '...';
    sheet.getRange(row, 3).setValue(patternsDisplay)
      .setFontSize(8).setFontColor(COLORS.SLATE).setFontStyle('italic');
    
    // Column D: Count
    sheet.getRange(row, 4)
      .setFormula(buildOutcomeCountFormula(item))
      .setNumberFormat('#,##0');
    
    // Column E: Vendor Cost
    sheet.getRange(row, 5)
      .setFormula(buildOutcomeCostFormula(item))
      .setNumberFormat('$#,##0.00');
    
    // Column F: % of Total
    sheet.getRange(row, 6)
      .setFormula("=IFERROR(D" + row + "/(COUNTA('MVR_Ticket_History'!A:A)-1),0)")
      .setNumberFormat('0.0%');
    
    // Column G: Status
    sheet.getRange(row, 7)
      .setValue(statusText).setFontColor(color).setFontWeight('bold');
    
    row++;
    rowIndex++;
  });
  
  // Add row for Unclassified/Empty outcomes (catch-all)
  var isAltRow = (rowIndex % 2 === 1);
  if (isAltRow) {
    sheet.getRange(row, 2, 1, 6).setBackground(COLORS.LIGHT_GRAY);
  }
  sheet.getRange(row, 2).setValue('⚠️ Unclassified / Empty')
    .setFontWeight('bold').setFontColor('#9ca3af');
  sheet.getRange(row, 3).setValue('No outcome assigned')
    .setFontSize(8).setFontColor(COLORS.SLATE).setFontStyle('italic');
  sheet.getRange(row, 4)
    .setFormula('=IFERROR((COUNTA(\'MVR_Ticket_History\'!A:A)-1)-SUM(D' + billableStartRow + ':D' + (row-1) + '),0)')
    .setNumberFormat('#,##0');
  sheet.getRange(row, 5)
    .setValue('-')
    .setHorizontalAlignment('center');
  sheet.getRange(row, 6)
    .setFormula("=IFERROR(D" + row + "/(COUNTA('MVR_Ticket_History'!A:A)-1),0)")
    .setNumberFormat('0.0%');
  sheet.getRange(row, 7)
    .setValue('Review').setFontColor('#9ca3af').setFontWeight('bold');
  row++;
  
  var billableEndRow = row - 1;
  
  // Totals row for billable analysis
  sheet.getRange(row, 2, 1, 2).merge().setValue('TOTALS')
    .setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setHorizontalAlignment('right');
  sheet.getRange(row, 4)
    .setFormula("=SUM(D" + billableStartRow + ":D" + billableEndRow + ")")
    .setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setNumberFormat('#,##0');
  sheet.getRange(row, 5)
    .setFormula("=SUM(E" + billableStartRow + ":E" + billableEndRow + ")")
    .setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setNumberFormat('$#,##0.00');
  sheet.getRange(row, 6)
    .setFormula("=SUM(F" + billableStartRow + ":F" + billableEndRow + ")")
    .setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setNumberFormat('0.0%');
  sheet.getRange(row, 7).setBackground(COLORS.NAVY);
  row += 2;
  
  // Key Metrics Row
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('🎯 KEY BILLABLE METRICS')
    .setBackground('#7c3aed')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);
  row++;
  
  // Build dynamic billable sum formula (sum rows where status = "Revenue")
  // Since we track billableRowsForMetrics array, build formula from it
  var billableCountFormula, billableCostFormula;
  if (billableRowsForMetrics.length > 0) {
    var countParts = billableRowsForMetrics.map(function(r) { return 'D' + r; });
    var costParts = billableRowsForMetrics.map(function(r) { return 'E' + r; });
    billableCountFormula = '=IFERROR(' + countParts.join('+') + ',0)';
    billableCostFormula = '=IFERROR(' + costParts.join('+') + ',0)';
  } else {
    billableCountFormula = '=0';
    billableCostFormula = '=0';
  }
  
  // Success Rate
  sheet.getRange(row, 2, 1, 2).merge().setValue('Success Rate (Billable %)');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR((" + billableCountFormula.replace('=IFERROR(', '').replace(',0)', '') + ")/(COUNTA('MVR_Ticket_History'!A:A)-1),0)")
    .setNumberFormat('0.0%').setFontSize(14).setFontWeight('bold').setFontColor('#059669');
  sheet.getRange(row, 5, 1, 3).merge().setValue('% of checks that generate revenue')
    .setFontColor(COLORS.SLATE).setFontStyle('italic').setFontSize(9);
  row++;
  
  // Cost per Billable Check (vendor only)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Cost per Billable Check (Vendor)')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4, 1, 4).setBackground(COLORS.LIGHT_GRAY);
  // Total vendor cost for all tickets ÷ billable count
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(SUM(E" + billableStartRow + ":E" + billableEndRow + ")/(" + billableCountFormula.replace('=IFERROR(', '').replace(',0)', '') + "),0)")
    .setNumberFormat('$#,##0.00').setFontSize(14).setFontWeight('bold').setFontColor(COLORS.NAVY);
  sheet.getRange(row, 5, 1, 3).merge().setValue('Total vendor cost ÷ billable checks (includes absorbed)')
    .setFontColor(COLORS.SLATE).setFontStyle('italic').setFontSize(9);
  row++;
  
  // Cost per Billable Check (with labor)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Cost per Billable Check (+ Labor)');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(D" + (row-1) + "+IFERROR(PARAM_LABOR_COST_PER_TICKET,10.21),0)")
    .setNumberFormat('$#,##0.00').setFontSize(14).setFontWeight('bold').setFontColor('#be185d');
  sheet.getRange(row, 5, 1, 3).merge().setValue('Fully-loaded cost including labor per ticket')
    .setFontColor(COLORS.SLATE).setFontStyle('italic').setFontSize(9);
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PRICING MODEL COMPARISON
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💰 PRICING MODEL COMPARISON: At-Cost vs. Full Price')
    .setBackground('#0d9488')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 All drivers pay $2/mo subscription. Billable checks add $5/check. Full Price = $15/check all-in. Each row = 1 driver checked.')
    .setBackground('#f0fdfa')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  // Pricing comparison headers
  var pricingHeaders = ['Metric', '', 'At-Cost Model', 'Full Price Model', 'Difference', ''];
  sheet.getRange(row, 2, 1, 6).setValues([pricingHeaders])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  // Revenue per Billable Check (only billable checks get this fee)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Check Fee (billable only)');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(PARAM_PRICE_PER_BILLABLE_CHECK,5)")
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 5)
    .setFormula("=IFERROR(PARAM_MARKET_PRICE_PER_CHECK,15)")
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#dc2626');
  row++;
  
  // Monthly Subscription Revenue (ALL drivers pay subscription)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Driver Subscription Revenue (ALL rows)')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4, 1, 3).setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4)
    .setFormula("=IFERROR((COUNTA('MVR_Ticket_History'!A:A)-1)*IFERROR(PARAM_DRIVER_MONTHLY_SUBSCRIPTION,2),0)")
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 5)
    .setValue(0)
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 6)
    .setFormula("=D" + row + "-E" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#059669');
  row++;
  
  // Total Monthly Revenue = (All drivers × subscription) + (Billable × check fee)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Total Monthly Revenue');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(D" + (row-1) + "+((D" + billableStartRow + "+D" + (billableStartRow+1) + ")*IFERROR(PARAM_PRICE_PER_BILLABLE_CHECK,5)),0)")
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 5)
    .setFormula("=IFERROR((COUNTA('MVR_Ticket_History'!A:A)-1)*PARAM_MARKET_PRICE_PER_CHECK,0)")
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(12);
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#dc2626').setFontWeight('bold');
  row++;
  
  // Total Cost (all checks incl absorbed)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Total Vendor Cost (incl absorbed)')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4, 1, 3).setBackground(COLORS.LIGHT_GRAY);
  var totalCostRow = row;
  sheet.getRange(row, 4)
    .setFormula("=E" + (billableEndRow + 1))
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 5)
    .setFormula("=D" + row)
    .setNumberFormat('$#,##0.00').setFontWeight('bold');
  sheet.getRange(row, 6)
    .setValue('Same')
    .setFontColor(COLORS.SLATE).setFontStyle('italic');
  row++;
  
  // Gross Margin
  sheet.getRange(row, 2, 1, 2).merge().setValue('Gross Margin (Revenue - Vendor Cost)');
  sheet.getRange(row, 4)
    .setFormula("=D" + (row-2) + "-D" + (row-1))
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(14)
    .setFontColor("=IF(D" + row + ">=0,'#059669','#dc2626')");
  sheet.getRange(row, 5)
    .setFormula("=E" + (row-2) + "-E" + (row-1))
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(14).setFontColor('#059669');
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#dc2626').setFontWeight('bold');
  row++;
  
  // Margin with Labor
  sheet.getRange(row, 2, 1, 2).merge().setValue('Net Margin (after Labor)')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4, 1, 3).setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4)
    .setFormula("=D" + (row-1) + "-(D" + (billableEndRow + 1) + "*IFERROR(PARAM_LABOR_COST_PER_TICKET,10.21))")
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(14).setFontColor('#be185d');
  sheet.getRange(row, 5)
    .setFormula("=E" + (row-1) + "-(D" + (billableEndRow + 1) + "*IFERROR(PARAM_LABOR_COST_PER_TICKET,10.21))")
    .setNumberFormat('$#,##0.00').setFontWeight('bold').setFontSize(14).setFontColor('#059669');
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#dc2626').setFontWeight('bold');
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PARTNER INVOICE IMPACT
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('🧾 PARTNER INVOICE IMPACT (Example)')
    .setBackground('#7c3aed')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 Example: Partner checks 100 drivers. ALL 100 pay $2/mo subscription. Only 10 are billable (uploaded) and add $5/check.')
    .setBackground('#faf5ff')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  // Example partner table
  var invoiceHeaders = ['Line Item', '', 'At-Cost', 'Full Price', 'Partner Savings', ''];
  sheet.getRange(row, 2, 1, 6).setValues([invoiceHeaders])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  // Subscription fee (ALL 100 drivers pay subscription)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Driver Subscriptions (ALL 100 × $2)');
  sheet.getRange(row, 4)
    .setFormula("=100*IFERROR(PARAM_DRIVER_MONTHLY_SUBSCRIPTION,2)")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 5)
    .setValue(0)
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 6)
    .setFormula("=D" + row + "-E" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#059669');
  row++;
  
  // Check Fee (only 10 billable checks)
  sheet.getRange(row, 2, 1, 2).merge().setValue('Billable Check Fees (10 × $5)')
    .setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4, 1, 3).setBackground(COLORS.LIGHT_GRAY);
  sheet.getRange(row, 4)
    .setFormula("=10*IFERROR(PARAM_PRICE_PER_BILLABLE_CHECK,5)")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 5)
    .setFormula("=100*IFERROR(PARAM_MARKET_PRICE_PER_CHECK,15)")
    .setNumberFormat('$#,##0.00');
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setFontColor('#dc2626').setFontWeight('bold');
  row++;
  
  // Total Invoice
  sheet.getRange(row, 2, 1, 2).merge().setValue('TOTAL MONTHLY INVOICE')
    .setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=D" + (row-1) + "+D" + (row-2))
    .setNumberFormat('$#,##0.00').setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setFontSize(14);
  sheet.getRange(row, 5)
    .setFormula("=E" + (row-1) + "+E" + (row-2))
    .setNumberFormat('$#,##0.00').setBackground(COLORS.NAVY).setFontColor(COLORS.WHITE).setFontWeight('bold').setFontSize(14);
  sheet.getRange(row, 6)
    .setFormula("=E" + row + "-D" + row)
    .setNumberFormat('$#,##0.00').setBackground('#059669').setFontColor(COLORS.WHITE).setFontWeight('bold').setFontSize(14);
  sheet.getRange(row, 7).setBackground(COLORS.NAVY);
  row++;
  
  // Savings percentage
  sheet.getRange(row, 2, 1, 2).merge().setValue('Partner Discount %');
  sheet.getRange(row, 4, 1, 2).merge()
    .setFormula("=IFERROR(1-(D" + (row-1) + "/E" + (row-1) + "),0)")
    .setNumberFormat('0.0% savings')
    .setFontWeight('bold').setFontColor('#059669').setFontSize(14).setHorizontalAlignment('center');
  sheet.getRange(row, 6).setValue('vs Full Price').setFontColor(COLORS.SLATE).setFontStyle('italic');
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // 12-MONTH FINANCIAL HISTORY (REAL DATES)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📈 12-MONTH FINANCIAL HISTORY')
    .setBackground('#2563eb')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 Actual monthly data from Last Updated (Col K). Subscription = ALL drivers × $2/mo | Check Fee = Billable only × $5')
    .setBackground('#eff6ff')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  // Real date headers
  var histHeaders = ['Month', 'Tickets', 'Vendor Cost', 'Subscription Rev', 'Check Fee Rev', 'Net Margin'];
  sheet.getRange(row, 2, 1, 6).setValues([histHeaders])
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  var histStartRow = row;
  
  // Generate last 12 months with REAL dates from Last Updated column (K)
  var today = new Date();
  // Add helper column to MVR_Ticket_History for native date conversion
  var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MVR_Ticket_History');
  var lastUpdatedCol = 11; // K is column 11 (1-based)
  var helperCol = historySheet.getLastColumn() + 1;
  historySheet.getRange(1, helperCol).setValue('LastUpdatedNative');
  var lastUpdatedValues = historySheet.getRange(2, lastUpdatedCol, historySheet.getLastRow()-1, 1).getValues();
  var nativeDates = lastUpdatedValues.map(function(row) {
    var iso = row[0];
    if (!iso) return '';
    var d = new Date(iso);
    return isNaN(d.getTime()) ? '' : d;
  });
  historySheet.getRange(2, helperCol, nativeDates.length, 1).setValues(nativeDates.map(function(d){return [d];}));

  // Use helperCol for all date-based formulas
  var helperColLetter = String.fromCharCode(64 + helperCol); // Only works for columns < 27
  for (var m = 11; m >= 0; m--) {
    var targetDate = new Date(today.getFullYear(), today.getMonth() - m, 1);
    var targetMonth = targetDate.getMonth() + 1; // 1-12
    var targetYear = targetDate.getFullYear();
    var monthName = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM yyyy');
    var bgColor = (11 - m) % 2 === 0 ? COLORS.WHITE : COLORS.LIGHT_GRAY;
    sheet.getRange(row, 2).setValue(monthName).setFontWeight('bold').setBackground(bgColor);
    // Total Tickets this month
    var today = new Date();
    var historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('MVR_Ticket_History');
    var data = historySheet.getDataRange().getValues();
    var headers = data[0];
    var lastUpdatedIdx = headers.indexOf('Last Updated');
    var vendorIdx = headers.indexOf('Vendor Group');
    var tagsIdx = headers.indexOf('Tags');
    var outcomeIdx = headers.indexOf('MVR Outcome');
    var ticketCountByMonth = {};
    var vendorCostByMonth = {};
    var subscriptionRevByMonth = {};
    var checkFeeRevByMonth = {};
    for (var m = 11; m >= 0; m--) {
      var targetDate = new Date(today.getFullYear(), today.getMonth() - m, 1);
      var targetMonth = targetDate.getMonth() + 1;
      var targetYear = targetDate.getFullYear();
      var monthKey = targetYear + '-' + String(targetMonth).padStart(2, '0');
      ticketCountByMonth[monthKey] = 0;
      vendorCostByMonth[monthKey] = 0;
      subscriptionRevByMonth[monthKey] = 0;
      checkFeeRevByMonth[monthKey] = 0;
    }
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var lastUpdated = row[lastUpdatedIdx];
      var vendor = row[vendorIdx];
      var tags = row[tagsIdx];
      var outcome = row[outcomeIdx];
      var d = lastUpdated ? new Date(lastUpdated) : null;
      if (!d || isNaN(d.getTime())) continue;
      var year = d.getFullYear();
      var month = d.getMonth() + 1;
      var monthKey = year + '-' + String(month).padStart(2, '0');
      if (ticketCountByMonth.hasOwnProperty(monthKey)) {
        ticketCountByMonth[monthKey]++;
        // Vendor cost logic (example, adjust as needed)
        if (vendor === 'CERTN') vendorCostByMonth[monthKey] += 4.5;
        else if (vendor === 'INFORMDATA') vendorCostByMonth[monthKey] += 3.75;
        else if (vendor === 'PENNDOT') vendorCostByMonth[monthKey] += 0;
        // Subscription revenue
        subscriptionRevByMonth[monthKey] += 2;
        // Check fee revenue (billable logic)
        if (tags && (tags.indexOf('Updated') !== -1 || tags.indexOf('uploaded') !== -1) ||
            outcome === 'Suspension Confirmed' || outcome === 'Expired License' || outcome === 'Invalid License') {
          checkFeeRevByMonth[monthKey] += 5;
        }
      }
    }
    row = histStartRow;
    for (var m = 11; m >= 0; m--) {
      var targetDate = new Date(today.getFullYear(), today.getMonth() - m, 1);
      var targetMonth = targetDate.getMonth() + 1;
      var targetYear = targetDate.getFullYear();
      var monthKey = targetYear + '-' + String(targetMonth).padStart(2, '0');
      var monthName = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM yyyy');
      var bgColor = (11 - m) % 2 === 0 ? COLORS.WHITE : COLORS.LIGHT_GRAY;
      sheet.getRange(row, 2).setValue(monthName).setFontWeight('bold').setBackground(bgColor);
      sheet.getRange(row, 3).setValue(ticketCountByMonth[monthKey]).setNumberFormat('#,##0').setBackground(bgColor);
      sheet.getRange(row, 4).setValue(vendorCostByMonth[monthKey]).setNumberFormat('$#,##0.00').setBackground(bgColor);
      sheet.getRange(row, 5).setValue(subscriptionRevByMonth[monthKey]).setNumberFormat('$#,##0.00').setBackground(bgColor);
      sheet.getRange(row, 6).setValue(checkFeeRevByMonth[monthKey]).setNumberFormat('$#,##0.00').setBackground(bgColor);
      sheet.getRange(row, 7).setValue(subscriptionRevByMonth[monthKey] + checkFeeRevByMonth[monthKey] - vendorCostByMonth[monthKey]).setNumberFormat('$#,##0.00').setBackground(bgColor).setFontWeight('bold');
      row++;
    }
  
  row++;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // ROI SUMMARY - FINAL VERDICT
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('✅ FINAL VERDICT — DOES THE AT-COST MODEL WORK?')
    .setBackground('#be185d')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(13);
  sheet.setRowHeight(row, 36);
  row++;
  
  // ───────────────────────────────────────────────────────────────────────────
  // SUMMARY METRICS TABLE (All values pull directly from MVR_Ticket_History)
  // ───────────────────────────────────────────────────────────────────────────
  var totalRow = histStartRow + 12; // TOTAL row from 12-Month History table
  
  // Section header
  var summaryHeaders = ['METRIC', '', 'VALUE', 'HOW IT\'S CALCULATED', '', ''];
  sheet.getRange(row, 2, 1, 6).setValues([summaryHeaders])
    .setBackground('#1e293b')
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 28);
  row++;
  
  // Row 1: Total Tickets
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('📋 Total Tickets (All Time)')
    .setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(COUNTA('MVR_Ticket_History'!A:A)-1,0)")
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ COUNT of all rows in MVR_Ticket_History (minus header)')
    .setFontColor('#64748b')
    .setFontSize(9);
  var totalTicketsRow = row;
  row++;
  
  // Row 2: Billable Tickets - use direct formula instead of referencing billable section rows
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('✓ Billable Tickets')
    .setFontWeight('bold')
    .setBackground('#f0fdf4');
  // Direct calculation: Updated/uploaded tags + violation outcomes
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(SUMPRODUCT((ISNUMBER(SEARCH(\"Updated\",'MVR_Ticket_History'!U:U))+ISNUMBER(SEARCH(\"uploaded\",'MVR_Ticket_History'!U:U))>0)*1)+COUNTIF('MVR_Ticket_History'!AI:AI,\"Suspension Confirmed\")+COUNTIF('MVR_Ticket_History'!AI:AI,\"Expired License\")+COUNTIF('MVR_Ticket_History'!AI:AI,\"Invalid License\"),0)")
    .setNumberFormat('#,##0')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#059669')
    .setHorizontalAlignment('center')
    .setBackground('#f0fdf4');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Tags contain "Updated" OR "uploaded" OR violations')
    .setFontColor('#64748b')
    .setFontSize(9)
    .setBackground('#f0fdf4');
  var billableTicketsRow = row;
  row++;
  
  // Row 3: Billable Rate
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('📊 Billable Success Rate')
    .setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(D" + billableTicketsRow + "/D" + totalTicketsRow + ",0)")
    .setNumberFormat('0.0%')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#2563eb')
    .setHorizontalAlignment('center');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Billable ÷ Total (what % of checks we can bill)')
    .setFontColor('#64748b')
    .setFontSize(9);
  row++;
  
  // Divider
  sheet.getRange(row, 2, 1, 6).merge().setBackground('#e2e8f0');
  sheet.setRowHeight(row, 4);
  row++;
  
  // Row 4: Total Vendor Costs (use separate SUMPRODUCT per vendor to avoid nested IF issues)
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('💸 Total Vendor Costs')
    .setFontWeight('bold')
    .setBackground('#fef2f2');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(COUNTIF('MVR_Ticket_History'!Z:Z,\"CERTN\")*IFERROR(PARAM_COST_PER_CHECK_CERTN,4.5)+COUNTIF('MVR_Ticket_History'!Z:Z,\"INFORMDATA\")*IFERROR(PARAM_COST_PER_CHECK_INFORMDATA,3.75)+COUNTIF('MVR_Ticket_History'!Z:Z,\"PENNDOT\")*IFERROR(PARAM_COST_PER_CHECK_PENNDOT,0),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#dc2626')
    .setHorizontalAlignment('center')
    .setBackground('#fef2f2');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Each row × vendor rate (CERTN $4.50, INFORMDATA $3.75, PENNDOT $0)')
    .setFontColor('#64748b')
    .setFontSize(9)
    .setBackground('#fef2f2');
  var vendorCostRow = row;
  row++;
  
  // Row 5: Total Labor Costs (Derived from 10% allocation, not per-ticket assumption)
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('👷 Total Labor Costs')
    .setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(IFERROR(PARAM_TOTAL_MONTHLY_LABOR,150)*12,1800)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#dc2626')
    .setHorizontalAlignment('center');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Monthly Labor ($150) × 12 months (10% agent allocation)')
    .setFontColor('#64748b')
    .setFontSize(9);
  var laborCostRow = row;
  row++;
  
  // Divider
  sheet.getRange(row, 2, 1, 6).merge().setBackground('#e2e8f0');
  sheet.setRowHeight(row, 4);
  row++;
  
  // Row 6: Subscription Revenue - reference totalTicketsRow directly
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('💰 Subscription Revenue')
    .setFontWeight('bold')
    .setBackground('#f0fdf4');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(D" + totalTicketsRow + "*IFERROR(PARAM_DRIVER_MONTHLY_SUBSCRIPTION,2),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#059669')
    .setHorizontalAlignment('center')
    .setBackground('#f0fdf4');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ ALL Tickets × $2/ticket (all recheck tickets)')
    .setFontColor('#64748b')
    .setFontSize(9)
    .setBackground('#f0fdf4');
  var subRevRow = row;
  row++;
  
  // Row 7: Check Fee Revenue - reference billableTicketsRow directly
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('💵 Check Fee Revenue')
    .setFontWeight('bold');
  sheet.getRange(row, 4)
    .setFormula("=IFERROR(D" + billableTicketsRow + "*IFERROR(PARAM_PRICE_PER_BILLABLE_CHECK,5),0)")
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontColor('#059669')
    .setHorizontalAlignment('center');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Billable Tickets only × $5/check')
    .setFontColor('#64748b')
    .setFontSize(9);
  var checkRevRow = row;
  row++;
  
  // Divider - thicker
  sheet.getRange(row, 2, 1, 6).merge().setBackground('#1e293b');
  sheet.setRowHeight(row, 6);
  row++;
  
  // Row 8: TOTAL REVENUE
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('📈 TOTAL REVENUE')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#dcfce7');
  sheet.getRange(row, 4)
    .setFormula("=D" + subRevRow + "+D" + checkRevRow)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#059669')
    .setHorizontalAlignment('center')
    .setBackground('#dcfce7');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Subscription + Check Fee Revenue')
    .setFontColor('#166534')
    .setFontSize(9)
    .setFontWeight('bold')
    .setBackground('#dcfce7');
  var totalRevRow = row;
  row++;
  
  // Row 9: TOTAL COSTS
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('📉 TOTAL COSTS')
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#fee2e2');
  sheet.getRange(row, 4)
    .setFormula("=D" + vendorCostRow + "+D" + laborCostRow)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(14)
    .setFontColor('#dc2626')
    .setHorizontalAlignment('center')
    .setBackground('#fee2e2');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Vendor Costs + Labor Costs')
    .setFontColor('#991b1b')
    .setFontSize(9)
    .setFontWeight('bold')
    .setBackground('#fee2e2');
  var totalCostRow = row;
  row++;
  
  // Row 10: NET PROFIT/LOSS
  sheet.getRange(row, 2, 1, 2).merge()
    .setValue('🎯 NET PROFIT / LOSS')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#fef3c7');
  sheet.getRange(row, 4)
    .setFormula("=D" + totalRevRow + "-D" + totalCostRow)
    .setNumberFormat('$#,##0.00')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setBackground('#fef3c7');
  sheet.getRange(row, 5, 1, 2).merge()
    .setValue('→ Total Revenue − Total Costs')
    .setFontColor('#92400e')
    .setFontSize(9)
    .setFontWeight('bold')
    .setBackground('#fef3c7');
  var netProfitRow = row;
  sheet.setRowHeight(row, 32);
  row++;
  
  // Apply conditional formatting to net profit cell
  var profitCell = sheet.getRange(netProfitRow, 4);
  var profitPositive = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setFontColor('#059669')
    .setRanges([profitCell])
    .build();
  var profitNegative = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor('#dc2626')
    .setRanges([profitCell])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(profitPositive);
  rules.push(profitNegative);
  sheet.setConditionalFormatRules(rules);
  
  row++;
  
  // ───────────────────────────────────────────────────────────────────────────
  // VERDICT BOX
  // ───────────────────────────────────────────────────────────────────────────
  sheet.getRange(row, 2, 5, 6).merge()
    .setFormula("=IF(D" + netProfitRow + ">=0," +
      "\"✅ YES, IT WORKS!\\n\\n\"&" +
      "\"The at-cost pricing model is PROFITABLE.\\n\"&" +
      "\"Net Profit: $\"&TEXT(D" + netProfitRow + ",\"#,##0\")&\"\\n\"&" +
      "\"Profit per Ticket: $\"&TEXT(D" + netProfitRow + "/D" + (vendorCostRow-4) + ",\"0.00\")&\"\\n\\n\"&" +
      "\"Revenue ($\"&TEXT(D" + totalRevRow + ",\"#,##0\")&\") exceeds Costs ($\"&TEXT(D" + totalCostRow + ",\"#,##0\")&\").\"," +
      "\"⚠️ WARNING: MODEL IS LOSING MONEY\\n\\n\"&" +
      "\"The at-cost pricing model is NOT profitable.\\n\"&" +
      "\"Net Loss: $\"&TEXT(ABS(D" + netProfitRow + "),\"#,##0\")&\"\\n\"&" +
      "\"Loss per Ticket: $\"&TEXT(ABS(D" + netProfitRow + "/D" + (vendorCostRow-4) + "),\"0.00\")&\"\\n\\n\"&" +
      "\"Options: Raise prices, reduce labor costs, or improve billable rate.\")")
    .setBackground('#fdf2f8')
    .setFontColor('#1e293b')
    .setFontSize(12)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true)
    .setBorder(true, true, true, true, false, false, '#be185d', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.setRowHeight(row, 28);
  sheet.setRowHeight(row+1, 28);
  sheet.setRowHeight(row+2, 28);
  sheet.setRowHeight(row+3, 28);
  sheet.setRowHeight(row+4, 28);
  row += 6;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FOOTER
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('💡 Blue-bordered cells accept input. After entering costs, run "🔄 Refresh Finance Costs → Assumptions" from the menu.')
    .setFontColor(COLORS.ROYAL_BLUE)
    .setFontSize(9)
    .setFontStyle('italic');
  row++;
  
  sheet.getRange(row, 2, 1, 6).merge()
    .setValue('📊 Update pricing values in Assumptions → Pricing Model to recalculate ROI projections. Each row = 1 applicant check.')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic');
  
  // Freeze header rows
  sheet.setFrozenRows(15);
  
  // Store data row info for refresh function
  PropertiesService.getScriptProperties().setProperty('FINANCE_DATA_START_ROW', actualDataStartRow.toString());
  PropertiesService.getScriptProperties().setProperty('FINANCE_DATA_END_ROW', actualDataEndRow.toString());
  
  // Update KPI formulas with actual row numbers
  sheet.getRange(kpiRow, 4).setFormula("=IFERROR(SUM(G" + actualDataStartRow + ":G" + actualDataEndRow + "),0)");
  
  ui.alert('✅ Finance Dashboard Created',
    'Dashboard features:\n' +
    '• All 51 states with vendor dropdowns\n' +
    '• Editable True Cost column (blue border)\n' +
    '• Costs persist between sessions\n' +
    '• Vendor summary with weighted averages\n\n' +
    'To update Assumptions:\n' +
    '1. Enter your actual invoice costs\n' +
    '2. Menu → MVR Reports → Dashboards\n' +
    '3. Click "🔄 Refresh Finance Costs → Assumptions"',
    ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function createFinanceKPICard(sheet, row, col, valueFormula, label, numberFormat) {
  sheet.getRange(row, col)
    .setFormula(valueFormula)
    .setFontSize(24)
    .setFontWeight('bold')
    .setFontColor(COLORS.NAVY)
    .setHorizontalAlignment('center')
    .setNumberFormat(numberFormat || '#,##0');
  
  sheet.getRange(row + 1, col)
    .setValue(label)
    .setFontSize(10)
    .setFontColor(COLORS.SLATE)
    .setHorizontalAlignment('center');
  
  sheet.getRange(row, col, 2, 1)
    .setBorder(true, true, true, true, false, false, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
}

function addFinanceExplanationBox(sheet, row, col, span, lines) {
  sheet.getRange(row, col, 1, span).merge()
    .setValue(lines[0])
    .setBackground('#f0f9ff')
    .setFontColor(COLORS.NAVY)
    .setFontWeight('bold')
    .setFontSize(10);
  
  for (var i = 1; i < lines.length; i++) {
    sheet.getRange(row + i, col, 1, span).merge()
      .setValue(lines[i])
      .setBackground('#f0f9ff')
      .setFontColor(COLORS.SLATE)
      .setFontSize(9);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// HIDDEN SHEET FOR PERSISTING COSTS
// ═══════════════════════════════════════════════════════════════════════════════

function ensureFinanceCostsSheet(ss) {
  var sheet = ss.getSheetByName(FINANCE_COSTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FINANCE_COSTS_SHEET);
    sheet.hideSheet();
    
    // Set up headers
    sheet.getRange(1, 1, 1, 3).setValues([['State', 'Vendor', 'Cost']]);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}

function loadSavedCosts(ss) {
  var sheet = ss.getSheetByName(FINANCE_COSTS_SHEET);
  var costs = {};
  
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    data.forEach(function(row) {
      if (row[0]) {
        costs[row[0]] = { vendor: row[1] || 'INFORMDATA', cost: row[2] || 0 };
      }
    });
  }
  
  return costs;
}

function saveCostsToHiddenSheet(costsData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureFinanceCostsSheet(ss);
  
  // Clear existing data (except header)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clear();
  }
  
  // Write new data
  if (costsData.length > 0) {
    sheet.getRange(2, 1, costsData.length, 3).setValues(costsData);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// REFRESH FUNCTION - Updates Assumptions with Weighted Averages
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Refresh Finance Dashboard:
 * 1. Reads all True Cost entries from the dashboard
 * 2. Saves them to hidden sheet for persistence
 * 3. Calculates weighted average cost per vendor
 * 4. Updates Assumptions sheet with new averages
 */
function refreshFinanceDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Finance Dashboard');
  var assumptions = ss.getSheetByName('Assumptions');
  
  if (!dashboard) {
    SpreadsheetApp.getUi().alert('❌ Finance Dashboard not found.\n\nPlease create it first from:\nMenu → MVR Reports → Dashboards → Finance Dashboard');
    return;
  }
  
  // Get data row range from properties
  var props = PropertiesService.getScriptProperties();
  var startRow = parseInt(props.getProperty('FINANCE_DATA_START_ROW')) || 16;
  var endRow = parseInt(props.getProperty('FINANCE_DATA_END_ROW')) || 66;
  
  // Read all state data: State(B), StateName(C), Vendor(D), True Cost(E), Volume(F)
  var dataRange = dashboard.getRange(startRow, 2, endRow - startRow + 1, 5);
  var data = dataRange.getValues();
  
  // Calculate weighted averages per vendor
  var vendorTotals = {
    'CERTN': { costSum: 0, volumeSum: 0 },
    'INFORMDATA': { costSum: 0, volumeSum: 0 },
    'PENNDOT': { costSum: 0, volumeSum: 0 }
  };
  
  var costsToSave = [];
  
  data.forEach(function(row) {
    var state = row[0];      // Column B
    var vendor = row[2];     // Column D
    var trueCost = parseFloat(row[3]) || 0; // Column E
    var volume = parseInt(row[4]) || 0;     // Column F
    
    if (state && vendor && vendorTotals[vendor] !== undefined) {
      costsToSave.push([state, vendor, trueCost]);
      
      if (trueCost > 0 && volume > 0) {
        vendorTotals[vendor].costSum += trueCost * volume;
        vendorTotals[vendor].volumeSum += volume;
      }
    }
  });
  
  // Save costs to hidden sheet
  saveCostsToHiddenSheet(costsToSave);
  
  // Calculate weighted averages
  var weightedAvgs = {};
  for (var vendor in vendorTotals) {
    if (vendorTotals[vendor].volumeSum > 0) {
      weightedAvgs[vendor] = vendorTotals[vendor].costSum / vendorTotals[vendor].volumeSum;
    }
  }
  
  // Update Assumptions sheet if we have valid averages
  var updated = [];
  
  if (assumptions) {
    // Find and update the cost cells in Assumptions
    var assumpData = assumptions.getDataRange().getValues();
    
    for (var i = 0; i < assumpData.length; i++) {
      var paramName = assumpData[i][1]; // Column B = Parameter name
      
      if (paramName === 'Cost per Check - CERTN' && weightedAvgs['CERTN']) {
        assumptions.getRange(i + 1, 3).setValue(weightedAvgs['CERTN']);
        updated.push('CERTN: $' + weightedAvgs['CERTN'].toFixed(2));
      }
      else if (paramName === 'Cost per Check - INFORMDATA' && weightedAvgs['INFORMDATA']) {
        assumptions.getRange(i + 1, 3).setValue(weightedAvgs['INFORMDATA']);
        updated.push('INFORMDATA: $' + weightedAvgs['INFORMDATA'].toFixed(2));
      }
      else if (paramName === 'Cost per Check - PENNDOT' && weightedAvgs['PENNDOT']) {
        assumptions.getRange(i + 1, 3).setValue(weightedAvgs['PENNDOT']);
        updated.push('PENNDOT: $' + weightedAvgs['PENNDOT'].toFixed(2));
      }
    }
  }
  
  // Force refresh of NOW() formula to show update time
  var nowFormula = dashboard.getRange('B2').getFormula();
  if (nowFormula) {
    dashboard.getRange('B2').setFormula(nowFormula);
  }
  
  // Show result
  if (updated.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Updated Assumptions:\n' + updated.join('\n'),
      '✅ Costs Saved & Assumptions Updated',
      8
    );
    Logger.log('Finance Dashboard refreshed. Updated: ' + updated.join(', '));
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Your True Cost entries have been saved.\n\nTo update Assumptions, enter costs > $0 for states with volume.',
      '💾 Costs Saved',
      5
    );
    Logger.log('Finance Dashboard refreshed. No averages to update (all costs are $0 or no volume).');
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU ENTRY POINT
// ═══════════════════════════════════════════════════════════════════════════════

function runCreateFinanceDashboard() {
  createFinanceDashboard();
}

// ═══════════════════════════════════════════════════════════════════════════════
// CHART CREATION FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Creates visual charts for the Finance Dashboard
 * Call after dashboard is created to add charts
 */
function addFinanceDashboardCharts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Finance Dashboard');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Finance Dashboard not found. Please create it first.');
    return;
  }
  
  // Remove any existing charts
  var existingCharts = sheet.getCharts();
  existingCharts.forEach(function(c) { sheet.removeChart(c); });
  
  // Find billable section by searching for header
  var data = sheet.getDataRange().getValues();
  var billableStartRow = -1;
  var pricingRow = -1;
  var projectionRow = -1;
  
  for (var i = 0; i < data.length; i++) {
    var cellValue = String(data[i][1] || '');
    if (cellValue.indexOf('Billable - Violations Found') > -1) {
      billableStartRow = i + 1;
    }
    if (cellValue.indexOf('At-Cost Model') > -1) {
      pricingRow = i + 1;
    }
    if (cellValue.indexOf('Month 1') > -1) {
      projectionRow = i + 1;
    }
  }
  
  if (billableStartRow < 0) {
    SpreadsheetApp.getUi().alert('❌ Could not find billable analysis section. Please recreate the dashboard.');
    return;
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CHART 1: Pie Chart - Cost Distribution (Billable vs Absorbed)
  // ═══════════════════════════════════════════════════════════════════════════
  var pieDataRange = sheet.getRange(billableStartRow, 3, 6, 2); // Category + Count
  
  var pieChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieDataRange)
    .setPosition(billableStartRow, 9, 0, 0)
    .setOption('title', 'Cost Distribution by Category')
    .setOption('pieHole', 0.4)
    .setOption('colors', ['#059669', '#10b981', '#dc2626', '#ef4444', '#f97316', '#f59e0b'])
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('pieSliceText', 'percentage')
    .setOption('width', 350)
    .setOption('height', 250)
    .build();
  
  sheet.insertChart(pieChart);
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CHART 2: Bar Chart - Revenue vs Cost Comparison
  // ═══════════════════════════════════════════════════════════════════════════
  if (pricingRow > 0) {
    // Create a temporary data range for the bar chart
    // We'll use columns I-K for chart data
    sheet.getRange(pricingRow, 9).setValue('Model');
    sheet.getRange(pricingRow, 10).setValue('Revenue');
    sheet.getRange(pricingRow, 11).setValue('Cost');
    sheet.getRange(pricingRow + 1, 9).setValue('At-Cost');
    sheet.getRange(pricingRow + 1, 10).setFormula('=D' + (pricingRow + 3)); // At-cost revenue
    sheet.getRange(pricingRow + 1, 11).setFormula('=D' + (pricingRow + 4)); // Cost
    sheet.getRange(pricingRow + 2, 9).setValue('Full Price');
    sheet.getRange(pricingRow + 2, 10).setFormula('=E' + (pricingRow + 3)); // Full price revenue
    sheet.getRange(pricingRow + 2, 11).setFormula('=E' + (pricingRow + 4)); // Cost
    
    // Format as currency
    sheet.getRange(pricingRow + 1, 10, 2, 2).setNumberFormat('$#,##0');
    
    var barDataRange = sheet.getRange(pricingRow, 9, 3, 3);
    
    var barChart = sheet.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(barDataRange)
      .setPosition(pricingRow + 5, 9, 0, 0)
      .setOption('title', 'Revenue vs Cost by Pricing Model')
      .setOption('colors', ['#059669', '#dc2626'])
      .setOption('legend', { position: 'top' })
      .setOption('hAxis', { format: '$#,##0', title: 'Amount ($)' })
      .setOption('width', 350)
      .setOption('height', 200)
      .setOption('isStacked', false)
      .build();
    
    sheet.insertChart(barChart);
  }
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CHART 3: Line Chart - 12-Month Projection
  // ═══════════════════════════════════════════════════════════════════════════
  if (projectionRow > 0) {
    var lineDataRange = sheet.getRange(projectionRow - 1, 2, 13, 5); // Headers + 12 months
    
    var lineChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(lineDataRange)
      .setPosition(projectionRow, 9, 0, 0)
      .setOption('title', '12-Month Cost Offset Projection')
      .setOption('curveType', 'function')
      .setOption('colors', ['#0d9488', '#2563eb', '#dc2626'])
      .setOption('legend', { position: 'bottom', textStyle: { fontSize: 9 } })
      .setOption('hAxis', { title: 'Month', slantedText: true, slantedTextAngle: 45 })
      .setOption('vAxis', { format: '$#,##0', title: 'Cumulative Amount ($)' })
      .setOption('width', 400)
      .setOption('height', 280)
      .setOption('pointSize', 5)
      .build();
    
    sheet.insertChart(lineChart);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Charts added to Finance Dashboard', '📊 Charts Created', 3);
}

/**
 * Wrapper to create dashboard with charts
 */
function createFinanceDashboardWithCharts() {
  createFinanceDashboard();
  // Small delay to let formulas calculate
  Utilities.sleep(2000);
  addFinanceDashboardCharts();
}

// Extra closing brace to ensure file braces are balanced (fixes unexpected EOF parser error)
}
