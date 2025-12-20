/**
 * SIMPLIFIED MVR MONTHLY REPORT FOR FINANCE
 * 
 * Purpose: Clean, comprehensible report showing which MVRs were requested (billable) vs not
 * 
 * DYNAMIC RULE (from Tag_Outcome_Mappings sheet):
 * - Tags matching "Is Billable = TRUE" outcomes = REQUESTED (billable)
 * - Everything else = NOT REQUESTED with clear reason
 * 
 * Report shows:
 * - Identity: Month, Year, Partner
 * - Metrics: Total Tickets, MVR REQUESTED, MVR NOT Requested
 * - Reasons: Why NOT Requested (Already Valid, DMV Down, Missing Data, Still Processing, Other)
 * - Billing: Turn IDs to Bill
 */

// ═══════════════════════════════════════════════════════════════════════════════
// TAG CLASSIFICATION (Uses Dynamic Mappings)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Classify whether MVR was requested based on tags
 * Uses dynamic tag mappings from Tag_Outcome_Mappings sheet for consistency
 * 
 * RULE: If outcome is billable in Tag_Outcome_Mappings → REQUESTED
 * 
 * @param {string} tagsString - Comma-separated tags string from ticket
 * @return {Object} Classification with requested (boolean) and reason (string)
 */
function classifyMVRRequest(tagsString) {
  if (!tagsString || tagsString.trim() === '') {
    return { requested: false, reason: 'Still Processing' };
  }
  
  // Use shared classification logic from OutcomeTracking.gs
  const tags = tagsString.split(',').map(t => t.trim());
  const result = classifyOutcome(tags, []);
  
  // Billable = Requested
  if (result.isBillable) {
    return { requested: true, reason: result.outcome };
  }
  
  // Map outcome to reason for non-billable
  const reasonMap = {
    'Clear': 'Already Valid',
    'DMV Unavailable': 'DMV Down',
    'Cannot Process': 'Missing Data',
    'Still Processing': 'Still Processing',
    'Unknown': 'Other'
  };
  
  return { 
    requested: false, 
    reason: reasonMap[result.outcome] || result.outcome 
  };
}

// ═══════════════════════════════════════════════════════════════════════════════
// MONTHLY REPORT GENERATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate monthly MVR request report by partner
 * Groups tickets by month + partner and lists requested turn IDs for finance approval
 * 
 * @return {Object} Monthly breakdown by partner with requested turn IDs
 */
function generateMonthlyRequestReport() {
  Logger.log('\n📊 Generating monthly MVR request report by partner...\n');
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found. Run data fetch first.');
    }
    
    const data = historySheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('⚠️ No ticket data in history');
      return { partners: [], total: 0 };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Group by month-partner (format: "2025-11|PartnerName")
    const partnerMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      // Parse ISO date string to Date object for correct month/year calculations
      const createdRaw = data[i][colMap['Created']];
      const created = createdRaw ? new Date(createdRaw) : null;
      const type = data[i][colMap['Type']];
      const tags = data[i][colMap['Tags']] || '';
      const partner = data[i][colMap['Partner']] || 'Unknown';
      const turnId = data[i][colMap['Turn_ID']] || '';
      
      // Get month key (YYYY-MM)
      const year = created.getFullYear();
      const month = String(created.getMonth() + 1).padStart(2, '0');
      const monthKey = `${year}-${month}`;
      const partnerKey = `${monthKey}|${partner}`;
      
      // Initialize partner entry if not exists
      if (!partnerMap.has(partnerKey)) {
        partnerMap.set(partnerKey, {
          month: month,
          monthName: created.toLocaleString('en-US', { month: 'long' }),
          year: year,
          partner: partner,
          totalTickets: 0,
          requested: 0,
          notRequested: 0,
          alreadyValid: 0,
          dmvDown: 0,
          missingData: 0,
          stillProcessing: 0,
          other: 0,
          requestedTurnIds: []
        });
      }
      
      const entry = partnerMap.get(partnerKey);
      
      // Update counts
      entry.totalTickets++;
      
      // Classify request status
      const classification = classifyMVRRequest(tags);
      
      if (classification.requested) {
        // REQUESTED (suspended tags) - THIS IS BILLABLE
        entry.requested++;
        
        // Add turn ID to requested list (for finance to approve charges)
        if (turnId) {
          entry.requestedTurnIds.push(turnId);
        }
      } else {
        // NOT REQUESTED - count the reason
        entry.notRequested++;
        
        if (classification.reason === 'Already Valid') {
          entry.alreadyValid++;
        } else if (classification.reason === 'DMV Down') {
          entry.dmvDown++;
        } else if (classification.reason === 'Missing Data') {
          entry.missingData++;
        } else if (classification.reason === 'Still Processing') {
          entry.stillProcessing++;
        } else {
          entry.other++;
        }
      }
    }
    
    // Convert to array and calculate request rate
    const partners = Array.from(partnerMap.values()).map(entry => {
      // Calculate request rate percentage
      const requestRate = entry.totalTickets > 0 
        ? ((entry.requested / entry.totalTickets) * 100).toFixed(1)
        : '0.0';
      
      // Join turn IDs with comma
      const turnIdsString = entry.requestedTurnIds.join(', ');
      
      return {
        ...entry,
        requestRate: requestRate,
        requestedTurnIdsString: turnIdsString
      };
    });
    
    // Sort by year-month descending, then by partner
    partners.sort((a, b) => {
      const keyA = `${a.year}-${a.month}`;
      const keyB = `${b.year}-${b.month}`;
      const dateCompare = keyB.localeCompare(keyA);
      if (dateCompare !== 0) return dateCompare;
      return a.partner.localeCompare(b.partner);
    });
    
    Logger.log(`✅ Generated report for ${partners.length} partner-months`);
    
    return {
      partners: partners,
      total: partners.length,
      generated: new Date()
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating monthly report: ${e.message}`);
    throw e;
  }
}

/**
 * Write monthly report to sheet
 * @param {Object} report - Monthly report object
 */
function writeMonthlyReport(report) {
  Logger.log('\n📝 Writing monthly report to sheet...\n');
  
  try {
    // CLEAN, SIMPLE HEADERS
    const headers = [
      'Month',
      'Year', 
      'Partner',
      'Total Tickets',
      'MVR REQUESTED (Billable)',
      'MVR NOT Requested',
      'Why: Already Valid',
      'Why: DMV Down',
      'Why: Missing Data',
      'Why: Still Processing',
      'Why: Other',
      'Request Rate %',
      'Billable Turn IDs (For Finance Approval)',
      'Report Date'
    ];
    
    const reportSheet = getOrCreateSheet(SHEET_NAMES.MONTHLY_REPORT, headers);
    
    // Always update headers in case structure changed
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Clear existing data and old columns
    const lastRow = reportSheet.getLastRow();
    const lastCol = reportSheet.getMaxColumns();
    if (lastRow > 1) {
      reportSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      reportSheet.getRange(2, 1, lastRow - 1, lastCol).clearFormat();
    }
    
    // Clear any extra columns beyond current headers
    if (lastCol > headers.length) {
      reportSheet.deleteColumns(headers.length + 1, lastCol - headers.length);
    }
    
    if (report.partners.length === 0) {
      Logger.log('⚠️ No monthly data to write');
      return;
    }
    
    // Convert partners to rows - clean and simple
    const rows = report.partners.map(p => [
      p.monthName,
      p.year,
      p.partner,
      p.totalTickets,
      p.requested,
      p.notRequested,
      p.alreadyValid,
      p.dmvDown,
      p.missingData,
      p.stillProcessing,
      p.other,
      p.requestRate + '%',
      p.requestedTurnIdsString,
      report.generated
    ]);
    
    // Write rows
    reportSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Format sheet
    formatMonthlyReportSheet(reportSheet, rows.length);
    
    Logger.log(`✅ Wrote ${rows.length} monthly summaries to sheet`);
    
  } catch (e) {
    Logger.log(`❌ Error writing monthly report: ${e.message}`);
    throw e;
  }
}

/**
 * Format monthly report sheet with colors and styles
 * @param {Sheet} sheet - Monthly report sheet
 * @param {number} rowCount - Number of data rows
 */
function formatMonthlyReportSheet(sheet, rowCount) {
  if (rowCount === 0) return;
  
  const dataRows = rowCount;
  
  // Header row - bold and colored (14 columns)
  const headerRange = sheet.getRange(1, 1, 1, 14);
  headerRange
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  sheet.setRowHeight(1, 50);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  
  // Simple color coding - only the important columns
  if (dataRows > 0) {
    // MVR REQUESTED (column 5) - GREEN - THIS IS BILLABLE
    sheet.getRange(2, 5, dataRows, 1)
      .setBackground('#d4edda')
      .setFontWeight('bold')
      .setFontSize(11);
    
    // MVR NOT Requested (column 6) - LIGHT RED
    sheet.getRange(2, 6, dataRows, 1)
      .setBackground('#f8d7da')
      .setFontSize(10);
    
    // Reason columns (columns 7-11) - LIGHT GREY - subtle
    sheet.getRange(2, 7, dataRows, 5)
      .setBackground('#f5f5f5')
      .setFontSize(9);
    
    // Request Rate % (column 12) - BOLD with center alignment
    sheet.getRange(2, 12, dataRows, 1)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setFontSize(10);
    
    // Billable Turn IDs (column 13) - YELLOW HIGHLIGHT - THIS IS FOR BILLING
    sheet.getRange(2, 13, dataRows, 1)
      .setBackground('#fff9e6')
      .setFontFamily('Courier New')
      .setFontSize(9)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
  
  // Auto-resize identity columns
  sheet.autoResizeColumns(1, 4);
  
  // Set specific widths for clarity
  sheet.setColumnWidth(5, 140); // MVR REQUESTED (Billable)
  sheet.setColumnWidth(6, 120); // MVR NOT Requested
  sheet.setColumnWidth(12, 100); // Request Rate %
  sheet.setColumnWidth(13, 350); // Billable Turn IDs
  
  Logger.log('✅ Monthly report sheet formatted - clean and comprehensible');
}

// ═══════════════════════════════════════════════════════════════════════════════
// TESTING & VALIDATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test tag classification with various tag patterns
 */
function testTagClassification() {
  Logger.log('\n🧪 Testing MVR request classification...\n');
  
  const testCases = [
    { tags: 'valid, driver', expected: 'NOT requested (valid tags)' },
    { tags: 'suspended note, no action needed', expected: 'REQUESTED (suspended)' },
    { tags: 'dmv down, retry later', expected: 'NOT requested (DMV down)' },
    { tags: 'mvr requested, pending', expected: 'REQUESTED' },
    { tags: 'valid tags', expected: 'NOT requested (valid)' },
    { tags: 'suspension-note', expected: 'REQUESTED (suspended)' },
    { tags: 'DMV_Down', expected: 'NOT requested (DMV down)' },
    { tags: '', expected: 'REQUESTED (default)' },
    { tags: 'random tag, other stuff', expected: 'REQUESTED (default)' },
    { tags: 'already valid mvr on file', expected: 'NOT requested (valid)' },
    { tags: 'MVR order placed', expected: 'REQUESTED' },
    { tags: 'suspended', expected: 'REQUESTED (suspended)' }
  ];
  
  Logger.log('Testing classification logic:');
  Logger.log('─'.repeat(80));
  
  testCases.forEach(test => {
    const result = classifyMVRRequest(test.tags);
    const status = result.requested ? '✅ REQUESTED' : '❌ NOT REQUESTED';
    
    Logger.log(`Tags: "${test.tags}"`);
    Logger.log(`  Result: ${status}`);
    Logger.log(`  Category: ${result.category}`);
    Logger.log(`  Reason: ${result.reason}`);
    Logger.log(`  Expected: ${test.expected}`);
    Logger.log('');
  });
  
  Logger.log('─'.repeat(80));
  Logger.log('✅ Tag classification test complete\n');
}

/**
 * Quick test of monthly report generation
 */
function testMonthlyReportGeneration() {
  Logger.log('\n🧪 Testing monthly report generation...\n');
  
  try {
    const report = generateMonthlyRequestReport();
    
    Logger.log(`✅ Generated report with ${report.total} months:`);
    
    report.months.forEach(month => {
      Logger.log(`\n${month.monthName} ${month.year}:`);
      Logger.log(`  Total: ${month.total}`);
      Logger.log(`  SC: ${month.sc_count}, RC: ${month.rc_count}`);
      Logger.log(`  Requested: ${month.requested_count}`);
      Logger.log(`  Not Requested: ${month.not_requested_count}`);
      Logger.log(`  - Valid Tags: ${month.valid_tags_count}`);
      Logger.log(`  - Suspended: ${month.suspended_count}`);
      Logger.log(`  - DMV Down: ${month.dmv_down_count}`);
      Logger.log(`  Request Rate: ${month.request_rate}%`);
    });
    
    Logger.log('\n✅ Monthly report test complete\n');
    
  } catch (e) {
    Logger.log(`❌ Test failed: ${e.message}`);
  }
}
