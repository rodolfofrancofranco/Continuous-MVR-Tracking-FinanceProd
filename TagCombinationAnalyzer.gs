/**
 * TAG COMBINATION ANALYZER (STANDALONE UTILITY)
 * 
 * Purpose: Analyze tag co-occurrence patterns to design complex regex mappings
 * Usage: Add to menu and run to generate pattern report
 * 
 * This is NOT part of the main pipeline - it's a planning tool
 */

// ═══════════════════════════════════════════════════════════════════════════════
// TAG COMBINATION ANALYSIS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Analyze all tag combinations in raw tickets
 * Shows which tags appear together and how frequently
 * Helps design multi-condition regex patterns
 * 
 * @return {Object} Analysis results with combinations and frequencies
 */
function analyzeTagCombinations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
  
  if (!rawSheet || rawSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Error', 'No raw tickets found to analyze.', SpreadsheetApp.getUi().ButtonSet.OK);
    return { success: false, error: 'No data' };
  }
  
  Logger.log('\n🔍 === TAG COMBINATION ANALYSIS ===');
  
  // Get tags column
  const headers = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  const tagsColIdx = headers.indexOf('Tags');
  
  if (tagsColIdx === -1) {
    Logger.log('❌ Tags column not found');
    return { success: false, error: 'Tags column not found' };
  }
  
  const lastRow = rawSheet.getLastRow();
  const tagsData = rawSheet.getRange(2, tagsColIdx + 1, lastRow - 1, 1).getValues();
  
  // Track combinations
  const combinations = {};
  const tagPairs = {};
  const individualTags = {};
  let totalTickets = 0;
  
  tagsData.forEach(row => {
    const tagsStr = row[0] || '';
    if (!tagsStr.trim()) return;
    
    totalTickets++;
    
    // Parse tags
    const tags = tagsStr.split(',').map(t => t.trim()).filter(t => t).sort();
    
    if (tags.length === 0) return;
    
    // Track individual tags
    tags.forEach(tag => {
      if (!individualTags[tag]) individualTags[tag] = 0;
      individualTags[tag]++;
    });
    
    // Track full combination (sorted for consistency)
    const combo = tags.join(' + ');
    if (!combinations[combo]) {
      combinations[combo] = {
        tags: tags,
        count: 0
      };
    }
    combinations[combo].count++;
    
    // Track pairs (for 2+ tag combinations)
    if (tags.length >= 2) {
      for (let i = 0; i < tags.length; i++) {
        for (let j = i + 1; j < tags.length; j++) {
          const pair = [tags[i], tags[j]].sort().join(' + ');
          if (!tagPairs[pair]) tagPairs[pair] = 0;
          tagPairs[pair]++;
        }
      }
    }
  });
  
  // Sort combinations by frequency
  const sortedCombos = Object.entries(combinations)
    .map(([combo, data]) => ({
      pattern: combo,
      tags: data.tags,
      count: data.count,
      percentage: ((data.count / totalTickets) * 100).toFixed(2)
    }))
    .sort((a, b) => b.count - a.count);
  
  // Sort pairs by frequency
  const sortedPairs = Object.entries(tagPairs)
    .map(([pair, count]) => ({
      pattern: pair,
      count: count,
      percentage: ((count / totalTickets) * 100).toFixed(2)
    }))
    .sort((a, b) => b.count - a.count);
  
  Logger.log(`\n📊 Analysis of ${totalTickets} tickets with tags:`);
  Logger.log(`   Unique tag combinations: ${sortedCombos.length}`);
  Logger.log(`   Unique tag pairs: ${sortedPairs.length}`);
  Logger.log(`   Individual tags: ${Object.keys(individualTags).length}`);
  
  // Create output sheets
  createCombinationAnalysisSheet(sortedCombos, sortedPairs, individualTags, totalTickets);
  const recommendations = generateMappingRecommendations(sortedCombos, totalTickets);
  
  return {
    success: true,
    totalTickets: totalTickets,
    combinations: sortedCombos,
    pairs: sortedPairs,
    individualTags: individualTags,
    recommendations: recommendations
  };
}

/**
 * Create/update sheet with tag combination analysis
 * Provides actionable regex patterns for mapping
 */
function createCombinationAnalysisSheet(combinations, pairs, individualTags, totalTickets) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or clear sheet
  let sheet = ss.getSheetByName('Tag_Combination_Analysis');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Tag_Combination_Analysis');
  }
  
  // === SECTION 1: FULL COMBINATIONS ===
  let row = 1;
  
  // Header
  sheet.getRange(row, 1, 1, 6).setValues([['TAG COMBINATION ANALYSIS', '', '', '', '', '']]);
  sheet.getRange(row, 1, 1, 6).merge().setFontWeight('bold').setFontSize(14).setBackground('#4285f4').setFontColor('#ffffff');
  row += 2;
  
  sheet.getRange(row, 1).setValue(`Total Tickets Analyzed: ${totalTickets}`);
  sheet.getRange(row, 1).setFontWeight('bold');
  row += 2;
  
  // Combinations table
  sheet.getRange(row, 1, 1, 6).setValues([['FULL TAG COMBINATIONS (Top 50)', '', '', '', '', '']]);
  sheet.getRange(row, 1, 1, 6).merge().setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  sheet.getRange(row, 1, 1, 6).setValues([['Rank', 'Pattern', 'Tag Count', 'Frequency', '% of Total', 'Suggested Regex']]);
  sheet.getRange(row, 1, 1, 6).setFontWeight('bold').setBackground('#f3f3f3');
  row++;
  
  const top50Combos = combinations.slice(0, 50);
  top50Combos.forEach((combo, idx) => {
    const tagCount = combo.tags.length;
    
    // Generate regex pattern
    let regexPattern = '';
    if (tagCount === 1) {
      regexPattern = `^${escapeRegexChars(combo.tags[0])}$`;
    } else {
      // Positive lookahead for each tag (order-independent)
      const lookaheads = combo.tags.map(tag => `(?=.*\\b${escapeRegexChars(tag)}\\b)`).join('');
      regexPattern = `${lookaheads}.*`;
    }
    
    sheet.getRange(row, 1, 1, 6).setValues([[
      idx + 1,
      combo.pattern,
      tagCount,
      combo.count,
      `${combo.percentage}%`,
      regexPattern
    ]]);
    
    // Color code by frequency
    if (combo.count >= 50) {
      sheet.getRange(row, 4).setBackground('#fce5cd'); // High volume - orange
    } else if (combo.count >= 10) {
      sheet.getRange(row, 4).setBackground('#fff2cc'); // Medium - yellow
    }
    
    row++;
  });
  
  row += 2;
  
  // === SECTION 2: TAG PAIRS (Co-occurrence) ===
  sheet.getRange(row, 1, 1, 5).setValues([['COMMON TAG PAIRS (Top 30)', '', '', '', '']]);
  sheet.getRange(row, 1, 1, 5).merge().setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  sheet.getRange(row, 1, 1, 5).setValues([['Rank', 'Tag 1 + Tag 2', 'Frequency', '% of Total', 'Suggested Regex']]);
  sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
  row++;
  
  const top30Pairs = pairs.slice(0, 30);
  top30Pairs.forEach((pair, idx) => {
    const tags = pair.pattern.split(' + ');
    const regexPattern = `(?=.*\\b${escapeRegexChars(tags[0])}\\b)(?=.*\\b${escapeRegexChars(tags[1])}\\b).*`;
    
    sheet.getRange(row, 1, 1, 5).setValues([[
      idx + 1,
      pair.pattern,
      pair.count,
      `${pair.percentage}%`,
      regexPattern
    ]]);
    
    if (pair.count >= 20) {
      sheet.getRange(row, 3).setBackground('#fce5cd');
    } else if (pair.count >= 5) {
      sheet.getRange(row, 3).setBackground('#fff2cc');
    }
    
    row++;
  });
  
  row += 2;
  
  // === SECTION 3: CONDITIONAL MAPPINGS GUIDE ===
  sheet.getRange(row, 1, 1, 5).setValues([['CONDITIONAL MAPPING EXAMPLES', '', '', '', '']]);
  sheet.getRange(row, 1, 1, 5).merge().setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  const examples = [
    ['Pattern Type', 'Example Tags', 'Regex Pattern', 'Outcome', 'Logic'],
    ['Single Tag', 'Clear', '^Clear$', 'Clear', 'Exact match'],
    ['Any Tag', 'updated, emailed', 'updated|emailed', 'Various', 'OR condition'],
    ['Both Tags Required', 'Consider + requested', '(?=.*\\bConsider\\b)(?=.*\\brequested\\b).*', 'Still Processing', 'AND condition (order-independent)'],
    ['All 3 Tags Required', 'Consider + requested + Updated', '(?=.*\\bConsider\\b)(?=.*\\brequested\\b)(?=.*\\bUpdated\\b).*', 'Clear', 'Multiple AND conditions'],
    ['Tag + Not Another', 'updated BUT NOT rejected', '(?=.*\\bupdated\\b)(?!.*\\brejected\\b).*', 'Clear', 'Positive + negative lookahead'],
    ['Wildcard Match', 'Pending*', '^Pending', 'Pending', 'Starts with'],
    ['Contains Substring', '*same info*', '.*same info.*', 'Clear', 'Anywhere in tags'],
  ];
  
  examples.forEach((example, idx) => {
    sheet.getRange(row, 1, 1, 5).setValues([example]);
    if (idx === 0) {
      sheet.getRange(row, 1, 1, 5).setFontWeight('bold').setBackground('#f3f3f3');
    }
    row++;
  });
  
  row += 2;
  
  // === SECTION 4: INDIVIDUAL TAG FREQUENCY ===
  sheet.getRange(row, 1, 1, 4).setValues([['INDIVIDUAL TAG FREQUENCY', '', '', '']]);
  sheet.getRange(row, 1, 1, 4).merge().setFontWeight('bold').setBackground('#e8f0fe');
  row++;
  
  sheet.getRange(row, 1, 1, 4).setValues([['Rank', 'Tag', 'Frequency', '% of Total']]);
  sheet.getRange(row, 1, 1, 4).setFontWeight('bold').setBackground('#f3f3f3');
  row++;
  
  const sortedIndividual = Object.entries(individualTags)
    .map(([tag, count]) => ({
      tag: tag,
      count: count,
      percentage: ((count / totalTickets) * 100).toFixed(2)
    }))
    .sort((a, b) => b.count - a.count);
  
  sortedIndividual.forEach((tag, idx) => {
    sheet.getRange(row, 1, 1, 4).setValues([[
      idx + 1,
      tag.tag,
      tag.count,
      `${tag.percentage}%`
    ]]);
    row++;
  });
  
  // Format columns
  sheet.setColumnWidth(1, 60);  // Rank
  sheet.setColumnWidth(2, 300); // Pattern/Tags
  sheet.setColumnWidth(3, 100); // Count
  sheet.setColumnWidth(4, 100); // Percentage
  sheet.setColumnWidth(5, 80);  // Tag count
  sheet.setColumnWidth(6, 400); // Regex
  
  // Freeze header
  sheet.setFrozenRows(1);
  
  Logger.log(`✅ Created Tag_Combination_Analysis sheet`);
}

/**
 * Generate specific combination patterns for planning
 * Focuses on high-value combinations for mapping decisions
 */
function generateMappingRecommendations(combinations, totalTickets) {
  const recommendations = [];
  
  // Analyze high-frequency combinations
  combinations.slice(0, 30).forEach(combo => {
    const tags = combo.tags;
    const pattern = combo.pattern;
    
    // Pattern detection logic based on your business rules
    let suggestedOutcome = 'Unknown';
    let billable = 'No';
    let reasoning = '';
    
    // Multi-tag analysis
    if (tags.includes('Consider') && tags.includes('requested') && tags.includes('Updated')) {
      suggestedOutcome = 'Clear';
      billable = 'Yes';
      reasoning = 'Report had hits, reviewed, and uploaded';
    } else if (tags.includes('same information') || tags.includes('same info')) {
      suggestedOutcome = 'Clear';
      billable = 'Yes';
      reasoning = 'No changes found - successful check';
    } else if (tags.includes('Updated') && tags.length === 1) {
      suggestedOutcome = 'Clear';
      billable = 'Yes';
      reasoning = 'Report uploaded to employer';
    } else if (tags.includes('emailed') && !tags.includes('Updated')) {
      suggestedOutcome = 'Still Processing';
      billable = 'No';
      reasoning = 'Waiting for applicant information';
    } else if (tags.includes('Record not found') && tags.length === 1) {
      suggestedOutcome = 'Record Not Found';
      billable = 'Maybe';
      reasoning = 'NEEDS REVIEW: Could be vendor issue or missing info';
    } else if (tags.includes('Review identity')) {
      suggestedOutcome = 'Still Processing';
      billable = 'No';
      reasoning = 'Waiting for identity verification from applicant';
    } else if (tags.includes('Withdrawn')) {
      suggestedOutcome = 'Withdrawn';
      billable = 'No';
      reasoning = 'NEEDS CONTEXT: Check other tags for withdrawal reason';
    } else if (tags.includes('approved') && tags.includes('Clear')) {
      suggestedOutcome = 'Clear';
      billable = 'Yes';
      reasoning = 'Report cleared after review';
    } else if (tags.includes('DL Expired') || tags.includes('expired')) {
      suggestedOutcome = 'Expired License';
      billable = 'Yes';
      reasoning = 'Violation found';
    } else if (tags.includes('pending') || pattern.startsWith('Pending')) {
      suggestedOutcome = 'Pending';
      billable = 'No';
      reasoning = 'Awaiting information';
    } else if (tags.includes('requested') && tags.length === 1) {
      suggestedOutcome = 'Still Processing';
      billable = 'No';
      reasoning = 'MVR requested but not yet received';
    }
    
    recommendations.push({
      pattern: pattern,
      count: combo.count,
      percentage: combo.percentage,
      outcome: suggestedOutcome,
      billable: billable,
      reasoning: reasoning
    });
  });
  
  // Output to sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Mapping_Recommendations');
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet('Mapping_Recommendations');
  }
  
  sheet.getRange(1, 1, 1, 7).setValues([['MAPPING RECOMMENDATIONS (AUTO-GENERATED)', '', '', '', '', '', '']]);
  sheet.getRange(1, 1, 1, 7).merge().setFontWeight('bold').setFontSize(14).setBackground('#34a853').setFontColor('#ffffff');
  
  sheet.getRange(3, 1, 1, 7).setValues([['Tag Pattern', 'Frequency', '% Total', 'Suggested Outcome', 'Billable?', 'Reasoning', 'Regex Pattern']]);
  sheet.getRange(3, 1, 1, 7).setFontWeight('bold').setBackground('#f3f3f3');
  
  let row = 4;
  recommendations.forEach(rec => {
    const tags = rec.pattern.split(' + ');
    let regexPattern = '';
    if (tags.length === 1) {
      regexPattern = `^${escapeRegexChars(tags[0])}$`;
    } else {
      const lookaheads = tags.map(tag => `(?=.*\\b${escapeRegexChars(tag)}\\b)`).join('');
      regexPattern = `${lookaheads}.*`;
    }
    
    sheet.getRange(row, 1, 1, 7).setValues([[
      rec.pattern,
      rec.count,
      `${rec.percentage}%`,
      rec.outcome,
      rec.billable,
      rec.reasoning,
      regexPattern
    ]]);
    
    // Color code by confidence
    if (rec.billable === 'Yes') {
      sheet.getRange(row, 4).setBackground('#d9ead3'); // Green - billable
    } else if (rec.billable === 'Maybe') {
      sheet.getRange(row, 4).setBackground('#fff2cc'); // Yellow - needs review
    }
    
    // Highlight patterns needing manual review
    if (rec.reasoning.includes('NEEDS')) {
      sheet.getRange(row, 6).setBackground('#f4cccc'); // Red - manual review
      sheet.getRange(row, 6).setFontWeight('bold');
    }
    
    row++;
  });
  
  sheet.setColumnWidth(1, 350);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 450);
  sheet.setColumnWidth(7, 400);
  
  sheet.activate();
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Created ${recommendations.length} mapping recommendations. Check "Mapping_Recommendations" sheet.`,
    '✅ Analysis Complete',
    5
  );
  
  Logger.log(`✅ Created Mapping_Recommendations sheet with ${recommendations.length} suggestions`);
  
  return recommendations;
}

/**
 * Helper: Escape special regex characters
 */
function escapeRegexChars(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Create/reset Tag_Outcome_Mappings sheet based on analysis
 * Only includes high-confidence mappings (billable='Yes' or high frequency)
 * 
 * @param {Array} recommendations - Analysis recommendations
 * @return {number} Number of mappings created
 */
function createTagMappingsFromAnalysis(recommendations) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('\n🔄 Creating Tag_Outcome_Mappings from analysis...');
  
  // Delete existing sheet if it exists
  let mappingSheet = ss.getSheetByName('Tag_Outcome_Mappings');
  if (mappingSheet) {
    ss.deleteSheet(mappingSheet);
    Logger.log('   Deleted existing Tag_Outcome_Mappings sheet');
  }
  
  // Create new sheet
  mappingSheet = ss.insertSheet('Tag_Outcome_Mappings');
  
  // Create header
  const headers = [
    'Tag Pattern',
    'Match Type',
    'Outcome Type',
    'Priority',
    'Is Billable',
    'Is Active',
    'Notes'
  ];
  
  mappingSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  mappingSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');
  
  // Filter high-confidence recommendations
  const highConfidence = recommendations.filter(rec => {
    // Only include "Yes" billable or strong patterns
    return rec.billable === 'Yes' || 
           (rec.count >= 10 && !rec.reasoning.includes('NEEDS'));
  });
  
  Logger.log(`   Found ${highConfidence.length} high-confidence patterns to map`);
  
  // Create mappings with appropriate priorities
  const mappings = [];
  
  highConfidence.forEach((rec, idx) => {
    const tags = rec.pattern.split(' + ');
    
    // Determine priority based on specificity and frequency
    let priority;
    if (rec.billable === 'Yes' && tags.length === 1) {
      // Single-tag billable
      if (rec.outcome.includes('Violation') || rec.outcome.includes('Expired') || rec.outcome.includes('Suspended')) {
        priority = 100; // Highest - violations
      } else if (rec.outcome === 'Clear') {
        priority = 90; // High - clear completions
      } else {
        priority = 80;
      }
    } else if (tags.length > 1) {
      priority = 95; // Very high - multi-tag patterns (more specific)
    } else if (rec.billable === 'No') {
      priority = 50; // Lower - non-billable states
    } else {
      priority = 70; // Medium - everything else
    }
    
    // Determine match type and pattern
    let matchType = 'contains';
    let pattern = tags[0]; // Default to first tag
    
    if (tags.length === 1) {
      // Single tag - use contains
      pattern = tags[0];
      matchType = 'contains';
    } else if (tags.length > 1) {
      // Multi-tag combination - use regex
      const lookaheads = tags.map(tag => `(?=.*\\b${escapeRegexChars(tag)}\\b)`).join('');
      pattern = `${lookaheads}.*`;
      matchType = 'regex';
    }
    
    // Create note
    const note = `${rec.reasoning} (${rec.count} tickets, ${rec.percentage}% of total)`;
    
    mappings.push([
      pattern,
      matchType,
      rec.outcome,
      priority,
      rec.billable === 'Yes', // Is Billable checkbox
      true, // Is Active checkbox
      note
    ]);
    
    Logger.log(`   [${priority}] ${pattern} → ${rec.outcome} (${matchType})`);
  });
  
  // Sort by priority (highest first)
  mappings.sort((a, b) => b[3] - a[3]);
  
  // Write to sheet
  if (mappings.length > 0) {
    mappingSheet.getRange(2, 1, mappings.length, headers.length).setValues(mappings);
    
    // Format columns
    mappingSheet.setColumnWidth(1, 300); // Tag Pattern
    mappingSheet.setColumnWidth(2, 100); // Match Type
    mappingSheet.setColumnWidth(3, 150); // Outcome Type
    mappingSheet.setColumnWidth(4, 80);  // Priority
    mappingSheet.setColumnWidth(5, 100); // Is Billable
    mappingSheet.setColumnWidth(6, 80);  // Is Active
    mappingSheet.setColumnWidth(7, 400); // Notes
    
    // Add dropdown validation for Match Type
    const matchTypeRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['contains', 'regex', 'exact'], true)
      .build();
    mappingSheet.getRange(2, 2, mappings.length, 1).setDataValidation(matchTypeRule);
    
    // Add checkbox validation for Is Billable and Is Active
    const checkboxRule = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .build();
    mappingSheet.getRange(2, 5, mappings.length, 2).setDataValidation(checkboxRule);
    
    // Color code by billable status
    for (let i = 0; i < mappings.length; i++) {
      const isBillable = mappings[i][4];
      const bgColor = isBillable ? '#d9ead3' : '#fff2cc'; // Green vs Yellow
      mappingSheet.getRange(i + 2, 1, 1, headers.length).setBackground(bgColor);
    }
    
    // Freeze header row
    mappingSheet.setFrozenRows(1);
  }
  
  Logger.log(`✅ Created Tag_Outcome_Mappings with ${mappings.length} mappings`);
  
  return mappings.length;
}

/**
 * Menu function to run analysis
 * Call this from your Code.gs menu
 */
function runTagCombinationAnalysis() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    'Tag Combination Analysis',
    'This will analyze all tag combinations and:\n\n' +
    '✅ CREATE/RESET Tag_Outcome_Mappings sheet\n' +
    '✅ Load high-confidence mappings (>80% certainty)\n' +
    '✅ Create Tag_Combination_Analysis (all patterns)\n' +
    '✅ Create Mapping_Recommendations (suggestions)\n\n' +
    '⚠️ WARNING: This will ERASE existing Tag_Outcome_Mappings!\n\n' +
    'Use "INITIAL SETUP" from Tag Management menu for proven mappings.\n' +
    'This tool is for comprehensive pattern analysis.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Analyzing tag combinations...', 'Please wait', -1);
  
  const results = analyzeTagCombinations();
  
  if (results.success && results.recommendations) {
    // Create/reset Tag_Outcome_Mappings based on analysis
    const mappingsCreated = createTagMappingsFromAnalysis(results.recommendations);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Analysis complete! Created ${mappingsCreated} mappings from ${results.totalTickets} tickets. Check the new sheets.`,
      '✅ Done',
      8
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('Analysis failed. Check the logs.', '❌ Error', 5);
  }
}
