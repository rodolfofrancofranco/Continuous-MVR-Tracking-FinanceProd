# Tag Management System - Complete Guide

## Overview
The MVR Tracking system now includes an intelligent tag management system that helps classify tickets based on Freshdesk tags. This guide covers the complete end-to-end workflow optimized for efficiency and ease of use.

---

## 🎯 Key Principles

1. **ONE-TIME SETUP**: Initial mappings are loaded once, giving you 90% coverage
2. **MANUAL CONTROL**: You always decide what gets mapped - no automatic changes
3. **SMART SUGGESTIONS**: System suggests outcomes based on 602-ticket analysis
4. **AUTOMATIC DISCOVERY**: New tags are discovered after every pull, but not auto-mapped

---

## 📋 Complete Workflow

### Step 1: Initial Setup (Do Once)
**Menu: 🆕 INITIAL SETUP: Load Proven Mappings**

Loads 18 proven tag mappings based on 602-ticket analysis:

#### Billable Patterns (Clear = 90-95 priority)
- `same information` → Clear
- `Updated` → Clear  
- `uploaded` → Clear
- `approved` → Clear

#### Violations (Priority 100)
- `DL Expired` → Expired License
- `suspended` → Suspended
- `invalid` → Invalid License

#### Not Billable (Priority 60-75)
- `emailed` → Still Processing
- `pending` → Pending
- `Verifying` → Still Processing
- `Review identity` → Still Processing

#### Special Cases
- `Record not found` → Record Not Found
- `Withdrawn` → Withdrawn

**Coverage**: ~90% of your current 602 tickets

---

### Step 2: Pull Tickets (Automatic Discovery)
**What Happens Automatically:**
1. You run Data Pull (menu or refresh)
2. System scans all new tickets for tags
3. Discovered_Tags sheet is updated with:
   - Tag name
   - Frequency count
   - Current mapping (if exists)
   - Suggested outcome (smart AI)
   - First seen date

**Important**: Tags are discovered but NOT automatically mapped. You stay in control.

---

### Step 3: Review Unmapped Tags (When Notified)
**Menu: ⚠️ View Unmapped Tags**

System shows you tags that need your decision:
- Tag name
- How many tickets have this tag
- Smart suggestion based on patterns
- Your options: Accept, Modify, or Skip

#### Smart Suggestions Explained

The system uses **priority-based matching** from 602-ticket analysis:

**Priority 1 - Exact Matches** (Most reliable)
- `same information` → Clear
- `updated` → Clear
- `requested` → Still Processing
- `expired` → Expired License

**Priority 2 - Substring Matches** (Specific before general)
- Multi-word first: `dl expired`, `record not found`, `review identity`
- Single-word after: `suspend`, `pending`, `emailed`

**Priority 3 - Pattern Detection**
- State codes (CA, WI, PA) → Unknown (metadata, not outcomes)
- Typos (`requesed`) → Still Processing
- System tags (json, dmv, api) → Unknown (technical tags)

**Result**: You get intelligent suggestions that you can accept or modify.

---

### Step 4: Map Tags Manually
**Menu: ➕ Map Single Tag**

When you want to map a specific tag:
1. Enter the tag name (or pattern like `*pending*`)
2. System shows smart suggestion
3. Select final outcome from dropdown
4. Set priority (100=violations → 50=requested)
5. Optionally enable regex for complex patterns

**Priority Guidelines:**
- 100: Clear violations (expired, suspended)
- 90-95: Completions (clear, updated, approved)
- 60-75: Pending states (requested, verifying)
- 50: Generic requests

---

### Step 5: Test Your Mappings
**Menu: 🧪 Test Classification**

Before applying to all tickets:
1. Enter a tag name
2. System shows which outcome it will map to
3. Shows matching pattern and priority
4. Verify it matches your expectation

**Use Cases:**
- Testing regex patterns
- Checking priority order (specific before general)
- Verifying new mappings work correctly

---

### Step 6: Apply to History
**Menu: 🔄 Reapply All Rules**

Once you're happy with your mappings:
1. Click "Reapply All Rules"
2. System reclassifies ALL tickets in Raw_Tickets
3. Updates MVR Outcome column based on current mappings
4. Progress shown every 100 tickets
5. Summary displayed when complete

**Important**: This uses your current mappings. If you add/change mappings later, run this again to update all historical tickets.

---

## 🛠️ Advanced Tools

### Pattern Discovery Analysis
**Menu: 📊 Analyze Combinations**

Comprehensive analysis tool that creates mappings from your actual data:

**What It Does:**
- Scans ALL tag combinations in your tickets
- Identifies high-confidence patterns (billable='Yes' or 10+ occurrences)
- **CREATES/RESETS Tag_Outcome_Mappings** sheet with data-driven mappings
- Creates `Tag_Combination_Analysis` sheet with all patterns
- Creates `Mapping_Recommendations` sheet with suggestions

**⚠️ WARNING: This ERASES existing Tag_Outcome_Mappings!**

**Use This When:**
- Starting fresh with comprehensive data-driven mappings
- You want the system to analyze YOUR specific data patterns
- You have significant ticket history (100+ tickets)
- Planning a complete mapping overhaul

**Output - Tag_Outcome_Mappings Created:**
- **Priority 100**: Violations (expired, suspended, invalid)
- **Priority 95**: Multi-tag combinations (most specific)
- **Priority 90**: Single-tag clear completions
- **Priority 80**: Other billable patterns
- **Priority 50**: Non-billable states (pending, requested)

**Confidence Threshold:**
- Includes: Billable='Yes' OR (count >= 10 AND no "NEEDS REVIEW" warnings)
- Only high-confidence patterns become mappings
- Low-confidence patterns appear in recommendations only

**vs. INITIAL SETUP:**
- **INITIAL SETUP**: Loads 18 proven mappings, preserves existing
- **Analyze Combinations**: Creates fresh mappings from YOUR data, erases existing

Choose based on your need:
- Fresh start with your data → Use Analyze Combinations
- Quick start with proven patterns → Use INITIAL SETUP

---

### View Current Mappings
**Menu: 📋 View Current Mappings**

Opens the `Tag_Outcome_Mappings` sheet showing:
- All your mapping rules
- Priority order (100 → 50)
- Pattern types (exact, wildcard, regex)
- Which outcome each maps to

**Direct Editing**: You can edit this sheet directly if you prefer, just be careful with:
- Priority numbers (higher = matched first)
- Pattern syntax (use `*` for wildcards, regex if enabled)
- Outcome types (must match valid OUTCOME_TYPES)

---

### Reset to Defaults
**Menu: ♻️ Reset to Defaults**

Complete reset:
- Clears all mappings
- Restarts with basic defaults only
- Does NOT delete discovered tags
- Use if you want to start mapping from scratch

---

## 📊 Monitoring & Quality

### Data Quality Alerts

The system automatically monitors:
- **>5% Unknown Outcomes**: Alerts you if too many tickets lack classifications
- **New Tags Discovered**: Toast notification after each pull
- **Duplicate Mappings**: Warns before overwriting existing mappings

### Sheets You'll Use

1. **Discovered_Tags**: All tags with stats and suggestions
2. **Tag_Outcome_Mappings**: Your classification rules (priority-ordered)
3. **Raw_Tickets**: Source data with MVR Outcome column
4. **Tag_Combination_Analysis**: Advanced pattern analysis (optional)
5. **Mapping_Recommendations**: AI-generated suggestions (optional)

---

## 💡 Best Practices

### 1. Start with Initial Setup
Run "INITIAL SETUP" first - gives you 90% coverage immediately.

### 2. Review Unmapped After Each Pull
Check "View Unmapped Tags" weekly to stay on top of new patterns.

### 3. Use Smart Suggestions
The suggestions are based on real data analysis - they're usually correct. Modify only when needed.

### 4. Group Similar Tags
Use wildcards like `*pending*` to catch variations (`pending`, `is pending`, `still pending`).

### 5. Test Before Reapplying
Use "Test Classification" to verify complex patterns before running "Reapply All Rules".

### 6. Set Priorities Correctly
- Specific patterns = higher priority (100)
- General patterns = lower priority (50)
- This ensures "dl expired" matches before just "expired"

### 7. Document Your Decisions
Add notes in the Tag_Outcome_Mappings sheet explaining why certain tags map certain ways.

---

## 🚨 Common Issues & Solutions

### Issue: Tag not matching as expected
**Solution**: 
1. Check priority order - specific patterns must be higher priority
2. Use "Test Classification" to see what's matching
3. Verify pattern syntax (wildcards, regex)

### Issue: Too many Unknown outcomes
**Solution**:
1. Run "View Unmapped Tags"
2. Review smart suggestions
3. Map high-frequency tags first (biggest impact)
4. Consider using wildcards for variations

### Issue: Duplicate mappings
**Solution**:
System warns you before overwriting. Choose:
- Overwrite: Replace with new mapping
- Cancel: Keep existing mapping

### Issue: Mapping changed, history not updated
**Solution**:
Run "Reapply All Rules" - this reclassifies all tickets with current mappings.

---

## 📈 Success Metrics

Track your progress:
- **Coverage**: % of tickets with known outcomes (goal: >95%)
- **Efficiency**: # of unmapped tags (goal: <10 at any time)  
- **Accuracy**: Manual review of classifications (spot-check weekly)

---

## 🔄 Maintenance Schedule

### Daily (Automatic)
- Tag discovery after each pull
- Data quality alerts if >5% unknown

### Weekly (Manual - 5 minutes)
- Check "View Unmapped Tags"
- Map 3-5 high-frequency tags
- Verify no data quality issues

### Monthly (Manual - 15 minutes)
- Run "Analyze Combinations" for patterns
- Review mapping priorities
- Reapply rules if needed
- Spot-check classification accuracy

---

## 🎓 Quick Reference

| Task | Menu Item | Frequency | Erases Existing? |
|------|-----------|-----------|------------------|
| First-time setup (proven patterns) | 🆕 INITIAL SETUP | Once | No - preserves |
| Comprehensive data-driven reset | 📊 Analyze Combinations | Rarely | ⚠️ YES - erases all |
| See what needs mapping | ⚠️ View Unmapped Tags | Weekly | No |
| Map one tag | ➕ Map Single Tag | As needed | No |
| Test a pattern | 🧪 Test Classification | Before complex mappings | No |
| Apply changes to all tickets | 🔄 Reapply All Rules | After mapping changes | No |
| Start over manually | ♻️ Reset to Defaults | Rarely | Yes - clears to basics |

---

## 📞 Support

If you encounter issues:
1. Check "Common Issues & Solutions" above
2. Review the Tag_Outcome_Mappings sheet for conflicts
3. Test individual tags with "Test Classification"
4. Use "View Current Mappings" to verify your rules

---

## 🎯 Remember

✅ **Setup once** - Initial 18 mappings give 90% coverage  
✅ **Control always** - You decide every mapping, no auto-updates  
✅ **Smart suggestions** - Based on 602-ticket real data analysis  
✅ **Discover automatically** - New tags found after every pull  
✅ **Test before applying** - Verify patterns work correctly first  

**You're in control. The system is here to help, not to decide for you.**
