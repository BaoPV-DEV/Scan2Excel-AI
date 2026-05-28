# TAB 5 - Team-Based Formula Deployment (v2.2)

## Overview
Enhanced salary deployment system with support for team-specific formula configurations. This allows different salary calculation formulas based on employee team/group (e.g., different formulas for "Cắt" team vs "Hoàn thiện" team).

## Features Added (v2.2)

### 1. Dynamic Team Recognition
- Automatically extracts team name from file path
- Example: `sx/Cắt - 04.xlsx` → Team: "Cắt"
- Example: `sx/Hoàn thiện - 04.xlsx` → Team: "Hoàn thiện"

### 2. Template Configuration
Enhanced `salary_sources.json` with new `team_formula` config type:

```json
"columns": {
  "AD": {
    "type": "team_formula",
    "description": "Column description",
    "team_configs": {
      "Cắt|Cut": {
        "formula": "=760*BF{row}",
        "description": "Formula for cut team"
      },
      "Hoàn thiện|Finishing": {
        "formula": "=680*BF{row}",
        "description": "Formula for finishing team"
      },
      "default": {
        "formula": "=400*BF{row}",
        "description": "Default formula for other teams"
      }
    }
  }
}
```

### 3. Pattern Matching
- Supports multiple name patterns separated by `|`
- Case-insensitive matching
- Example: `"Cắt|Cut"` matches "Cắt", "cut", "CUT", etc.
- Fallback to `default` config if no match found

## Supported Templates

### sx/to_may_template.xlsx (Tổ may - Sewing Teams)

#### Column AD (Sheet: "Bang TH nop")
Productivity/Salary calculation based on team:

**Tổ Cắt (Cut team):**
```
AD row 9:  =760*BF9
AD row 11: =760*BF11
... etc (every 2 rows)
```

**Tổ Hoàn thiện (Finishing team):**
```
AD row 9:  =680*BF9
AD row 11: =680*BF11
... etc
```

**Other teams (default):**
```
AD row 9:  =400*BF9
AD row 11: =400*BF11
... etc
```

#### Column I (Sheet: "Lương {mm}")
Overtime pay calculation based on team:

**Tổ Cắt (Cut team) - Rate: 358**
```
I9:  =F9*358*(((I8-N8)*8+N8*7))/(((I8-N8)*8+N8*7+J8))*$I$4
I11: =F11*358*(((I10-N10)*8+N10*7))/(((I10-N10)*8+N10*7+J10))*$I$4
... etc
```

**Tổ Hoàn thiện (Finishing team) - Rate: 400**
```
I9:  =F9*400*(((I8-N8)*8+N8*7))/(((I8-N8)*8+N8*7+J8))*$I$4
I11: =F11*400*(((I10-N10)*8+N10*7))/(((I10-N10)*8+N10*7+J10))*$I$4
... etc
```

**Other teams (default) - Rate: 680**
```
I9:  =F9*680*(((I8-N8)*8+N8*7))/(((I8-N8)*8+N8*7+J8))*$I$4
I11: =F11*680*(((I10-N10)*8+N10*7))/(((I10-N10)*8+N10*7+J10))*$I$4
... etc
```

## Files Modified

### 1. Template/salary_sources.json
- Added `team_formula` type support
- Configured sx/to_may_template.xlsx with:
  - Column AD (Bang TH nop): Team-based productivity formulas
  - Column I (Lương {mm}): Team-based overtime pay formulas
- Supports: Cắt, Hoàn thiện, and other default teams

### 2. logic/salary_deployer.py
**New functions:**
- `_extract_team_name(file_path)`: Extracts team name from file path
  - Input: `"sx/Cắt - 04.xlsx"`
  - Output: `"Cắt"`
  - Regex pattern: Removes month suffix (e.g., " - 04")

- `_match_team_config(team_name, team_configs)`: Matches team to correct formula config
  - Case-insensitive matching
  - Pattern support with `|` separator
  - Fallback to `default` config

**Modified functions:**
- `deploy_file()`: Added `team_name` parameter for team-specific formula deployment
- `deploy_salary_from_metadata()`: Extracts team name before calling `deploy_file()`

**Feature:**
- Team name automatically logged in deployment output for visibility

### 3. logic/salary_config_parser.py
**Enhanced function:**
- `parse_column_config()`: Now recognizes `"type": "team_formula"` config entries
- Returns new fields: `team_configs` dict, `description`

## Processing Workflow

```
User runs salary deployment (Tab 5)
    ↓
Read metadata.json for file list
    ↓
For each file (e.g., "sx/Cắt - 04.xlsx"):
    1. Extract team name → "Cắt"
    2. Load template config for sx/to_may_template.xlsx
    3. For each column in template:
       - If "team_formula" type:
         * Match team name to team_configs
         * Select appropriate formula based on team
         * Apply formula to all employee rows
       - If "source" type:
         * Link to external file (existing behavior)
    4. Save file with applied formulas
```

## Deployment Log Example

```
📄 [1/3] Xử lý: sx/Cắt - 04
   📋 Template: sx/to_may_template.xlsx
   📁 File: sx/Cắt - 04.xlsx

   📋 Sheet: Bang TH nop (25 nhân viên)
      ✏️ Cột AD: Công suất (phụ thuộc vào tổ) (Tổ: Cắt)
         ↳ Đã ghi 25 dòng cột AD
      ✏️ Cột AO: công thức liên kết file...

   📋 Sheet: Lương 04 (25 nhân viên)
      ✏️ Cột I: Tiền thêm giờ tính theo lương tương ứng (phụ thuộc vào tổ) (Tổ: Cắt)
         ↳ Đã ghi 25 dòng cột I
      ✏️ Cột AG: công thức liên kết file...
```

## Formula Patterns Explained

### Column AD Formula: `=760*BF{row}`
- Multiplies a constant (760 for Cắt, 680 for Hoàn thiện, 400 for others)
- By column BF value at the current row
- `{row}` placeholder replaced with actual row number (9, 11, 13, etc.)

### Column I Formula: Complex Overtime Calculation
```
=F{row}*RATE*(((I{row-1}-N{row-1})*8+N{row-1}*7))/(((I{row-1}-N{row-1})*8+N{row-1}*7+J{row-1}))*$I$4
```
Where:
- `F{row}`: Base salary (same row)
- `RATE`: Team-specific rate (358, 400, or 680)
- `I{row-1}`: Previous row value
- `N{row-1}`: Previous row value
- `J{row-1}`: Previous row value
- `$I$4`: Fixed cell reference (absolute)

Breakdown:
- Calculates overtime hours based on working hours vs scheduled hours
- Applies team-specific rate multiplier
- Adjusts by base salary and fixed allowance

## Extending to Other Teams

To add new team configurations:

1. Update `salary_sources.json`:
```json
"team_configs": {
  "Your Team Name|Alternative Name": {
    "formula": "=VALUE*BF{row}",
    "description": "Description for this team"
  }
}
```

2. Update file naming to match pattern:
   - File: `sx/Your Team Name - MM.xlsx`
   - Team extracted: `Your Team Name`

3. Pattern matching is case-insensitive and supports alternatives

## Backward Compatibility

- Existing `formula` type configs (fixed formulas) still work
- Existing `source` type configs (file linking) still work
- Mixed template configs supported (some columns team-based, others fixed)
- Files without team markers use `default` config

## Testing Checklist

- [x] JSON configuration parsing
- [x] Team name extraction from file paths
- [x] Team config matching (including pattern matching)
- [x] Formula placeholder substitution
- [x] Backward compatibility with existing configs
- [ ] End-to-end deployment test (awaits actual data)

## Troubleshooting

### Issue: Formula not applied
- Check file name contains team identifier (e.g., "Cắt", "Hoàn thiện")
- Verify team name matches pattern in `salary_sources.json`
- Check `default` config exists as fallback

### Issue: Wrong formula applied
- Verify team name extraction: should match between file name and config pattern
- Check pattern case-sensitivity (matching is case-insensitive)
- Review deployment log for actual team name used

### Issue: Column not found in sheet
- Verify sheet name matches (e.g., "Bang TH nop", "Lương {mm}")
- Check template column definition in `salary_sources.json`

## Future Enhancements

1. Support for more team-specific templates
2. Admin interface to manage team configurations
3. Team override mapping (e.g., map "Team A" → use "Cắt" config)
4. Rate configuration per team (currently hardcoded in formulas)

---

**Version:** 2.2 (May 25, 2026)
**Status:** Production Ready
**Tested:** Configuration parsing and team matching logic
**Pending:** End-to-end deployment with actual data
