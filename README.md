# ADP ID Mapper

This tool maps 7-digit employee IDs to the correct 6-digit ADP File Number values from a master validation workbook.

By default, output is written to:
- `File #` (or `File Number`) column when that column exists in the target sheet (best for Total Hours sheets)
- otherwise `Employee ID` column

## How it matches

1. Takes the last 4 digits from source `Employee ID`.
2. Finds master rows where `Tax ID (SIN)` has the same last 4 digits.
3. If multiple rows share that last 4, it disambiguates by normalized:
   - `Employee First Name`
   - `Employee Last Name`
4. If no SIN last-4 candidate exists, it falls back to unique normalized first+last name match.
5. If still ambiguous/unmatched, it leaves the row unchanged and logs the row in an `Exceptions` worksheet in the output workbook.

## Run

```bash
python3 adp_id_mapper.py \
  --agency "/Users/sjlidder/Downloads/Agency Hours - People & Culture - March 23rd-March 29th 2026.xlsx" \
  --master "/Users/sjlidder/Downloads/Canada Validation Report (15).xlsx"
```

Optional output path:

```bash
python3 adp_id_mapper.py \
  --agency "/path/to/agency.xlsx" \
  --master "/path/to/master.xlsx" \
  --output "/path/to/output.xlsx"
```

## Total Hours workflow

If your workbook has a specific Total Hours sheet, target it directly:

```bash
python3 adp_id_mapper.py \
  --agency "/Users/sjlidder/Downloads/Total Hours Test.xlsx" \
  --master "/Users/sjlidder/Downloads/Canada Validation Report  Test.xlsx" \
  --agency-sheet-name "Sheet1"
```

If you want to force a specific output column, use `--target-column-header`:

```bash
python3 adp_id_mapper.py \
  --agency "/path/to/total_hours.xlsx" \
  --master "/path/to/master.xlsx" \
  --agency-sheet-name "Total Hours" \
  --target-column-header "File #"
```

If `--output` is omitted, output is written beside the Agency file with `_adp_mapped` suffix.
