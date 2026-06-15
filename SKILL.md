---
name: mus-sampler
description: >
  Performs Monetary Unit Sampling (MUS) on a transaction population for statutory audit purposes.
  Use this skill whenever the user wants to sample a ledger or population for audit, mentions MUS,
  monetary unit sampling, audit sampling, selecting a sample from a population, or provides a
  transaction file with a materiality and risk factor. Also trigger when the user says things like
  "run sampling on this population", "select my audit sample", "do MUS on this ledger", or
  "generate a sample workpaper". The skill reads an Excel or CSV population, asks the user to
  confirm the population number and value columns, splits into positive and negative populations,
  runs full MUS logic (individually significant items first, then systematic MUS on the remainder),
  and outputs a formatted Excel workpaper with six tabs: Summary, Population, Sampling form
  positive, Sampling form negative, Positive samples, Negative samples.
---

# MUS Sampler Skill

Produces a statutory audit MUS workpaper from a transaction population file.

## Scripts

All sampling logic lives in a private GitHub repo. Claude must clone this repo before running
any scripts. The scripts are not regenerated -- the exact code from the repo is always used.

**Repo:** https://github.com/toml8888/MUS-Sampling-Tool---Claude-Skill

Scripts location after clone:
```
MUS-Sampling-Tool---Claude-Skill/Scripts/run_mus.py
MUS-Sampling-Tool---Claude-Skill/Scripts/mus_engine.py
MUS-Sampling-Tool---Claude-Skill/Scripts/excel_builder.py
MUS-Sampling-Tool---Claude-Skill/Scripts/image1.png
MUS-Sampling-Tool---Claude-Skill/Scripts/image2.png
```

The two image files are bundled alongside the scripts and are used by excel_builder.py for
the sampling form logo. They do not need to be supplied by the user.

## What the skill does

1. Clones the GitHub repo to get the verified sampling scripts and images
2. Reads an Excel (.xlsx / .xlsm) or CSV population file
3. Asks the user to confirm the population number and value columns, and the decimal system
4. Splits the population into positive and negative sub-populations (zeros and unparseable rows excluded)
5. Runs MUS on each:
   - Pass 1: all items >= PM are individually significant (IS), selected 100%
   - Pass 2: MUS on the residual -- n = CEILING(|residual total| x risk factor / PM, 1)
   - Systematic selection: interval = residual total / n, random start fixed (seed 42)
6. Outputs MUS_Sample.xlsx with six tabs (see Workpaper structure below)

## Inputs required from the user

Before running, Claude must have all of:
- **Population file** -- Excel (.xlsx / .xlsm) or CSV
- **Performance materiality** -- numeric value (e.g. 5000)
- **Risk factor** -- numeric value supplied by the user (e.g. 1.6)
- **Population number column name** -- the column that numbers every item in the population (used as the identifier column in the workpaper)
- **Value column name** -- the column containing monetary values
- **Decimal system** -- Normal (1,234.56) or EUR (1.234,56)

If the user does not state the population number column, value column, or decimal system, Claude
must ask before proceeding. Do not guess silently -- always confirm with the user.

### Why a population number column, not a reference

Every item in the population must carry a unique sequential number. The engine keys selections by
this number, so it must be unique per row. Transaction references (invoice numbers, journal IDs)
are not reliable for this because they repeat (split journals, recurring invoices, blanks), which
would corrupt the IS and MUS flags. The user numbers the population so every row is distinct.

## Workpaper structure (6 tabs)

### Tab 1: Summary
- PM (hardcoded input from user, blue text)
- Risk factor (hardcoded input from user, blue text)
- Total positive: SUMIF on the Classification column of Population tab
- Total negative: SUMIF on the Classification column of Population tab
- Total population: sum of the above two
- Total samples: sum of E37 from both sampling form tabs

The Classification column letter is determined dynamically at build time based on how many
source columns the population has. Summary formulas always reference the correct column.

### Tab 2: Population
- **Column A**: identifier (the population number column the user identified)
- **Column B**: value (the value column the user identified, parsed to numeric)
- **Columns C onwards**: remaining source columns (up to 15; truncated after that)
- **Above PM column**: `=IF(ABS(B{row})>Summary!$C$2,"Yes","")` -- dynamic column position
- **Classification column**: `=IF(B{row}>0,"Positive",IF(B{row}<0,"Negative","Zero"))` -- dynamic column position

The population number column appears only in column A. It is never duplicated in the other
columns. Above PM and Classification are entirely formula-driven -- no manual entries.
The builder returns the actual column letters for Above PM and Classification so all downstream
formulas (Summary, sampling forms) always reference the correct columns.

### Tab 3: Sampling form positive / Tab 4: Sampling form negative
Each tab contains:
- Logo image (top-left, rows 1-4)
- Blue accent line (row 5) and blue title bar (row 8)
- Medium border around the full form body (rows 8-39, cols B-J)
- Key indicator showing blue = cells requiring input (row 10)
- Random seed: 42 (row 12, for reproducibility)
- Guidance notes (rows 14-15)
- Risk factor lookup table (rows 19-25) with five risk levels across columns E-J
- Sampling risk factor linked to Summary: `=Summary!C7` (E27, blue input cell)
- Performance materiality linked to Summary: `=Summary!C2` (E29, blue input cell)
- Total population value: SUMIF on Classification column (E31)
- Value of items above PM: SUMIFS on Above PM and Classification columns (E32)
- Number of items above PM: COUNTIFS on Above PM and Classification columns (E33)
- Value of residual population: SUMIFS excluding Above PM items (E34)
- Calculated sample size: `=CEILING(ABS(E34)*E27/E29,1)` (E37)

All formula cells in E31-E34 reference the dynamically determined Above PM and Classification
column letters, so they are always correct regardless of how many source columns the population has.
No manual entries anywhere on the sampling form.

### Tab 5: Positive samples / Tab 6: Negative samples
- Same column layout as Population tab (A=identifier, B=value, then other source cols)
- Final column: Selection Status (IS or MUS)
- IS rows highlighted in salmon, MUS rows in green
- TOTAL SELECTED row at the bottom with a SUM formula on the value column

## Running the skill

### Step 1: Confirm inputs

Check the user's message for all required inputs. If any are missing:

```
To run MUS sampling I need the following:
1. Your population file (Excel or CSV) -- please upload it
2. Performance materiality (numeric)
3. Risk factor (numeric)
4. Population number column (please number every item in the population)
5. Which column is the monetary value?
6. Which decimal system are you using: Normal (1,234.56) or EUR (1.234,56)?
```

If the population is not yet numbered, ask the user to add a column numbering every item 1 to N
and re-upload. Selections cannot run safely on a non-unique key.

### Step 2: Install dependencies

```bash
pip install openpyxl pandas --break-system-packages -q
```

### Step 3: Clone the repo

```bash
cd /home/claude
git clone https://github.com/toml8888/MUS-Sampling-Tool---Claude-Skill.git
```

If the repo has already been cloned in this session, skip this step.

### Step 4: Save the uploaded file

The user's file will be at `/mnt/user-data/uploads/<filename>`.
Copy it to a clean working directory (keep the cloned repo pristine):

```bash
mkdir -p /home/claude/mus-run
cp "/mnt/user-data/uploads/<filename>" /home/claude/mus-run/
```

### Step 5: Run the script

```bash
cd /home/claude/mus-run
python /home/claude/MUS-Sampling-Tool---Claude-Skill/Scripts/run_mus.py \
  "<filename>" <pm> <risk_factor> "<popnum_col>" "<value_col>" <decimal_system>
```

`<decimal_system>` is `normal` or `eur`.

Example:
```bash
python /home/claude/MUS-Sampling-Tool---Claude-Skill/Scripts/run_mus.py \
  "population.xlsx" 5000 1.6 "Population Number" "Net Amount" normal
```

### Step 6: Check for errors

If the script errors:
- File not found: check the upload path with `ls /mnt/user-data/uploads/`
- Column not found: re-read the headers printed by the script and confirm with user
- Parse errors on value column: the script lists skipped rows -- report the count to the user
- Duplicate population numbers: the script will flag these -- ask the user to fix the numbering and re-upload
- Git clone fails: check network access or ask user to verify repo permissions

### Step 7: Copy output and present

```bash
cp /home/claude/mus-run/MUS_Sample.xlsx /mnt/user-data/outputs/MUS_Sample.xlsx
```

Then use `present_files` to deliver the file.

### Step 8: Report summary to user

After presenting the file, give a brief summary:

```
Positive population: X items, Y IS, Z MUS selected (N total)
Negative population: X items, Y IS, Z MUS selected (N total)
Output: MUS_Sample.xlsx
  - Summary: PM, risk factor, population totals, total sample count (all formula-linked)
  - Population: all rows with Above PM and Classification columns (formula-driven)
  - Sampling form positive / negative: risk factor table, seed, and sample size calcs
  - Positive samples / Negative samples: selected items with IS/MUS status
```

## Sampling methodology

**Individually significant (IS):** any item where |value| >= performance materiality.
Selected 100%. Not included in the MUS walk.

**MUS sample size:** n = CEILING(|residual_total| x risk_factor / PM, 1) -- always rounded up

**Sampling interval:** residual_total / n

**Random start:** uniform random between 0 and interval, fixed seed 42 for reproducibility.

**Selection walk:** cumulative total built across residual items. Each hit point
(start, start + interval, start + 2*interval ...) selects the transaction whose
cumulative range the hit point falls within. Duplicate hits on a single transaction
do not double-count it. Selections are keyed by the unique population number.

**Positive and negative populations run separately.** Zeros excluded from both and still
written to the Population tab where the Classification formula returns "Zero".

## Value parsing

- Decimal system is set by the user: Normal treats `.` as the decimal separator, EUR treats `,` as the decimal separator
- Currency symbols stripped: £, $, €
- Bracketed negatives `(1,234.56)` parsed as negative values
- Blank or unparseable cells are not silently dropped: they are reported to the console and excluded from the population
- The engine and the excel builder use one shared parse function so values always match

## Edge cases handled

- Sample size >= residual count: entire residual population selected
- No residual population: only IS items selected, MUS section has no effect
- No negative transactions: Negative samples tab shows a message
- Column names with spaces or mixed case: handled directly
- Blank, zero, and unparseable rows: excluded from sampling; zero rows still appear in the Population tab with Classification = "Zero"
