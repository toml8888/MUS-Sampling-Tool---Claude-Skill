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
  and outputs a formatted Excel workpaper with a cover tab, population reconciliation tab,
  exclusions tab, positive and negative population tabs, and positive and negative sample tabs.
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
```

## What the skill does

1. Clones the GitHub repo to get the verified sampling scripts
2. Reads an Excel (.xlsx / .xlsm) or CSV population file
3. Asks the user to confirm the population number and value columns, and the decimal system
4. Splits the population into positive and negative sub-populations (zeros and unparseable rows excluded and evidenced)
5. Runs MUS on each:
   - Pass 1: all items >= PM are individually significant (IS), selected 100%
   - Pass 2: MUS on the residual -- n = (residual total x risk factor) / PM, rounded up
   - Systematic selection: interval = residual total / n, random start fixed (seed 42)
6. Outputs MUS_Sample.xlsx with seven tabs: Cover, Population Reconciliation, Exclusions, Positive Population, Negative Population, Positive Sample, Negative Sample

## Inputs required from the user

Before running, Claude must have all of:
- **Population file** -- Excel (.xlsx / .xlsm) or CSV
- **Performance materiality** -- numeric value (e.g. 5000)
- **Risk factor** -- numeric value supplied by the user (e.g. 1.6)
- **Population number column name** -- the column that numbers every item in the population
- **Value column name** -- the column containing monetary values
- **Decimal system** -- Normal (1,234.56) or EUR (1.234,56)

If the user does not state the population number column, value column, or decimal system, Claude
must ask before proceeding. Do not guess silently -- always confirm with the user.

### Why a population number column, not a reference

Every item in the population must carry a unique sequential number. The engine keys selections by
this number, so it must be unique per row. Transaction references (invoice numbers, journal IDs)
are not reliable for this because they repeat (split journals, recurring invoices, blanks), which
would corrupt the IS and MUS flags. The user numbers the population so every row is distinct.

## Running the skill

### Step 1: Confirm inputs

Check the user's message for all required inputs. If any are missing:

```
To run MUS sampling I need the following:
1. Your population file (Excel or CSV) -- please upload it
2. Performance materiality (numeric)
3. Risk factor (numeric)
4. Population number column (please number every item in the population):
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
- Parse errors on value column: the script lists skipped rows and writes them to the Exclusions tab -- report the count to the user
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
Population reconciliation: source rows = positives + negatives + zeros + skipped (must tie)
Positive population: X items, Y IS, Z MUS selected (N total)
Negative population: X items, Y IS, Z MUS selected (N total)
Excluded: P zero-value rows, Q unparseable rows (see Exclusions tab)
Output: MUS_Sample.xlsx
  - Cover: all workings and inputs documented with live Excel formulas
  - Population Reconciliation: row counts and value totals reconciling to the source file
  - Exclusions: every zero-value and unparseable row with the reason
  - Positive Population / Negative Population: full population with IS/MUS flags and running totals
  - Positive Sample / Negative Sample: selected items only
```

## Sampling methodology

**Individually significant (IS):** any item where |value| >= performance materiality.
Selected 100%. Not included in the MUS walk.

**MUS sample size:** n = residual_total x risk_factor / PM (always rounded up)

**Sampling interval:** residual_total / n

**Random start:** uniform random between 0 and interval, fixed seed 42 for reproducibility.

**Selection walk:** cumulative total built across residual items. Each hit point
(start, start + interval, start + 2*interval ...) selects the transaction whose
cumulative range the hit point falls within. Duplicate hits on a single transaction
do not double-count it. Selections are keyed by the unique population number.

**Positive and negative populations run separately.** Zeros excluded from both and evidenced
on the Exclusions tab.

**Running total in output:** formula-driven using SUMPRODUCT to exclude IS items
so the auditor can trace each interval against the cumulative residual total.

## Value parsing

- Decimal system is set by the user: Normal treats `.` as the decimal separator, EUR treats `,` as the decimal separator
- Currency symbols stripped: £, $, €
- Bracketed negatives `(1,234.56)` parsed as negative values
- Blank or unparseable cells are not silently dropped: they route to the Exclusions tab with a reason
- The engine and the Excel builder use one shared parse function so the workpaper formulas always match the engine result

## Edge cases handled

- Sample size >= residual count: entire residual population selected, Cover shows "N/A: full residual selected" for interval and random start
- No residual population: only IS items selected, MUS section skipped
- No negative transactions: Negative Population tab shows a message
- Column names with spaces or mixed case: handled directly, no reordering needed
- Blank, zero, and unparseable rows: excluded from sampling and listed on the Exclusions tab so the population reconciles to source
