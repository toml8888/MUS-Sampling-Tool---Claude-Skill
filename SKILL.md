---
name: mus-sampler
description: >
  Performs Monetary Unit Sampling (MUS) on a transaction population for statutory audit purposes.
  Use this skill whenever the user wants to sample a ledger or population for audit, mentions MUS,
  monetary unit sampling, audit sampling, selecting a sample from a population, or provides a
  transaction file with a materiality and risk factor. Also trigger when the user says things like
  "run sampling on this population", "select my audit sample", "do MUS on this ledger", or
  "generate a sample workpaper". The skill reads an Excel or CSV population, asks the user to
  confirm the reference and value columns, splits into positive and negative populations, runs
  full MUS logic (individually significant items first, then systematic MUS on the remainder),
  and outputs a formatted Excel workpaper with a cover tab, positive population tab, negative
  population tab, positive sample tab, and negative sample tab.
---

# MUS Sampler Skill

Produces a statutory audit MUS workpaper from a transaction population file.

## Scripts

All sampling logic lives in a private GitHub repo. Claude must clone this repo before running
any scripts. The scripts are not regenerated -- the exact code from the repo is always used.

**Repo:** https://github.com/toml8888/MUS-Sampling-Tool---Claude-Skill

Scripts location after clone:
```
MUS-Sampling-Tool---Claude-Skill/scripts/run_mus.py
MUS-Sampling-Tool---Claude-Skill/scripts/mus_engine.py
MUS-Sampling-Tool---Claude-Skill/scripts/excel_builder.py
```

## What the skill does

1. Clones the GitHub repo to get the verified sampling scripts
2. Reads an Excel (.xlsx / .xls / .xlsm) or CSV population file
3. Asks the user to confirm the reference and value columns
4. Splits the population into positive and negative sub-populations (zeros excluded)
5. Runs MUS on each:
   - Pass 1: all items >= PM are individually significant (IS), selected 100%
   - Pass 2: MUS on the residual -- n = (residual total x risk factor) / PM, rounded up
   - Systematic selection: interval = residual total / n, random start fixed (seed 42)
6. Outputs MUS_Sample.xlsx with five tabs: Cover, Positive Population, Negative Population, Positive Sample, Negative Sample

## Inputs required from the user

Before running, Claude must have all of:
- **Population file** -- Excel or CSV
- **Performance materiality** -- numeric value (e.g. 5000)
- **Risk factor** -- numeric value supplied by the user (e.g. 1.6)
- **Reference column name** -- the column containing transaction references
- **Value column name** -- the column containing monetary values

If the user does not state the reference or value column, Claude must ask before proceeding.
Do not guess silently -- always confirm with the user.

## Running the skill

### Step 1: Confirm inputs

Check the user's message for all required inputs. If any are missing:

```
To run MUS sampling I need the following:
1. Your population file (Excel or CSV) -- please upload it
2. Performance materiality (numeric)
3. Risk factor (numeric)
4. Which column is the transaction reference?
5. Which column is the monetary value?
```

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
Copy it to the scripts directory:

```bash
cp "/mnt/user-data/uploads/<filename>" /home/claude/MUS-Sampling-Tool---Claude-Skill/scripts/
```

### Step 5: Run the script

```bash
cd /home/claude/MUS-Sampling-Tool---Claude-Skill/scripts
python run_mus.py "<filename>" <pm> <risk_factor> "<ref_col>" "<value_col>"
```

Example:
```bash
python run_mus.py "population.xlsx" 5000 1.6 "Ref" "Net Amount"
```

### Step 6: Check for errors

If the script errors:
- File not found: check the upload path with `ls /mnt/user-data/uploads/`
- Column not found: re-read the headers printed by the script and confirm with user
- Parse errors on value column: the script will list skipped rows -- report them to the user
- Git clone fails: check network access or ask user to verify repo permissions

### Step 7: Copy output and present

```bash
cp /home/claude/MUS-Sampling-Tool---Claude-Skill/scripts/MUS_Sample.xlsx /mnt/user-data/outputs/MUS_Sample.xlsx
```

Then use `present_files` to deliver the file.

### Step 8: Report summary to user

After presenting the file, give a brief summary:

```
Positive population: X items, Y IS, Z MUS selected (N total)
Negative population: X items, Y IS, Z MUS selected (N total)
Output: MUS_Sample.xlsx
  - Cover: all workings and inputs documented with live Excel formulas
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
do not double-count it.

**Positive and negative populations run separately.** Zeros excluded from both.

**Running total in output:** formula-driven using SUMPRODUCT to exclude IS items
so the auditor can trace each interval against the cumulative residual total.

## Edge cases handled

- Sample size >= residual count: entire residual population selected
- No residual population: only IS items selected, MUS section skipped
- No negative transactions: Negative Population tab shows a message
- Column names with spaces or mixed case: handled directly, no reordering needed
- Currency symbols and commas in values: stripped automatically (£, $, €, commas)
