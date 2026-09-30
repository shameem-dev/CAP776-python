# CAP776 – Minor Project #1: My Data, My Story

**Student:** Shameem Ali T
**Registration No.:** 12600934
**Section:** 496
**Course:** CAP776 – Programming in Python

---

## 1. What this project does

This Python program reads a daily activity tracker from an Excel file, checks which days are valid, and calculates a set of personal activity indices and relationships from the valid days.

It uses only core Python plus `openpyxl` (to read the Excel file). **NumPy and Pandas are not used.**

The program:

1. Reads the `Daily Log` sheet of the Excel file into a list of dictionaries (one per day).
2. Checks each day for completeness and for the recording period.
3. Counts valid and missing days against the 40 expected days.
4. Converts the text answers (Feeling, Satisfaction, Energy) into numbers.
5. Calculates average minutes per day for each activity.
6. Calculates the indices TPI, AAI, PhAI, SRI, ABI, TUI, EI, DCI and PAI.
7. Calculates three Pearson correlations (Coding–Energy, Sleep–Energy, Study–Satisfaction).
8. Prints a final summary.

---

## 2. Files

| File | Purpose |
|---|---|
| `12600934.py` | The Python program |
| `12600934.xlsx` | The input Excel tracker (must contain a sheet named `Daily Log`) |
| `README.md` | This file |

Keep the `.py` file and the `.xlsx` file in the **same folder**.

---

## 3. Requirements

- Python 3.8 or newer
- The `openpyxl` library

Install `openpyxl` (Ubuntu):

```
sudo apt install python3-openpyxl
```

or, with pip:

```
pip install openpyxl
```

---

## 4. How to run

1. Open a terminal in the folder that contains the two files.
2. Run:

```
python3 12600934.py
```

3. Read the results printed on the screen. The last part, `--- FINAL SUMMARY ---`, lists the valid days and all the index values.

**If you rename the Excel file**, change the filename in this line of the program:

```python
all_days = load_data("12600934.xlsx")
```

---

## 5. Input file format

The program reads the sheet **`Daily Log`**, starting from **row 6**, with the columns in this order:

| Column | Field | Type |
|---|---|---|
| A | Date | Excel date |
| B | Sleep (min) | number |
| C | Fitness (min) | number |
| D | Study (min) | number |
| E | Coding (min) | number |
| F | Class (min) | number |
| G | Classes Attended | number (count) |
| H | Other Activities (min) | number |
| I | Total Tracked (min) | number (formula in sheet) |
| J | Free / Unaccounted (min) | number (formula in sheet) |
| K | Day's Feeling | text |
| L | Satisfaction Level | text |
| M | Energy Level | text |

Text values must be spelled exactly as listed in section 7, otherwise the program stops with a `KeyError`.

Formula cells (columns I and J) are read with `data_only=True`, so the workbook must have been **saved in a spreadsheet program at least once** so the calculated values are stored.

---

## 6. Which days count as valid

A day is valid only when all of these are true:

- It has a date.
- The date is between **13 August 2026** and **21 September 2026** (inclusive).
- Sleep, Fitness, Study, Coding, Class, Other, Total Tracked, Free time, Feeling, Satisfaction and Energy are all filled in.

Days that fail any check are skipped in all averages, indices and correlations.

**Expected days = 40** (13 Aug to 21 Sep 2026).

---

## 7. Scoring scales

Text answers are converted to numbers so they can be averaged.

**Day's Feeling**

| Answer | Score |
|---|---|
| Stressed | 1 |
| Low | 2 |
| Neutral | 3 |
| Good | 4 |
| Excellent | 5 |

**Satisfaction Level**

| Answer | Score |
|---|---|
| Very Unsatisfied | 1 |
| Unsatisfied | 2 |
| Neutral | 3 |
| Satisfied | 4 |
| Very Satisfied | 5 |

**Energy Level**

| Answer | Score |
|---|---|
| Very Low | 1 |
| Low | 2 |
| Medium | 3 |
| High | 4 |
| Very High | 5 |

The scales are stored in the dictionaries `feeling_scale`, `satisfaction_scale` and `energy_scale` in the code. If a different scale is required, change the numbers there and run the program again.

---

## 8. Indices calculated

All averages are taken over **valid days only**.

| Index | Meaning | Formula |
|---|---|---|
| TPI | Tech Productivity | average Coding time |
| AAI | Academic Activity | average (Study + Class) time |
| PhAI | Physical Activity | average Fitness time |
| SRI | Sleep & Recovery | average Sleep time |
| ABI | Activity Balance | average Free / Unaccounted time |
| TUI | Time Utilization | average Total Tracked time |
| EI | Experience Index | average of (Feeling + Satisfaction + Energy) ÷ 3 |
| DCI | Data Continuity | (valid days ÷ expected days) × 100 |
| PAI | Personal Activity Index | 0.15·TPI + 0.20·AAI + 0.15·PhAI + 0.20·SRI + 0.15·TUI + 0.10·EI + 0.05·DCI |

The PAI weights add up to 1.00.

---

## 9. Relationships analysed

The program calculates the Pearson correlation coefficient (without NumPy) for:

1. Coding ↔ Energy
2. Sleep ↔ Energy
3. Study ↔ Satisfaction

How to read the value:

- close to **+1**: the two values tend to rise together
- close to **0**: almost no linear relationship
- close to **-1**: when one rises, the other tends to fall

A correlation shows association only. It does not prove that one thing causes the other.

---

## 10. Program structure

| Function | Job |
|---|---|
| `load_data(filename)` | Opens the workbook and returns a list of day dictionaries |
| `check_valid(day)` | Returns `True` if a day is complete and inside the recording period |
| `count_days(all_days, expected_days)` | Counts valid and missing days |
| `calculate_averages(all_days)` | Averages of each activity and of the experience score |
| `calculate_indices(averages, valid_days, expected_days)` | Builds all nine indices |
| `calculate_correlation(x_values, y_values)` | Pearson correlation of two lists |
| `get_pairs(all_days, field1, field2, scale2)` | Collects matching value pairs from valid days |

Data structures used: **lists** (all days, value lists), **dictionaries** (one per day, the scoring scales, averages and indices).

---

## 11. Output

The program prints, in order:

1. Number of rows read and the first row
2. `Valid`, `Missing`, `Expected` day counts
3. The averages and the indices as dictionaries
4. The three correlation values
5. A `FINAL SUMMARY` with valid days and TPI, AAI, PhAI, SRI, ABI, TUI, EI, DCI, PAI

---

## 12. Notes and limitations

- The printed **Missing** count includes empty rows at the bottom of the sheet, so it can be larger than the real number of skipped days. Use `40 − Valid` for the number of missing days in the recording period.
- Only `FileNotFoundError` is handled when opening the file. A wrong sheet name or a misspelled text answer stops the program with an error message.
- If there are no valid days, the averages cannot be calculated.
- The results describe only the days recorded in the input file.

---

## 13. How to reproduce the results

1. Place `12600934.py` and `12600934.xlsx` in one folder.
2. Install `openpyxl`.
3. Run `python3 12600934.py`.
4. Compare the printed values with the report.
