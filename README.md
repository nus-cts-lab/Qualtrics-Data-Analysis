# Qualtrics Data Analysis

Analysis scripts for processing questionnaire data from Qualtrics surveys.

## Files

- `Analysis/wsap_analysis.py` - Analyzes WSAP (Word-Sentence Association Paradigm) data
- `Analysis/questionnaire_analysis.py` - Analyzes QIDS, GAD, and MASQ questionnaires

## How to Run

### 1. Update Input File Name

In each Python file, update the `file_name` variable (around line 10-11):

```python
# For questionnaire_analysis.py
file_name = "1_labels_excel.xlsx"  # Change to your file name

# For wsap_analysis.py
file_name = "1_values_excel.xlsx"  # Change to your file name
```

### 2. Run the Script

```bash
cd Analysis
python3 questionnaire_analysis.py
# or
python3 wsap_analysis.py
```

Output will be saved to respective results directories.

## Questionnaire Analysis

### Input
- **File:** Excel file with questionnaire responses (e.g., `1_labels_excel.xlsx`)
- **Required columns:**
  - QIDS: Q2-Q16 (15 items)
  - GAD: Q1_1 to Q1_7 (7 items)
  - MASQ: Q1_1.1 to Q1_26 (26 items)

### Output
- **File:** `questionnaire_analysis_results/questionnaire_analysis_results.xlsx`
- **6 sheets:**
  1. QIDS Results - Individual participant scores
  2. QIDS Summary - Descriptive statistics
  3. GAD Results - Individual participant scores
  4. GAD Summary - Descriptive statistics
  5. MASQ Results - Individual participant scores (GD, AA, AD subscales)
  6. MASQ Summary - Descriptive statistics for all subscales

### Scoring Methods

**QIDS (Quick Inventory of Depressive Symptomatology):**
- Sum of Q2-Q16 (15 items)
- Higher score = more depressive symptoms

**GAD (Generalized Anxiety Disorder):**
- Sum of Q1_1 to Q1_7 (7 items)
- Higher score = more anxiety symptoms

**MASQ (Mood and Anxiety Symptom Questionnaire):**
- **Reverse scoring:** Items 1, 9, 15, 19, 23, 25 are scored as `6 - original_score`
- **Three subscales:**
  - **GD (General Distress):** Items 2, 3, 7, 12, 13, 17, 20, 21
  - **AA (Anxious Arousal):** Items 4, 6, 8, 10, 14, 16, 18, 22, 24, 26
  - **AD (Anhedonic Depression):** Items 1, 5, 9, 11, 15, 19, 23, 25
- Higher scores = greater symptoms

### Important Notes

- **Missing data handling:** Statistics include all participants with at least one valid response
- **Average completion rate:** Calculated as `(total valid items across all participants) / (participants × total items)`
- Data quality metrics (valid items, missing items, completion rate) are included for each participant
- Empty cells in Excel are treated as missing data (NaN)

## WSAP Analysis

### Input
- **File:** Excel file with WSAP task data (e.g., `1_values_excel.xlsx`)
- **Required columns:**
  - Original WSAP: `__js_responses`, `__js_reaction_times`, `__js_scenario_types`, `__js_word_types`
  - New WSAP: `__js_reaction_time`, `__js_valence`, `__js_stimulus`, `__js_response`

### Output
- **File:** `wsap_analysis_results/wsap_complete_analysis.xlsx`
- **4 sheets:**
  1. Original WSAP Results - Individual participant scores
  2. Original WSAP Summary - Descriptive statistics
  3. New WSAP Results - Individual participant scores
  4. New WSAP Summary - Descriptive statistics
- **Additional files:**
  - `original_wsap_ddm_data.csv` - DDM-ready trial data (Original WSAP)
  - `new_wsap_ddm_data.csv` - DDM-ready trial data (New WSAP)
  - `wsap_data_quality_report.csv` - Data quality report

### Scoring Methods

**Response Selection Score:**
- Calculated as: (Proportion negative endorsed) - (Proportion benign endorsed)
- Range: -1 to 1
- Higher scores indicate greater negative bias

**RT Bias Index:**
- Calculated as: (Mean RT endorse negative) - (Mean RT reject negative)
- Positive values indicate slower endorsement of negative items
- Negative values indicate faster endorsement of negative items

### Important Notes

- File expects header row with 2 metadata rows to skip (specified as `skiprows=[1, 2]`)
- Comma-separated values in cells are automatically parsed
- Participants with no valid data are marked as "No data" in results
- Statistics include all participants with at least one valid response
