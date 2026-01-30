# Qualtrics Data Analysis

This directory contains Python scripts for analyzing questionnaire and behavioral task data from Qualtrics surveys.

## Directory Structure

```
Analysis/
├── 1_labels_excel.xlsx          # Input data file (labeled responses)
├── 1_values_excel.xlsx          # Input data file (numeric values)
├── README.md                    # This file
├── AST/                         # Ambiguous Scenarios Task analysis
│   ├── ast_analysis.py
│   └── ast_analysis_results.xlsx
├── SST/                         # Scrambled Sentences Test analysis
│   ├── sst_analysis.py
│   └── sst_analysis_results.xlsx
├── Questionnaire/               # Questionnaire analysis (QIDS, GAD, MASQ)
│   ├── questionnaire_analysis.py
│   └── questionnaire_analysis_results.xlsx
└── WSAP/                        # Word Sentence Association Paradigm analysis
    ├── wsap_analysis.py
    ├── wsap_complete_analysis.xlsx
    ├── original_wsap_ddm_data.csv
    ├── new_wsap_ddm_data.csv
    └── wsap_data_quality_report.csv
```

## Requirements

- Python 3.x
- pandas
- numpy
- openpyxl

Install dependencies:
```bash
pip install pandas numpy openpyxl
```

## Usage

Each analysis script is located in its own subfolder. Navigate to the subfolder and run the script to generate results in the same location.

### 1. AST Analysis (Ambiguous Scenarios Task)

**Purpose:** Analyze interpretation bias using reverse-scored pleasantness ratings and prepare outcome descriptions for manual qualitative coding.

**Components:**
1. **Quantitative Analysis:** Reverse-score pleasantness ratings
2. **Qualitative Coding Preparation:** Prepare outcome descriptions for manual coding by independent raters

**Formula:**
```
Reverse Score = 10 - Original Score
Mean Reverse-Scored Rating = Average of all reverse-scored items
```

**Input:** `../1_values_excel.xlsx`
- Columns: `main_pleasantness_ratings`, `main_outcome_descriptions`, `main_scenario_ids`, `main_scenario_themes`

**How to run:**
```bash
cd AST
python3 ast_analysis.py
```

**Output:** `ast_analysis_results.xlsx` (3 sheets)
- **Sheet 1: Reverse-Scored Ratings**
  - Participant-level mean reverse-scored ratings
  - Summary statistics (mean, SD, min, max, median)
  - Count of valid ratings and descriptions

- **Sheet 2: Data Quality**
  - List of participants with/without valid data
  - Count of valid ratings and descriptions per participant
  - Data status: Complete, Partial, or No Data

- **Sheet 3: Coding Template**
  - Ready for manual coding by independent raters
  - Columns: Subject, Main_Outcome_Descriptions, Coder_1, Coder_2, Final, Coder_3
  - 16 outcome descriptions per participant
  - Categories to code: negative, neutral/unclear, positive

**Coding Instructions:**
- Two independent coders rate each description as: negative, neutral/unclear, or positive
- If there is a discrepancy, a third rater (Coder_3) evaluates
- Final consensus is recorded in the "Final" column
- After manual coding is complete, calculate index scores by summing responses in each category

---

### 2. SST Analysis (Scrambled Sentences Test)

**Purpose:** Calculate negativity scores for anxiety and depression stimuli.

**Formula:**
```
Negativity Score = (Negative_D + Negative_GA) / Total_Completed_Sentences
```

**Input:** `../1_values_excel.xlsx`
- Columns: `list_assignment`, `main_total_completed`, `main_sentence_interpretations`

**How to run:**
```bash
cd SST
python3 sst_analysis.py
```

**Output:** `sst_analysis_results.xlsx` (3 sheets)
- **SST Results:** Participant-level data with negativity scores
- **SST Summary:** Overall summary statistics
- **Summary by List:** Results broken down by list assignment

**Interpretation Types:**
- `negative_D` - Depression-related negative interpretation
- `negative_GA` - General Anxiety-related negative interpretation
- `positive` - Positive interpretation
- `mixed` - Mixed interpretation
- `unclear` - Unclear interpretation

---

### 3. Questionnaire Analysis (QIDS, GAD-7, MASQ)

**Purpose:** Analyze questionnaire responses for depression, anxiety, and mood symptoms.

**Questionnaires:**
- **QIDS (Q2-Q16):** Quick Inventory of Depressive Symptomatology
- **GAD-7 (Q1_1-Q1_7):** Generalized Anxiety Disorder scale
- **MASQ (Q1_1.1-Q1_26):** Mood and Anxiety Symptom Questionnaire
  - General Distress (GD)
  - Anxious Arousal (AA)
  - Anhedonic Depression (AD) - with reverse scoring

**Input:** `../1_labels_excel.xlsx`

**How to run:**
```bash
cd Questionnaire
python3 questionnaire_analysis.py
```

**Output:** `questionnaire_analysis_results.xlsx` (6 sheets)
- **QIDS Results:** Participant-level QIDS scores
- **QIDS Summary:** QIDS summary statistics
- **GAD Results:** Participant-level GAD-7 scores
- **GAD Summary:** GAD-7 summary statistics
- **MASQ Results:** Participant-level MASQ subscale scores
- **MASQ Summary:** MASQ summary statistics

---

### 4. WSAP Analysis (Word Sentence Association Paradigm)

**Purpose:** Analyze interpretation bias using two versions of WSAP.

**Metrics:**
- **Response Selection Score:** Prop(Negative Endorsed) - Prop(Benign Endorsed)
- **RT Bias Index:** RT(Endorse Negative) - RT(Reject Negative)

**Input:** `../1_values_excel.xlsx`

**How to run:**
```bash
cd WSAP
python3 wsap_analysis.py
```

**Output:**
- `wsap_complete_analysis.xlsx` (4 sheets)
  - **Original WSAP Results:** Original task participant-level data
  - **Original WSAP Summary:** Original task summary statistics
  - **New WSAP Results:** New task participant-level data
  - **New WSAP Summary:** New task summary statistics
- `original_wsap_ddm_data.csv` - Original WSAP data formatted for Drift Diffusion Modeling
- `new_wsap_ddm_data.csv` - New WSAP data formatted for Drift Diffusion Modeling
- `wsap_data_quality_report.csv` - Data quality metrics for both tasks

---

## Data Quality

Each analysis script includes:
- Participant-level data quality indicators
- Count of valid vs. missing responses
- Completion rates
- Data validation checks

## Notes

- All scripts read input data files from the parent `Analysis/` directory using relative paths (`../`)
- Output files are generated in the same subfolder as each script
- Existing result files will be overwritten when scripts are re-run
- Scripts handle missing data gracefully with appropriate NaN values

## Support

For questions or issues with the analysis scripts, please refer to the inline comments in each `.py` file for detailed documentation of the analysis procedures.
