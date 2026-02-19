# Qualtrics Data Analysis

**Summary:** Analysis scripts for processing questionnaire and behavioral task data from Qualtrics surveys.

For detailed documentation, see [Analysis/README.md](Analysis/README.md)

## Directory Structure

```
Qualtrics-Data-Analysis/
└── Analysis/
    ├── 1_labels_excel.xlsx          # Input data (labeled responses)
    ├── 1_values_excel.xlsx          # Input data (numeric values)
    ├── README.md                    # Detailed documentation
    ├── AST/                         # Ambiguous Scenarios Task
    ├── SST/                         # Scrambled Sentences Test
    ├── Questionnaire/               # QIDS, GAD-7, MASQ
    └── WSAP/                        # Word Sentence Association Paradigm
```

## Quick Start

Each analysis has its own subfolder with a Python script and output files.

```bash
cd Analysis/AST
python3 ast_analysis.py

cd ../SST
python3 sst_analysis.py

cd ../Questionnaire
python3 questionnaire_analysis.py

cd ../WSAP
python3 wsap_analysis.py
```

## Documentation

**For detailed documentation, formulas, and usage instructions, see [Analysis/README.md](Analysis/README.md)**

## Requirements

```bash
pip install pandas numpy openpyxl
```

## Available Analyses

1. **AST (Ambiguous Scenarios Task)**
   - Reverse-scored pleasantness ratings
   - Qualitative coding template for outcome descriptions

2. **SST (Scrambled Sentences Test)**
   - Negativity scores for anxiety and depression stimuli

3. **Questionnaire Analysis**
   - QIDS (Quick Inventory of Depressive Symptomatology)
   - GAD-7 (Generalized Anxiety Disorder)
   - MASQ (Mood and Anxiety Symptom Questionnaire)

4. **WSAP (Word Sentence Association Paradigm)**
   - Response selection scores and RT bias indices
   - DDM-ready data export

## Notes

- All scripts read input files from the `Analysis/` directory
- Output files are generated in the same subfolder as each script
- Existing result files will be overwritten when scripts are re-run
- Scripts handle missing data gracefully with appropriate NaN values
