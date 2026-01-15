import pandas as pd
import numpy as np
import os

# Create results directory
results_dir = "questionnaire_analysis_results"
os.makedirs(results_dir, exist_ok=True)

# Read the Excel file
file_name = "1_labels_excel.xlsx"
df = pd.read_excel(file_name, header=0, skiprows=[1, 2])

# ============================================================================
# QIDS ANALYSIS - Columns S-AG (Q2-Q16)
# ============================================================================

# Define the columns to sum (S-AG = Q2-Q16)
questionnaire_cols = ['Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9',
                      'Q10', 'Q11', 'Q12', 'Q13', 'Q14', 'Q15', 'Q16']

results = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']

    # Extract questionnaire responses
    responses = row[questionnaire_cols]

    # Convert to numeric, replacing any non-numeric with NaN
    responses_numeric = pd.to_numeric(responses, errors='coerce')

    # Calculate total score
    total_score = responses_numeric.sum()

    # Count valid (non-NaN) responses
    valid_responses = responses_numeric.notna().sum()

    # Count missing responses
    missing_responses = responses_numeric.isna().sum()

    # Calculate mean score (average per item)
    mean_score = responses_numeric.mean() if valid_responses > 0 else np.nan

    results.append({
        'ResponseId': participant_id,
        'Questionnaire_Total_Score': total_score,
        'Questionnaire_Mean_Score': mean_score,
        'Valid_Items': valid_responses,
        'Missing_Items': missing_responses,
        'Total_Items': len(questionnaire_cols),
        'Completion_Rate': f"{(valid_responses/len(questionnaire_cols)*100):.1f}%"
    })

results_df = pd.DataFrame(results)

# ============================================================================
# CALCULATE SUMMARY STATISTICS
# ============================================================================

# Filter out any participants with no valid data
valid_participants = results_df[results_df['Valid_Items'] > 0]

if len(valid_participants) > 0:
    # Create summary statistics DataFrame
    summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            None,
            'Total Score - Mean',
            'Total Score - SD',
            'Total Score - Min',
            'Total Score - Max',
            'Total Score - Median',
            None,
            'Mean Score (per item) - Mean',
            'Mean Score (per item) - SD',
            'Mean Score (per item) - Min',
            'Mean Score (per item) - Max',
            None,
            'Average Completion Rate',
            'Participants with Complete Data',
            'Participants with Incomplete Data'
        ],
        'Value': [
            str(len(results_df)),
            str(len(valid_participants)),
            str(len(results_df) - len(valid_participants)),
            '',
            f"{valid_participants['Questionnaire_Total_Score'].mean():.2f}",
            f"{valid_participants['Questionnaire_Total_Score'].std():.2f}",
            f"{valid_participants['Questionnaire_Total_Score'].min():.2f}",
            f"{valid_participants['Questionnaire_Total_Score'].max():.2f}",
            f"{valid_participants['Questionnaire_Total_Score'].median():.2f}",
            '',
            f"{valid_participants['Questionnaire_Mean_Score'].mean():.2f}",
            f"{valid_participants['Questionnaire_Mean_Score'].std():.2f}",
            f"{valid_participants['Questionnaire_Mean_Score'].min():.2f}",
            f"{valid_participants['Questionnaire_Mean_Score'].max():.2f}",
            '',
            f"{(valid_participants['Valid_Items'].sum()/(len(valid_participants)*len(questionnaire_cols))*100):.1f}%",
            str(len(valid_participants[valid_participants['Valid_Items'] == len(questionnaire_cols)])),
            str(len(valid_participants[valid_participants['Valid_Items'] < len(questionnaire_cols)]))
        ]
    }
else:
    summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(results_df)), '0']
    }

summary_df = pd.DataFrame(summary_data)

# ============================================================================
# GAD ANALYSIS - Columns AH-AN (Q1_1-Q1_7)
# ============================================================================

# Define the GAD columns (AH-AN = Q1_1-Q1_7)
gad_cols = ['Q1_1', 'Q1_2', 'Q1_3', 'Q1_4', 'Q1_5', 'Q1_6', 'Q1_7']

gad_results = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']

    # Extract GAD responses
    responses = row[gad_cols]

    # Convert to numeric, replacing any non-numeric with NaN
    responses_numeric = pd.to_numeric(responses, errors='coerce')

    # Calculate total score
    total_score = responses_numeric.sum()

    # Count valid (non-NaN) responses
    valid_responses = responses_numeric.notna().sum()

    # Count missing responses
    missing_responses = responses_numeric.isna().sum()

    # Calculate mean score (average per item)
    mean_score = responses_numeric.mean() if valid_responses > 0 else np.nan

    gad_results.append({
        'ResponseId': participant_id,
        'GAD_Total_Score': total_score,
        'GAD_Mean_Score': mean_score,
        'Valid_Items': valid_responses,
        'Missing_Items': missing_responses,
        'Total_Items': len(gad_cols),
        'Completion_Rate': f"{(valid_responses/len(gad_cols)*100):.1f}%"
    })

gad_results_df = pd.DataFrame(gad_results)

# ============================================================================
# CALCULATE GAD SUMMARY STATISTICS
# ============================================================================

# Filter out any participants with no valid data
gad_valid_participants = gad_results_df[gad_results_df['Valid_Items'] > 0]

if len(gad_valid_participants) > 0:
    # Create summary statistics DataFrame
    gad_summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            None,
            'Total Score - Mean',
            'Total Score - SD',
            'Total Score - Min',
            'Total Score - Max',
            'Total Score - Median',
            None,
            'Mean Score (per item) - Mean',
            'Mean Score (per item) - SD',
            'Mean Score (per item) - Min',
            'Mean Score (per item) - Max',
            None,
            'Average Completion Rate',
            'Participants with Complete Data',
            'Participants with Incomplete Data'
        ],
        'Value': [
            str(len(gad_results_df)),
            str(len(gad_valid_participants)),
            str(len(gad_results_df) - len(gad_valid_participants)),
            '',
            f"{gad_valid_participants['GAD_Total_Score'].mean():.2f}",
            f"{gad_valid_participants['GAD_Total_Score'].std():.2f}",
            f"{gad_valid_participants['GAD_Total_Score'].min():.2f}",
            f"{gad_valid_participants['GAD_Total_Score'].max():.2f}",
            f"{gad_valid_participants['GAD_Total_Score'].median():.2f}",
            '',
            f"{gad_valid_participants['GAD_Mean_Score'].mean():.2f}",
            f"{gad_valid_participants['GAD_Mean_Score'].std():.2f}",
            f"{gad_valid_participants['GAD_Mean_Score'].min():.2f}",
            f"{gad_valid_participants['GAD_Mean_Score'].max():.2f}",
            '',
            f"{(gad_valid_participants['Valid_Items'].sum()/(len(gad_valid_participants)*len(gad_cols))*100):.1f}%",
            str(len(gad_valid_participants[gad_valid_participants['Valid_Items'] == len(gad_cols)])),
            str(len(gad_valid_participants[gad_valid_participants['Valid_Items'] < len(gad_cols)]))
        ]
    }
else:
    gad_summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(gad_results_df)), '0']
    }

gad_summary_df = pd.DataFrame(gad_summary_data)

# ============================================================================
# MASQ ANALYSIS - Columns AO-BN (Q1_1.1-Q1_26)
# ============================================================================

# Define the MASQ columns (AO-BN = Q1_1.1-Q1_26)
masq_cols = ['Q1_1.1', 'Q1_2.1', 'Q1_3.1', 'Q1_4.1', 'Q1_5.1', 'Q1_6.1', 'Q1_7.1',
             'Q1_8', 'Q1_9', 'Q1_10', 'Q1_11', 'Q1_12', 'Q1_13', 'Q1_14', 'Q1_15',
             'Q1_16', 'Q1_17', 'Q1_18', 'Q1_19', 'Q1_20', 'Q1_21', 'Q1_22', 'Q1_23',
             'Q1_24', 'Q1_25', 'Q1_26']

# Define subscale items (using 1-based indexing as in the instructions)
# Negatively keyed items that need reverse scoring
negative_keyed_items = [1, 9, 15, 19, 23, 25]

# Subscale item numbers (1-based)
gd_items = [2, 3, 7, 12, 13, 17, 20, 21]  # General Distress
aa_items = [4, 6, 8, 10, 14, 16, 18, 22, 24, 26]  # Anxious Arousal
ad_positive_items = [5, 11]  # Anhedonic Depression - positively keyed
ad_negative_items = [1, 9, 15, 19, 23, 25]  # Anhedonic Depression - negatively keyed

masq_results = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']

    # Extract MASQ responses
    responses = row[masq_cols]

    # Convert to numeric, replacing any non-numeric with NaN
    responses_numeric = pd.to_numeric(responses, errors='coerce')

    # Create a copy for scoring (1-based indexing for easier mapping)
    # Index 0 will be unused, indices 1-26 correspond to items 1-26
    scored_items = [np.nan] + list(responses_numeric.values)

    # Reverse score negatively keyed items
    for item_num in negative_keyed_items:
        if not pd.isna(scored_items[item_num]):
            scored_items[item_num] = 6 - scored_items[item_num]

    # Calculate GD (General Distress) score
    gd_values = [scored_items[i] for i in gd_items]
    gd_valid = sum(1 for v in gd_values if not pd.isna(v))
    gd_total = sum(v for v in gd_values if not pd.isna(v))

    # Calculate AA (Anxious Arousal) score
    aa_values = [scored_items[i] for i in aa_items]
    aa_valid = sum(1 for v in aa_values if not pd.isna(v))
    aa_total = sum(v for v in aa_values if not pd.isna(v))

    # Calculate AD (Anhedonic Depression) score
    ad_values = [scored_items[i] for i in ad_positive_items + ad_negative_items]
    ad_valid = sum(1 for v in ad_values if not pd.isna(v))
    ad_total = sum(v for v in ad_values if not pd.isna(v))

    # Overall data quality
    total_valid = sum(1 for v in scored_items[1:] if not pd.isna(v))
    total_missing = len(masq_cols) - total_valid

    masq_results.append({
        'ResponseId': participant_id,
        'GD_Total_Score': gd_total,
        'GD_Valid_Items': gd_valid,
        'AA_Total_Score': aa_total,
        'AA_Valid_Items': aa_valid,
        'AD_Total_Score': ad_total,
        'AD_Valid_Items': ad_valid,
        'Total_Valid_Items': total_valid,
        'Total_Missing_Items': total_missing,
        'Total_Items': len(masq_cols),
        'Completion_Rate': f"{(total_valid/len(masq_cols)*100):.1f}%"
    })

masq_results_df = pd.DataFrame(masq_results)

# ============================================================================
# CALCULATE MASQ SUMMARY STATISTICS
# ============================================================================

# Filter out any participants with no valid data
masq_valid_participants = masq_results_df[masq_results_df['Total_Valid_Items'] > 0]

if len(masq_valid_participants) > 0:
    # Create summary statistics DataFrame for all three subscales
    masq_summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            'Average Completion Rate',
            '',
            'GD Total Score - Mean',
            'GD Total Score - SD',
            'GD Total Score - Min',
            'GD Total Score - Max',
            'GD Total Score - Median',
            '',
            'AA Total Score - Mean',
            'AA Total Score - SD',
            'AA Total Score - Min',
            'AA Total Score - Max',
            'AA Total Score - Median',
            '',
            'AD Total Score - Mean',
            'AD Total Score - SD',
            'AD Total Score - Min',
            'AD Total Score - Max',
            'AD Total Score - Median'
        ],
        'Value': [
            str(len(masq_results_df)),
            str(len(masq_valid_participants)),
            str(len(masq_results_df) - len(masq_valid_participants)),
            f"{(masq_valid_participants['Total_Valid_Items'].sum()/(len(masq_valid_participants)*len(masq_cols))*100):.1f}%",
            '',
            f"{masq_valid_participants['GD_Total_Score'].mean():.2f}",
            f"{masq_valid_participants['GD_Total_Score'].std():.2f}",
            f"{masq_valid_participants['GD_Total_Score'].min():.2f}",
            f"{masq_valid_participants['GD_Total_Score'].max():.2f}",
            f"{masq_valid_participants['GD_Total_Score'].median():.2f}",
            '',
            f"{masq_valid_participants['AA_Total_Score'].mean():.2f}",
            f"{masq_valid_participants['AA_Total_Score'].std():.2f}",
            f"{masq_valid_participants['AA_Total_Score'].min():.2f}",
            f"{masq_valid_participants['AA_Total_Score'].max():.2f}",
            f"{masq_valid_participants['AA_Total_Score'].median():.2f}",
            '',
            f"{masq_valid_participants['AD_Total_Score'].mean():.2f}",
            f"{masq_valid_participants['AD_Total_Score'].std():.2f}",
            f"{masq_valid_participants['AD_Total_Score'].min():.2f}",
            f"{masq_valid_participants['AD_Total_Score'].max():.2f}",
            f"{masq_valid_participants['AD_Total_Score'].median():.2f}"
        ]
    }
else:
    masq_summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(masq_results_df)), '0']
    }

masq_summary_df = pd.DataFrame(masq_summary_data)

# ============================================================================
# EXPORT RESULTS TO EXCEL
# ============================================================================

output_file = os.path.join(results_dir, "questionnaire_analysis_results.xlsx")

# Write to Excel with multiple sheets
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: QIDS Participant-level results
    results_df.to_excel(writer, sheet_name='QIDS Results', index=False)

    # Sheet 2: QIDS Summary statistics
    summary_df.to_excel(writer, sheet_name='QIDS Summary', index=False)

    # Sheet 3: GAD Participant-level results
    gad_results_df.to_excel(writer, sheet_name='GAD Results', index=False)

    # Sheet 4: GAD Summary statistics
    gad_summary_df.to_excel(writer, sheet_name='GAD Summary', index=False)

    # Sheet 5: MASQ Participant-level results
    masq_results_df.to_excel(writer, sheet_name='MASQ Results', index=False)

    # Sheet 6: MASQ Summary statistics
    masq_summary_df.to_excel(writer, sheet_name='MASQ Summary', index=False)

print(f"Questionnaire analysis complete. Results saved to: {output_file}")
