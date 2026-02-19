import pandas as pd
import numpy as np

# Read the Excel file from parent directory
file_name = "../1_values_excel.xlsx"
df = pd.read_excel(file_name, header=0, skiprows=[1, 2])

# ============================================================================
# SST ANALYSIS - Negativity Score Calculation
# ============================================================================
# Formula: Negativity score = Total negative sentences / (Total negative + Total positive sentences)
# Mixed and unclear sentences are excluded from both numerator and denominator
# Participants where mixed > (positive + negative) are excluded
# For anxiety and depression stimuli only

sst_results = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']
    list_assignment = row['list_assignment']
    main_total_completed = row['main_total_completed']
    main_sentence_interpretations = row['main_sentence_interpretations']

    # Check if participant has valid SST data
    if pd.isna(main_sentence_interpretations) or pd.isna(main_total_completed):
        sst_results.append({
            'ResponseId': participant_id,
            'List_Assignment': list_assignment,
            'Total_Completed_Sentences': main_total_completed,
            'Negative_D_Count': np.nan,
            'Negative_GA_Count': np.nan,
            'Total_Negative_Count': np.nan,
            'Positive_Count': np.nan,
            'Mixed_Count': np.nan,
            'Unclear_Count': np.nan,
            'Negativity_Score': np.nan,
            'Data_Quality': "No data"
        })
        continue

    # Parse interpretations
    interpretations = str(main_sentence_interpretations).split(';')

    # Count each interpretation type
    negative_d_count = interpretations.count('negative_D')
    negative_ga_count = interpretations.count('negative_GA')
    total_negative_count = negative_d_count + negative_ga_count
    positive_count = interpretations.count('positive')
    mixed_count = interpretations.count('mixed')
    unclear_count = interpretations.count('unclear')

    # Exclude participant if mixed > total positive + negative
    valid_denominator = positive_count + total_negative_count
    if mixed_count > valid_denominator:
        sst_results.append({
            'ResponseId': participant_id,
            'List_Assignment': list_assignment,
            'Total_Completed_Sentences': main_total_completed,
            'Negative_D_Count': negative_d_count,
            'Negative_GA_Count': negative_ga_count,
            'Total_Negative_Count': total_negative_count,
            'Positive_Count': positive_count,
            'Mixed_Count': mixed_count,
            'Unclear_Count': unclear_count,
            'Negativity_Score': np.nan,
            'Data_Quality': f"Excluded: Mixed ({mixed_count}) > Positive + Negative ({valid_denominator})"
        })
        continue

    # Calculate negativity score
    # Negativity score = Total negative sentences / (Total negative + Total positive)
    # Mixed and unclear sentences are excluded from both numerator and denominator
    if valid_denominator > 0:
        negativity_score = total_negative_count / valid_denominator
    else:
        negativity_score = np.nan

    # Data quality check
    total_interpretations = len(interpretations)
    if total_interpretations == main_total_completed:
        data_quality = f"Complete: {total_interpretations}/{main_total_completed}"
    else:
        data_quality = f"Mismatch: {total_interpretations} interpretations vs {main_total_completed} completed"

    sst_results.append({
        'ResponseId': participant_id,
        'List_Assignment': list_assignment,
        'Total_Completed_Sentences': main_total_completed,
        'Negative_D_Count': negative_d_count,
        'Negative_GA_Count': negative_ga_count,
        'Total_Negative_Count': total_negative_count,
        'Positive_Count': positive_count,
        'Mixed_Count': mixed_count,
        'Unclear_Count': unclear_count,
        'Negativity_Score': negativity_score,
        'Data_Quality': data_quality
    })

sst_results_df = pd.DataFrame(sst_results)

# ============================================================================
# CALCULATE SUMMARY STATISTICS
# ============================================================================

# Filter out any participants with no valid data
valid_participants = sst_results_df[sst_results_df['Negativity_Score'].notna()]

if len(valid_participants) > 0:
    # Create summary statistics DataFrame
    summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            None,
            'Negativity Score - Mean',
            'Negativity Score - SD',
            'Negativity Score - Min',
            'Negativity Score - Max',
            'Negativity Score - Median',
            None,
            'Total Completed Sentences - Mean',
            'Total Completed Sentences - SD',
            'Total Completed Sentences - Min',
            'Total Completed Sentences - Max',
            None,
            'Negative Sentences (D) - Mean',
            'Negative Sentences (D) - SD',
            'Negative Sentences (GA) - Mean',
            'Negative Sentences (GA) - SD',
            'Total Negative Sentences - Mean',
            'Total Negative Sentences - SD',
            None,
            'Positive Sentences - Mean',
            'Positive Sentences - SD',
            'Mixed Sentences - Mean',
            'Mixed Sentences - SD',
            'Unclear Sentences - Mean',
            'Unclear Sentences - SD'
        ],
        'Value': [
            str(len(sst_results_df)),
            str(len(valid_participants)),
            str(len(sst_results_df) - len(valid_participants)),
            '',
            f"{valid_participants['Negativity_Score'].mean():.4f}",
            f"{valid_participants['Negativity_Score'].std():.4f}",
            f"{valid_participants['Negativity_Score'].min():.4f}",
            f"{valid_participants['Negativity_Score'].max():.4f}",
            f"{valid_participants['Negativity_Score'].median():.4f}",
            '',
            f"{valid_participants['Total_Completed_Sentences'].mean():.2f}",
            f"{valid_participants['Total_Completed_Sentences'].std():.2f}",
            f"{valid_participants['Total_Completed_Sentences'].min():.0f}",
            f"{valid_participants['Total_Completed_Sentences'].max():.0f}",
            '',
            f"{valid_participants['Negative_D_Count'].mean():.2f}",
            f"{valid_participants['Negative_D_Count'].std():.2f}",
            f"{valid_participants['Negative_GA_Count'].mean():.2f}",
            f"{valid_participants['Negative_GA_Count'].std():.2f}",
            f"{valid_participants['Total_Negative_Count'].mean():.2f}",
            f"{valid_participants['Total_Negative_Count'].std():.2f}",
            '',
            f"{valid_participants['Positive_Count'].mean():.2f}",
            f"{valid_participants['Positive_Count'].std():.2f}",
            f"{valid_participants['Mixed_Count'].mean():.2f}",
            f"{valid_participants['Mixed_Count'].std():.2f}",
            f"{valid_participants['Unclear_Count'].mean():.2f}",
            f"{valid_participants['Unclear_Count'].std():.2f}"
        ]
    }
else:
    summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(sst_results_df)), '0']
    }

summary_df = pd.DataFrame(summary_data)

# ============================================================================
# ANALYSIS BY LIST ASSIGNMENT
# ============================================================================

# Calculate summary statistics for each list assignment
list_assignments = valid_participants['List_Assignment'].dropna().unique()
list_summary_data = []

for list_num in sorted(list_assignments):
    list_data = valid_participants[valid_participants['List_Assignment'] == list_num]

    if len(list_data) > 0:
        list_summary_data.append({
            'List_Assignment': f"List {int(list_num)}",
            'N_Participants': len(list_data),
            'Negativity_Score_Mean': f"{list_data['Negativity_Score'].mean():.4f}",
            'Negativity_Score_SD': f"{list_data['Negativity_Score'].std():.4f}",
            'Negativity_Score_Median': f"{list_data['Negativity_Score'].median():.4f}",
            'Total_Negative_Mean': f"{list_data['Total_Negative_Count'].mean():.2f}",
            'Total_Negative_SD': f"{list_data['Total_Negative_Count'].std():.2f}"
        })

if list_summary_data:
    list_summary_df = pd.DataFrame(list_summary_data)
else:
    list_summary_df = pd.DataFrame({'Message': ['No list assignment data available']})

# ============================================================================
# EXPORT RESULTS TO EXCEL
# ============================================================================

output_file = "sst_analysis_results.xlsx"

# Write to Excel with multiple sheets
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: Participant-level results
    sst_results_df.to_excel(writer, sheet_name='SST Results', index=False)

    # Sheet 2: Overall summary statistics
    summary_df.to_excel(writer, sheet_name='SST Summary', index=False)

    # Sheet 3: Summary by list assignment
    list_summary_df.to_excel(writer, sheet_name='Summary by List', index=False)

print(f"SST analysis complete. Results saved to: {output_file}")
print(f"\nSummary:")
print(f"  Total participants: {len(sst_results_df)}")
print(f"  Participants with valid data: {len(valid_participants)}")
if len(valid_participants) > 0:
    print(f"  Mean negativity score: {valid_participants['Negativity_Score'].mean():.4f} (SD: {valid_participants['Negativity_Score'].std():.4f})")
    print(f"  Range: {valid_participants['Negativity_Score'].min():.4f} - {valid_participants['Negativity_Score'].max():.4f}")
