import pandas as pd
import numpy as np

# Read the Excel file from parent directory
file_name = "../1_values_excel.xlsx"
df = pd.read_excel(file_name, header=0, skiprows=[1, 2])

# ============================================================================
# AST ANALYSIS - Reverse-Scored Pleasantness Ratings
# ============================================================================
# Formula: Reverse score = 10 - original score
# Calculate mean of all reverse-scored items

ast_results = []
coding_data = []
subject_number = 1

for idx, row in df.iterrows():
    participant_id = row['ResponseId']
    main_ratings = row['main_pleasantness_ratings']
    main_descriptions = row['main_outcome_descriptions']

    # Parse ratings
    if pd.isna(main_ratings):
        ratings_list = []
    else:
        ratings_list = str(main_ratings).split(';')

    # Parse descriptions
    if pd.isna(main_descriptions):
        descriptions_list = []
    else:
        descriptions_list = str(main_descriptions).split('|')

    # Reverse score ratings: 10 - x
    reverse_scored_ratings = []
    for rating in ratings_list:
        try:
            original_score = float(rating.strip())
            reverse_score = 10 - original_score
            reverse_scored_ratings.append(reverse_score)
        except (ValueError, AttributeError):
            reverse_scored_ratings.append(np.nan)

    # Calculate mean of reverse-scored ratings
    valid_reverse_scores = [score for score in reverse_scored_ratings if not pd.isna(score)]

    if len(valid_reverse_scores) > 0:
        mean_reverse_score = np.mean(valid_reverse_scores)
        has_ratings_data = True
    else:
        mean_reverse_score = np.nan
        has_ratings_data = False

    # Check if descriptions are valid (not just x, -, ?, etc.)
    valid_descriptions = []
    invalid_markers = ['x', '-', '?', 'nan', '']

    for desc in descriptions_list:
        desc_clean = str(desc).strip().lower()
        # Consider valid if it has more than 2 characters and isn't just a marker
        if len(desc_clean) > 2 and desc_clean not in invalid_markers:
            valid_descriptions.append(str(desc).strip())
        else:
            valid_descriptions.append(None)

    has_description_data = any(d is not None for d in valid_descriptions)

    # Store results
    ast_results.append({
        'ResponseId': participant_id,
        'Mean_Reverse_Scored_Rating': mean_reverse_score,
        'Total_Ratings': len(ratings_list),
        'Valid_Ratings': len(valid_reverse_scores),
        'Total_Descriptions': len(descriptions_list),
        'Valid_Descriptions': sum(1 for d in valid_descriptions if d is not None),
        'Has_Ratings_Data': 'Yes' if has_ratings_data else 'No',
        'Has_Description_Data': 'Yes' if has_description_data else 'No'
    })

    # Prepare coding template data (only if participant has valid descriptions)
    if has_description_data:
        for desc_idx, description in enumerate(descriptions_list, 1):
            desc_clean = str(description).strip()
            # Add to coding template even if individual description might be invalid
            # Coders can mark these as unclear if needed
            coding_data.append({
                'Subject': subject_number,
                'Main_Outcome_Descriptions': desc_clean,
                'Coder_1': '',
                'Coder_2': '',
                'Final': '',
                'Coder_3': ''
            })
        subject_number += 1

ast_results_df = pd.DataFrame(ast_results)
coding_template_df = pd.DataFrame(coding_data)

# ============================================================================
# CALCULATE SUMMARY STATISTICS FOR REVERSE-SCORED RATINGS
# ============================================================================

valid_participants = ast_results_df[ast_results_df['Mean_Reverse_Scored_Rating'].notna()]

if len(valid_participants) > 0:
    summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Ratings',
            'Participants with Missing Ratings',
            None,
            'Mean Reverse-Scored Rating - Mean',
            'Mean Reverse-Scored Rating - SD',
            'Mean Reverse-Scored Rating - Min',
            'Mean Reverse-Scored Rating - Max',
            'Mean Reverse-Scored Rating - Median',
            None,
            'Total Ratings per Participant - Mean',
            'Valid Ratings per Participant - Mean',
            None,
            'Participants with Description Data',
            'Participants without Description Data'
        ],
        'Value': [
            str(len(ast_results_df)),
            str(len(valid_participants)),
            str(len(ast_results_df) - len(valid_participants)),
            '',
            f"{valid_participants['Mean_Reverse_Scored_Rating'].mean():.4f}",
            f"{valid_participants['Mean_Reverse_Scored_Rating'].std():.4f}",
            f"{valid_participants['Mean_Reverse_Scored_Rating'].min():.4f}",
            f"{valid_participants['Mean_Reverse_Scored_Rating'].max():.4f}",
            f"{valid_participants['Mean_Reverse_Scored_Rating'].median():.4f}",
            '',
            f"{valid_participants['Total_Ratings'].mean():.2f}",
            f"{valid_participants['Valid_Ratings'].mean():.2f}",
            '',
            str(len(ast_results_df[ast_results_df['Has_Description_Data'] == 'Yes'])),
            str(len(ast_results_df[ast_results_df['Has_Description_Data'] == 'No']))
        ]
    }
else:
    summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Ratings'],
        'Value': [str(len(ast_results_df)), '0']
    }

summary_df = pd.DataFrame(summary_data)

# Combine participant results with summary
results_sheet_df = pd.concat([
    ast_results_df,
    pd.DataFrame([{}]),  # Empty row separator
    summary_df
], ignore_index=True)

# ============================================================================
# DATA QUALITY REPORT
# ============================================================================

quality_data = []

for idx, row in ast_results_df.iterrows():
    quality_data.append({
        'ResponseId': row['ResponseId'],
        'Has_Ratings_Data': row['Has_Ratings_Data'],
        'Total_Ratings': row['Total_Ratings'],
        'Valid_Ratings': row['Valid_Ratings'],
        'Has_Description_Data': row['Has_Description_Data'],
        'Total_Descriptions': row['Total_Descriptions'],
        'Valid_Descriptions': row['Valid_Descriptions'],
        'Data_Status': 'Complete' if row['Has_Ratings_Data'] == 'Yes' and row['Has_Description_Data'] == 'Yes'
                       else 'Partial' if row['Has_Ratings_Data'] == 'Yes' or row['Has_Description_Data'] == 'Yes'
                       else 'No Data'
    })

quality_df = pd.DataFrame(quality_data)

# ============================================================================
# EXPORT RESULTS TO EXCEL
# ============================================================================

output_file = "ast_analysis_results.xlsx"

# Write to Excel with 3 sheets
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: Reverse-Scored Ratings (includes participant-level and summary)
    ast_results_df.to_excel(writer, sheet_name='Reverse-Scored Ratings', index=False)

    # Add summary to same sheet with spacing
    summary_df.to_excel(writer, sheet_name='Reverse-Scored Ratings',
                       startrow=len(ast_results_df) + 2, index=False)

    # Sheet 2: Data Quality Report
    quality_df.to_excel(writer, sheet_name='Data Quality', index=False)

    # Sheet 3: Coding Template for manual coding
    coding_template_df.to_excel(writer, sheet_name='Coding Template', index=False)

print(f"AST analysis complete. Results saved to: {output_file}")
print(f"\nSummary:")
print(f"  Total participants: {len(ast_results_df)}")
print(f"  Participants with valid ratings: {len(valid_participants)}")
if len(valid_participants) > 0:
    print(f"  Mean reverse-scored rating: {valid_participants['Mean_Reverse_Scored_Rating'].mean():.4f} (SD: {valid_participants['Mean_Reverse_Scored_Rating'].std():.4f})")
    print(f"  Range: {valid_participants['Mean_Reverse_Scored_Rating'].min():.4f} - {valid_participants['Mean_Reverse_Scored_Rating'].max():.4f}")
print(f"\n  Participants in coding template: {subject_number - 1}")
print(f"  Total descriptions to code: {len(coding_template_df)}")
