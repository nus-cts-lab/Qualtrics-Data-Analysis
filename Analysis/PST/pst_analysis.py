import pandas as pd
import numpy as np

# Read the Excel file from parent directory
file_name = "../1_values_excel.xlsx"
df = pd.read_excel(file_name, header=0, skiprows=[1, 2])

# ============================================================================
# PST ANALYSIS - RT Bias Index Calculation
# ============================================================================
# Only correctly resolved scenarios (main_word_accuracy == true) are included
# Negative scenarios = anxiety + depression
# Positive scenarios = positive
# Formula: RT bias index = Negative mean RT - Positive mean RT
# The smaller the RT bias index, the faster the formation of negative interpretations

pst_results = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']
    list_assignment = row['list_assignment']
    main_scenarios_completed = row['main_scenarios_completed']

    main_rts = row['main_reaction_times']
    main_word_acc = row['main_word_accuracy']
    main_comp_acc = row['main_comprehension_accuracy']
    main_scenario_types = row['main_scenario_types']

    # Check if participant has valid PST data
    if pd.isna(main_rts) or pd.isna(main_scenario_types) or pd.isna(main_word_acc):
        pst_results.append({
            'ResponseId': participant_id,
            'List_Assignment': list_assignment,
            'Main_Scenarios_Completed': main_scenarios_completed,
            'N_Correctly_Resolved': np.nan,
            'N_Negative_Valid': np.nan,
            'N_Positive_Valid': np.nan,
            'Mean_RT_Negative': np.nan,
            'Mean_RT_Positive': np.nan,
            'RT_Bias_Index': np.nan,
            'Data_Quality': "No data"
        })
        continue

    # Parse semicolon-separated values, filtering out empty strings from trailing semicolons
    rts_raw = [x.strip() for x in str(main_rts).split(';') if x.strip() != '']
    word_accs = [x.strip().lower() for x in str(main_word_acc).split(';') if x.strip() != '']
    comp_accs = [x.strip().lower() for x in str(main_comp_acc).split(';') if x.strip() != ''] if not pd.isna(main_comp_acc) else []
    scenario_types = [x.strip().lower() for x in str(main_scenario_types).split(';') if x.strip() != '']

    # Parse RTs to float
    rts = []
    for val in rts_raw:
        try:
            rts.append(float(val))
        except ValueError:
            rts.append(np.nan)

    # Build trial-level data aligned by index
    n_trials = max(len(rts), len(word_accs), len(scenario_types))

    negative_rts = []
    positive_rts = []
    n_correctly_resolved = 0

    for i in range(n_trials):
        rt = rts[i] if i < len(rts) else np.nan
        word_acc = word_accs[i] if i < len(word_accs) else None
        scenario_type = scenario_types[i] if i < len(scenario_types) else None

        # Only include correctly resolved scenarios (word fragment correctly filled)
        if word_acc != 'true':
            continue

        n_correctly_resolved += 1

        if pd.isna(rt) or scenario_type is None:
            continue

        if scenario_type in ('anxiety', 'depression'):
            negative_rts.append(rt)
        elif scenario_type == 'positive':
            positive_rts.append(rt)

    # Calculate mean RTs
    mean_rt_negative = np.mean(negative_rts) if len(negative_rts) > 0 else np.nan
    mean_rt_positive = np.mean(positive_rts) if len(positive_rts) > 0 else np.nan

    # RT bias index = Negative mean RT - Positive mean RT
    if not (pd.isna(mean_rt_negative) or pd.isna(mean_rt_positive)):
        rt_bias_index = mean_rt_negative - mean_rt_positive
    else:
        rt_bias_index = np.nan

    # Data quality
    if not pd.isna(main_scenarios_completed) and n_correctly_resolved == int(main_scenarios_completed):
        data_quality = f"Complete: {n_correctly_resolved}/{int(main_scenarios_completed)}"
    else:
        data_quality = f"Correctly resolved: {n_correctly_resolved} of {int(main_scenarios_completed) if not pd.isna(main_scenarios_completed) else '?'} completed"

    pst_results.append({
        'ResponseId': participant_id,
        'List_Assignment': list_assignment,
        'Main_Scenarios_Completed': main_scenarios_completed,
        'N_Correctly_Resolved': n_correctly_resolved,
        'N_Negative_Valid': len(negative_rts),
        'N_Positive_Valid': len(positive_rts),
        'Mean_RT_Negative': mean_rt_negative,
        'Mean_RT_Positive': mean_rt_positive,
        'RT_Bias_Index': rt_bias_index,
        'Data_Quality': data_quality
    })

pst_results_df = pd.DataFrame(pst_results)

# ============================================================================
# CALCULATE SUMMARY STATISTICS
# ============================================================================

valid_participants = pst_results_df[pst_results_df['RT_Bias_Index'].notna()]

if len(valid_participants) > 0:
    summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            None,
            'RT Bias Index - Mean',
            'RT Bias Index - SD',
            'RT Bias Index - Min',
            'RT Bias Index - Max',
            'RT Bias Index - Median',
            None,
            'Mean RT Negative - Mean',
            'Mean RT Negative - SD',
            'Mean RT Positive - Mean',
            'Mean RT Positive - SD',
            None,
            'Correctly Resolved Scenarios - Mean',
            'Correctly Resolved Scenarios - SD',
            'Correctly Resolved Scenarios - Min',
            'Correctly Resolved Scenarios - Max'
        ],
        'Value': [
            str(len(pst_results_df)),
            str(len(valid_participants)),
            str(len(pst_results_df) - len(valid_participants)),
            '',
            f"{valid_participants['RT_Bias_Index'].mean():.3f}",
            f"{valid_participants['RT_Bias_Index'].std():.3f}",
            f"{valid_participants['RT_Bias_Index'].min():.3f}",
            f"{valid_participants['RT_Bias_Index'].max():.3f}",
            f"{valid_participants['RT_Bias_Index'].median():.3f}",
            '',
            f"{valid_participants['Mean_RT_Negative'].mean():.3f}",
            f"{valid_participants['Mean_RT_Negative'].std():.3f}",
            f"{valid_participants['Mean_RT_Positive'].mean():.3f}",
            f"{valid_participants['Mean_RT_Positive'].std():.3f}",
            '',
            f"{valid_participants['N_Correctly_Resolved'].mean():.2f}",
            f"{valid_participants['N_Correctly_Resolved'].std():.2f}",
            f"{valid_participants['N_Correctly_Resolved'].min():.0f}",
            f"{valid_participants['N_Correctly_Resolved'].max():.0f}"
        ]
    }
else:
    summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(pst_results_df)), '0']
    }

summary_df = pd.DataFrame(summary_data)

# ============================================================================
# ANALYSIS BY LIST ASSIGNMENT
# ============================================================================

list_assignments = valid_participants['List_Assignment'].dropna().unique()
list_summary_data = []

for list_num in sorted(list_assignments):
    list_data = valid_participants[valid_participants['List_Assignment'] == list_num]

    if len(list_data) > 0:
        list_summary_data.append({
            'List_Assignment': f"List {int(list_num)}",
            'N_Participants': len(list_data),
            'RT_Bias_Index_Mean': f"{list_data['RT_Bias_Index'].mean():.3f}",
            'RT_Bias_Index_SD': f"{list_data['RT_Bias_Index'].std():.3f}",
            'RT_Bias_Index_Median': f"{list_data['RT_Bias_Index'].median():.3f}",
            'Mean_RT_Negative_Mean': f"{list_data['Mean_RT_Negative'].mean():.3f}",
            'Mean_RT_Positive_Mean': f"{list_data['Mean_RT_Positive'].mean():.3f}"
        })

if list_summary_data:
    list_summary_df = pd.DataFrame(list_summary_data)
else:
    list_summary_df = pd.DataFrame({'Message': ['No list assignment data available']})

# ============================================================================
# EXPORT RESULTS TO EXCEL
# ============================================================================

output_file = "pst_analysis_results.xlsx"

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    pst_results_df.to_excel(writer, sheet_name='PST Results', index=False)
    summary_df.to_excel(writer, sheet_name='PST Summary', index=False)
    list_summary_df.to_excel(writer, sheet_name='Summary by List', index=False)

print(f"PST analysis complete. Results saved to: {output_file}")
print(f"\nSummary:")
print(f"  Total participants: {len(pst_results_df)}")
print(f"  Participants with valid data: {len(valid_participants)}")
if len(valid_participants) > 0:
    print(f"  Mean RT bias index: {valid_participants['RT_Bias_Index'].mean():.3f} (SD: {valid_participants['RT_Bias_Index'].std():.3f})")
    print(f"  Range: {valid_participants['RT_Bias_Index'].min():.3f} - {valid_participants['RT_Bias_Index'].max():.3f}")
    print(f"  Mean RT Negative: {valid_participants['Mean_RT_Negative'].mean():.3f}")
    print(f"  Mean RT Positive: {valid_participants['Mean_RT_Positive'].mean():.3f}")
