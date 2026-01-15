import pandas as pd
import numpy as np
import os

def safe_parse_comma_data(data_str, data_type='str'):
    """
    Safely parse comma-separated data, handling empty values and type conversion
    """
    if pd.isna(data_str) or str(data_str).strip() == '':
        return []
    
    values = str(data_str).split(',')
    parsed_values = []
    
    for val in values:
        val = val.strip()
        if val == '':
            parsed_values.append(np.nan)
        else:
            try:
                if data_type == 'float':
                    parsed_values.append(float(val))
                else:
                    parsed_values.append(val)
            except ValueError:
                parsed_values.append(np.nan)
    
    return parsed_values

def create_trial_dataframe(responses, rts, scenario_types, word_types=None):
    """
    Create trial-level DataFrame preserving all available data
    """
    # Find maximum length
    max_len = max(len(responses) if responses else 0, 
                  len(rts) if rts else 0,
                  len(scenario_types) if scenario_types else 0,
                  len(word_types) if word_types else 0)
    
    if max_len == 0:
        return pd.DataFrame()
    
    # Pad arrays to same length
    def pad_array(arr, target_len):
        if not arr:
            return [np.nan] * target_len
        padded = arr + [np.nan] * (target_len - len(arr))
        return padded[:target_len]
    
    trial_data = {
        'response': pad_array(responses, max_len),
        'rt': pad_array(rts, max_len),
        'scenario_type': pad_array(scenario_types, max_len)
    }
    
    if word_types is not None:
        trial_data['word_type'] = pad_array(word_types, max_len)
    
    return pd.DataFrame(trial_data)

# Create results directory
results_dir = "wsap_analysis_results"
os.makedirs(results_dir, exist_ok=True)

# Read the Excel file
file_name = "1_values_excel.xlsx"
df = pd.read_excel(file_name, header=0, skiprows=[1, 2])

# ============================================================================
# PART 1: ORIGINAL WSAP ANALYSIS (Columns DO-DT)
# ============================================================================

original_results = []
original_ddm_data = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']
    
    try:
        # Parse comma-separated values safely
        responses = safe_parse_comma_data(row['__js_responses'])
        rts = safe_parse_comma_data(row['__js_reaction_times'], 'float')
        scenario_types = safe_parse_comma_data(row['__js_scenario_types'])
        word_types = safe_parse_comma_data(row['__js_word_types'])
        
        # Create trial-level data preserving all available information
        trials = create_trial_dataframe(responses, rts, scenario_types, word_types)
        
        if len(trials) == 0:
            raise ValueError("No trial data available")
    
    except Exception as e:
        # Create "No data" record for this participant
        original_results.append({
            'ResponseId': participant_id,
            'Original_Response_Selection_Score': "No data",
            'Original_Prop_Negative_Endorsed': "No data",
            'Original_Prop_Benign_Endorsed': "No data", 
            'Original_RT_Bias_Index': "No data",
            'Original_Mean_RT_Endorse_Negative': "No data",
            'Original_Mean_RT_Reject_Negative': "No data",
            'Original_N_Trials': 0,
            'Original_N_Valid_Trials': 0,
            'Original_N_Depression_Trials': 0,
            'Original_N_Anxiety_Trials': 0,
            'Original_N_Positive_Trials': 0,
            'Original_Data_Quality': f"Error: {str(e)}"
        })
        continue
    
    # Filter for valid trials with available data
    valid_response_trials = trials[~trials['response'].isna()]
    valid_rt_trials = trials[(~trials['response'].isna()) & (~trials['rt'].isna())]
    
    # Calculate Response Selection Score (uses all trials with valid responses)
    negative_trials = valid_response_trials[valid_response_trials['scenario_type'].isin(['depression', 'anxiety'])]
    benign_trials = valid_response_trials[valid_response_trials['scenario_type'] == 'positive']
    
    prop_negative_endorsed = (negative_trials['response'] == 'r').sum() / len(negative_trials) if len(negative_trials) > 0 else np.nan
    prop_benign_endorsed = (benign_trials['response'] == 'r').sum() / len(benign_trials) if len(benign_trials) > 0 else np.nan
    
    response_selection_score = prop_negative_endorsed - prop_benign_endorsed if not (pd.isna(prop_negative_endorsed) or pd.isna(prop_benign_endorsed)) else np.nan
    
    # Calculate RT Bias Index (uses only trials with both response and RT)
    valid_negative_trials = valid_rt_trials[valid_rt_trials['scenario_type'].isin(['depression', 'anxiety'])]
    negative_endorsed = valid_negative_trials[valid_negative_trials['response'] == 'r']
    negative_rejected = valid_negative_trials[valid_negative_trials['response'] == 'u']
    
    mean_rt_endorse_negative = negative_endorsed['rt'].mean() if len(negative_endorsed) > 0 else np.nan
    mean_rt_reject_negative = negative_rejected['rt'].mean() if len(negative_rejected) > 0 else np.nan
    
    rt_bias_index = mean_rt_endorse_negative - mean_rt_reject_negative if not (pd.isna(mean_rt_endorse_negative) or pd.isna(mean_rt_reject_negative)) else np.nan
    
    # Prepare clean data for DDM (complete trials only)
    ddm_trials = valid_rt_trials[(~valid_rt_trials['scenario_type'].isna())].copy()
    ddm_trials['response_binary'] = (ddm_trials['response'] == 'r').astype(int)
    ddm_trials['participant_id'] = participant_id
    
    # Separate by stimulus type for DDM
    ddm_depression = ddm_trials[ddm_trials['scenario_type'] == 'depression'].copy()
    ddm_anxiety = ddm_trials[ddm_trials['scenario_type'] == 'anxiety'].copy() 
    ddm_positive = ddm_trials[ddm_trials['scenario_type'] == 'positive'].copy()
    
    # Add to DDM dataset for export
    if len(ddm_trials) > 0:
        original_ddm_data.append(ddm_trials)
    
    original_results.append({
        'ResponseId': participant_id,
        'Original_Response_Selection_Score': response_selection_score,
        'Original_Prop_Negative_Endorsed': prop_negative_endorsed,
        'Original_Prop_Benign_Endorsed': prop_benign_endorsed,
        'Original_RT_Bias_Index': rt_bias_index,
        'Original_Mean_RT_Endorse_Negative': mean_rt_endorse_negative,
        'Original_Mean_RT_Reject_Negative': mean_rt_reject_negative,
        'Original_N_Trials': len(trials),
        'Original_N_Valid_Trials': len(valid_rt_trials),
        'Original_N_Depression_Trials': len(ddm_depression),
        'Original_N_Anxiety_Trials': len(ddm_anxiety),
        'Original_N_Positive_Trials': len(ddm_positive),
        'Original_Data_Quality': f"Valid: {len(valid_rt_trials)}/{len(trials)} trials"
    })

original_df = pd.DataFrame(original_results)

# ============================================================================
# PART 2: NEW WSAP ANALYSIS (Columns BW-BZ)
# ============================================================================

new_results = []
new_ddm_data = []

for idx, row in df.iterrows():
    participant_id = row['ResponseId']
    
    # Column names for New WSAP (adjust if needed)
    rt_col = '__js_reaction_time'
    valence_col = '__js_valence'
    stimulus_col = '__js_stimulus'
    response_col = '__js_response'
    
    try:
        # Parse comma-separated values safely
        rts = safe_parse_comma_data(row[rt_col], 'float')
        valences = safe_parse_comma_data(row[valence_col])
        responses = safe_parse_comma_data(row[response_col])
        
        if not any([rts, valences, responses]):
            raise ValueError("No trial data available")
        
        # Create trial-level data with choice interpretation
        trials_list = []
        max_len = max(len(rts) if rts else 0, len(valences) if valences else 0, len(responses) if responses else 0)
        
        for i in range(max_len):
            # Get values or NaN if missing
            rt = rts[i] if i < len(rts) else np.nan
            valence = valences[i] if i < len(valences) else np.nan
            response = responses[i] if i < len(responses) else np.nan
            
            if pd.isna(response):
                chosen_valence = np.nan
            else:
                response = str(response).strip()
                
                # Determine which valence was chosen
                # j = left option (first valence), f = right option (second valence)
                if pd.isna(valence):
                    chosen_valence = np.nan
                else:
                    valence_str = str(valence)
                    valence_pair = valence_str.split(',') if ',' in valence_str else [valence_str]
                    
                    if len(valence_pair) == 2:
                        chosen_valence = valence_pair[0].strip() if response == 'j' else valence_pair[1].strip()
                    else:
                        chosen_valence = valence_pair[0].strip()
            
            trials_list.append({
                'rt': rt,
                'response': response,
                'chosen_valence': chosen_valence
            })
        
        trials = pd.DataFrame(trials_list)
        
        if len(trials) == 0:
            raise ValueError("No trial data created")
    
    except Exception as e:
        # Create "No data" record for this participant
        new_results.append({
            'ResponseId': participant_id,
            'New_Response_Selection_Score': "No data",
            'New_Prop_Negative_Chosen': "No data",
            'New_Prop_Benign_Chosen': "No data",
            'New_RT_Bias_Index': "No data",
            'New_Mean_RT_Negative': "No data",
            'New_Mean_RT_Benign': "No data",
            'New_N_Trials': 0,
            'New_N_Valid_Trials': 0,
            'New_N_Depression_Chosen': 0,
            'New_N_Anxiety_Chosen': 0,
            'New_N_Positive_Chosen': 0,
            'New_Data_Quality': f"Error: {str(e)}"
        })
        continue
    
    # Filter for valid trials
    valid_choice_trials = trials[~trials['chosen_valence'].isna()]
    valid_rt_trials = trials[(~trials['chosen_valence'].isna()) & (~trials['rt'].isna())]
    
    # Calculate Response Selection Score (uses all trials with valid choices)
    negative_chosen = valid_choice_trials[valid_choice_trials['chosen_valence'].isin(['anxiety', 'depression'])].shape[0]
    benign_chosen = valid_choice_trials[valid_choice_trials['chosen_valence'].isin(['benign', 'positive'])].shape[0]
    
    total_valid_choices = len(valid_choice_trials)
    prop_negative_chosen = negative_chosen / total_valid_choices if total_valid_choices > 0 else np.nan
    prop_benign_chosen = benign_chosen / total_valid_choices if total_valid_choices > 0 else np.nan
    
    response_selection_score = prop_negative_chosen - prop_benign_chosen if not (pd.isna(prop_negative_chosen) or pd.isna(prop_benign_chosen)) else np.nan
    
    # Calculate RT Bias Index (uses only trials with both choice and RT)
    rt_negative = valid_rt_trials[valid_rt_trials['chosen_valence'].isin(['anxiety', 'depression'])]['rt'].mean()
    rt_benign = valid_rt_trials[valid_rt_trials['chosen_valence'].isin(['benign', 'positive'])]['rt'].mean()
    
    rt_bias_index = rt_negative - rt_benign if not (pd.isna(rt_negative) or pd.isna(rt_benign)) else np.nan
    
    # Prepare clean data for DDM (complete trials only)
    ddm_trials = valid_rt_trials.copy()
    ddm_trials['participant_id'] = participant_id
    # For new WSAP, response_binary represents choosing negative vs positive (1=negative)
    ddm_trials['response_binary'] = ddm_trials['chosen_valence'].isin(['anxiety', 'depression']).astype(int)
    
    # Separate by stimulus type for DDM
    depression_trials = ddm_trials[ddm_trials['chosen_valence'] == 'depression']
    anxiety_trials = ddm_trials[ddm_trials['chosen_valence'] == 'anxiety']
    positive_trials = ddm_trials[ddm_trials['chosen_valence'] == 'positive']
    
    # Add to DDM dataset for export
    if len(ddm_trials) > 0:
        new_ddm_data.append(ddm_trials)
    
    new_results.append({
        'ResponseId': participant_id,
        'New_Response_Selection_Score': response_selection_score,
        'New_Prop_Negative_Chosen': prop_negative_chosen,
        'New_Prop_Benign_Chosen': prop_benign_chosen,
        'New_RT_Bias_Index': rt_bias_index,
        'New_Mean_RT_Negative': rt_negative,
        'New_Mean_RT_Benign': rt_benign,
        'New_N_Trials': len(trials),
        'New_N_Valid_Trials': len(valid_rt_trials),
        'New_N_Depression_Chosen': len(depression_trials),
        'New_N_Anxiety_Chosen': len(anxiety_trials),
        'New_N_Positive_Chosen': len(positive_trials),
        'New_Data_Quality': f"Valid: {len(valid_rt_trials)}/{len(trials)} trials"
    })

new_df = pd.DataFrame(new_results)

# ============================================================================
# COMBINE RESULTS
# ============================================================================

combined_df = pd.merge(original_df, new_df, on='ResponseId', how='outer')

# ============================================================================
# CALCULATE SUMMARY STATISTICS
# ============================================================================

# Original WSAP Summary
orig_rss_numeric = pd.to_numeric(original_df['Original_Response_Selection_Score'], errors='coerce').dropna()
orig_rt_numeric = pd.to_numeric(original_df['Original_RT_Bias_Index'], errors='coerce').dropna()

if len(orig_rss_numeric) > 0:
    original_summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            '',
            'Response Selection Score - Mean',
            'Response Selection Score - SD',
            'Response Selection Score - Min',
            'Response Selection Score - Max',
            'Response Selection Score - Median',
            '',
            'RT Bias Index - Mean',
            'RT Bias Index - SD',
            'RT Bias Index - Min',
            'RT Bias Index - Max',
            'RT Bias Index - Median'
        ],
        'Value': [
            str(len(original_df)),
            str(len(orig_rss_numeric)),
            str(len(original_df) - len(orig_rss_numeric)),
            '',
            f"{orig_rss_numeric.mean():.3f}",
            f"{orig_rss_numeric.std():.3f}",
            f"{orig_rss_numeric.min():.3f}",
            f"{orig_rss_numeric.max():.3f}",
            f"{orig_rss_numeric.median():.3f}",
            '',
            f"{orig_rt_numeric.mean():.3f}" if len(orig_rt_numeric) > 0 else 'No data',
            f"{orig_rt_numeric.std():.3f}" if len(orig_rt_numeric) > 0 else 'No data',
            f"{orig_rt_numeric.min():.3f}" if len(orig_rt_numeric) > 0 else 'No data',
            f"{orig_rt_numeric.max():.3f}" if len(orig_rt_numeric) > 0 else 'No data',
            f"{orig_rt_numeric.median():.3f}" if len(orig_rt_numeric) > 0 else 'No data'
        ]
    }
else:
    original_summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(original_df)), '0']
    }

original_summary_df = pd.DataFrame(original_summary_data)

# New WSAP Summary
new_rss_numeric = pd.to_numeric(new_df['New_Response_Selection_Score'], errors='coerce').dropna()
new_rt_numeric = pd.to_numeric(new_df['New_RT_Bias_Index'], errors='coerce').dropna()

if len(new_rss_numeric) > 0:
    new_summary_data = {
        'Metric': [
            'Total Participants',
            'Participants with Valid Data',
            'Participants with Missing Data',
            '',
            'Response Selection Score - Mean',
            'Response Selection Score - SD',
            'Response Selection Score - Min',
            'Response Selection Score - Max',
            'Response Selection Score - Median',
            '',
            'RT Bias Index - Mean',
            'RT Bias Index - SD',
            'RT Bias Index - Min',
            'RT Bias Index - Max',
            'RT Bias Index - Median'
        ],
        'Value': [
            str(len(new_df)),
            str(len(new_rss_numeric)),
            str(len(new_df) - len(new_rss_numeric)),
            '',
            f"{new_rss_numeric.mean():.3f}",
            f"{new_rss_numeric.std():.3f}",
            f"{new_rss_numeric.min():.3f}",
            f"{new_rss_numeric.max():.3f}",
            f"{new_rss_numeric.median():.3f}",
            '',
            f"{new_rt_numeric.mean():.3f}" if len(new_rt_numeric) > 0 else 'No data',
            f"{new_rt_numeric.std():.3f}" if len(new_rt_numeric) > 0 else 'No data',
            f"{new_rt_numeric.min():.3f}" if len(new_rt_numeric) > 0 else 'No data',
            f"{new_rt_numeric.max():.3f}" if len(new_rt_numeric) > 0 else 'No data',
            f"{new_rt_numeric.median():.3f}" if len(new_rt_numeric) > 0 else 'No data'
        ]
    }
else:
    new_summary_data = {
        'Metric': ['Total Participants', 'Participants with Valid Data'],
        'Value': [str(len(new_df)), '0']
    }

new_summary_df = pd.DataFrame(new_summary_data)

# ============================================================================
# EXPORT RESULTS TO EXCEL
# ============================================================================

output_file = os.path.join(results_dir, "wsap_complete_analysis.xlsx")

# Write to Excel with multiple sheets
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Sheet 1: Original WSAP Results
    original_df.to_excel(writer, sheet_name='Original WSAP Results', index=False)

    # Sheet 2: Original WSAP Summary
    original_summary_df.to_excel(writer, sheet_name='Original WSAP Summary', index=False)

    # Sheet 3: New WSAP Results
    new_df.to_excel(writer, sheet_name='New WSAP Results', index=False)

    # Sheet 4: New WSAP Summary
    new_summary_df.to_excel(writer, sheet_name='New WSAP Summary', index=False)

# Export DDM-ready datasets
if original_ddm_data:
    original_ddm_combined = pd.concat(original_ddm_data, ignore_index=True)
    ddm_file = os.path.join(results_dir, "original_wsap_ddm_data.csv")
    original_ddm_combined.to_csv(ddm_file, index=False)

if new_ddm_data:
    new_ddm_combined = pd.concat(new_ddm_data, ignore_index=True)
    ddm_file = os.path.join(results_dir, "new_wsap_ddm_data.csv")
    new_ddm_combined.to_csv(ddm_file, index=False)

# Export data quality report
quality_report = combined_df[['ResponseId', 'Original_Data_Quality', 'New_Data_Quality']].copy()
quality_file = os.path.join(results_dir, "wsap_data_quality_report.csv")
quality_report.to_csv(quality_file, index=False)

print(f"WSAP analysis complete. Results saved to: {output_file}")