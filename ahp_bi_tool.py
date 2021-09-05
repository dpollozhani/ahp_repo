import ahpy
import pandas as pd
from pprint import pprint

def load_overview_from_excel(file, sheet):
    df = pd.read_excel(io=file, sheet_name=sheet)
    return df

def load_pairwise_from_excel(file, sheet):
    df = pd.read_excel(io=file, sheet_name=sheet)
    return df

def split_levels(df, level_col):
    criteria = df[df[level_col] == 'Criteria']
    subcriteria = df[df[level_col] == 'Subcriteria']
    alternatives = df[df[level_col] == 'Alternatives']
    return criteria, subcriteria, alternatives

def criteria_to_tuple(df):
    comparisons = {}
    for _, row in df.iterrows():
        comparisons[(row['A'], row['B'])] = row['Relative importance']
    return comparisons

def subcriteria_alternatives_to_tuples(df, all_criteria):
    comparisons = {}
    for c in all_criteria:
        if c in df['Parent'].values:
            s_df = df[df['Parent'] == c].iterrows()
            comparisons[c] = {(row['A'],row['B']):row['Relative importance'] for _, row in s_df}
    return comparisons

def calculate_weights(level, comparison_name):
    weights = ahpy.Compare(name=comparison_name, comparisons=level)
    return weights

if __name__ == '__main__':
    file_path = 'data/AHP test.xlsx'
    
    overview = load_overview_from_excel(file=file_path, sheet='Overview')
    all_criteria = overview[(overview['Criteria'].str.len()>0) & (overview['Criteria']!='None')]['Criteria'].values
    #all_subcriteria = overview[overview['Subcriteria'].str.len()>0]['Subcriteria'].values
    #all_alternatives = overview[overview['Alternatives'].str.len()>0]['Alternatives'].values

    pairwise = load_pairwise_from_excel(file=file_path, sheet='Comparison')
    criteria_base, subcriteria_base, alternatives_base = split_levels(pairwise, 'Hierarchy level')
    
    criteria_comparisons = criteria_to_tuple(criteria_base)
    sub_comparisons = subcriteria_alternatives_to_tuples(subcriteria_base, all_criteria)
    #alt_comparisons = subcriteria_alternatives_to_tuples(alternatives_base, all_subcriteria)

    ### Calculating the weights ###
    criteria = calculate_weights(criteria_comparisons, 'Criteria')
    
    subcriteria = []
    for k,v in sub_comparisons.items():
        subcriteria.append(calculate_weights(v, k))

    criteria.add_children(subcriteria)

    # alternatives = []
    # last_subcriteria = subcriteria[0]
    # for i, (k,v) in enumerate(alt_comparisons.items()):
    #     current_subcriteria = subcriteria[i]
    #     if current_subcriteria != last_subcriteria:
    #         last_subcriteria.add_children(alternatives)
    #         alternatives = []
    #     alternatives.append(calculate_weights(v, k))
    #     last_subcriteria = current_subcriteria

    criteria.report(show=True)