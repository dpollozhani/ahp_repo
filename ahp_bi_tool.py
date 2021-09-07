import json
import ahpy
import pandas as pd
from datetime import date, datetime
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

def save_report(level, to_path):
    report = level.report(show=False)
    clean_report = {k:v for k,v in report.items() if k in ['name', 'weights', 'consistency_ratio']}
    if '.json' not in to_path:
        to_path += '.json'
    #pprint(clean_report)
    with open(to_path, 'w') as file:
        json.dump(clean_report, file)


def main():
    file_path = r'G:\Business Intelligence\BI strategy\AHP.xlsx'
    
    overview = load_overview_from_excel(file=file_path, sheet='Overview')
    all_criteria = overview[(overview['Criteria'].str.len()>0) & (overview['Criteria']!='None')]['Criteria'].values
    all_subcriteria = overview[overview['Subcriteria'].str.len()>0]['Subcriteria'].values
    all_parents = overview[overview['Subcriteria parent'].str.len()>0][['Subcriteria', 'Subcriteria parent']].set_index('Subcriteria')
    
    pairwise = load_pairwise_from_excel(file=file_path, sheet='Comparison')
    criteria_base, subcriteria_base, alternatives_base = split_levels(pairwise, 'Hierarchy level')
    
    criteria_comparisons = criteria_to_tuple(criteria_base)
    sub_comparisons = subcriteria_alternatives_to_tuples(subcriteria_base, all_criteria)
    alt_comparison_parents = list(all_subcriteria) + list(all_criteria)
    alt_comparisons = subcriteria_alternatives_to_tuples(alternatives_base, alt_comparison_parents)
    
    ### Calculating the weights and connecting hierarchy ###
    criteria = calculate_weights(criteria_comparisons, 'Criteria') #Top criteria
    
    subcriteria = [] #Subcriteria
    for k,v in sub_comparisons.items():
        subcriteria.append(calculate_weights(v, k))
    
    #Alternatives - connecting to subcriteria and - where applicable - top criteria
    parent_models = subcriteria + [criteria]
    parent_model_names = [s.name for s in subcriteria] + [criteria.name]
    alternatives_tmp, alternatives = [], []
    last_parent_name = all_parents.iloc[0,0]
    max_iter = len(alt_comparisons.keys())
    for x, (k,v) in enumerate(alt_comparisons.items()):
        current_parent_name = all_parents.at[k, 'Subcriteria parent']
        i = parent_model_names.index(current_parent_name)
        j = parent_model_names.index(last_parent_name)

        if parent_models[i].name != last_parent_name:
            parent_models[j].add_children(alternatives_tmp)
            alternatives_tmp=[]
        
        weights = calculate_weights(v, k)
        alternatives_tmp.append(weights)
        alternatives.append(weights)
        
        if x==max_iter-1:
            if current_parent_name!='Criteria':
                parent_models[i].add_children(alternatives_tmp)
            else:
                criteria_children = subcriteria+alternatives_tmp
                criteria.add_children(criteria_children)
        #print(k, ";", current_parent_name)

        last_parent_name = current_parent_name

    #criteria.report(show=True)
    file_id = int(datetime.today().timestamp())
    save_report(criteria, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Top_level_report_{file_id}')
    for s in subcriteria:
        save_report(s, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Subcriteria_level_report_{s.name}_{file_id}')
    for a in alternatives:
        save_report(a, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Alternatives_level_report_{a.name}_{file_id}')

if __name__ == '__main__':
    main()