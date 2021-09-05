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

def generate_tabular_report(criteria, subcriteria, alternatives):
    pass

def main():
    file_path = 'data/AHP test.xlsx'
    
    overview = load_overview_from_excel(file=file_path, sheet='Overview')
    all_criteria = overview[(overview['Criteria'].str.len()>0) & (overview['Criteria']!='None')]['Criteria'].values
    all_subcriteria = overview[overview['Subcriteria'].str.len()>0]['Subcriteria'].values
    all_subcriteria_parents = overview[overview['Subcriteria parent'].str.len()>0][['Subcriteria', 'Subcriteria parent']].set_index('Subcriteria')

    pairwise = load_pairwise_from_excel(file=file_path, sheet='Comparison')
    criteria_base, subcriteria_base, alternatives_base = split_levels(pairwise, 'Hierarchy level')
    
    criteria_comparisons = criteria_to_tuple(criteria_base)
    sub_comparisons = subcriteria_alternatives_to_tuples(subcriteria_base, all_criteria)
    alt_comparisons = subcriteria_alternatives_to_tuples(alternatives_base, all_subcriteria)
    
    ### Calculating the weights and connecting hierarchy ###
    criteria = calculate_weights(criteria_comparisons, 'Criteria')
    
    subcriteria = []
    for k,v in sub_comparisons.items():
        subcriteria.append(calculate_weights(v, k))

    criteria.add_children(subcriteria)
   
    alternatives_tmp, alternatives = [], []
    last_subcriteria_name = all_subcriteria_parents.iloc[0,0]
    subcriteria_model_names = [s.name for s in subcriteria]
    max_iter = len(alt_comparisons.keys())
    for x, (k,v) in enumerate(alt_comparisons.items()):
        current_subcriteria_name = all_subcriteria_parents.at[k, 'Subcriteria parent']
        i = subcriteria_model_names.index(current_subcriteria_name)
        j = subcriteria_model_names.index(last_subcriteria_name)
    
        if subcriteria[i].name != last_subcriteria_name:
            subcriteria[j].add_children(alternatives_tmp)
            #print([x.name for x in alternatives_tmp])
            #print("->", subcriteria[j].name) 
            alternatives_tmp=[]
        
        weights = calculate_weights(v, k)
        alternatives_tmp.append(weights)
        alternatives.append(weights)
        
        if x==max_iter-1:
            subcriteria[i].add_children(alternatives_tmp)
            #print([x.name for x in alternatives_tmp])
            #print("->", subcriteria[j].name) 
      
        last_subcriteria_name = current_subcriteria_name
        
    criteria.report(show=True)

if __name__ == '__main__':
    main()