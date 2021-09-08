import ahpy
import pandas as pd
from datetime import datetime
import json
from pprint import pprint

def load_comparisons(file_name, sheet_name) -> dict:
    "Load symmetric matrices of comparisons and return as dict of tuples"
    df = df = pd.read_excel(io=file_name, sheet_name=sheet_name)
    cols_to_use = [c for c in df.columns if any([isinstance(c, float), isinstance(c, int)]) or "Partition" in c]
    df = df[cols_to_use]
    df = df[df["Partition id"] > 0]
    comparisons = matrix_to_tuples(df)
    return comparisons

def matrix_to_tuples(df: pd.DataFrame) -> dict:
    "Return dict of tuples with pairwise criteria/alternatives and each weight as keys"
    comparisons = {}
    for i in set(df["Partition id"].values):
        df_part_base = df[df["Partition id"]==i]
        partition = df_part_base["Partition name"].values[0]
        first_n_cols = list(df_part_base.columns).index("Partition name")
        new_index = df_part_base.iloc[0,:first_n_cols].values
        df_part = df_part_base.iloc[1:,:first_n_cols]
        df_part.index, df_part.columns = new_index, new_index
        comparisons_base = df_part.unstack().reset_index()
        comparisons[partition]  = {(row.loc['level_1'], row.loc['level_0']): row.loc[0] for _,row in comparisons_base.iterrows() if row.loc[0] > 0 and row.loc['level_1'] != row.loc['level_0']}
    return comparisons

def calculate_weights(comparisons, comparison_name):
    weights = ahpy.Compare(name=comparison_name, comparisons=comparisons)
    return weights

def save_report(level, to_path):
    report = level.report(show=False)
    clean_report = {k:v for k,v in report.items() if k in ['name', 'weights', 'consistency_ratio']}
    if '.json' not in to_path:
        to_path += '.json'
    with open(to_path, 'w') as file:
        json.dump(clean_report, file)

def main():
    file_path = r'G:\Business Intelligence\BI strategy\AHP.xlsx'
    overview = df = pd.read_excel(io=file_path, sheet_name='Overview')

    all_parents = overview[overview['Subcriteria parent'].str.len()>0][['Subcriteria', 'Subcriteria parent']].set_index('Subcriteria')
    
    criteria_comparisons = load_comparisons(file_path, sheet_name='Criteria comparisons') 
    subcriteria_comparisons = load_comparisons(file_path, sheet_name='Subcriteria comparisons') 
    alternatives_comparisons = load_comparisons(file_path, sheet_name='Alternative comparisons') 
    
    ### Calculating the weights and connecting hierarchy ###
    criteria = calculate_weights(criteria_comparisons['Criteria'], 'Criteria') #Top criteria
    
    subcriteria = [] #Subcriteria
    for k,v in subcriteria_comparisons.items():
        subcriteria.append(calculate_weights(v, k))
    
    #Alternatives - connecting to subcriteria and - where applicable - top criteria
    parent_models = subcriteria + [criteria]
    parent_model_names = [p.name for p in parent_models] 
    alternatives_tmp, alternatives = [], []
    last_parent_name = all_parents.iloc[0,0]
    max_iter = len(alternatives_comparisons.keys())-1
    for x, (k,v) in enumerate(alternatives_comparisons.items()):
        current_parent_name = all_parents.at[k, 'Subcriteria parent']
        current_parent = parent_model_names.index(current_parent_name)
        last_parent = parent_model_names.index(last_parent_name)

        if parent_models[current_parent].name != last_parent_name:
            parent_models[last_parent].add_children(alternatives_tmp)
            alternatives_tmp=[]
        
        weights = calculate_weights(v, k)
        alternatives_tmp.append(weights) #used only temporarily for adding children to top/subcriteria
        alternatives.append(weights) #used for saving all weight data for reporting
        
        if x==max_iter:
            if current_parent_name != 'Criteria':
                parent_models[current_parent].add_children(alternatives_tmp)
            else:
                criteria_children = subcriteria+alternatives_tmp #since children can only be added once together, we add everything to top criteria in the end
                criteria.add_children(criteria_children)

        last_parent_name = current_parent_name

    file_id = int(datetime.today().timestamp())
    save_report(criteria, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Top_level_report_{file_id}')
    for s in subcriteria:
        save_report(s, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Subcriteria_level_report_{s.name}_{file_id}')
    for a in alternatives:
        save_report(a, to_path=f'G:\Business Intelligence\BI strategy\AHP reports\Alternatives_level_report_{a.name}_{file_id}')

if __name__ == '__main__':
    main()
    