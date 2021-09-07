import ahpy
import pandas as pd
from pprint import pprint

def matrix_to_tuples(df: pd.DataFrame) -> dict:
    "Return dict of tuples with pairwise criteria/alternatives and each weight as keys"
    first_n_cols = list(df.columns).index("Partition name")
    new_index = df.iloc[0,:first_n_cols].values
    comparisons = {}
    for i in set(df["Partition id"].values):
        df_part_base = df[df["Partition id"]==i]
        partition = df_part_base["Partition name"].values[0]
        df_part = df_part_base.iloc[1:,:first_n_cols]
        df_part.index, df_part.columns = new_index, new_index
        comparisons_base = df_part.unstack().reset_index()
        comparisons[partition]  = {(row.loc['level_0'], row.loc['level_1']): row.loc[0] for _,row in comparisons_base.iterrows() if row.loc[0] > 0}
    return comparisons

def load_comparisons(file, sheet_name='Test') -> dict:
    "Load symmetric matrices of comparisons and return as dict of tuples"
    df = pd.read_excel(io=file, sheet_name='Test')
    cols_to_use = [c for c in df.columns if isinstance(c, float) or "Partition" in c]
    df = df[cols_to_use]
    df = df[df["Partition id"] > 0]
    comparisons = matrix_to_tuples(df)
    return comparisons

def calculate_weights(comparisons, comparison_name):
    weights = ahpy.Compare(name=comparison_name, comparisons=comparisons)
    return weights

if __name__ == '__main__':
    file_path = 'data/AHP test.xlsx'
    alternatives = load_comparisons(file_path)
    pprint(alternatives)
   
