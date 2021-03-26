import pandas as pd
import numpy as np


def cartesian_product(*arrays):
    la = len(arrays)
    dtype = np.result_type(*arrays)
    arr = np.empty([len(a) for a in arrays] + [la], dtype=dtype)
    for i, a in enumerate(np.ix_(*arrays)):
        arr[..., i] = a
    return arr.reshape(-1, la)


def cartesian_product_multi(*dfs):
    idx = cartesian_product(*[np.ogrid[:len(df)] for df in dfs])
    return pd.DataFrame(
        np.column_stack([df.values[idx[:, i]] for i, df in enumerate(dfs)]))


def select_avant_realloc(row):
    if not row["Unité de gestion / Direction"] in ['nan', np.nan, 'None', None]:
        return row["Unité de gestion / Direction"]
    else:
        return row["Unité de gestion après réallocation"]
