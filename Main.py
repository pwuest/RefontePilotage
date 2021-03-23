import pandas as pd
import numpy as np
import openpyxl as pyxl
import warnings
import re
import string

folder = "C:/Users/admin/Documents/Pilotage LBP"
path_referentiel = folder + "/Dictionnaire de données - 019.xlsx"
path_excel = folder + "/Exemple - Dashboard PPH - 009.xlsx"

referentiel_address_dict = {'Indicateurs_Bilan': ["Indicateurs - Bilan comptable", "D4:H"],
                            'Indicateurs_PL': ["Indicateurs - P&L", "D4:H"],
                            'Indicateurs_Activite': ["Indicateurs - Activité", "D4:H"],
                            'Indicateurs_Risques': ["Indicateurs - Risques", "D4:H"],
                            'Produit': ["Produits", "D4:H"],
                            'Entité juridique': ["EJ & UG", "F4:I"],
                            'Unité de gestion / Direction': ["EJ & UG", "K4:P"],
                            'Distribution': ["EJ & UG", "R4:T"],
                            'Segment client / marché': ["EJ & UG", "V4:Z"]}


# %% Load from ws excel 'range_string' range in the form of 'A1:B2' or 'A:B' or ''
def Load_Workbook_Range(range_string, ws, bFirstRowasHeader = False):
    # Reformat input range string
    if not (re.match('[A-Z]+:[A-Z]+', range_string)) is None:
        a, b = range_string.split(':')
        range_string = a + '1:' + b + str(ws.max_row)
    elif not (re.match('[A-Z]+[0-9]+:[A-Z]+$', range_string)) is None:
        a, b = range_string.split(':')
        range_string = a + ':' + b + str(ws.max_row)

    # Load data
    data_rows = []
    with warnings.catch_warnings():
        # Openpyxl does not recognize Data Validation extension in excel files
        warnings.simplefilter('ignore')
        if not (re.match('[A-Z]+[0-9]+:[A-Z]+[0-9]+', range_string)) is None:
            for row in ws[range_string]:
                data_rows.append([cell.value for cell in row])
        else:
            for row in ws.iter_rows():
                data_rows.append([cell.value for cell in row])

    if len(data_rows) > 0:
        if bFirstRowasHeader:
            return pd.DataFrame(data_rows[1:], columns=data_rows[0])
        else:
            return pd.DataFrame(data_rows[0:])
    else:
        return None


# %% For each table in parameter load the input range from input sheet
def Load_Table_From_Excel(table_dict, workbook_excel, bFirstRowasHeader = False):
    # Basic load from excel and return dictionary
    tables_output = dict()
    for table_name in table_dict.keys():
        sheet = workbook_excel[table_dict[table_name][0]]
        rng = table_dict[table_name][1]

        tables_output[table_name] = Load_Workbook_Range(rng, sheet, bFirstRowasHeader)

    return tables_output


def Clean_Table(table, table_name):
    if table is None:
        return None

    table = table[list(table)[0:300]]  # limit the number of columns

    # Rename columns like excel
    letters = list(string.ascii_uppercase)
    excel_letters = letters + [x + y for x in letters for y in letters]
    table.columns = excel_letters[0:table.shape[1]]

    table = table.dropna(how="all").dropna(how="all", axis=1)
    table = table.head(10000)
    table = table.astype(object).replace(np.nan, 'None')
    table = table.astype(str)

    # Loop table and write address at intersection where there is a line indicator and column is not empty at line 6
    n = 9
    table_col = list(table)
    N = table.shape[0]
    for col in table_col[n:]:
        for i in list(range(n, N)):
            indicators = table[table_col[0:n-1]].iloc[i]
            if (table[col][n-1] != 'None') and not((indicators == 'None').all()) and table[col].iloc[i] != "" \
                    and table[col].iloc[i] != None and table[col].iloc[i] != 'None':
                table[col].iloc[i] = "'" + table_name + "'!$" + col + "$" + str(int(table.iloc[i].name) + 1)
            else:
                table[col].iloc[i] = 'None'

    # Now select only those cells that have been marked with an address in previous loop
    col_select = [c for c in table if table[c].str.contains('\$').any()]
    row_select = table.select_dtypes(include=[object]).applymap(lambda x: '$' in x if pd.notnull(x) else False)

    if (len(col_select) > 0) and (len(row_select[row_select.any(axis=1)])):
        # Row
        temp = table.iloc[:, 0:n-1].copy()
        temp["ColNew"] = "None"
        KPI_rows = pd.concat([temp, table[col_select]], axis=1)
        KPI_rows = KPI_rows[row_select.any(axis=1)]
        KPI_rows = KPI_rows.melt(id_vars=list(KPI_rows.columns[0:n])).drop(["variable"], axis=1)

        # Column
        temp = table[col_select]
        KPI_columns = pd.concat([temp.iloc[0:n].transpose().reset_index(drop=True),
                                 temp[row_select.any(axis=1)].transpose().reset_index(drop=True)], axis=1)
        KPI_columns = KPI_columns.melt(id_vars=list(KPI_columns.columns[0:n])).drop(["variable"], axis=1)
        KPI_columns.columns = KPI_rows.columns

        # Output : merge and ffill and remove duplicates
        Out = pd.concat([KPI_rows, KPI_columns], axis=0)
        Out.replace("None", np.nan, inplace=True)
        Out.dropna(subset=["value"], inplace=True)
        temp = Out.groupby("value").apply(lambda x: x.fillna(method='ffill'))
        Out = temp.dropna(subset=["ColNew"])

        return Out

    else:
        return None


# Load data dictionnary
wb = pyxl.load_workbook(path_referentiel, read_only=True, keep_vba=False, data_only=True, keep_links=False)
Referentiel = Load_Table_From_Excel(referentiel_address_dict, wb, bFirstRowasHeader=True)
wb.close()

Referentiel['Indicateur'] = pd.concat([Referentiel['Indicateurs_Bilan'],
                                       Referentiel['Indicateurs_PL'],
                                       Referentiel['Indicateurs_Activite'],
                                       Referentiel['Indicateurs_Risques']])

Referentiel['Unité de gestion après réallocation'] = Referentiel['Unité de gestion / Direction']


# Load workbook
wb = pyxl.load_workbook(path_excel, read_only=True, keep_vba=False, data_only=True, keep_links=False)

table_address_dict = dict()
for name in wb.sheetnames:
    if ((not name[0].isdigit()) and name[1] == "." and name[2].isdigit()):
    # if name == "C.1.1.a. PPH dashboard":
        table_address_dict[name] = [name, ""]

Table = Load_Table_From_Excel(table_address_dict, wb)
wb.close()


Out = pd.DataFrame()
for table_name in Table.keys():
    print(table_name)
    Out = pd.concat([Out, Clean_Table(Table[table_name], table_name)])

Out.columns = ["Entité juridique", "Unité de gestion / Direction", "Unité de gestion après réallocation", "Segment client / marché", "Distribution", "Produit", "Indicateur", "Type d'indicateur", "Phase", "Adresse"]

Out.to_csv(folder + "/Carto_out.csv", encoding='utf-8-sig', index=False)


def cartesian_product(*arrays):
    la = len(arrays)
    dtype = np.result_type(*arrays)
    arr = np.empty([len(a) for a in arrays] + [la], dtype=dtype)
    for i, a in enumerate(np.ix_(*arrays)):
        arr[...,i] = a
    return arr.reshape(-1, la)

def cartesian_product_multi(*dfs):
    idx = cartesian_product(*[np.ogrid[:len(df)] for df in dfs])
    return pd.DataFrame(
        np.column_stack([df.values[idx[:,i]] for i,df in enumerate(dfs)]))


Out_eclate = dict()
Out_columns = [col for sublist in [Referentiel[dim].columns.tolist() for dim in Out.columns[0:7]] for col in sublist]
Out_iter = Out[Out.columns[0:7]].drop_duplicates().reset_index(drop=True)
for index, row in Out_iter.iterrows():
    print(index)
    row_ = pd.DataFrame(row).transpose()
    Out_ = row_
    for dimension in row_.columns:
        item = row[dimension]

        if item == 'Tous':
            match_ = Referentiel[dimension]
        else:
            match_ = Referentiel[dimension][Referentiel[dimension].isin([item]).any(axis=1)].copy()

        if match_.shape[0] == 0 or Out_.shape[0] == 0:
            # Cartesian product
            Out_['key'] = 0
            match_['key'] = 0
            Out_ = Out_.merge(match_, on='key', how='outer').drop(columns=['key'])
        else:
            Out_ = cartesian_product_multi(*[Out_, match_])

    Out_.columns = [col for col in row_.columns.tolist() if col != 'key'] + Out_columns
    Out_eclate[index] = Out_

test = pd.concat([v for k, v in Out_eclate.items()], axis=0)


