import pandas as pd
import numpy as np
import openpyxl as pyxl
import warnings
import re
import string

folder = "C:/Users/admin/Documents/Pilotage LBP"
path_excel = "C:/Users/admin/Documents/Pilotage LBP/Consolidation 017 - test.xlsx"
path_excel = "C:/Users/admin/Documents/Pilotage LBP/Exemple - Dashboard PPH.xlsx"
folder = "C:/Users/admin/Documents/ST EBA ALM 2021/Template"
path_excel = folder + "/Template Conso v120.xlsm"


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


# Load workbook
wb = pyxl.load_workbook(path_excel, read_only=True, keep_vba=False, data_only=True, keep_links=False)
#
# table_address_dict = dict()
# for name in wb.sheetnames:
#     # if name[0].isdigit():
#     if name == "C.1.1.a. PPH dashboard":
#         table_address_dict[name] = [name, ""]
#
# Table = Load_Table_From_Excel(table_address_dict, wb)
# wb.close()
#
#
# Out = pd.DataFrame()
# for table_name in Table.keys():
#     print(table_name)
#     Out = pd.concat([Out, Clean_Table(Table[table_name], table_name)])
#
# Out.columns = ["Entit?? juridique", "Unit?? de gestion / Direction", "Unit?? de gestion apr??s r??allocation", "Segment client / march??", "Distribution", "Produit", "Indicateur", "Type d'indicateur", "Phase", "Adresse"]
#
# Out.to_csv(folder + "/Carto_out.csv", encoding='utf-8-sig', index=False)


# Template EBA load
table_address_dict = dict()
table_address_dict["CSV_NII_CALC"] = ["CSV_NII_CALC", "B3:JJ848"]

NII_calc = Load_Table_From_Excel(table_address_dict, wb, bFirstRowasHeader =True)["CSV_NII_CALC"]
wb.close()

# Index rows
col_name_row = NII_calc.iloc[11][0:8]
NII_calc.columns = col_name_row.tolist() + NII_calc.columns.tolist()[len(col_name_row):]

NII_calc = NII_calc[[col for col in NII_calc.columns.tolist() if col != None]]

# Index mapping columns
mapping_col = NII_calc.iloc[0:6, 8:].T.reset_index()
mapping_col.columns = ["variable", "Fixed/Floating", "Actual/Proj", "Performing/NPE", "Time Bucket", "Indicator", "SubIndicator"]

# Melt and map columns
NII_calc = NII_calc.iloc[12:].melt(id_vars=col_name_row.tolist())
NII_calc = NII_calc.merge(mapping_col, how="left", on="variable")

# Output
NII_calc.rename(columns={'variable':'ColNum'}, inplace=True)
NII_calc = NII_calc[[c for c in NII_calc if c != "value"] + ["value"]]

NII_calc.to_csv(folder + "Template_NII_Calc_CSV.csv")
