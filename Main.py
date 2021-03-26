import pandas as pd
import numpy as np
import openpyxl as pyxl

from ExcelLoad import Load_Table_From_Excel, Clean_Table
from ProduitCartésien import select_avant_realloc, cartesian_product_multi

folder = "C:/Users/admin/Documents/Pilotage LBP"
path_referentiel = folder + "/Dictionnaire de données - 024.xlsx"
path_excel = folder + "/Cartographie - PPH - 002.xlsx"

referentiel_address_dict = {'Indicateurs_Bilan': ["Indicateurs - Bilan comptable", "D4:H"],
                            'Indicateurs_PL': ["Indicateurs - P&L", "D4:H"],
                            'Indicateurs_Activite': ["Indicateurs - Activité", "D4:H"],
                            'Indicateurs_Risques': ["Indicateurs - Risques", "D4:H"],
                            'Produit': ["Produits", "D4:H"],
                            'Entité juridique': ["EJ & UG", "D4:G"],
                            'Unité de gestion / Direction': ["EJ & UG", "I4:N"],
                            'Distribution': ["EJ & UG", "P4:R"],
                            'Segment client / marché': ["EJ & UG", "T4:X"]}


# Load référentiel ----------------------------------------------------------------------------------------------------
wb = pyxl.load_workbook(path_referentiel, read_only=True, keep_vba=False, data_only=True, keep_links=False)
Referentiel = Load_Table_From_Excel(referentiel_address_dict, wb, bFirstRowasHeader=True)
wb.close()

Referentiel['Indicateur'] = pd.concat([Referentiel['Indicateurs_Bilan'],
                                       Referentiel['Indicateurs_PL'],
                                       Referentiel['Indicateurs_Activite'],
                                       Referentiel['Indicateurs_Risques']])

Referentiel['Unité de gestion après réallocation'] = Referentiel['Unité de gestion / Direction']


# Load Tableau de bord ------------------------------------------------------------------------------------------------
wb = pyxl.load_workbook(path_excel, read_only=True, keep_vba=False, data_only=True, keep_links=False)

table_address_dict = dict()
for name in wb.sheetnames:
    if ((not name[0].isdigit()) and name[1] == "." and name[2].isdigit()):
        # if name == "C.1.1.a. PPH dashboard":
        table_address_dict[name] = [name, ""]

table_address_dict["Ratios"] = ["Ratios", "A1:"]
table_address_dict["Exclusions"] = ["Exclusions", "A1:"]

Table = Load_Table_From_Excel(table_address_dict, wb)
wb.close()

# -------------------------------------------------------------------------------------------------------------------- #
# 1ere étape - Rassembler le tout en une seule
# -------------------------------------------------------------------------------------------------------------------- #
Out = pd.DataFrame()
for table_name in Table.keys():
    if table_name not in ["Ratios", "Exclusions"]:
        print(table_name)
        Out = pd.concat([Out, Clean_Table(Table[table_name], table_name)])

Out.columns = ["Entité juridique", "Unité de gestion / Direction", "Unité de gestion après réallocation", "Segment client / marché", "Distribution", "Produit", "Indicateur", "Type d'indicateur", "Phase", "Adresse"]

Out.to_csv(folder + "/Carto_out.csv", encoding='utf-8-sig', index=False, sep=";")
Out_original = Out.copy()

# -------------------------------------------------------------------------------------------------------------------- #
# 2eme étape - Merger avec le référentiel pour décomposer en indicateurs élémentaires
# -------------------------------------------------------------------------------------------------------------------- #

# Ne faire qu'une colonne sur l'unité de gestion : si avant réallocation non vide alors prendre celle-là
Out["Unité de gestion après réallocation"] = Out.apply(lambda row: select_avant_realloc(row), axis=1)
Out.drop(["Unité de gestion / Direction"], axis=1, inplace=True)

Ref_columns = [col for sublist in [Referentiel[dim].columns.tolist() for dim in Out.columns[0:6]] for col in sublist]

# On éclate en indicateurs élémentaires
Out_eclate = dict()
Out_iter = Out[Out.columns[0:7]].drop_duplicates().reset_index(drop=True)
for index, row in Out_iter.iterrows():
    print(index)
    row_ = pd.DataFrame(row[:-1]).transpose()
    Out_ = row_.copy()
    if not row["Type d'indicateur"] in ["Ratio", "Total"]:
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
        # else:
        #     Out_[Referentiel[dimension].columns] = np.nan

        Out_.columns = [col for col in row_.columns.tolist() if col != 'key'] + Ref_columns
        Out_eclate[index] = Out_

# On rassemble les bases éclatées
Base = pd.concat([v for k, v in Out_eclate.items()], axis=0)

# Vision plus ramassée
col_not0 = [col for col in Base.columns.tolist() if ((not col.endswith("N0")) and (col not in ["Segment N2", "Segment N1"]))]
Base_not0 = Base[col_not0].drop_duplicates()

# -------------------------------------------------------------------------------------------------------------------- #
# 3eme étape - On rassemble avec la base initiale sans certains niveaux
# -------------------------------------------------------------------------------------------------------------------- #
Base_complete = pd.merge(Base_not0, Out_original[~Out_original["Type d'indicateur"].isin(["Ratio", "Total"])],
                         on=Base_not0.columns[0:6].tolist(), how='left')
Base_complete = pd.concat([Base_complete, Out_original[Out_original["Type d'indicateur"].isin(["Ratio", "Total"])]])
Base_complete = Base_complete.drop(["Phase"], axis=1).reset_index(drop=True)

Base_complete = Base_complete.astype(object).replace(np.nan, 'None')
Base_complete = Base_complete.astype(str)

# Version où l'on rassemble les adresses ensemble
Base_dimensions = Base_complete.groupby([col for col in Base_complete.columns if col != "Adresse"])['Adresse'].apply('/'.join).reset_index()
Base_dimensions.replace("None", np.nan, inplace=True)


# On exclut certaines lignes
for index, row in Table["Exclusions"].iterrows():
    row_ = pd.DataFrame(row).transpose().dropna(axis=1, how='all')

Base_dimensions.to_csv(folder + "/Base_dimensions.csv", encoding='utf-8-sig', index=False, sep=",")

