import datetime
import os
import shutil

import numpy as np
import pandas
import pandas as pd
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


# Fournis par : Laboratoire
path_Mapping = "0-Input/GMM_Mapping.xlsx"  # Fichier de mapping
path_GMMUpdate = "0-Input/GMM_Update.xlsx"  # Fichier d'update
# Liste des équipements maîtres
path_MastEqt = "0-Input/MasterEquipments.xlsx"
# Fournis par : Steeven
path_GMM1 = "0-Input/GMM.csv"  # Fichier général (1)
path_GMM2 = "0-Input/GMM - Instruments Plastic Omnium Alphatech.csv"  # Fichier général (2)
path_SubEqt = "0-Input/GMM - Instruments liés.csv"  # Instruments liés
# Fournis par : Deltamu
path_Inter = "0-Input/PlasticOmnium.xls"  # Liste des interventions
path_InterDocs = "0-Input/PLAST-60/"  # Fichiers

# Liste des chemins de sortie
path_Export = "1-Output/Export_Equipment.xlsm"
path_Interventions = "1-Output/Export_Interventions.xlsm"
path_Derogations = "1-Output/Export_Derogations.xlsx"
path_Category = "1-Output/Export_Category.xlsx"
path_StorageArea = "1-Output/Export_StorageArea.xlsx"

path_Analyze = "1-Output/Analyze.xlsx"
path_Export_Template = r'1-Output/0-Template/2023-06-07_15-25-50_Equipments.xlsm'
path_Export_Calibration_Template = r'1-Output/0-Template/2023-06-07_15-34-35_Calibrations.xlsm'

# Suppression des fichiers d'export existants
if os.path.exists(path_Export):
    os.remove(path_Export)
if os.path.exists(path_Analyze):
    os.remove(path_Analyze)
if os.path.exists(path_Category):
    os.remove(path_Category)
if os.path.exists(path_StorageArea):
    os.remove(path_StorageArea)
if os.path.exists(path_Interventions):
    os.remove(path_Interventions)
if os.path.exists(path_Derogations):
    os.remove(path_Derogations)

# Création du dataframe : mapping(CalibrationStatus)
df_Mapping_CalibrationStatus = pandas.read_excel(path_Mapping, sheet_name="CalibrationStatus")
df_Mapping_CalibrationStatus = df_Mapping_CalibrationStatus.reset_index()

# Création du dataframe : mapping(InterventionType)
df_Mapping_InterventionType = pandas.read_excel(path_Mapping, sheet_name="InterventionType")
df_Mapping_InterventionType = df_Mapping_InterventionType.reset_index()

# Création du dataframe : mapping(Status)
df_Mapping_Status = pandas.read_excel(path_Mapping, sheet_name="Status")
df_Mapping_Status = df_Mapping_Status.reset_index()

# Création du dataframe : mapping(CalibrationType_Supplier)
df_Mapping_CalibrationSupplier = pandas.read_excel(path_Mapping, sheet_name="CalibrationType_Supplier")
df_Mapping_CalibrationSupplier = df_Mapping_CalibrationSupplier.reset_index()

# Création du dataframe : mapping(EquipmentType_Chain)
df_Mapping_EquipmentType = pandas.read_excel(path_Mapping, sheet_name="EquipmentType_Chain")
df_Mapping_EquipmentType = df_Mapping_EquipmentType.reset_index()

# Création du dataframe : mapping(GMM)
df_Mapping_Update = pandas.read_excel(path_GMMUpdate, sheet_name="GMM")
df_Mapping_Update = df_Mapping_Update.reset_index()

# Création de la liste des équipements
list_to_keep = ['Identification', 'Gamme', 'Tolérance', 'Avis',
                "Date d'intervention", "Prochaine date d'intervention", 'Périodicité']
df_GMM2 = pd.read_csv(path_GMM2, sep=';', usecols=list_to_keep)

list_to_keep = ['Identification', 'Domaine', 'Désignation', "Etat d'utilisation", 'Statut', 'Localisation']

df_GMM1 = pd.read_csv(path_GMM1, sep=';', encoding="ISO-8859-1", usecols=list_to_keep)
df_GMM1.rename(columns={'Domaine': 'Domaine de mesure'}, inplace=True)

df_GMM = pd.merge(df_GMM1, df_GMM2, on='Identification')  # Fusion des 2 dataframes
# endregion

# A : Remplacement des valeurs d'OPTIMU avec le travail d'épuration réalisé par le laboratoire *27/01/2023
# region A-01 Remplacement: Désignation > Equipment name
# Création de la colonne 'Equipment name' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Equipment name'].to_dict()
df_GMM['Equipment name'] = df_GMM['Identification'].replace(replace_dict)

for new_itr, select_row in df_GMM.iterrows():  # Comparaison des colonnes 'Désignation' et 'Equipment name'
    if pd.isnull(select_row["Equipment name"]):
        df_GMM.at[new_itr, 'Equipment name'] = str(select_row["Désignation"])
# endregion
# region A-02 Ajout: Equipment category
# Création de la colonne 'Equipment category' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Equipment category'].to_dict()
df_GMM['Equipment category'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-03 Remplacement: Fabricant > Manufacturer
# Création de la colonne 'Manufacturer' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Fabricant'].to_dict()
df_GMM['Manufacturer'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-04 Remplacement: Modèle > Model
# Création de la colonne 'Model' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Model'].to_dict()
df_GMM['Model'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-05 Remplacement: N° de série > Serial number
# Création de la colonne 'Serial number' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Serial number'].to_dict()
df_GMM['Serial number'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-06 Remplacement: Gestionnaire > Affected person
# Création de la colonne 'Affected perso' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Affected person'].to_dict()
df_GMM['Affected person'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-07 Remplacement: Localisation > Storage area
# Création de la colonne 'Storage area' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Storage area'].to_dict()
df_GMM['Storage area'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-08 Ajout: Intermediate controls
# Création de la colonne 'Intermediate controls' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Intermediate controls'].to_dict()
df_GMM['Intermediate controls'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-09 Ajout: Number of intermediate controls
# Création de la colonne 'Number of intermediate controls' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Number of controls'].to_dict()
df_GMM['Number of intermediate controls'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-10 Ajout: Est. Downtime
# Création de la colonne 'Est. Downtime' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Downtime'].to_dict()
df_GMM['Est. downtime'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-11 Ajout: Est. Calibration time
# Création de la colonne 'Est. Calibration time' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Calibration time'].to_dict()
df_GMM['Est. calibration time'] = df_GMM['Identification'].replace(replace_dict)
# endregion
# region A-12 Ajout: Est. Cost
# Création de la colonne 'Est. Cost' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Cost'].to_dict()
df_GMM['Est. cost'] = df_GMM['Identification'].replace(replace_dict)
# Création de la colonne 'Est. Cost' dans GMM
replace_dict = df_Mapping_Update.set_index('Identification')['Code'].to_dict()
df_GMM['Business group'] = df_GMM['Identification'].replace(replace_dict)
# endregion

df_GMM['Plannable'] = 'Yes'
replace_dict = df_Mapping_Update.set_index('Identification')['Plannable'].to_dict()
df_GMM['Plannable'] = df_GMM['Identification'].replace(replace_dict)


# region A-13 Ajout: Naming rule
for new_itr, select_row in df_GMM.iterrows():  # Création de la colonne 'Naming rule' dans GMM
    df_GMM.at[new_itr, 'Naming rule'] = str(select_row["Identification"]) + "_" + str(select_row["Equipment name"])
# endregion

# Concaténation de la liste des interventions
# Lecture du fichier contenant les interventions
df_Inter = pd.read_excel(path_Inter, sheet_name='Documents')
df_Inter = df_Inter.reset_index()
df_Inter.pop('CODE')  # Epurage des colonnes
# Retouche des noms des colonnes
df_Inter.rename(columns={'IDENTIFICATION': 'Identification'}, inplace=True)
df_Inter.rename(columns={'INTER': 'Intervention'}, inplace=True)
df_Inter.rename(columns={'DATE_INTER': "End date"}, inplace=True)
df_Inter.rename(columns={'TITRE': 'Calibration certificat'}, inplace=True)
df_Inter.rename(columns={'AVIS': 'Compliance status'}, inplace=True)
# Ajout des colonnes nécessaires
df_Inter['Naming rule'] = 0
df_Inter['Calibration type'] = 0
df_Inter['Operation status'] = "Closed"
df_Inter['Compliance comment(s)'] = "Imported from Deltamu/Optimu software."
df_Inter['Start date'] = df_Inter['End date']  # Ajout de la date de début
# Ajout: Calibration certificat path
for new_itr, select_row in df_Inter.iterrows():
    df_Inter.at[new_itr, 'Calibration certificat path'] = path_InterDocs + select_row['Calibration certificat']
# Remplacement: Intervention type
# Remplacer les valeurs d'apprès le fichier de mapping : col[Intervention type]
replace_dict = df_Mapping_InterventionType.set_index('GMM values')['Intervention type'].to_dict()
df_Inter['Intervention type'] = df_Inter['Intervention'].replace(replace_dict)
df_Inter = df_Inter[~df_Inter['Intervention type'].str.contains("NA", na=True)]  # Supprime les lignes inutiles
# endregion
# region B-03 Modification de format: End date; Start date
# Changement de format des dates
df_Inter["End date"] = pd.to_datetime(df_Inter["End date"], errors='coerce', format='%Y-%m-%d %H:%M:%S')
df_Inter["Start date"] = pd.to_datetime(df_Inter["Start date"], errors='coerce', format='%Y-%m-%d %H:%M:%S')
# Tri des dates d'interventions pour préparer la concaténation
df_Inter = df_Inter.sort_values(by="End date", ascending=False)
df_Inter['Year'] = df_Inter['End date'].dt.strftime('%Y')
df_Inter['Month'] = df_Inter['End date'].dt.strftime('%m')
df_Inter['Day'] = df_Inter['End date'].dt.strftime('%d')
# endregion
# region B-04 Ajout: Calibration type
for new_itr, select_row in df_Inter.iterrows():
    if select_row['Intervention type'] == 'Calibration':
        df_Inter.at[new_itr, 'Calibration type'] = "External calibration"
# endregion
# region B-05 Ajout: Naming rule
# Charger la liste des 'Equipment category' du GMM
df_Inter = pd.merge(df_GMM[['Identification', 'Equipment category', 'Equipment name']], df_Inter, on='Identification',
                    how='right', suffixes=('_x', ''))
df_Inter['Datum'] = df_Inter['End date'].dt.strftime('%Y/%m/%d')
for new_itr, select_row in df_Inter.iterrows():
    df_Inter.at[new_itr, 'Naming rule'] = str(select_row['Datum']) + "_" + str(select_row['Identification']) + "_" + \
                                          str(select_row['Equipment category'])
# endregion
# region #>B-06 Ajout: Associated equipment
for new_itr, select_row in df_Inter.iterrows():
    df_Inter.at[new_itr, 'Associated equipment'] = str(select_row['Identification']) + "_" + \
                                                   str(select_row['Equipment name'])
# endregion
# region B-07 Création du fichier de sortie : path_Interventions
list_to_pop = ['Intervention', 'index', 'Datum', 'Equipment category', 'Equipment name']  # Epurage des colonnes
[df_Inter.pop(col) for col in list_to_pop]
df_Inter1 = df_Inter[df_Inter['Calibration type'] == 'External calibration']
df_Inter1 = df_Inter1[['Naming rule', 'Calibration type', 'Operation status', 'Associated equipment', 'Start date',
                       'End date', 'Compliance status', 'Compliance comment(s)', 'Calibration certificat',
                       'Calibration certificat path']]

df_Inter1.rename(columns={'Naming rule': 'Attributes:'}, inplace=True)

# Création du fichier de sortie final
# Importation du template de BASSETTI dans un dataframe
sheets_dict = pd.read_excel(path_Export_Calibration_Template, sheet_name='Calibration operation', skiprows=9, nrows=1)
sheets_dict.columns = sheets_dict.iloc[0]
sheets_dict = sheets_dict[1:]
# Match l'ordre des colonnes avec le template de BASSETTI
for i in df_Inter1:
    found = False
    for j in sheets_dict:
        if i == j:
            sheets_dict[j] = df_Inter1[i]
            found = True
    if not found:
        print(f'"{i}" not found')
# Suppression des colonnes inutiles
sheets_dict.drop(sheets_dict.columns[0], axis=1, inplace=True)
# Création du fichier de sortie: path_Export

shutil.copyfile(path_Export_Calibration_Template, path_Interventions)
writer = pd.ExcelWriter(path_Interventions, engine='openpyxl', mode='a', if_sheet_exists='overlay',
                        engine_kwargs={'keep_vba': True})
sheets_dict.to_excel(writer, sheet_name='Calibration operation', startrow=11, startcol=1, header=False, index=False)
writer.close()

# df_Inter1.to_excel(path_Interventions)

# endregion
# region B-08 Création du fichier de sortie : path_Derogations
df_Inter.rename(columns={'Operation status': 'Status'}, inplace=True)
df_Inter.rename(columns={'Compliance comment(s)': 'Comment(s)'}, inplace=True)
df_Inter.rename(columns={'Calibration certificat': 'Validation proof'}, inplace=True)
df_Inter.rename(columns={'Calibration certificat path': 'Validation proof path'}, inplace=True)
df_Inter['Datum'] = df_Inter['Year'] + df_Inter['Month']
for new_itr, select_row in df_Inter.iterrows():
    df_Inter.at[new_itr, 'Datum'] = select_row['Datum'][-4:]
for new_itr, select_row in df_Inter.iterrows():
    df_Inter.at[new_itr, 'Naming rule'] = "DER" + str(select_row['Datum']) + "." + str('{:0>4}'.format(new_itr))
df_Inter2 = df_Inter[df_Inter['Intervention type'] == 'Derogation']
df_Inter2 = df_Inter2[['Naming rule', 'Status', 'Associated equipment', 'Start date', 'End date', 'Comment(s)',
                       'Validation proof', 'Validation proof path']]
df_Inter2.to_excel(path_Derogations)
# Changement du format des dates (à faire après le tri)
df_Inter["End date"] = df_Inter["End date"].dt.strftime("%d/%m/%Y")
# endregion
# region B-09 Ajout: List of events
for new_itr, select_row in df_Inter.iterrows():  # Concaténation de l'historique dans une seule colonne
    df_Inter.at[new_itr, 'Historique'] = "[" + str(select_row["End date"]) + "] [Optimu] [" + \
                                         str(select_row['Intervention type']) + "]"
# Concaténation des interventions par équipement
df_Event = df_Inter.groupby(['Identification']).agg({'Historique': "\n".join})
# Remplacer les valeurs 'nan'
df_Event['Historique'] = df_Event['Historique'].str.replace("nan", "NA")
# Concaténation dans le dataframe principal
df_GMM = pd.merge(df_GMM, df_Event, on='Identification', how='left', suffixes=('', '_y'))
# endregion
# region #>B-10 Ajout: Last calibration status
df_Inter = df_Inter[df_Inter['Intervention type'] == 'Calibration']
for new_itr, select_row in df_Inter.iterrows():
    df_LstCal = df_Inter[df_Inter['Identification'] == select_row['Identification']]
    maxYear = df_LstCal['Year'].max()
    df_LstCal = df_LstCal[df_LstCal['Year'] == maxYear]
    maxMonth = df_LstCal['Month'].max()
    df_LstCal = df_LstCal[df_LstCal['Month'] == maxMonth]
    maxDay = df_LstCal['Day'].max()
    if select_row['Year'] == maxYear:
        if select_row['Month'] == maxMonth:
            if select_row['Day'] == maxDay:
                df_Inter.at[new_itr, 'Last calibration status'] = select_row['Compliance status']
df_Inter = df_Inter[['Identification', 'Last calibration status']]  # Suppression de toutes les colonnes inutiles
df_Inter = df_Inter[df_Inter['Last calibration status'].notnull()]  # Suppression des lignes inutiles
# Concaténation par équipement
df_Inter = df_Inter.groupby(['Identification']).agg({'Last calibration status': "\n".join})
for new_itr, select_row in df_Inter.iterrows():  # On écrème lorsque l'on a eu 2 interventions le même jour
    if len(str(select_row['Last calibration status'])) > 12:
        if 'Conforme' in select_row['Last calibration status']:
            df_Inter.at[new_itr, 'Last calibration status'] = 'Conforme'
        else:
            df_Inter.at[new_itr, 'Last calibration status'] = 'Non-conforme'
# Concaténation dans le dataframe principal
df_GMM = pd.merge(df_GMM, df_Inter, on='Identification', how='left', suffixes=('', '_y'))
# endregion

# region ###> C Concaténation de la liste des catégories *27/01/2023
output = df_GMM['Equipment category'].drop_duplicates()  # Création du dataframe et fichier de sortie
output.to_excel(path_Category)
# endregion

# region ###> D Concaténation de la liste des zones de stockage *27/01/2023
output = df_GMM['Storage area'].drop_duplicates()  # Création du dataframe et fichier de sortie
output.to_excel(path_StorageArea)
# endregion

# E : Concaténation de la liste des sous-équipements *27/01/2023
# region 000 Lecture du fichier contenant les sous-équipements
df_SubEqt = pd.read_csv(path_SubEqt, sep=';', encoding='latin_1')  # Création du dataframe
df_SubEqt = df_SubEqt.reset_index()
list_to_pop = ['Code', "Code de l'instrument lié", 'Valeur nominale', 'Unité', 'Désignation', 'Type de relation',
               'Commentaire', 'Étalonné', 'Groupe', 'N° de voie']  # Epurage des colonnes
[df_SubEqt.pop(col) for col in list_to_pop]
# endregion
# region 000 Retouche des noms des colonnes
df_SubEqt.rename(columns={'Identification': 'Sub-equipment(s)'}, inplace=True)  # Renommage des colonnes
df_SubEqt.rename(columns={"Ident. de l'instrument lié": 'Identification'}, inplace=True)
# endregion
# region E-01 Remplacement: Identification > Naming rule (Master equipment)
# Charger la liste des 'Equipment name' du GMM
df_SubEqt = pd.merge(df_GMM[['Identification', 'Equipment name']], df_SubEqt, on='Identification', how='right',
                     suffixes=('_x', ''))
for new_itr, select_row in df_SubEqt.iterrows():  # Ajout de 'Equipment name' après le numéro d'identification
    df_SubEqt.at[new_itr, 'Identification'] = str(select_row['Identification']) + "_" + \
                                              str(select_row['Equipment name'])
list_to_pop = ['Equipment name', 'index']  # Epurage des colonnes
[df_SubEqt.pop(col) for col in list_to_pop]
# endregion
# region E-02 Remplacement: Identification > Naming rule (Sub-equipment)
df_SubEqt.rename(columns={"Identification": 'Naming rule'}, inplace=True)  # Renommage des colonnes
df_SubEqt.rename(columns={'Sub-equipment(s)': 'Identification'}, inplace=True)
# Charger la liste des 'Equipment name' du GMM
df_SubEqt = pd.merge(df_GMM[['Identification', 'Equipment name']], df_SubEqt, on='Identification', how='right',
                     suffixes=('_x', ''))
for new_itr, select_row in df_SubEqt.iterrows():  # Ajout de 'Equipment name' après le numéro d'identification
    df_SubEqt.at[new_itr, 'Identification'] = str(select_row['Identification']) + "_" + \
                                              str(select_row['Equipment name'])
list_to_pop = ['Equipment name']  # Epurage des colonnes
[df_SubEqt.pop(col) for col in list_to_pop]
df_SubEqt.rename(columns={'Identification': 'Sub-equipment(s)'}, inplace=True)
# endregion
# region #>E-03 Ajout: Sub-equipment(s)
df_SubEqt = df_SubEqt.groupby(['Naming rule']).agg({"Sub-equipment(s)": '<v>'.join})  # Concaténation par 'Naming rule'
# Concaténation dans le dataframe principal
df_GMM = pd.merge(df_GMM, df_SubEqt, on='Naming rule', how='left', suffixes=('', '_y'))
# endregion

# F Retouche des attributs de la liste des équipements *01/02/2023
# region F01 Conversion: Période (unité > jours)
# Sépare les périodes de leurs unités
df_GMM[['Calibration period', 'Calibration period unit']] = df_GMM['Périodicité'].str.split(' ', n=1, expand=True)
# Remplacer les unités par leurs valeurs de convertion en jours
df_GMM = df_GMM.replace({'Calibration period unit': {'Mois': 31, 'An(s)': 365, 'Nb jours sortis actifs': 1}})
df_GMM['Calibration period'] = df_GMM['Calibration period'].fillna(0).astype(int)  # Convertion des colonnes en entier
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].fillna(0).astype(int)
df_GMM = df_GMM.astype({"Calibration period": "int", "Calibration period unit": "int"})
df_GMM['Calibration period'] = df_GMM['Calibration period'] * df_GMM['Calibration period unit']  # Multiplication
# Remplacement des unités en 'Day(s)'
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].replace(
    {31: 'Day(s)', 365: 'Day(s)', 1: 'Day(s)'})
df_GMM['Calibration period'] = df_GMM['Calibration period'].replace(0, "")
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].replace(0, "")
df_GMM.pop('Périodicité')  # Suppression de la colonne 'Périodicité'
# endregion
# region F02 Remplacement: Etat d'utilisation > Equipment status
# Création de la colonne 'Equipment status' dans GMM
replace_dict = df_Mapping_Status.set_index('GMM values')['Equipment status'].to_dict()
df_GMM['Equipment status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)
# endregion
# region F03 Remplacement: Etat d'utilisation > Maintenance status
# Création de la colonne 'Maintenance status' dans GMM
replace_dict = df_Mapping_Status.set_index('GMM values')['Maintenance status'].to_dict()
df_GMM['Maintenance status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)
# endregion
# region F04 Remplacement: Etat d'utilisation > Calibration status
# Création de la colonne 'Calibration status' dans GMM // 'No calibration status' and 'Scraped'
replace_dict = df_Mapping_Status.set_index('GMM values')['Calibration status'].to_dict()
df_GMM['Calibration status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)
# endregion
# region F05 Remplacement: Avis > Calibration status
replace_dict = df_Mapping_CalibrationStatus.set_index('GMM values')['Calibration status'].to_dict()
df_GMM['Calibration status'] = df_GMM["Avis"].replace(replace_dict)
df_GMM.pop("Avis")
# endregion
# region F06 Remplacement: Calibration status (vide)  > Last calibration status
for new_itr, select_row in df_GMM.iterrows():
    # Résultats du dernier étalonnage lorsqu'il est disponible // 'Lost' and 'No calibration status'
    if select_row['Equipment status'] == "Lost":
        df_GMM.at[new_itr, 'Calibration status'] = str(select_row['Last calibration status'])
    if select_row['Calibration status'] == "Refer to attached docs and certificats":
        df_GMM.at[new_itr, 'Calibration status'] = str(select_row['Last calibration status'])
# endregion
# region #>F07 Remplacement: Calibration status ("Refer ...")  > Last calibration status
for new_itr, select_row in df_GMM.iterrows():
    if select_row['Calibration status'] == "Refer to attached docs and certificats":
        df_GMM.at[new_itr, 'Calibration status'] = str(select_row['Last calibration status'])
# endregion
# region #>F08 Remplacement: Statut > Approach device
for new_itr, select_row in df_GMM.iterrows():
    if select_row["Calibration status"] == "nan":
        df_GMM.at[new_itr, 'Statut'] = "Moyen d'approche"
        df_GMM.at[new_itr, 'Calibration status'] = "No calibration status"
    if select_row["Etat d'utilisation"] == "Non soumis à l'étalonnage":
        df_GMM.at[new_itr, 'Statut'] = "Moyen d'approche"
        df_GMM.at[new_itr, 'Calibration status'] = "No calibration status"
    else:
        df_GMM.at[new_itr, 'Statut'] = select_row['Statut']
# endregion
# Vérification: Calibration status == Out of date
df_GMM.rename(columns={"Prochaine date d'intervention": 'Date of next calibration'}, inplace=True)
# Vérification de la prochaine date d'étalonnage et changement du statut de calibration
for new_itr, select_row in df_GMM.iterrows():
    if not pd.isnull(select_row['Date of next calibration']):
        if datetime.datetime.strptime(select_row['Date of next calibration'], '%d/%m/%Y %H:%M') < \
                datetime.datetime.today():
            df_GMM.at[new_itr, 'Calibration status'] = "Out of date"
# Création de la colonne 'Equipment type' dans GMM
replace_dict = df_Mapping_EquipmentType.set_index('GMM values')['Equipment type'].to_dict()
df_GMM['Equipment type'] = df_GMM['Statut'].replace(replace_dict)
# Création de la colonne 'Measuring chain' dans GMM
replace_dict = df_Mapping_EquipmentType.set_index('GMM values')['Measuring chain'].to_dict()
df_GMM['Measuring chain'] = df_GMM['Statut'].replace(replace_dict)
# Création de la colonne 'Calibration type' dans GMM
replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration type'].to_dict()
df_GMM['Calibration type'] = df_GMM['Equipment category'].replace(replace_dict)
# Création de la colonne 'Calibration supplier' dans GMM
replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration supplier'].to_dict()
df_GMM['Calibration suppliers'] = df_GMM['Equipment category'].replace(replace_dict)
# Création de la colonne 'Applicable norms' dans GMM
df_GMM['Applicable norms'] = df_GMM['Equipment category']
# Création de la colonne 'Standard gages' dans GMM
df_GMM['Standard gages'] = df_GMM['Equipment category']
# Création de la colonne 'Laboratory' dans GMM
# replace_dict = df_Mapping_Laboratory.set_index('GMM values')['Laboratory'].to_dict()
df_GMM['Laboratory'] = df_GMM['Affected person']
# endregion
# Création de la colonne 'Business group' dans GMM
# replace_dict = df_Mapping_Laboratory.set_index('GMM values')['Business group'].to_dict()
# df_GMM['Business group'] = df_GMM['Affected person'].replace(replace_dict)
# Epurage des colonnes
list_to_pop = ["Etat d'utilisation", 'Statut', 'Calibration period unit', 'Désignation']  # Epurage des colonnes
[df_GMM.pop(col) for col in list_to_pop]

# Mise en accord avec le dictionnaire
# Renommage des colonnes en fonction du dictionnaire
df_GMM.rename(columns={'Naming rule': 'Attributes:'}, inplace=True)
df_GMM.rename(columns={'Identification': 'Equipment number'}, inplace=True)
df_GMM.rename(columns={'Equipment status': 'Equipment Status'}, inplace=True)
df_GMM.rename(columns={'Sub-equipment(s)': 'Sub-equipment'}, inplace=True)
df_GMM.rename(columns={'Calibration supplier': 'Calibration supplier(s)'}, inplace=True)
df_GMM.rename(columns={'Manufacturer': 'Name'}, inplace=True)
df_GMM.rename(columns={'SAP number': 'N°SAP (Finance/Maintenance)'}, inplace=True)
df_GMM.rename(columns={'Domaine de mesure': 'Measurement domain 1'}, inplace=True)
df_GMM.rename(columns={'Gamme': 'Range 1'}, inplace=True)
df_GMM.rename(columns={'Tolérance': 'Tolerance 1'}, inplace=True)
df_GMM.rename(columns={"Date d'intervention": 'Date of last calibration'}, inplace=True)
df_GMM.rename(columns={"Number of intermediate controls": 'Number of controls between calibrations'}, inplace=True)
df_GMM.rename(columns={"Est. downtime": 'Estimated downtime'}, inplace=True)
df_GMM.rename(columns={"Est. calibration time": 'Estimated time of  calibration'}, inplace=True)
df_GMM.rename(columns={"Est. cost": 'Estimated cost of calibration'}, inplace=True)
df_GMM.rename(columns={'Historique': 'Historic'}, inplace=True)
# Création des colonnes du dictionnaire non existantes
df_GMM['Approval necessary for calibration'] = 'No'
df_GMM['Legal Entity'] = 'CES'
df_GMM['Method(s)'] = ''
df_GMM['Measured points'] = ''
df_GMM['Criteria defined by'] = ''
df_GMM['Plan of laboratory'] = ''

df_GMM['Measuring chain'] = np.where(df_GMM['Measuring chain'], 'true', 'false')
df_GMM['Date of last calibration'] = df_GMM['Date of last calibration'].fillna("To be determined")

df_GMM.to_excel('Export.xlsx')

# Création du fichier de sortie final
# Importation du template de BASSETTI dans un dataframe
sheets_dict = pd.read_excel(path_Export_Template, sheet_name='Equipment', skiprows=9, nrows=1)
sheets_dict.columns = sheets_dict.iloc[0]
sheets_dict = sheets_dict[1:]
# Match l'ordre des colonnes avec le template de BASSETTI
for i in df_GMM:
    found = False
    for j in sheets_dict:
        if i == j:
            sheets_dict[j] = df_GMM[i]
            found = True
    if not found:
        print(f'"{i}" not found')
# Suppression des colonnes inutiles
sheets_dict.drop(sheets_dict.columns[0], axis=1, inplace=True)
# Création du fichier de sortie: path_Export

shutil.copyfile(path_Export_Template, path_Export)
writer = pd.ExcelWriter(path_Export, engine='openpyxl', mode='a', if_sheet_exists='overlay',
                        engine_kwargs={'keep_vba': True})
sheets_dict.to_excel(writer, sheet_name='Equipment', startrow=11, startcol=1, header=False, index=False)
writer.close()
