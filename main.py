# Exemple de traitement de données laboratoire
#05/01/2023: Modification du calibration status en l189-205 (remplacement de la valeur 'Avis' par dernière valeur de l'historique d'intervention) > Le remplacement ne se fait pas correctement.
#13/01/2023: Règle en l201-206 pour la vérification que la date de prochaine étalonnage soit supérieure à la date d'aujourd'hui.
#24/01/2023: Création de la liste des opérations sur la BD à valider par Eric PIERRE.

import datetime
from dateutil import parser
import os
import pandas
import pandas as pd

#************************************************************************************************ CHEMINS
#Liste des chemins d'entrée *16/12/2022
### Fournis par : Laboratoire
path_Mapping = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM_Mapping.xlsx" #Fichier de mapping
path_GMMUpdate = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM_Update.xlsx" #Fichier d'update
path_MastEqt = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/MasterEquipments.xlsx" #Liste des équipements maîtres
### Fournis par : Steeven
path_GMM1 = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM.xlsx" #Fichier général (1)
path_GMM2 = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM - Instruments Plastic Omnium Alphatech.csv" #Fichier général (2)
path_SubEqt = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM Instruments liés.csv" #Instruments liés
### Fournis par : Deltamu
path_Inter = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/PlasticOmnium.xlsx" #Liste des interventions
path_InterDocs = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/PLAST-60/" #Fichiers

#Liste des chemins de sortie
path_Export = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Export.xlsx"
path_Category = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Export_Category.xlsx"
path_StorageArea = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Export_StorageArea.xlsx"
path_Analyze = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Analyze.xlsx"
path_Analyze2 = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Analyze2.xlsx"

#************************************************************************************************ CODE
###> Suppression des fichiers d'export existants *16/12/2022
if os.path.exists(path_Export):
    os.remove(path_Export)

###> Création des dataframes de mapping *16/12/2022
df_Mapping_CalibrationStatus = pandas.read_excel(path_Mapping,sheet_name="CalibrationStatus")  # Création du dataframe : mapping(CalibrationStatus)
df_Mapping_CalibrationStatus = df_Mapping_CalibrationStatus.reset_index()

df_Mapping_InterventionType = pandas.read_excel(path_Mapping,sheet_name="InterventionType")  # Création du dataframe : mapping(InterventionType)
df_Mapping_InterventionType = df_Mapping_InterventionType.reset_index()

df_Mapping_Status = pandas.read_excel(path_Mapping,sheet_name="Status")  # Création du dataframe : mapping(Status)
df_Mapping_Status = df_Mapping_Status.reset_index()

df_Mapping_CalibrationSupplier = pandas.read_excel(path_Mapping,sheet_name="CalibrationType_Supplier")  # Création du dataframe : mapping(CalibrationType_Supplier)
df_Mapping_CalibrationSupplier = df_Mapping_CalibrationSupplier.reset_index()

df_Mapping_EquipmentType = pandas.read_excel(path_Mapping, sheet_name="EquipmentType_Chain")  # Création du dataframe : mapping(EquipmentType_Chain)
df_Mapping_EquipmentType = df_Mapping_EquipmentType.reset_index()

df_Mapping_Domain = pandas.read_excel(path_Mapping, sheet_name="Domain")  # Création du dataframe : mapping(Domain)
df_Mapping_Domain = df_Mapping_Domain.reset_index()

df_Mapping_Laboratory = pandas.read_excel(path_Mapping, sheet_name="Laboratory")  # Création du dataframe : mapping(Laboratory)
df_Mapping_Laboratory = df_Mapping_Laboratory.reset_index()

df_Mapping_Update = pandas.read_excel(path_GMMUpdate, sheet_name="GMM")  # Création du dataframe : mapping(GMM)
df_Mapping_Update = df_Mapping_Update.reset_index()

###> Création de la liste des équipements *16/12/2022
df_GMM2 = pd.read_csv(path_GMM2, sep=';') #Création du dataframe : liste des équipements (1)
df_GMM2 = df_GMM2.reset_index()
list_to_pop = ['Affectation', 'Sous-Affectation', 'Désignation', 'Informations liées (Identification, Désignation, N° de série)',"Etat d'utilisation",'N° de série','Commentaire'] #Epurage des colonnes
[df_GMM2.pop(col) for col in list_to_pop]

df_GMM1 = pd.read_excel(path_GMM1, sheet_name="GMM")  #Création du dataframe : liste des équipements (2)
df_GMM1 = df_GMM1.reset_index()
list_to_pop = ['Code', "Famille d'instrument",'Type',"Nombre d'éléments",'Fournisseur',"Prix d'achat","Date d'achat",'Constructeur','Référence constructeur',"N° de série",'Mode de lecture','Particularité','Matière','Commentaire','Valeur nominale','Unité','Référentiel','Résolution','Désignation littérale',"N° d'immobilisation",'Code instrument','Gestionnaire','Ident. secondaire','Localisation temporaire','Unité de la résolution','Localisation'] #Epurage des colonnes
[df_GMM1.pop(col) for col in list_to_pop]
df_GMM1.rename(columns={"État d'utilisation": "Etat d'utilisation"}, inplace=True)

df_GMM = pd.merge(df_GMM1, df_GMM2, on='Identification') #Fusion des 2 dataframes

###> A Remplacement des valeurs d'OPTIMU avec le travail d'épuration réalisé par le laboratoire *04/01/2023
#> A-01 Remplacement: Désignation > Equipment name
replace_dict = df_Mapping_Update.set_index('Identification')['Equipment name'].to_dict()  #Création de la colonne 'Equipment name' dans GMM
df_GMM['Equipment name'] = df_GMM['Identification'].replace(replace_dict)

for new_itr, select_row in df_GMM.iterrows(): #Comparaison des colonnes 'Désignation' et 'Equipment name'
    if pd.isnull(select_row["Equipment name"]):
        df_GMM.at[new_itr, 'Equipment name'] = str(select_row["Désignation"])
#>A-02 Ajout: Equipment category
replace_dict = df_Mapping_Update.set_index('Identification')['Equipment category'].to_dict()  #Création de la colonne 'Equipment category' dans GMM
df_GMM['Equipment category'] = df_GMM['Identification'].replace(replace_dict)
#>A-03 Remplacement: Fabricant > Manufacturer
replace_dict = df_Mapping_Update.set_index('Identification')['Fabricant'].to_dict()  #Création de la colonne 'Manufacturer' dans GMM
df_GMM['Manufacturer'] = df_GMM['Identification'].replace(replace_dict)
#>A-04 Remplacement: Modèle > Model
replace_dict = df_Mapping_Update.set_index('Identification')['Model'].to_dict()  #Création de la colonne 'Model' dans GMM
df_GMM['Model'] = df_GMM['Identification'].replace(replace_dict)
#>A-05 Remplacement: N° de série > Serial number
replace_dict = df_Mapping_Update.set_index('Identification')['Serial number'].to_dict()  #Création de la colonne 'Serial number' dans GMM
df_GMM['Serial number'] = df_GMM['Identification'].replace(replace_dict)
#>A-06 Remplacement: Gestionnaire > Affected person
replace_dict = df_Mapping_Update.set_index('Identification')['Affected person'].to_dict()  #Création de la colonne 'Affected perso' dans GMM
df_GMM['Affected person'] = df_GMM['Identification'].replace(replace_dict)
#>A-07 Remplacement: Localisation > Storage area
replace_dict = df_Mapping_Update.set_index('Identification')['Storage area'].to_dict()  #Création de la colonne 'Storage area' dans GMM
df_GMM['Storage area'] = df_GMM['Identification'].replace(replace_dict)
#>A-08 Ajout: Intermediate controls
replace_dict = df_Mapping_Update.set_index('Identification')['Intermediate controls'].to_dict()  #Création de la colonne 'Intermediate controls' dans GMM
df_GMM['Intermediate controls'] = df_GMM['Identification'].replace(replace_dict)
#>A-09 Ajout: Number of controls
replace_dict = df_Mapping_Update.set_index('Identification')['Number of controls'].to_dict()  #Création de la colonne 'Number of controls' dans GMM
df_GMM['Number of controls'] = df_GMM['Identification'].replace(replace_dict)
#>A-10 Ajout: Est. Downtime
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Downtime'].to_dict()  #Création de la colonne 'Est. Downtime' dans GMM
df_GMM['Est. downtime'] = df_GMM['Identification'].replace(replace_dict)
#>A-11 Ajout: Est. Calibration time
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Calibration time'].to_dict()  #Création de la colonne 'Est. Calibration time' dans GMM
df_GMM['Est. calibration time'] = df_GMM['Identification'].replace(replace_dict)
#>A-12 Ajout: Est. Cost
replace_dict = df_Mapping_Update.set_index('Identification')['Est. Cost'].to_dict()  #Création de la colonne 'Est. Cost' dans GMM
df_GMM['Est. cost'] = df_GMM['Identification'].replace(replace_dict)
#>A-13 Ajout: Naming rule
for new_itr, select_row in df_GMM.iterrows(): #Création de la colonne 'Naming rule' dans GMM
    df_GMM.at[new_itr, 'Naming rule'] = str(select_row["Identification"]) + "_" + str(select_row["Equipment name"])

###> B Ajout de la liste des interventions *04/01/2023
df_Inter = pd.read_excel(path_Inter, sheet_name='Documents') #Création du dataframe
df_Inter = df_Inter.reset_index()
df_Inter.pop('CODE') #Epurage des colonnes
df_Inter.rename(columns={'IDENTIFICATION': 'Identification'}, inplace=True) #Renommage des colonnes
df_Inter.rename(columns={'INTER': 'Intervention'}, inplace=True)
df_Inter.rename(columns={'DATE_INTER': "Date d'intervention"}, inplace=True)
df_Inter.rename(columns={'TITRE': 'Document'}, inplace=True)
df_Inter.rename(columns={'AVIS': 'Calibration status'}, inplace=True)
for new_itr, select_row in df_Inter.iterrows():  #Retouche des dates et chemins
    df_Inter.at[new_itr, 'Documents'] = path_InterDocs + select_row['Document']
    # df_Inter.at[new_itr, "Date d'intervention"] = datetime.datetime.strptime(select_row["Date d'intervention"], '%m/%d/%Y %H:%M:%S').strftime('%d/%m/%Y')
replace_dict = df_Mapping_InterventionType.set_index('GMM values')['Intervention type'].to_dict()  #Remplacer les valeurs d'apprès le fichier de mapping : col[Intervention type]
df_Inter['Intervention type'] = df_Inter['Intervention'].replace(replace_dict)
df_Inter = df_Inter[df_Inter['Intervention type'].str.contains("NA") == False] #Supprime les lignes inutiles
for new_itr, select_row in df_Inter.iterrows():  #Création d'une colonne qui servira de mapping pour la concaténation
    df_Inter.at[new_itr, 'Map'] = str(select_row['Identification']) + "." + str(select_row['Intervention type']) + "." + str(select_row["Date d'intervention"])

df_Inter["Date d'intervention"] = pd.to_datetime(df_Inter["Date d'intervention"],errors='coerce', format='%d/%m/%Y') #Changement de format des dates
df_Inter = df_Inter.sort_values(by="Date d'intervention",ascending=False) #Tri des dates d'interventions pour préparer la concaténation
df_Inter["Date d'intervention"] = df_Inter["Date d'intervention"].dt.strftime("%d/%m/%Y") #Changement du format des dates (à faire après le tri)

for new_itr, select_row in df_Inter.iterrows():  #Concaténation de l'historique dans une seule colonne
    df_Inter.at[new_itr, 'Historique'] = str(select_row["Date d'intervention"]) + " : " + str(select_row['Intervention type']) + ", Compliance: " + str(select_row['Calibration status']) + ". Document: " + str(select_row['Document']) #+ ". Comments: " + str(select_row['COMMENTAIRE']).replace("\n",". ").replace("\r","")

# list_to_pop = ['Intervention', 'DATE_INTER','CHEMIN_DOC',"Date d'intervention",'Intervention type'] #Epurage des colonnes
list_to_pop = ['Intervention', "Date d'intervention", 'Intervention type', 'Calibration status'] # Epurage des colonnes
[df_Inter.pop(col) for col in list_to_pop]
# df_Inter = df_Inter.groupby(['Identification']).agg({'Documents': '<v>'.join})  #Concaténation des documents d'interventions par équipement

df_Inter = df_Inter.groupby(['Identification']).agg({'Historique': "\n".join})  #Concaténation des interventions par équipement
# df_Inter = pd.merge(df_Inter, df_Inter1, on='Identification', how='left',suffixes=('','_y')) #Concaténation des 2 dataframes
df_Inter['Historique'] = df_Inter['Historique'].str.replace("nan", "NA") #Remplacer les valeurs 'nan'

for new_itr, select_row in df_Inter.iterrows():  #Création de la colonne 'Last calibration status'
    my_list = select_row['Historique'].split('Compliance:')
    my_list = my_list[1].split('. Document:')
    df_Inter.at[new_itr, 'Last calibration status'] = my_list[0].lstrip()

df_GMM = pd.merge(df_GMM,df_Inter,on='Identification',how='left', suffixes=('','_y')) #Concaténation dans le dataframe principal
list_to_pop = ['index_x', 'index_y'] #Epurage des colonnes
[df_GMM.pop(col) for col in list_to_pop]

###> Création de la liste des catégories *16/12/2022
output = df_GMM['Equipment category'].drop_duplicates()  #Création du dataframe et fichier de sortie
output.to_excel(path_Category)

###> Zone de stockage *16/12/2022
output = df_GMM['Storage area'].drop_duplicates()  #Création du dataframe et fichier de sortie
output.to_excel(path_StorageArea)

###> Equipements
#df_GMM['Gamme'] = df_GMM['Gamme'].str.replace("SANS", "")
#df_GMM['Tolérance'] = df_GMM['Tolérance'].str.replace("SANS", "")

df_GMM[['Calibration period', 'Calibration period unit']] = df_GMM['Périodicité'].str.split(' ', 1, expand=True) #Sépare les périodes de leurs unités
df_GMM = df_GMM.replace({'Calibration period unit' : {'Mois' : 31, 'An(s)' : 365, 'Nb jours sortis actifs' : 1}}) #Remplacer les unités par leurs valeurs de convertion en jours
df_GMM['Calibration period'] = df_GMM['Calibration period'].fillna(0).astype(int) #Convertion des colonnes en entier
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].fillna(0).astype(int)
df_GMM = df_GMM.astype({"Calibration period":"int","Calibration period unit":"int"})
df_GMM['Calibration period'] = df_GMM['Calibration period'] * df_GMM['Calibration period unit'] #Multiplication
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].replace({31 : 'Day(s)', 365 : 'Day(s)', 1 : 'Day(s)'}) #Remplacement des unités en 'Day(s)'
df_GMM['Calibration period'] = df_GMM['Calibration period'].replace(0,"")
df_GMM['Calibration period unit'] = df_GMM['Calibration period unit'].replace(0,"")
df_GMM.pop('Périodicité') #Suppression de la colonne 'Périodicité'

replace_dict = df_Mapping_Status.set_index('GMM values')['Equipment status'].to_dict()  #Création de la colonne 'Equipment status' dans GMM
df_GMM['Equipment status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)

#>Création-Modification du 'Calibration status'
replace_dict = df_Mapping_Status.set_index('GMM values')['Calibration status'].to_dict()  #Création de la colonne 'Calibration status' dans GMM // 'No calibration status' and 'Scraped'
df_GMM['Calibration status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)

for new_itr, select_row in df_GMM.iterrows(): #Résultats du dernier étalonnage lorsqu'il est disponible // 'Lost' and 'No calibration status'
    if select_row['Equipment status'] == "Lost":
        df_GMM.at[new_itr, 'Avis'] = str(select_row['Last calibration status'])

replace_dict = df_Mapping_CalibrationStatus.set_index('GMM values')['Calibration status'].to_dict()
df_GMM['Calibration status'] = df_GMM["Avis"].replace(replace_dict)
df_GMM.pop("Avis")

df_GMM.rename(columns={"Prochaine date d'intervention": 'Date of next calibration'}, inplace=True)

for new_itr, select_row in df_GMM.iterrows(): #Vérification de la prochaine date d'étalonnage et changement du statut de calibration
    if not pd.isnull(select_row['Date of next calibration']):
        if datetime.datetime.strptime(select_row['Date of next calibration'], '%d/%m/%Y %H:%M:%S') < datetime.datetime.today():
            df_GMM.at[new_itr, 'Calibration status'] = "Out of date"

for new_itr, select_row in df_GMM.iterrows(): #Mise à jour de la colonne 'Statut' et de la colonne 'Equipment type'
    if select_row['Calibration status'] == "Refer to attached docs and certificats":
        df_GMM.at[new_itr, 'Avis'] = str(select_row['Last calibration status'])
    if select_row["Calibration status"] == "nan":
        df_GMM.at[new_itr, 'Statut'] = "Moyen d'approche"
        df_GMM.at[new_itr, 'Calibration status'] = "No calibration status"
    if select_row["Etat d'utilisation"] == "Non soumis à l'étalonnage":
        df_GMM.at[new_itr, 'Statut'] = "Moyen d'approche"
    else:
        df_GMM.at[new_itr, 'Statut'] = select_row['Statut']


replace_dict = df_Mapping_Status.set_index('GMM values')['Maintenance status'].to_dict()  #Création de la colonne 'Maintenance status' dans GMM
df_GMM['Maintenance status'] = df_GMM["Etat d'utilisation"].replace(replace_dict)

replace_dict = df_Mapping_EquipmentType.set_index('GMM values')['Measuring chain'].to_dict()  #Création de la colonne 'Measuring chain' dans GMM
df_GMM['Measuring chain'] = df_GMM['Statut'].replace(replace_dict)

replace_dict = df_Mapping_EquipmentType.set_index('GMM values')['Equipment type'].to_dict()  #Création de la colonne 'Equipment type' dans GMM
df_GMM['Equipment type'] = df_GMM['Statut'].replace(replace_dict)

df_GMM.to_excel(path_Analyze)

replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration type'].to_dict()  #Création de la colonne 'Calibration type' dans GMM
df_GMM['Calibration type'] = df_GMM['Equipment category'].replace(replace_dict)

replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration supplier'].to_dict()  #Création de la colonne 'Calibration supplier' dans GMM
df_GMM['Calibration suppliers'] = df_GMM['Equipment category'].replace(replace_dict)

replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Applicable norms'].to_dict()  #Création de la colonne 'Applicable norms' dans GMM
df_GMM['Applicable norms'] = df_GMM['Equipment category'].replace(replace_dict)

replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Standard gages'].to_dict()  #Création de la colonne 'Standard gages' dans GMM
df_GMM['Standard gages'] = df_GMM['Equipment category'].replace(replace_dict)

replace_dict = df_Mapping_Laboratory.set_index('GMM values')['Laboratory'].to_dict()  #Création de la colonne 'Laboratory' dans GMM
df_GMM['Laboratory'] = df_GMM['Affected person'].replace(replace_dict)

replace_dict = df_Mapping_Laboratory.set_index('GMM values')['Business group'].to_dict()  #Création de la colonne 'Business group' dans GMM
df_GMM['Business group'] = df_GMM['Affected person'].replace(replace_dict)


list_to_pop = ["Etat d'utilisation", 'Statut','Calibration period unit','Désignation'] #Epurage des colonnes
[df_GMM.pop(col) for col in list_to_pop]

#replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration type'].to_dict()  #Création de la colonne 'Calibration type' dans GMM
#df_GMM['Calibration type'] = df_GMM['Equipment category'].replace(replace_dict)

#replace_dict = df_Mapping_CalibrationSupplier.set_index('GMM values')['Calibration supplier'].to_dict()  #Création de la colonne 'Calibration supplier(s)' dans GMM
#df_GMM['Calibration supplier(s)'] = df_GMM['Equipment category'].replace(replace_dict)

###> Sous-équipements
df_SubEqt = pd.read_csv(path_SubEqt, sep=';', encoding='latin_1')  #Création du dataframe
df_SubEqt = df_SubEqt.reset_index()

df_SubEqt = pd.merge(df_GMM[['Identification','Equipment name']],df_SubEqt,on='Identification', how='right', suffixes=('_x','')) #Charger la liste des 'Equipment name' du GMM
for new_itr, select_row in df_SubEqt.iterrows():  #Ajout de 'Equipment name' après le numéro d'identification
    df_SubEqt.at[new_itr, 'Identification'] = str(select_row['Identification']) + "_" + str(select_row['Equipment name'])

df_SubEqt.rename(columns={'Identification': 'Sub-equipment(s)'}, inplace=True) #Renommage des colonnes
df_SubEqt.rename(columns={"Ident. de l'instrument lié": 'Identification'}, inplace=True)

df_SubEqt = df_SubEqt.groupby(['Identification']).agg({"Sub-equipment(s)": '<v>'.join}) #Concaténation par 'Identification'
df_GMM = pd.merge(df_GMM,df_SubEqt,on='Identification',how='left', suffixes=('','_y')) #Concaténation dans le dataframe principal

###> Renommage des colonnes en fonction du dictionnaire *16/12/2022
df_GMM.rename(columns={'Identification': 'Equipment number'}, inplace=True)
df_GMM.rename(columns={'Constructeur': 'Manufacturer'}, inplace=True)
df_GMM.rename(columns={'Référence constructeur': 'Model'}, inplace=True)
df_GMM.rename(columns={'N° de série': 'Serial number'}, inplace=True)
df_GMM.rename(columns={'Domaine de mesure': 'Measurement domain 1'}, inplace=True)
df_GMM.rename(columns={'Gamme': 'Range 1'}, inplace=True)
df_GMM.rename(columns={'Tolérance': 'Tolerance 1'}, inplace=True)
df_GMM.rename(columns={"Date d'intervention": 'Date of last calibration'}, inplace=True)
df_GMM.rename(columns={'Historique': 'Historic'}, inplace=True)
df_GMM.rename(columns={'Document': 'Certificats'}, inplace=True)

###> Création des colonnes du dictionnaire non existantes
df_GMM['Plannable'] = "No"
df_GMM['Approval necessary for calibration'] = "No"

###> Tri des colonnes
#df_GMM = df_GMM[['Equipment name', 'Equipment number', 'Equipment status', 'Plannable', 'Equipment category', 'Laboratory', 'Business group', 'Storage area', 'Sub-equipment', 'Attachment equipment', 'Affected person', 'Approval necessary for calibration', 'Characteristic', 'Manufacturer', 'Model', 'Serial number', "SAP number", 'Calibration status', 'Equipment type', 'Date of next calibration', 'Measurement domain 1', 'Range 1', 'Tolerance 1', 'Calibration type', 'Calibration suppliers', 'Measuring chain', 'Applicable norms', 'Standard gages', 'Date of last calibration', 'Calibration period', 'Number of controls', 'Est. downtime', 'Est. calibration time', 'Est. cost', 'Historic', 'Certificats', 'Intermediate controls']]

###> Attachment equipment
# df_MastEqt = pd.read_excel(path_MastEqt, sheet_name='Sheet1')
# df_MastEqt = df_MastEqt.reset_index()
#
# df_GMM = pd.concat(df_GMM,df_MastEqt,on='Identification')

# df_GMM['Plannable'] = "Yes"
# df_GMM['Approval necessary for calibration'] = "Yes"

###> IMPRESSION TEMPORAIRE
df_GMM = df_GMM[sorted(df_GMM)]
df_GMM.to_excel(path_Export)
###< IMPRESSION TEMPORAIRE