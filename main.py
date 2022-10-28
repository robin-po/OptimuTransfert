#Exemple de traitement de données laboratoire
import pandas
import pandas as pd
import os
import glob
import datetime
import time
from openpyxl import load_workbook
import xlrd

#Liste des fichiers utilisés
path_GMM = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM - Instruments Plastic Omnium Alphatech.csv"
path_Inter = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/INTERVENTIONS_DOCS.csv"
path_Mapping = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/0-Input/GMM_Mapping.xlsx"
path_RevInter = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/INTERVENTIONS_DOCS_revised.csv"
path_Import = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/2022-09-28_Import_Equipment_CES.xlsx"
path_Categories = "C:/Users/yoann.skoczek/PycharmProjects/OptimuTransfert/1-Output/Categories.csv"

#Fonction: Catégories d'équipement
def Categories(df_GMM,df_Mapping_Name):
    df_GMM['Equipment category'] = 0 #Création de la colonne 'Equipment category' dans GMM
    df_GMM['Equipment name'] = df_GMM['Désignation'] #Création de la colonne 'Equipment name' dans GMM

    #df_GMM['Equipment name'] = df_Mapping_Name[df_Mapping_Name['GMM values'].isin(df_GMM['Désignation'])]['Equipment name'].values
    #df_GMM['Equipment name'] = df_Mapping_Name[df_Mapping_Name['GMM values'].isin(df_GMM['Désignation'])]['Equipment name'].values
    df_Mapping_Name.set_index('GMM values').T.to_dict('list')


    output = df_GMM['Equipment category'].drop_duplicates() #Création du dataframe et fichier de sortie
    output = output.reset_index()
    output.to_csv(path_Categories, "\t", decimal=",")  # Export au format *.csv
    print(output)

#Fonction: Equipements
def Equipements():
    df = pd.read_csv(path_Inter,sep=';') #Création du dataframe
    df = df.reset_index()

    for new_itr, select_row in df.iterrows(): #Retouche des dates et chemins
        df.at[new_itr, 'DOC'] = os.path.basename(select_row['CHEMIN_DOC'])
        df.at[new_itr, 'DATE_INTER'] = select_row['DATE_INTER'][0:10]

    for new_itr, select_row in df.iterrows(): #Concaténation de l'historique dans une seule colonne
        df.at[new_itr, 'HISTORIQUE']= select_row['INTER'].replace("é","e") + " : " + select_row['DATE_INTER'] + " : " + select_row['DOC']

    output = df.groupby(['IDENTIFICATION']).agg({'HISTORIQUE':list,'DOC':list}) #Concaténation des interventions par équipement
    output.to_csv(path_RevInter,"\t", decimal=",") #Export au format *.csv

#Fonction: Sous-équipements
def SousEquipements():
    df = pd.read_csv(path_Inter,sep=';') #Création du dataframe
    df = df.reset_index()

    for new_itr, select_row in df.iterrows(): #Retouche des dates et chemins
        df.at[new_itr, 'DOC'] = os.path.basename(select_row['CHEMIN_DOC'])
        df.at[new_itr, 'DATE_INTER'] = select_row['DATE_INTER'][0:10]

    for new_itr, select_row in df.iterrows(): #Concaténation de l'historique dans une seule colonne
        df.at[new_itr, 'HISTORIQUE']= select_row['INTER'].replace("é","e") + " : " + select_row['DATE_INTER'] + " : " + select_row['DOC']

    output = df.groupby(['IDENTIFICATION']).agg({'HISTORIQUE':list,'DOC':list}) #Concaténation des interventions par équipement
    output.to_csv(path_RevInter,"\t", decimal=",") #Export au format *.csv

#Fonction: Zones de stockage
def Zones():
    df = pd.read_csv(path_Inter,sep=';') #Création du dataframe
    df = df.reset_index()

    for new_itr, select_row in df.iterrows(): #Retouche des dates et chemins
        df.at[new_itr, 'DOC'] = os.path.basename(select_row['CHEMIN_DOC'])
        df.at[new_itr, 'DATE_INTER'] = select_row['DATE_INTER'][0:10]

    for new_itr, select_row in df.iterrows(): #Concaténation de l'historique dans une seule colonne
        df.at[new_itr, 'HISTORIQUE']= select_row['INTER'].replace("é","e") + " : " + select_row['DATE_INTER'] + " : " + select_row['DOC']

    output = df.groupby(['IDENTIFICATION']).agg({'HISTORIQUE':list,'DOC':list}) #Concaténation des interventions par équipement
    output.to_csv(path_RevInter,"\t", decimal=",") #Export au format *.csv

#Fonction: Interventions
def Interventions():
    df = pd.read_csv(path_Inter,sep=';') #Création du dataframe
    df = df.reset_index()

    for new_itr, select_row in df.iterrows(): #Retouche des dates et chemins
        df.at[new_itr, 'DOC'] = os.path.basename(select_row['CHEMIN_DOC'])
        df.at[new_itr, 'DATE_INTER'] = select_row['DATE_INTER'][0:10]

    for new_itr, select_row in df.iterrows(): #Concaténation de l'historique dans une seule colonne
        df.at[new_itr, 'HISTORIQUE']= select_row['INTER'].replace("é","e") + " : " + select_row['DATE_INTER'] + " : " + select_row['DOC']

    output = df.groupby(['IDENTIFICATION']).agg({'HISTORIQUE':list,'DOC':list}) #Concaténation des interventions par équipement
    output.to_csv(path_RevInter,"\t", decimal=",") #Export au format *.csv

#Code
df_GMM = pd.read_csv(path_GMM,sep=';') #Création du dataframe : liste des équipements
df_GMM = df_GMM.reset_index()
df_Mapping_Name = pandas.read_excel(path_Mapping,sheet_name="Name") #Création du dataframe : mapping(name)
df_Mapping_Name = df_Mapping_Name.reset_index()
Categories(df_GMM,df_Mapping_Name) #Etape 1 : Catégories d'équipement
#Equipements #Etape 2 : Equipements
#SousEquipements() #Etape 3 : Sous-équipements
#Zones() #Etape 4 : Zones de stockage
#Interventions() #Etape 5 : Historique d'interventions