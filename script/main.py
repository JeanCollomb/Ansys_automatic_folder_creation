# -*- coding: utf-8 -*-
"""

"""


ansys_folder = "D:\Program Files\ANSYS Inc\v181\ansys\bin\winx64\ansys181"


#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                       IMPORT PACKAGE
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

import os
from distutils.dir_util import copy_tree
import time
import pandas as pd

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                     INFO.
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
print('\n')
print('############################')
print('-- Program Launch --')
print('############################ \n')

temps_ini_import = time.time()
combination             = pd.read_excel('tests_combination.xlsx', sheetname='tests')
test_number             = len(combination)
variables               = list(combination)

ansys_path = pd.read_excel('tests_combination.xlsx', sheetname='path')
ansys_folder = ansys_path['Ansys'][0]

temps_fin_import = time.time()

chemin_dossier_general = os.getcwd()

#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                     FOLDER CREATION
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

#-- Creation fichiers variables et Copier-Coller des fichiers Ansys et Python
temps_ini_dossiers_etudes = time.time()
for essai_number in range(test_number):
    print('ETUDE ' + str(essai_number + 1) + ' sur ' + str(test_number) + ' : ')
    
    #--Création des dossiers
    chemin_etude = chemin_dossier_general + '\Etude_' + str(essai_number + 1)
    if not os.path.exists(chemin_etude):
        os.makedirs(chemin_etude)
    print('... Création des études - création dossier : done')
    
    #--Copie des fichiers
    dossier_emetteur        = chemin_dossier_general + '\Dossier_initial'
    chemin_etude            = chemin_dossier_general + '\Etude_' + str(essai_number + 1) + '/'
    copy_tree(dossier_emetteur, chemin_etude)
    
    print('... Création des étude - copie des fichiers : done')
    
    
    #--Création fichier variables.mac
    file = open(chemin_etude + 'variables' + '.mac','w') 
    for variable in variables :
        file.write(variable + ' = ' + str(combination[variable][essai_number]) + 'E-03 \n')
    file.close()
    
    print('... Création des étude - création du fichier variables.mac : done')
    
    
print('--> CREATION DES ETUDES : OK \n')
temps_fin_dossiers_etudes = time.time()


#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#                  FICHIER BAT LANCEMENT AUTOMATIQUE ANSYS
#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

temps_ini_bat = time.time()
fichier_ansys = open(chemin_dossier_general + '\\' + 'lancement_ansys.bat','w')
fichier_ansys.write('TITLE = CALCULS ANSYS')
fichier_ansys.write('\n')
fichier_ansys.write('@echo off' + '\n')
for i in range(test_number):
    chemin_etude            = chemin_dossier_general + '\Etude_' + str(i + 1)
    fichier_ansys.write('echo Lancement du calcul ' + str(i+1) + '\n')
    fichier_ansys.write('echo Calcul ' + str(i+1) + ' en cours ...' + '\n')
    fichier_ansys.write(r'call ' + str(ansys_folder) + ' -b -dir ' + chemin_etude + '" -i "' + chemin_etude + '\main.mac" -o "' + chemin_etude + '\RUN.out"' + '\n')
    fichier_ansys.write('echo Calcul ' + str(i+1) + ' fini' + '\n')
    fichier_ansys.write('echo' + '\n')
    fichier_ansys.write('\n')
    fichier_ansys.write('\n')
fichier_ansys.close()

print('--> CREATION DU FICHIER .bat : OK \n')
temps_fin_bat = time.time()

print("Le temps d'import des données est de : " + str(round(temps_fin_import - temps_ini_import,3)) + " sec")
print("Le temps de création des dossiers est de : " + str(round(temps_fin_dossiers_etudes - temps_ini_dossiers_etudes,3)) + " sec")
print("Le temps de création du fichier .bat est de : " + str(round(temps_fin_bat - temps_ini_bat,3)) + " sec")
print("--> Le temps total est de : " + str(round(temps_fin_bat - temps_ini_import,3)) + " sec")

print('############################')
print('--    FIN DU PROGRAMME    --')
print('############################ \n')