# Python scripts for Revit

🗣️ FR

Ces fichiers sont utilisé via pyrevit. Pour les utiliser via RevitPythonShell supprimer ces lignes :
```
__doc__ =
__title__ =
__author__ =
```

### BOM_to_Excel_Ducts_script.py et BOM_to_Excel_Pipes_script.py

Ces scripts permettent d'extraire un quantitatif de gaines ou de tuyauterie depuis Revit pour l'écrire dans un fichier Excel, sans passer par les nomenclatures.
Cela permet de créer rapidement un bordereau de prix pour envoyer à un fournisseur.

Avant de lancer le script il faut ouvrir un fichier excel dans lequel le quantitatif sera écrit.

### Show_Parameters.py

Ce script vous permet de récupérer les noms des paramètres du premier objet Revit présent dans votre fichier. Pour ce faire vous pouvez ne laisser qu'un seul objet dans votre projet Revit pour récupérer ces données.
A noter la différence entre les éléments "type" et les éléments "instance" ou "non type". Ceci est important car il faut ajouter une ligne de code pour que le script récupère les données des élements "type".

Le fichier excel PA_Parameters_name.xlsx est un exemple de noms de paramètres Revit pour un "Pipe Accessories"

🗣️ EN

This files are used in pyrevit. For using in RevitPythonShell just delete this lines :
```
__doc__ =
__title__ =
__author__ =
```

### BOM_to_Excel_Ducts_script.py et BOM_to_Excel_Pipes_script.py

These scripts allow you to extract a duct or piping quantity from Revit and write it to an Excel file, without going through the schedules.
This allows you to quickly create a price list to send to a supplier.

Before running the script, an excel file must be opened in which the quantity will be written.

### Show_Parameters.py

This script allows you to retrieve the parameter names of the first Revit object in your file. To do this you can leave only one object in your Revit project to retrieve this data.
Note the difference between "type" elements and "instance" or "non-type" elements. This is important because you have to add a line of code so that the script can retrieve the data of the "type" elements.

The excel file PA_Parameters_name.xlsx is an example of Revit parameter names for a "Pipe Accessories".
