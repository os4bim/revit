# -*- coding: utf-8 -*-

__doc__ = "Création automatique d'un bordereau de prix pour tuyauterie\nPrérequis :\n- Ouvrir une feuille Excel"
__title__ = 'Export\nPipes'
__author__ = 'Yoann OBRY'

#BOM to Excel Pipes v2.0


from Autodesk.Revit.DB import *
import System
from System import Guid
import math
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document


#Fenêtre de confirmation
res = forms.alert("Le quantitatif va être exporté sur la feuille Excel active de votre espace de travail.\n"
                  "Voulez-vous continuer ?",
                  yes=True, no=True, exitscript=True)


#Shared parameter code circuit
code_cir = Guid(r'55934d0c-0246-4ce2-9bdf-57ed4244e11b')

#Shared parameter FMF_Angle
angle = Guid(r'a8b84336-4f16-462c-a50f-f0f8b2e4f7c2')

### PA : Création d'un BOM de PIPE ACCESSORIES sous forme de liste de tuple

#Collecte les Pipe Accessories
PAs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeAccessory).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
PA_code_circuit = []
PA_family_name = []
PA_description = []
PA_size = []

for PA in PAs:

	
	## Get Type Parameter value
    PA_type = doc.GetElement(PA.GetTypeId())
	
	# Element ID - Instance Parameter
	#print PA.Id

	# Code circuit - Instance Parameter (Shared Parameter)
    code_circuit = PA.get_Parameter(code_cir).AsString()
    if code_circuit == None or code_circuit == '':
        code_circuit = '_N/A'
    PA_code_circuit.append(code_circuit)

	# Family Name - Type Parameter
    family_name = PA_type.get_Parameter(
					BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
    PA_family_name.append(family_name.AsString())

	# Description - Type Parameter
    description = PA_type.get_Parameter(
					BuiltInParameter.ALL_MODEL_DESCRIPTION).AsString()
    if description == None:
        description = ''
    PA_description.append(description)
	
	# Size - Instance Parameter
    size = PA.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
    PA_size.append(size.AsString())


## Assemblage des listes de caractéristiques en une seule
PA_libelle = [PA_family_name[i] +"  "+ PA_description[i] +"  "+ PA_size[i] for i in range(len(PA_code_circuit))]


## Créer une liste par élément avec unité de mesure et count=1
lstPA = [[PA_code_circuit[i],PA_libelle[i],'u',1] for i in range(len(PA_code_circuit))]

## Compte le nombre d'éléments identique
PAcount=[]
for i in range(len(lstPA)):
    PAcount.append(lstPA.count(lstPA[i]))
## Incrémente les quantité tout en conservant les doublons
for i in range(len(lstPA)):
    lstPA[i][3]=PAcount[i]
    
## Supprime les doublons
setPA=set(tuple(row) for row in lstPA)
lstPA=list(setPA)
lstPA.sort()


print(lstPA)


### PI : Création d'un BOM de PIPE SEGMENTS sous forme de liste de tuple

#Collecte les Pipes
PIs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
PI_code_circuit = []
PI_type_name = []
PI_size = []
PI_length = []

for PI in PIs:

	
	## Get Type Parameter value
	PI_type = doc.GetElement(PI.GetTypeId())
	
	# Element ID - Instance Parameter
	#print PI.Id

	# Code circuit - Instance Parameter (Shared Parameter)
	code_circuit = PI.get_Parameter(code_cir).AsString()
	if code_circuit == None or code_circuit == '':
		code_circuit = '_N/A'
	PI_code_circuit.append(code_circuit)

	# Type Name - Type Parameter
	type_name = PI_type.get_Parameter(
					BuiltInParameter.SYMBOL_NAME_PARAM)
	PI_type_name.append(type_name.AsString())

	# Size - Instance Parameter
	size = PI.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
	PI_size.append(size.AsString())

	# Length - Instance Parameter
	length = PI.get_Parameter(
					BuiltInParameter.CURVE_ELEM_LENGTH)
	PI_length.append(length.AsDouble())


## Assemblage des listes de caractéristiques en une seule
PI_libelle = [PI_type_name[i] +"  "+ PI_size[i] for i in range(len(PI_code_circuit))]


## Créer une liste par élément avec unité de mesure et métré total
lstPI = [[PI_code_circuit[i],PI_libelle[i],PI_length[i]/3.2808] for i in range(len(PI_code_circuit))]

lstPI_unique = list(set([(element[0],element[1]) for element in lstPI]))

quantites = [sum([float(part[2]) for part in lstPI if (part[0],part[1]) == element]) for element in lstPI_unique]

lstPI = [list(lstPI_unique[element])+["{:01.1f}".format(quantites[element])] for element in range(0,len(lstPI_unique))]
lstPI = [[lstPI[i][0],lstPI[i][1],'m',lstPI[i][2]] for i in range(len(lstPI))]

lstPI.sort()

print(lstPI)


### PF : Création d'un BOM de PIPE FITTINGS sous forme de liste de tuple

#Collecte les Pipe Fittings
PFs = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeFitting).WhereElementIsNotElementType().ToElements()

#Créer des listes vides
PF_code_circuit = []
PF_family_name = []
PF_type_name = []
PF_size = []
PF_angle = []

for PF in PFs:

	
	## Get Type Parameter value
	PF_type = doc.GetElement(PF.GetTypeId())
	
	# Element ID - Instance Parameter
	#print PF.Id

	# Code circuit - Instance Parameter (Shared Parameter)
	code_circuit = PF.get_Parameter(code_cir).AsString()
	if code_circuit == None or code_circuit == '':
		code_circuit = '_N/A'
	PF_code_circuit.append(code_circuit)

	# Family Name - Type Parameter
	family_name = PF_type.get_Parameter(
					BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
	PF_family_name.append(family_name.AsString())

	# Type Name - Type Parameter
	type_name = PF_type.get_Parameter(
					BuiltInParameter.SYMBOL_NAME_PARAM)
	PF_type_name.append(type_name.AsString())
	
	# Size - Instance Parameter
	size = PF.get_Parameter(
					BuiltInParameter.RBS_CALCULATED_SIZE)
	PF_size.append(size.AsString())

	# Angle	- Instance Parameter (Shared Parameter)
	angle_coude = PF.get_Parameter(angle)
	if angle_coude:
		angle_coude = angle_coude.AsDouble()
        # Arrondi les angles des pipes fittings
		if 85 <= angle_coude <= 95:
			angle_coude = 90
		elif 55 <= angle_coude <= 65:
			angle_coude = 60
		elif 40 <= angle_coude <= 50:
			angle_coude = 45
		elif 25 <= angle_coude <= 35:
			angle_coude = 30
		elif 15 <= angle_coude <= 25:
			angle_coude = 20
      
		PF_angle.append(angle_coude * 180 / math.pi)
        
	else:
		PF_angle.append(0)



## Assemblage des listes de caractéristiques en une seule
PF_libelle = [PF_family_name[i] +"  "+ PF_type_name[i] +"  "+ PF_size[i] +"  "+ str("{:01.0f}".format(5 * round(PF_angle[i])/5)) +"°" for i in range(len(PF_code_circuit))]

## Efface les angles nuls dans le libellé
PF_libelle = [w.replace('  0°','') for w in PF_libelle]


## Créer une liste PF d'éléments avec unité de mesure et count=1
lstPF = [[PF_code_circuit[i],PF_libelle[i],'u',1] for i in range(len(PF_code_circuit))]

## Compte le nombre d'éléments identique
PFcount=[]
for i in range(len(lstPF)):
    PFcount.append(lstPF.count(lstPF[i]))
## Incrémente les quantité tout en conservant les doublons
for i in range(len(lstPF)):
    lstPF[i][3]=PFcount[i]
    
## Supprime les doublons
setPF=set(tuple(row) for row in lstPF)
lstPF=list(setPF)
lstPF.sort()

print(lstPF)



##Ajout des codes circuit manquant dans les categories pour l'écriture
#dans Excel

# Identification des codes circuits
circuit_unique = sorted(set(PA_code_circuit + PI_code_circuit + PF_code_circuit))

#Codes circuit manquants
def elements_absents(circuit_unique, lstPA):
    # Crée un ensemble (set) à partir des éléments de lstPA
    lstPA_elements = set(item[0] for item in lstPA)
    
    # Filtrage des éléments de circuit_unique qui ne sont pas dans lstPA
    elements_absents = [element for element in circuit_unique if element not in lstPA_elements]
    
    return elements_absents


code_absent_PA = elements_absents(circuit_unique, lstPA)
code_absent_PI = elements_absents(circuit_unique, lstPI)
code_absent_PF = elements_absents(circuit_unique, lstPF)


#Mise à jour de la liste des PA, PI et PF
def update_lst(code_absent,lstP):
    if len(code_absent) > 0:
        for i,item in enumerate(code_absent):
            lstP.append((code_absent[i],'N/A','N/A',0))
    
    lstP.sort()
    return update_lst
            
update_lst(code_absent_PA,lstPA)
update_lst(code_absent_PI,lstPI)
update_lst(code_absent_PF,lstPF)



		### Exporter les données dans Excel ###
 
#Accessing the Excel applications.
xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
 
#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 1
columnStart = 1
 
#Effacer la feuille excel
for i in range(100):
	for j in range(10):
		data = xlApp.Worksheets(worksheet).Cells(rowStart + i, columnStart + j)
		data.Value = ""
 
#Compteur de lignes excel
count_circuit = 0
saut_ligne = 0

#Fonction qui permet à i de commencer à 0 pour l'écriture des circuits suivants
def find(c,d):
	return [(i, premier.index(c)) for i, premier in enumerate(d) if c in premier]

	

for k, item in enumerate(circuit_unique):

	count_lstPA = 0
	count_lstPI = 0
	count_lstPF = 0

	## Numéro Circuit
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart + 1)
	data.Value = "Circuit - " + circuit_unique[k]


	## Ecriture des Pipe Accessories

	# Titre
	saut_ligne += 2
	
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + saut_ligne, columnStart + 1)
	data.Value = "Robinetterie et instrumentation"

	# Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstPA)[0][0]
	for i,item in enumerate(lstPA):

		if lstPA[i][0] == circuit_unique[k]:
			#Worksheet object specifying the cell location.
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 6)
			#Assigning a value to the cell.
			data.Value = lstPA[i][0]
		
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 1)
			data.Value = lstPA[i][1]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 2)
			data.Value = lstPA[i][2]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + i - decal + saut_ligne, columnStart + 3)
			data.Value = lstPA[i][3]

			count_lstPA += 1



	## Ecriture des Pipes

	# Titre
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.2
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne, columnStart + 1)
	data.Value = "Canalisations"

	#Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstPI)[0][0]
	for i,item in enumerate(lstPI):

		if lstPI[i][0] == circuit_unique[k]:

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne + i - decal, columnStart + 6)
			data.Value = lstPI[i][0]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne + i - decal, columnStart + 1)
			data.Value = lstPI[i][1]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne + i - decal, columnStart + 2)
			data.Value = lstPI[i][2]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + saut_ligne + i - decal, columnStart + 3)
			data.Value = lstPI[i][3]
			
			count_lstPI += 1


	## Ecriture des Pipe Fittings
	 
	# Titre
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne, columnStart)
	data.Value = range(len(circuit_unique))[k] + 1 + 0.3
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne, columnStart + 1)
	data.Value = "Raccords"

	#Eléments
	saut_ligne += 1
	decal = find(circuit_unique[k],lstPF)[0][0]
	for i,item in enumerate(lstPF):

		if lstPF[i][0] == circuit_unique[k]:

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne + i - decal, columnStart + 6)
			data.Value = lstPF[i][0]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne + i - decal, columnStart + 1)
			data.Value = lstPF[i][1]

			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne + i - decal, columnStart + 2)
			data.Value = lstPF[i][2]
		 
			data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + saut_ligne + i - decal, columnStart + 3)
			data.Value = lstPF[i][3]
			
			count_lstPF += 1

	## Sous total
	saut_ligne += 1
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + count_lstPF + saut_ligne, columnStart)
	data.Value = "ST" + str(range(len(circuit_unique))[k] + 1)
	data = xlApp.Worksheets(worksheet).Cells(rowStart + count_circuit + count_lstPA + count_lstPI + count_lstPF + saut_ligne, columnStart + 1)
	data.Value = "Total " + str(range(len(circuit_unique))[k] + 1) + " sous poste"

	count_circuit += count_lstPA + count_lstPI + count_lstPF
	saut_ligne += 2
	

##Afficher une console pour maintenance
#from rpw.ui.forms import Console
#Console(context=locals())
