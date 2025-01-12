Attribute VB_Name = "Importation"
Option Explicit
Option Base 1

Sub importation()

'Déclaration des classeurs
Dim wbSource As Workbook

'Déclaration des feuilles de calcul
Dim wsA As Worksheet
Dim wsO As Worksheet
Dim wsFS As Worksheet
Dim wsI As Worksheet
Dim wsR As Worksheet
Dim wsTS As Worksheet
Dim wsCours As Worksheet
Dim ws As Worksheet

'Déclaration des variables itératives
Dim j As Integer
Dim i As Integer
Dim k As Integer

'Déclaration des autres variables
Dim adresse As String
Dim nbCol As Long

'Initialisation des feuilles de calcul du fichier VBAProjet
adresse = Application.GetOpenFilename
Set wbSource = Workbooks.Open(adresse)
Set wsA = wbSource.Sheets("Actions")
Set wsFS = wbSource.Sheets("FrenchStocks")
Set wsI = wbSource.Sheets("Indexes")
Set wsTS = wbSource.Sheets("TechStocks")
Set wsR = wbSource.Sheets("Rates")
Set wsO = wbSource.Sheets("Obligations")

'Ajout des feuilles initialisées au nouveau classeur et affactations de nouveaux noms à Actions et Obligations
wsA.Copy after:=ThisWorkbook.Worksheets(1)
ThisWorkbook.Worksheets(2).Name = "Composition actions"
wsO.Copy after:=ThisWorkbook.Worksheets(2)
ThisWorkbook.Worksheets(3).Name = "Composition obligations"
wsFS.Copy after:=ThisWorkbook.Worksheets(3)
wsI.Copy after:=ThisWorkbook.Worksheets(4)
wsTS.Copy after:=ThisWorkbook.Worksheets(5)
wsR.Copy after:=ThisWorkbook.Worksheets(6)

'Fermeture de VBAProject
wbSource.Close

'Retrait des messages de fermeture
Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(1).Delete
Application.DisplayAlerts = True

'Suppression du tableau des runs et des rendements pour ne garder que les titres duportefeuille et leurs caractéristiques
ThisWorkbook.Worksheets(1).Rows("1:18").Delete

'Initialisation des feuilles du fichier Outil_projet
Set wsA = ThisWorkbook.Worksheets(1)
Set wsFS = ThisWorkbook.Worksheets(3)
Set wsI = ThisWorkbook.Worksheets(4)
Set wsTS = ThisWorkbook.Worksheets(5)
Set wsR = ThisWorkbook.Worksheets(6)

'Ajout et initialisation d'une nouvelle feuille
ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(6)
Set wsCours = ThisWorkbook.Worksheets(7)

'Boucle pour itérer sur les 10 titres
For k = 1 To 10

    'Boucle pour itérer sur les 4 feuilles de titres
    For i = 1 To 4
        
        'Initialisation de la feuille prise en compte
        Set ws = ThisWorkbook.Worksheets(2 + i)
        
        'Affectation du nombre de titres dans la feuille à nbCol
        nbCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column - 1
        
        'Boucle pour itérer sur les titres de la feuille
        For j = 1 To nbCol
        
            'Condition pour recopier le cours du titre sur une nouvelle feuille si il fait partie du portefeuille
            If ws.Cells(1, j + 1).Value = wsA.Cells(1, 1 + k).Value Then
                wsCours.Columns(1 + k).Value = ws.Columns(1 + j).Value
            End If
        Next j
        
    Next i
Next k

'Recopie des dates en colonne 1
wsCours.Columns(1).Value = wsFS.Columns(1).Value

'Suppression des 4 feuilles de cours initiales etsuppression des messages d'erreur qui s'en suivent
Application.DisplayAlerts = False
    wsFS.Delete
    wsI.Delete
    wsTS.Delete
    wsR.Delete
Application.DisplayAlerts = True

'Suppression des cours jusqu'au 30/12/2004, dernière valeur en date à l'achat des titres
wsCours.Rows("2:53").Delete

'Affectation dunom Cours à la nouvelle feuille
wsCours.Name = "Cours"

'Mise en page des titres des colonnes de la feuille Cours
With wsCours.Cells(1, 1).Resize(1, 11)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With

End Sub
