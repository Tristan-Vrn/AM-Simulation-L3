Attribute VB_Name = "Actions"
Option Explicit
Option Base 1
Sub Selection()

'DŽclaration des feuilles de calcul
Dim wsD As Worksheet

'DŽclaration des variables itŽratives
Dim i As Integer
Dim j As Integer
Dim k As Integer

'DŽclaration des autres variables
Dim nbCols As Integer
Dim nbCols2 As Integer
Dim ws As Worksheet
Dim titre As String
Dim adresse As Variant
Dim nbre_titres As Integer
Dim budget_tot As Long
Dim budget_titre As Long
Dim parts As Long
Dim nbre_parts As Long
Dim budget_investi As Double

Set wsD = ThisWorkbook.Worksheets("Actions")

'On fixe nbCols au nombres de colonnes du tableau sans compter la colonne Runs, soit aux nombres de titres o l'investissement est possible
nbCols = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
j = 2

'Boucle pour sŽlŽctionner les titres ayant au moins 75% de leurs runs rŽpartis entre 3 et 7, et qui les recopie dans la ligne "Titres"
For i = 2 To nbCols
    If wsD.Cells(11, i).Value >= 0.075 Then
        wsD.Cells(19, j).Value = wsD.Cells(1, i).Value
        j = j + 1
    End If
Next i
wsD.Cells(19, 1).Value = "Titre"

'Recherche de la valeurs de chaque titre sŽlŽctionnŽ parmis les 4 feuilles de donnŽes ˆ la date 30/12/2004, dernire valeur en dateau 01/01/2005
nbCols2 = wsD.Cells(19, wsD.Columns.Count).End(xlToLeft).Column
For i = 1 To nbCols2
    titre = wsD.Cells(19, i + 1).Value
    For k = 1 To 4
        Set ws = ThisWorkbook.Worksheets(1 + k)
        Set adresse = ws.Rows(1).Find(what:=titre, LookIn:=xlValues, LookAt:=xlWhole)
        If Not adresse Is Nothing Then
            wsD.Cells(20, i + 1).Value = adresse.Offset(53, 0).Value
        End If
    Next k
Next i

'Mise en forme de la ligne des valuers des titres
wsD.Cells(20, 2).Resize(1, nbCols2).NumberFormat = "0.000000"
wsD.Cells(20, 1).Value = "Valeur au 30/12/2004"

'On fixe le nombre de parts et le budget investi pour chaque titre.
'Comme nous voulons un portefeuille avec environ autant de titres d'actions que d'obligations, on augmente les parts jusqu'ˆ ce que la valeur totale soit supŽrieure ou Žgale ˆ 500 000
'On veut que les 10 titres soient ŽquipondŽrŽs - soit environ 50 000 par titre.
 budget_tot = 500000
 nbre_titres = nbCols2 - 1
 budget_titre = budget_tot / nbre_titres
 For i = 2 To nbCols2
    parts = 0
    nbre_parts = 0
    Do
        parts = wsD.Cells(20, i).Value + parts
        nbre_parts = nbre_parts + 1
    Loop Until parts >= budget_titre
    wsD.Cells(21, i).Value = nbre_parts
    wsD.Cells(22, i).Value = parts
Next i

'Mise en place et calcul de la colonne Total en fin de tableau
wsD.Cells(19, nbCols2 + 1).Value = "Total"
wsD.Cells(19, nbCols2 + 1).Font.Bold = True
wsD.Cells(21, nbCols2 + 1).Value = Application.WorksheetFunction.Sum(wsD.Cells(21, 2).Resize(1, nbCols2 - 1))
budget_investi = Application.WorksheetFunction.Sum(wsD.Cells(22, 2).Resize(1, nbCols2 - 1))
wsD.Cells(22, nbCols2 + 1).Value = budget_investi


'Mise en forme des lignes des "Nombres de parts" et "Budget investi" et du tableau
wsD.Cells(21, 2).Resize(1, nbCols2).NumberFormat = "0"
wsD.Cells(21, 1).Value = "Nombre de parts"
wsD.Cells(22, 2).Resize(1, nbCols2).NumberFormat = "0.00"
wsD.Cells(22, 1).Value = "Budget investi"
wsD.Cells(19, 1).Resize(4, 1).Font.Bold = True
wsD.Cells(19, 1).Resize(4, 1).Interior.Color = RGB(255, 192, 160)
wsD.Cells(19, 1).Resize(4, nbCols2 + 1).Borders.Color = RGB(0, 0, 0)
ActiveWindow.DisplayGridlines = False
wsD.Columns.AutoFit

'Appel de la macro Obli du module Obligations avec pour argument le montant du budget restant ˆ investir en obligations
Call Obli(1000000 - budget_investi)

End Sub
