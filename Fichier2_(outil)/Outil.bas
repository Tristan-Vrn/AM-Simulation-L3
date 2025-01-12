Attribute VB_Name = "Outil"
Option Explicit
Option Base 1

'Déclaration desvariables publiques
Public wsA As Worksheet
Public wsO As Worksheet
Public wsCours As Worksheet
Public nbAct As Integer
Public nbrow As Integer
Public adresse As Variant

Sub Outil()

'Déclaration ds variables
Dim date_ As Variant
Dim date_2 As Variant
Dim nbObli As Integer
Dim valeur_port As Double

'Déclaration des variables liées aux classes et des collections
Dim Action As ActionClass
Dim CompoAction As New Collection
Dim Obli As ObliValeur
Dim CompoObli As New Collection

'Déclaration des variables itératives
Dim i As Integer
Dim j As Integer
Dim ligne As Integer
Dim cellule As Range

'Initialisation des feuilles de calcul du classeur
Set wsA = ThisWorkbook.Worksheets(1)
Set wsO = ThisWorkbook.Worksheets(2)
Set wsCours = ThisWorkbook.Worksheets("Cours")

'Convervation de l'ancienne date saisie dans la cellule 13L et dans date_2
wsO.Cells(13, 12).Value = wsO.Cells(14, 12).Value
date_2 = wsO.Cells(14, 12).Value

'Demande de nouvelle date à saisir par l'utilisateur
date_ = InputBox("Veuillez renseigner une date au format MM/DD/YYYY: ", "Date", "07/07/2005")
wsO.Cells(14, 12).Value = date_

'Changement du format de la date en nombre
wsO.Cells(14, 12).NumberFormat = "0"

'Affectation du nombre de cours enregistrés à nbrow
nbrow = wsCours.Cells(wsCours.Rows.Count, 1).End(xlUp).Row - 1

'Changement du format des dates des cours en nombres
wsCours.Columns(1).NumberFormat = "0"

'Affectation du nombre d'actions à nbAct
nbAct = wsCours.Cells(1, wsCours.Columns.Count).End(xlToLeft).Column - 1

'Boucle sur les titres
For i = 1 To nbAct

    'Calcul de la valeur d'une action du titre à l'ancienne date rentrée
    wsA.Cells(7, 1 + i).Value = wsA.Cells(5, 1 + i).Value
    'Attribution d'un nom à la ligne
    wsA.Cells(7, 1).Value = wsA.Cells(5, 1).Value & " (ancienne date rentrée)"
    'Initialisation de la classe Action
    Set Action = New ActionClass
    'Initialisation de adresse à la date demandée par l'utilisateur
    Set adresse = wsCours.Columns(1).Find(What:=wsO.Cells(14, 12).Value, LookIn:=xlValues)
    
    'Condition pour que si la date demandée ne correspond pas à une date enregistrée, ce soit la dernière date enregistrée dans Cours qui soit prise en compte
    If adresse Is Nothing Then
        Do While adresse Is Nothing
             wsO.Cells(14, 12).Value = wsO.Cells(14, 12).Value - 1
             Set adresse = wsCours.Columns(1).Find(What:=wsO.Cells(14, 12).Value, LookIn:=xlValues)
        Loop
    End If
    
    'Condition pour le cas où il n'y a pas de valeur enregistrée à la date demandée
    If wsCours.Cells(adresse.Row, i + 1).Value = "" Then
    Else
        
        'Recopie de la valeur des titres dans la feuille Action
        If Not adresse Is Nothing Then
            wsA.Cells(5, i + 1).Value = wsCours.Cells(adresse.Row, 1).Offset(0, i).Value
        
        End If
    End If
    
    'Attribution des caractéristiques de l'Action avec InitAction
    Call Action.InitAction(wsA.Cells(1, 1 + i).Value, wsA.Cells(3, 1 + i).Value, wsA.Cells(4, 1 + i).Value, wsA.Cells(2, 1 + i).Value, wsA.Cells(5, 1 + i).Value, wsA.Cells(7, 1 + i).Value)
    'Calcul du rendement du titre
    wsA.Cells(6, 1 + i).Value = Action.Rend1
    
    'Calcul du rendement entre la date saisie précédemment et la nouvelle
    If date_2 > date_ Then
        wsA.Cells(8, 1 + i).Value = Action.Rend2
    Else
        wsA.Cells(8, 1 + i).Value = Action.Rend3
    End If
    
    'Ajout de l'action à la collection
    CompoAction.Add Action
    'Calcul de la valeur par poche à la date saisie
    wsA.Cells(10, 1 + i).Value = wsA.Cells(5, 1 + i).Value * wsA.Cells(3, 1 + i).Value
    'Titre de la ligne
    wsA.Cells(10, 1).Value = "Valeur par poche au " & date_
    'Somme
    wsA.Cells(10, nbAct + 2).Value = Application.WorksheetFunction.Sum(wsA.Cells(10, 2).Resize(1, nbAct))
    'Calcul de la valeur par poche à la date saisie précedemment
    wsA.Cells(11, 1 + i).Value = wsA.Cells(7, 1 + i).Value * wsA.Cells(3, 1 + i).Value
    'Titre de la ligne
    wsA.Cells(11, 1).Value = "Valeur par poche au " & date_2
    'Somme
    wsA.Cells(11, nbAct + 2).Value = Application.WorksheetFunction.Sum(wsA.Cells(11, 2).Resize(1, nbAct))
Next i

'Attribution du nombre d'obligations à nbObli
nbObli = wsO.Cells(wsO.Rows.Count, 9).End(xlUp).Row - 1

'Boucle sur les obligations
For i = 1 To nbObli
    
    'Initialisation de la classe Obli
    Set Obli = New ObliValeur
    'Affectation des caractéristique de l'obligation avec InitOblig
    Call Obli.InitOblig(wsO.Cells(1 + i, 3).Value, wsO.Cells(1 + i, 1).Value, wsO.Cells(1 + i, 4).Value, wsO.Cells(1 + i, 5).Value, wsO.Cells(1 + i, 6).Value, wsO.Cells(1 + i, 7).Value, wsO.Cells(1 + i, 2).Value, "30/12/2004", wsO.Cells(14, 12).Value)
    'Calcul de la valeur actuelle de l'obligation
    wsO.Cells(15 + i, 3).Value = Obli.ValeurAct
    
    'Calcul des cash flows générés parl'obligation à la date demandée et formatage des cellules
    With wsO.Cells(15 + i, 2)
        .Value = Obli.CashFlows
        .NumberFormat = "0.00"
    End With
    
    'Titres des obligations
    wsO.Cells(i + 15, 1).Value = "Obligation " & i
    
    'Calcul des durations de Macaulay et modifiée
    wsO.Cells(i + 15, 4).Value = Obli.MaucaulayDuration
    wsO.Cells(i + 15, 5).Value = Obli.ModifiedMacaulayDuration
    
    'Affectation des caractéristiques de l'obligation InitOblig avec l'ancienne date saisie
    Call Obli.InitOblig(wsO.Cells(1 + i, 3).Value, wsO.Cells(1 + i, 1).Value, wsO.Cells(1 + i, 4).Value, wsO.Cells(1 + i, 5).Value, wsO.Cells(1 + i, 6).Value, wsO.Cells(1 + i, 7).Value, wsO.Cells(1 + i, 2).Value, "30/12/2004", wsO.Cells(13, 12).Value)
    'Calcul de la valeur actuelle de l'obligation à l'ancienne date saisie
    wsO.Cells(29 + i, 3).Value = Obli.ValeurAct

    'Calcul des cash flows générés parl'obligation à l'ancienne date saisie et formatage des cellules
     With wsO.Cells(29 + i, 2)
        .Value = Obli.CashFlows
        .NumberFormat = "0.00"
    End With
    
    'Calcul des durations de Macaulay et modifiée à l'ancienne date saisie
    wsO.Cells(i + 29, 1).Value = "Obligation " & i
    wsO.Cells(i + 29, 4).Value = Obli.MaucaulayDuration
    wsO.Cells(i + 29, 5).Value = Obli.ModifiedMacaulayDuration
    
    'Formatage des cellules
    wsO.Cells(16, 2).Resize(nbObli * 5, 4).NumberFormat = "0.00"
    
    'Calcul descash flows perçus entre l'anncienne date saisie et la nouvelle
    If wsO.Cells(13, 12).Value < wsO.Cells(14, 12).Value Then
        wsO.Cells(29 + i, 6).Value = wsO.Cells(15 + i, 2).Value - wsO.Cells(29 + i, 2).Value
    Else
        wsO.Cells(29 + i, 6).Value = wsO.Cells(29 + i, 2).Value - wsO.Cells(15 + i, 2).Value
    End If
    
    'Ajout de l'obligation à la collection
    CompoObli.Add Obli
    
Next i

'Titre colonne F d'évolution des cash flows
If wsO.Cells(13, 12).Value < wsO.Cells(14, 12).Value Then
    wsO.Cells(29, 6).Value = "Cash flow perçus entre " & date_ & " et " & date_2
Else
    wsO.Cells(29, 6).Value = "Cash flow perçus entre " & date_2 & " et " & date_
End If

'Titres des lignes du tableau des dates colonne K
wsO.Cells(13, 11).Value = "Ancienne date saisie"
wsO.Cells(14, 11).Value = "Nouvelle date saisie"

'Titres colonne B
wsO.Cells(15, 2).Value = "Cash flows générés par l'obligation"
wsO.Cells(29, 2).Value = "Cash flows générés par l'obligation"

'Titre colonne C
wsO.Cells(15, 3).Value = "Valeur au " & date_
wsO.Cells(29, 3).Value = "Valeur au " & date_2

'Titres colonnes D et E
wsO.Cells(15, 4).Value = "Macaulay Duration"
wsO.Cells(15, 5).Value = "Modified Macaulay duration"
wsO.Cells(29, 4).Value = "Macaulay Duration"
wsO.Cells(29, 5).Value = "Modified Macaulay duration"

'Remise en frome des cellules de dates au format date
wsCours.Columns(1).NumberFormat = "dd/mm/yyyy"
wsO.Cells(14, 12).NumberFormat = "dd/mm/yyyy"
wsO.Cells(13, 12).NumberFormat = "dd/mm/yyyy"

'Titre des lignes de Total
wsO.Cells(nbObli + 16, 1).Value = "Total"
wsO.Cells(nbObli + 29, 1).Value = "Total"

'Calcul des totaux des tableaux
wsO.Cells(nbObli + 16, 2).Value = Application.WorksheetFunction.Sum(wsO.Cells(16, 2).Resize(nbObli, 1))
wsO.Cells(nbObli + 16, 3).Value = Application.WorksheetFunction.Sum(wsO.Cells(16, 3).Resize(nbObli, 1))

'Ajustement des colonnes
wsO.Columns.AutoFit

'Calcul des totaux du tableau des titres de la feuille Composition actions et mise en forme des titres des lignes
wsA.Cells(5, 1).Value = "Valeur au " & date_
wsA.Cells(5, nbAct + 2).Value = Application.WorksheetFunction.Sum(wsA.Cells(5, 2).Resize(1, nbAct))
wsA.Cells(7, nbAct + 2).Value = Application.WorksheetFunction.Sum(wsA.Cells(7, 2).Resize(1, nbAct))
wsA.Cells(6, 1).Value = "Rendement du 01/01/2005 au " & date_
wsA.Cells(6, nbAct + 2).Value = Application.WorksheetFunction.Average(wsA.Cells(6, 2).Resize(1, nbAct))
wsA.Cells(8, 1).Value = "Rendement entre la dernière date rentrée le " & date_
wsA.Cells(8, nbAct + 2).Value = Application.WorksheetFunction.Average(wsA.Cells(8, 2).Resize(1, nbAct))


'Mise en forme des cellules de la feuille Composition actions
For Each cellule In wsA.UsedRange
    cellule.NumberFormat = "0.00"
Next cellule
wsA.Columns.AutoFit

'Mise en page tableau budget
wsO.Cells(2, 11).Resize(3, 2).Borders.Color = RGB(0, 0, 0)
With wsO.Cells(2, 11).Resize(3, 1)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With

'Mise en page tableau des dates saisies
wsO.Cells(13, 11).Resize(2, 2).Borders.Color = RGB(0, 0, 0)
With wsO.Cells(13, 11).Resize(2, 1)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With

'Mise en page du tableau des statistiques des obligations à la date rentrée
wsO.Cells(15, 1).Resize(12, 5).Borders.Color = RGB(0, 0, 0)
With wsO.Cells(15, 2).Resize(1, 4)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With
With wsO.Cells(16, 1).Resize(11, 1)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With

'Mise en page du tableau des statistiques des obligations à la date rentrée précédemment
wsO.Cells(29, 1).Resize(11, 6).Borders.Color = RGB(0, 0, 0)
With wsO.Cells(29, 2).Resize(1, 5)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With
With wsO.Cells(30, 1).Resize(10, 1)
    .Font.Bold = True
    .Interior.Color = RGB(224, 224, 224)
End With

'Calcul et affichage de la valeur du portefeuille à la date saisie
valeur_port = wsA.Cells(10, nbAct + 2).Value + wsO.Cells(nbObli + 16, 3).Value
MsgBox "La valeur de votre portefeuille au " & date_ & " est " & valeur_port & " €"

'Appel de la fonction volat du module Volatilité pour calculer les volatiltiés des titres et du portefeuille - feuille Composition actions
Call volat

End Sub
