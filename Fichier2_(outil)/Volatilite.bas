Attribute VB_Name = "Volatilité"
Option Explicit
Option Base 1

Sub volat()

'Déclaratopon des variables itératives
Dim i As Long
Dim j As Long

'Déclaration des autres variables
Dim Parts As Variant
Dim Mat As Variant
Dim cellule As Range

'Initialisation des variables des feuilles de calcul
Set wsA = ThisWorkbook.Worksheets(1)
Set wsO = ThisWorkbook.Worksheets(2)
Set wsCours = ThisWorkbook.Worksheets("Cours")

'Formatage des cellules
wsO.Cells(14, 12).NumberFormat = "0"

'Affectation des valeurs
nbAct = wsCours.Cells(1, wsCours.Columns.Count).End(xlToLeft).Column - 1
nbrow = wsCours.Cells(wsCours.Rows.Count, 1).End(xlUp).Row - 1

'Formatage des cellules
wsCours.Columns(1).NumberFormat = "0"

'Boucle pour itérer sur les actions
For i = 1 To nbAct
    'Initialisation de adresse à la date demandée par l'utilisateur
    Set adresse = wsCours.Columns(1).Find(What:=wsO.Cells(14, 12).Value, LookIn:=xlValues)
    'Condition pour que si la date demandée ne correspond pas à une date enregistrée, ce soit la dernière date enregistrée dans Cours qui soit prise en compte
    If adresse Is Nothing Then
        Do While adresse Is Nothing
            wsO.Cells(14, 12).Value = wsO.Cells(14, 12).Value - 1
            Set adresse = wsCours.Columns(1).Find(What:=wsO.Cells(14, 12).Value, LookIn:=xlValues)
        Loop
    End If
    'Utilisation de la fonction cov() pour calculer la matrice variance-covariance des titres
    If Not adresse Is Nothing Then
        wsA.Cells(14, 2).Resize(nbAct, nbAct).Value = cov(wsCours.Cells(2, 2).Resize(adresse.Row, nbAct))
    End If
Next i

'Formatage des cellules de dates
wsCours.Columns(1).NumberFormat = "dd/mm/yyyy"
wsO.Cells(14, 12).NumberFormat = "dd/mm/yyyy"

For i = 1 To nbAct
    'Écriture des titres en colonnes de la matrice
    wsA.Cells(13 + i, 1).Value = wsA.Cells(1, 1 + i).Value
    'Écriture des titres en lignes de la matrice
    wsA.Cells(13, i + 1).Value = wsA.Cells(1, 1 + i).Value
    'Mise en gras des titres de la matrice
    wsA.Cells(13 + i, 1 + i).Font.Bold = True
    'Calcul de la part du nombres de titres dans la poche par rapport au nombre total de titres
    wsA.Cells(26, 1 + i).Value = wsA.Cells(4, 1 + i).Value / wsA.Cells(4, 12).Value
    'Recopie de la variance du titre
    wsA.Cells(24, 1 + i).Value = wsA.Cells(13 + i, 1 + i).Value
    'Calcul de la volatilité du titre
    wsA.Cells(25, 1 + i).Value = (wsA.Cells(24, 1 + i).Value) ^ (0.5)
Next i

'Calcul de la variance du portefeuille par produit matriciel
Parts = wsA.Cells(26, 2).Resize(1, nbAct).Value
Mat = wsA.Cells(14, 2).Resize(nbAct, nbAct)
wsA.Cells(27, 2).Value = Application.WorksheetFunction.MMult(Application.WorksheetFunction.MMult(Parts, Mat), Application.WorksheetFunction.Transpose(Parts))
'Calcul de la volatilité du portefeuille
wsA.Cells(28, 2).Value = wsA.Cells(27, 2) ^ (0.5)

'Ajustement des cellules
For Each cellule In wsA.UsedRange
    cellule.NumberFormat = "0.00"
Next cellule
wsA.Columns.AutoFit
wsA.Rows(26).NumberFormat = "0.00%"

'Mise en places des titres des tableaux de la matrice variance-covariance et des volatilités
wsA.Cells(13, 1).Value = "Matrice variance-covariance"
wsA.Cells(24, 1).Value = "Variance du titre"
wsA.Cells(25, 1).Value = "Volatilité du titre"
wsA.Cells(26, 1).Value = "Budget en pourcentage"
wsA.Cells(27, 1).Value = "Variance du portefeuille d'actions"
wsA.Cells(28, 1).Value = "Volatilité du portefeuille d'actions"

'Mise en page du tableau par titre
wsA.Cells(5, 1).Resize(4, 12).Borders.Color = RGB(0, 0, 0)
With wsA.Cells(5, 1).Resize(4, 1)
    .Font.Bold = True
    .Interior.Color = RGB(255, 192, 160)
End With

'Mise en page du tableau par poche
wsA.Cells(10, 1).Resize(2, 12).Borders.Color = RGB(0, 0, 0)
With wsA.Cells(10, 1).Resize(2, 1)
    .Font.Bold = True
    .Interior.Color = RGB(255, 192, 160)
End With

'Mise en page de la matrice variance-covariance et des volatilités des titres
wsA.Cells(13, 1).Resize(14, 11).Borders.Color = RGB(0, 0, 0)
With wsA.Cells(13, 1)
    .Font.Bold = True
    .Interior.Color = RGB(255, 192, 160)
End With
With wsA.Cells(24, 1).Resize(5, 1)
    .Font.Bold = True
    .Interior.Color = RGB(255, 192, 160)
End With

'Mise en page du tableau de la volatilité du portefeuille
wsA.Cells(27, 1).Resize(2, 2).Borders.Color = RGB(0, 0, 0)

End Sub

Function cov(plage As Range)

'Déclaration des variables itératives
Dim i As Long
Dim j As Long

'Déclaration des autres variables
Dim largeur As Long
Dim plage_1 As Range
Dim plage_2 As Range
Dim Result()

'Calcul du nombre de titres grace au nombre de colonne et affectation à largeur
largeur = plage.Columns.Count

'Redimmension deResult à la taille de la matrice variance-covariance de dimension (largeur, largeur)
ReDim Result(1 To largeur, 1 To largeur)

'Boucle pour calculer les valeurs des covariances - et variance lorsque largeur = largeur
For i = 1 To largeur

    'Boucle pour itérer sur les titres en calculant la covaraince standarde
    For j = 1 To largeur
        Set plage_1 = plage.Cells(1, i).Resize(plage.Rows.Count, 1)
        Set plage_2 = plage.Cells(1, j).Resize(plage.Rows.Count, 1)
        Result(i, j) = Application.WorksheetFunction.Covariance_S(plage_1, plage_2)
    Next j

Next i

'Résultat renvoyé sous forme matricielle
cov = Result
    
End Function
