Attribute VB_Name = "Runs"
Option Explicit
Option Base 0
Sub Runs()

'Déclaration des feuilles de calcul
Dim wsD As Worksheet
Dim wsFS As Worksheet
Dim wsR As Worksheet

'Déclaration des variables itératives
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

'Déclaration des autres variables
Dim ws As Worksheet
Dim nbcolD As Integer
Dim nbCols As Integer
Dim cell As Range
Dim lastrow As Long
Dim colonne As Integer
Dim colonne2 As Integer
Dim nbre_periodes As Integer
Dim max_duree As Integer

'Déclaration des vecteurs
Dim c() As Double
Dim x() As Variant
Dim nbCol As Integer
Dim stats() As Variant
Dim cours() As Double
Dim r() As Double

'Affectation des feuilles
Set wsD = ThisWorkbook.Worksheets(1)
Set wsFS = ThisWorkbook.Worksheets(2)

'Effacement des données sur Actions
wsD.Cells.Clear
wsD.Name = "Actions"

'Report en ligne 1 des intitulés
wsD.Cells(1, 1).Resize(1, 2).Value = Array("Runs", "freq")

'Calcul des dimensions de la série en s'arretant pour chaque action au 30/12/2004
lastrow = wsFS.Cells(wsFS.Rows.Count, 1).End(xlUp).Row - 1
For Each cell In wsFS.Range("A1:A" & lastrow)
    If cell.Value < DateSerial(2005, 1, 1) Then
        nbre_periodes = cell.Row - 1
    End If
Next cell

'Récupération des cours de chaque titre
colonne = 0

'Boucle pour itérer sur les 4 feuilles de titres
For i = 1 To 4
    colonne2 = 0
    Set ws = ThisWorkbook.Worksheets(1 + i)
    nbCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column - 1
        
        'Boucle pour itérer sur les colonnes de la feuille pour parcourir chaque titre
        For j = 1 To nbCol
            x = ws.Cells(2, j + 1).Resize(nbre_periodes, 1).Value
            ReDim c(0 To nbre_periodes - 1)
            
            'Boucler pour itérer sur le titre pour sa valeur à chaque date jusqu'au 30/12/2004
            For t = 1 To nbre_periodes
                c(t - 1) = x(t, 1)
            Next t
            
'%%%%%            'Redimensionner r pour qu'il contienne tous les noms de titres en l'ajoutant chaque fois jusqu'à ce qu'il n'y en ai plus sur la ligne 1
            colonne = colonne + 1
            ReDim r(1 To colonne)
            k = 2
            
            Do Until ws.Cells(k, j + 1).Value <> ""
                k = k + 1
            Loop
            
            'Calcul du rendement du titre et recopie sur la ligne 12, à titre informatif
            r(j) = (ws.Cells(nbre_periodes, j + 1).Value - ws.Cells(k, j + 1).Value) / ws.Cells(k, j + 1).Value
            wsD.Cells(12, colonne + 1).Value = r(j)

            'Calcul de la distribution des runs par appel de fnStatsRuns pour calculer le nombre et la fréquence de runs des cours
            stats = fnStatsRuns(c)

            'Calcul de max_duree
            max_duree = UBound(stats(1))

            'Boucle de report des valeurs des runs observés de 0 à max_duree sur la feuille Actions
            For t = 0 To max_duree
                wsD.Cells(2 + t, 1).Value = t
            Next t

            'Report des effectifs et des fréquences des runs avec mise en forme
            colonne2 = colonne2 + 1
            wsD.Cells(2, 1 + colonne).Resize(max_duree + 1, 1).Value = WorksheetFunction.Transpose(stats(1))
            wsD.Columns(1 + colonne).NumberFormat = "0.00%"
            wsD.Cells(1, 1 + colonne).Value = ws.Cells(1, 1 + colonne2).Value
            
            'Report des fréquences de runs ≥ 3
            wsD.Cells(11, colonne + 1).Value = Application.WorksheetFunction.Sum(wsD.Cells(5, colonne + 1).Resize(max_duree, 1))
            Next j

Next i

'Calcul du nombre de colonnes dans le tableau des runs pour mise en forme
nbCols = wsD.Cells(1, wsD.Columns.Count).End(xlToLeft).Column
    
'Style des intitulés
With wsD.Cells(1, 1).Resize(1, nbCols)
    .Font.Bold = True
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .HorizontalAlignment = xlCenter
End With
    
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
wsD.Cells(max_duree + 2, 7).Font.Bold = True
    
'Bordures des tableaux
With wsD.Cells(1, 1).Resize(1, nbCols)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
End With
    
With wsD.Cells(1, 1).Resize(max_duree + 2, nbCols)
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlRight).LineStyle = xlContinuous
End With

'Mise en forme des lignes "Runs" et "Rendements"
wsD.Cells(11, 1).Value = "Runs ≥ 3"
wsD.Cells(11, 1).Font.Bold = True
wsD.Cells(12, 1).Value = "Rendements"
wsD.Cells(12, 1).Font.Bold = True

'Mise en forme en des couleurs
wsD.Cells(11, 1).Resize(2, nbCols).Borders.Color = RGB(0, 0, 0)
wsD.Cells(1, 1).Resize(1, nbCols).Interior.Color = RGB(255, 192, 160)
       
'On appelle la sub séléction pour constituer les actions et obligations
 Call Selection
       
End Sub

Function fnStatsRuns(c() As Double) As Variant()

'Déclaration des variables itératives
Dim i As Integer
Dim j As Integer

'Déclaration des autres variables
Dim nbre_periodes As Integer
Dim Runs() As Double
Dim duree As Integer
Dim max_duree As Integer
Dim var_prec As Double
Dim var As Double
Dim freq() As Double
Dim tot As Integer

'Calcul du nombre de périodes consécutives à la première, cette dernière étant 0
nbre_periodes = UBound(c)

'Redimensionnement de runs
ReDim Runs(0 To nbre_periodes)

'Initialisation de prec_var
var_prec = c(1) - c(0)

'Boucle pour dénombrer les runs sur les cours de la période 2 à la dernière
For i = 2 To nbre_periodes

    'Calcul de la variation
    var = c(i) - c(i - 1)
    
    'Cas où les cours gardent la même tendance
    If var * var_prec > 0 Then
    
        'Incrémentation de duree
        duree = duree + 1
        
    'Cas où la tendance est rompue
    Else
    
        'Incrémentation du nombre de runs de longueur duree
        Runs(duree) = Runs(duree) + 1
        
        'Mofidicationn de max_duree dans le cas où la longueur est supérieure au max
        If duree > max_duree Then max_duree = duree
        
        duree = 0
        
    End If
    
    'Report de var dans var_prec
    var_prec = var
    
Next i


'Réduction de runs
ReDim Preserve Runs(0 To max_duree)

'Redimensionnement de freq
ReDim freq(0 To max_duree)

'Calcul de l'effectif des runs tot
tot = WorksheetFunction.Sum(Runs)

'Calcul des fréquences des runs
For j = 0 To max_duree
    freq(j) = Runs(j) / tot
Next j

'Résultat
fnStatsRuns = Array(Runs, freq)

End Function
