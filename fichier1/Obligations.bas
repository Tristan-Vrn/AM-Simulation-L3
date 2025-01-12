Attribute VB_Name = "Obligations"
Option Explicit
Option Base 1
Dim budget As Double

Sub Obli(budget As Double)

'Déclaration des feuilles de calcul
Dim wsD As Worksheet
Dim wsA As Worksheet

'Décalration des variables
Dim nbCols As Integer
Dim nbTitres As Integer
  
    Application.DisplayAlerts = False
    'ThisWorkbook.Worksheets("Obligations").Delete
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets.Add after:=ThisWorkbook.Worksheets(Worksheets.Count)
    Set wsD = ThisWorkbook.Worksheets(Worksheets.Count)
    wsD.Name = "Obligations"

'On permet au client de renseigner son budget (pas forcement 1 000 000) pour garder un code polyvalent
   'budget = InputBox("Quel est votre budget total ?", "Renseignement budget", 1000000)
   Set wsA = ThisWorkbook.Worksheets("Actions")
   'budget = 1000000 - wsA.Cells(22, 12).Value


    'On fixe le nombre d'obligations à 10
     nbTitres = 10

    'Mise en place des intitules initiaux (on definit arbitrairement une premiere fois le nombre de colonnes afin de garder en generalite ensuite en utilisant nbCols)
    nbCols = 9
    wsD.Cells(1, 1).Resize(1, nbCols).Value = Array("Nominal", "Coupon", "Maturité", "Taux de coupon", "Périodicité", "Taux sans risque", "Valeur", "Macaulay Duration", "Modified Macaulay Duration")
    wsD.Cells(3, nbCols + 2).Value = "Budget total :"
    wsD.Cells(3, nbCols + 3).Value = budget

    'Appel de la procedure de creation des obligations
    Call Bonds(nbTitres, budget)

    'Appel de la procedure de mise en place stylistique
    Call ApplyStyle(wsD)

End Sub
'Procedure qui calcule le prix des obligations dans la feuille
Sub Bonds(nbTitres As Integer, budget As Double)

    Dim WsData As Worksheet
    
    Dim i As Integer
    
    Dim nbCols As Integer
    
    Dim Bond As BondsClass
    Dim BondPtf As New Collection
    Dim PtfValue As Double

'Attribution
    Set WsData = ThisWorkbook.Worksheets(Worksheets.Count)

'Definition des dimensions
    nbCols = WsData.Cells(1, WsData.Columns.Count).End(xlToLeft).Column
    
'Cellules informatives pour le client
    WsData.Cells(2, 2 + nbCols).Value = "Nombre d'obligations :"
    WsData.Cells(2, 3 + nbCols).Value = nbTitres

'Creation des obligations en faisant en sorte que leur valeur totale soit inferieure a 500 000 (car au moins 50% du portefeuille doit être en actions) grace a une boucle DO LOOP
    Do
    
    'On intialise a 0 la valeur du portefuille avant la boucle de creation des obligations pour eviter une boucle DO infinie
    PtfValue = 0
        
        'Creation d'obligations partiellement aléatoires avec une boucle for
        For i = 1 To nbTitres - 1
            
            Set Bond = New BondsClass
       
            Randomize
            
            'Le nominal de chaque obligation depend du budget total ET de nbTitres (afin d'eviter que le code lag trop pour trouver une configuration qui respecte la conditon sur les 50% d'actions si notre client desire acheter beaucoup d'obligations ou reduire son budget )
            Bond.Nominal = ((budget / nbTitres) * (0.5 + (Rnd()))) * 10 / nbTitres
            
            'On simule differents taux de coupon et maturites avec la fonction rnd()
            Bond.Maturite = Int(Rnd() * 20) + 1
            Bond.TxCoupon = 0.01 + Int(Rnd() * 10) / 100
            
            Bond.RiskFreeRate = 0.05
            
            'On creee la moitie des obligations avec une periodicite de 2 a l'aide d'une condition if sur le numero de creation de l'obligation
            If i <= (nbTitres / 2) Then
                Bond.Periodicity = 2
            Else
                Bond.Periodicity = 1
            End If
        
            'Ajout de chaque obligation dans une collection
            BondPtf.Add Bond
            
            'Ecriture des caracteristiques de l'obligation dans la feuille
            WsData.Cells(1 + i, 1).Resize(1, nbCols).Value = Array(Bond.Nominal, Bond.Coupon, Bond.Maturite, Bond.TxCoupon, Bond.Periodicity, _
            Bond.RiskFreeRate, Bond.BondValue, Bond.MaucaulayDuration, Bond.ModifiedMacaulayDuration)
            
            
            'Calcul de la somme des valeurs des obligations
            PtfValue = PtfValue + Bond.BondValue
    
        Next i
    
    'On reitere tant que la valeur totale des obligations represente plus de la moitie de la valeur du portefeuille
    Loop Until PtfValue < budget
    '%%%%%%%%%%refaire

    'Ecriture de la valeur du ptf
    WsData.Cells(4, 3 + nbCols).Value = budget - PtfValue
    
    Set Bond = New BondsClass
       
   
            Bond.Maturite = 1
            Bond.TxCoupon = 0
            Bond.Periodicity = 1
            Bond.RiskFreeRate = 0.05
            Bond.Nominal = WsData.Cells(4, 3 + nbCols).Value * 1.05
            
        BondPtf.Add Bond
        
    WsData.Cells(1 + nbTitres, 1).Resize(1, nbCols).Value = Array(Bond.Nominal, Bond.Coupon, Bond.Maturite, Bond.TxCoupon, Bond.Periodicity, _
                                                                                                Bond.RiskFreeRate, Bond.BondValue, Bond.MaucaulayDuration, Bond.ModifiedMacaulayDuration)
                                                                                                
        PtfValue = PtfValue + Bond.BondValue
        
        WsData.Cells(nbTitres + 2, 7).Value = PtfValue
                                                                                                
    'Ecriture du budget restant une fois les obligations acquises
    WsData.Cells(4, 2 + nbCols).Value = "Budget restant après achat des obligations :"
    WsData.Cells(4, 3 + nbCols).Value = budget - PtfValue
    
End Sub

'Procedure qui applique le style de notre choix
Sub ApplyStyle(ws As Worksheet)

    Dim nbCols As Integer
    Dim nbTitres As Integer

'Determination de la taille du tableau
    ws.Activate
    nbCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    'On recupere le nombre de titres grace a la cellule informative remplie au début de la sub Bonds
    nbTitres = ws.Cells(2, nbCols + 3).Value

'Tri par duration
    ws.Cells(1, 1).Resize(nbTitres + 1, nbCols).Sort key1:=ws.Cells(1, 8), order1:=xlAscending, Header:=xlYes

'Application du style
    
    'Style des intitules
        With ws.Cells(1, 1).Resize(1, nbCols)
            .Font.Bold = True
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
        End With
        ws.Cells(nbTitres + 2, 7).Font.Bold = True
    
    'Bordures du tableau
        With ws.Cells(1, 1).Resize(nbTitres + 1, nbCols)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
        End With
    
    'Changement de couleur d'arriere plan des cellules importantes
        ws.Cells(1, 2).Resize(nbTitres + 1, 1).Interior.Color = RGB(224, 224, 224)
        ws.Cells(1, 8).Resize(nbTitres + 1, 1).Interior.Color = RGB(24, 224, 224)
        ws.Cells(1, 9).Resize(nbTitres + 1, 1).Interior.Color = RGB(24, 224, 224)
        ws.Cells(1, 7).Resize(nbTitres + 1, 1).Interior.Color = RGB(224, 224, 224)
    
    'Ajustement des colonnes
        ws.Columns.AutoFit
        ws.Columns(2).NumberFormat = "0 000.00"
        ws.Columns(4).NumberFormat = "0%"
        ws.Columns(6).NumberFormat = "0%"
        ws.Columns(1).NumberFormat = "00 000"
        ws.Columns(7).NumberFormat = "00 000"
        ws.Columns(8).NumberFormat = "0.00"
        ws.Columns(9).NumberFormat = "0.00"
        
        ws.Cells(4, 3 + nbCols).NumberFormat = "0 €"
        ws.Cells(3, 3 + nbCols).NumberFormat = "0 €"
        
        'Desactivation de la grille d'arriere plan
        ActiveWindow.DisplayGridlines = False
        

End Sub


