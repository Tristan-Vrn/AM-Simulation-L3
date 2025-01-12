Attribute VB_Name = "importation"
Sub importation()

'DŽclaration des classeurs
Dim projet As Workbook
Dim wbSource As Workbook

'DŽclaration des feuilles de calcul
Dim wsFS As Worksheet
Dim wsI As Worksheet
Dim wsR As Worksheet
Dim wsTS As Worksheet

'DŽclaration des autres variables
Dim adresse As String

'On importe en premier la feuille des frenchstocks
adresse = Application.GetOpenFilename
Set wbSource = Workbooks.Open(adresse)
Set wsFS = wbSource.Sheets(1)
wsFS.Copy after:=ThisWorkbook.Worksheets(1)
ThisWorkbook.Worksheets(2).Name = "FrenchStocks"
wbSource.Close

'Puis la feuille des indexes
adresse = Application.GetOpenFilename
Set wbSource = Workbooks.Open(adresse)
Set wsI = wbSource.Sheets(1)
wsI.Copy after:=ThisWorkbook.Worksheets(2)
ThisWorkbook.Worksheets(3).Name = "Indexes"
wbSource.Close

'Puis la feuille des techstocks
adresse = Application.GetOpenFilename
Set wbSource = Workbooks.Open(adresse)
Set wsTS = wbSource.Sheets(1)
wsTS.Copy after:=ThisWorkbook.Worksheets(3)
ThisWorkbook.Worksheets(4).Name = "TechStocks"
wbSource.Close

'Enfin, on importe la feuille des rates
adresse = Application.GetOpenFilename
Set wbSource = Workbooks.Open(adresse)
Set wsR = wbSource.Sheets(1)
wsR.Copy after:=ThisWorkbook.Worksheets(4)
ThisWorkbook.Worksheets(5).Name = "Rates"
wbSource.Close

End Sub
