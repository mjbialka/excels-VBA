Private Sub AutofitText()
    Dim wys As Worksheet
    Dim rng As Range
    Dim n As Long
    
    ' Ustawienie referencji do arkusza
    Set wys = ThisWorkbook.Sheets("scenariusz_testowy")
    
    ' Ustawienie referencji do zakresu, dla którego chcemy dostosować wysokość wierszy
    Set rng = wys.Range("A33:F100")
    
    ' Włączenie zawijania tekstu w zakresie
    rng.WrapText = True
    
    ' Dostosowanie wysokości wierszy tylko dla wierszy powyżej numeru 33
    For n = 32 To rng.Rows.Count
        rng.Rows(n).AutoFit
    Next n
End Sub





'Private Sub Workbook_Open()
'    Dim wys As Worksheet
'    Dim rng As Range
'    Dim n As Long
'    
    ' Ustawienie referencji do arkusza
'    Set wys = ThisWorkbook.Sheets("tablica_zgloszen")
    
    ' Ustawienie referencji do zakresu, dla którego chcemy dostosować wysokość wierszy
'    Set rng = wys.Range("A1:F100")
    
    ' Włączenie zawijania tekstu w zakresie
'    rng.WrapText = True
    
    ' Dostosowanie wysokości wierszy tylko dla wierszy powyżej numeru 3
'    For n = 4 To rng.Rows.Count
'        rng.Rows(n).AutoFit
'    Next n
'End Sub