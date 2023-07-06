Sub KolorowanieKomorek()
    Dim wys As Worksheet
    Dim cel As Range
    
    ' Ustawienie arkusza
    Set wys = ThisWorkbook.Sheets("rejestr_defektow")
    
    ' Iteracja przez komórki w zakresie E6:E115
    For Each cel In wys.Range("E6:E115")
        ' Sprawdzenie wartości w komórce
        Select Case cel.Value
            Case "Niski"
                cel.Interior.Color = RGB(51, 204, 204) ' Kolor #33CCCC
            Case "Średni"
                cel.Interior.Color = RGB(0, 153, 153) ' Kolor #009999
            Case "Wysoki"
                cel.Interior.Color = RGB(0, 102, 102) ' Kolor #006666
            ' Dodaj inne przypadki, jeśli są wymagane
        End Select
    Next cel
End Sub
