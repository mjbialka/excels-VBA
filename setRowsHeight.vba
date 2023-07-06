Sub UstawWysokoscWierszy()
    Dim wys As Worksheet
    Dim i As Long
    
    ' Ustawienie arkusza
    Set wys = ThisWorkbook.Sheets("rejestr_defektow")
    
    ' Iteracja przez wiersze od 6 do 115
    For i = 6 To 115
        wys.Rows(i).RowHeight = 36
    Next i
End Sub