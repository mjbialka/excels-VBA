
Sub PrzypiszNumer()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim numer As Long
    
    ' Ustawienie referencji do arkusza
    Set ws = ThisWorkbook.Sheets("tablica_zgloszen") ' Zastąp "NazwaArkusza" nazwą swojego arkusza
    
    ' Znalezienie ostatniego wiersza w kolumnie "C"
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Przypisanie numeru inkrementalnego w kolumnie "B"
    numer = 1 ' Początkowy numer inkrementalny
    
    For i = 3 To lastRow
        If Not IsEmpty(ws.Cells(i, "C").Value) Then
            ws.Cells(i, "B").Value = numer
            numer = numer + 1
        Else
            ws.Cells(i, "B").ClearContents
        End If
    Next i
End Sub
