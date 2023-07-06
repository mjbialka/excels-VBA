Sub DodajBlad()
    Dim wsFormularz As Worksheet
    Dim wsRejestr As Worksheet
    Dim lastRow As Long
    Dim firstEmptyRow As Long
    
    ' Ustawienie referencji do arkuszy
    Set wsFormularz = ThisWorkbook.Sheets("formularz_zgloszeniowy")
    Set wsRejestr = ThisWorkbook.Sheets("rejestr_defektow")
    
    ' Znalezienie ostatniego wiersza w rejestrze
    lastRow = wsRejestr.Cells(wsRejestr.Rows.Count, "C").End(xlUp).Row
    
    ' Znalezienie pierwszego wolnego wiersza w rejestrze
    firstEmptyRow = wsRejestr.Cells(6, "C").End(xlDown).Row + 1
    
    ' Przypisanie wartości z formularza do rejestrze
    wsRejestr.Cells(firstEmptyRow, "C").Value = wsFormularz.Range("E4").Value
    wsRejestr.Cells(firstEmptyRow, "D").Value = wsFormularz.Range("E6").Value
    wsRejestr.Cells(firstEmptyRow, "E").Value = wsFormularz.Range("E10").Value
    wsRejestr.Cells(firstEmptyRow, "F").Value = wsFormularz.Range("E11").Value
    wsRejestr.Cells(firstEmptyRow, "G").Value = wsFormularz.Range("E23").Value
    wsRejestr.Cells(firstEmptyRow, "H").Value = wsFormularz.Range("E30").Value
    
    ' Zapisanie zmian w pliku
    ThisWorkbook.Save
    
    ' Komunikat potwierdzający
    MsgBox "Dane zostały zapisane w rejestrze."
    
     ' Wywołanie procedury PrzypiszNumer() po wykonaniu DodajZgloszenie()
    Application.Run "PrzypiszNumerDefektu"
    
    ' Zwolnienie zasobów
    Set wsFormularz = Nothing
    Set wsRejestr = Nothing
End Sub

Sub PrzypiszNumerDefektu()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim numer As Long
    
    ' Ustawienie referencji do arkusza
    Set ws = ThisWorkbook.Sheets("rejestr_defektow")
    
    ' Znalezienie ostatniego wiersza w kolumnie "C"
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Początkowy numer inkrementalny
    numer = 1
    
    ' Przypisanie numeru inkrementalnego w kolumnie "B"
    For i = 6 To lastRow
        If Not IsEmpty(ws.Cells(i, "C").Value) Then
            ws.Cells(i, "B").Value = "D" & Format(numer, "000")
            numer = numer + 1
        Else
            ws.Cells(i, "B").ClearContents
        End If
    Next i
End Sub





