Sub DodajZgloszenie()
    Dim wysFormularz As Worksheet
    Dim wysRejestr As Worksheet
    Dim lastRow As Long
    Dim firstEmptyRow As Long
    
    ' Ustawienie referencji do arkuszy
    Set wysFormularz = ThisWorkbook.Sheets("formularz_zgloszeniowy")
    Set wysRejestr = ThisWorkbook.Sheets("tablica_zgloszen")
    
    ' Znalezienie ostatniego wiersza w rejestrze
    lastRow = wysRejestr.Cells(wysRejestr.Rows.Count, "C").End(xlUp).Row
    
    ' Znalezienie pierwszego wolnego wiersza w rejestrze
    firstEmptyRow = wysRejestr.Cells(3, "C").End(xlDown).Row + 1
    
    ' Przypisanie wartości z formularza do rejestrze
    wysRejestr.Cells(firstEmptyRow, "C").Value = wysFormularz.Range("E6").Value
    wysRejestr.Cells(firstEmptyRow, "D").Value = wysFormularz.Range("E23").Value
    wysRejestr.Cells(firstEmptyRow, "E").Value = wysFormularz.Range("E9").Value
    wysRejestr.Cells(firstEmptyRow, "F").Value = wysFormularz.Range("E30").Value
    
    ' Zapisanie zmian w pliku
    ThisWorkbook.Save
    
    ' Komunikat potwierdzający
    MsgBox "Dane zostały zapisane w tablicy zgłoszeń."
    
    
 ' Wywołanie procedury PrzypiszNumer() po wykonaniu DodajZgloszenie()
    Application.Run "PrzypiszNumer"
    
    ' Zwolnienie zasobów
    Set wysFormularz = Nothing
    Set wysRejestr = Nothing
End Sub

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