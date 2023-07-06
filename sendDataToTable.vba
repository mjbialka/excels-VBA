Sub DodajZgloszenie()
    Dim Formularz As Worksheet
    Dim Defekty As Worksheet
    'Dim Tablica As Worksheet
    Dim LastRow As Long
    
    ' Ustaw arkusze
    Set Formularz = ThisWorkbook.Worksheets("formularz_zgloszen")
    Set Defekty = ThisWorkbook.Worksheets("rejestr_defektow")
    'Set Tablica = ThisWorkbook.Worksheets("tablica_zgloszen")
    
    ' Sprawdź ostatni wiersz w zakładce "rejestr_defektow"
    LastRow = Defekty.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' Oblicz numer wiersza, do którego należy wpisać nowy rekord
    If Defekty.Cells(LastRow, "C").Value <> "" Then
    
    LastRow = LastRow + 1
    
    End If
    
    'Dim NewRow As Long
    'NewRow = (LastRow \ 7) * 7 + 7
        
    ' Przypisz wartość z komórki B10 do odpowiedniej komórki w zakładce "rejestr_defektow"
    Defekty.Cells(NewRow, "C").Value = Formularz.Range("E4").Value
    Defekty.Cells(NewRow, "D").Value = Formularz.Range("E6").Value
    Defekty.Cells(NewRow, "E").Value = Formularz.Range("E10").Value
    Defekty.Cells(NewRow, "F").Value = Formularz.Range("E11").Value
    Defekty.Cells(NewRow, "G").Value = Formularz.Range("E23").Value
    Defekty.Cells(NewRow, "I").Value = Formularz.Range("E30").Value
    
    
    ' Wyświetl komunikat o zapisaniu testu
    MsgBox "Zgłoszenie zostało dodane."
    
End Sub