Sub Drukuj()

    Dim ObszarDoDruku As Range
    
    ' Sprawdź, czy coś jest zaznaczone
    If Selection.Cells.Count > 0 Then
        ' Ustal zaznaczony obszar do druku
        Set ObszarDoDruku = Selection
        
        ' Wywołaj funkcję drukowania dla zaznaczonego obszaru
        ObszarDoDruku.PrintOut
    Else
        ' Jeśli nic nie jest zaznaczone, wyświetl komunikat
        MsgBox "Nie zaznaczono żadnego obszaru do druku."
    End If
End Sub
