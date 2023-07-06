Sub KopiujAktywnyArkuszIZapiszNowyPlik()
    Dim sourceWorkbook As Workbook
    Dim newWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim newWorksheet As Worksheet
    Dim newFileName As Variant
    
    ' Określenie pliku źródłowego
    Set sourceWorkbook = ActiveWorkbook
    
    ' Określenie arkusza źródłowego
    Set sourceWorksheet = sourceWorkbook.ActiveSheet
    
    ' Wyświetlanie okna dialogowego zapisu
    newFileName = Application.GetSaveAsFilename(FileFilter:="Plik Excel (*.xlsm), *.xlsm", _
                                                Title:="Zapisz jako plik Excel", _
                                                InitialFileName:="Zgloszenie_xxx.xlsm")
    
    ' Sprawdzenie, czy użytkownik wybrał plik do zapisu
    If newFileName <> False Then
        ' Sprawdzenie, czy w nowym pliku istnieje arkusz o nazwie "formularz_zgloszeniowy" i jeśli tak, to go usuń
        On Error Resume Next
        Application.DisplayAlerts = False
        Set newWorkbook = Workbooks.Add
        Set newWorksheet = newWorkbook.Sheets(1)
        If Not newWorksheet Is Nothing Then
            newWorksheet.Delete
        End If
        Application.DisplayAlerts = True
        On Error GoTo 0
        
        ' Kopiowanie zawartości, formatowania i skryptów z arkusza źródłowego do nowego arkusza
        sourceWorksheet.Copy Before:=newWorkbook.Sheets(1)
        newWorkbook.Sheets(1).Name = "formularz_zgloszeniowy"
        
        ' Zapisywanie nowego pliku Excel
        newWorkbook.SaveAs newFileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        newWorkbook.Close SaveChanges:=False
    End If
    
    ' Zwolnienie zasobów
    Set sourceWorksheet = Nothing
    Set sourceWorkbook = Nothing
    Set newWorksheet = Nothing
    Set newWorkbook = Nothing
End Sub


