Sub DrukujDoPDF()
    Dim area As Range
    Dim pdfFileName As String

    ' Określenie zaznaczonego obszaru
    Set area = Selection

    ' Utworzenie nazwy pliku PDF
    pdfFileName = "ścieżka_do_folderu\nazwa_pliku.pdf" ' Podaj właściwą ścieżkę i nazwę pliku PDF

    ' Ustawienia drukowania
    With ActiveSheet.PageSetup
        ' Ustawienie obszaru drukowania na zaznaczony obszar
        .PrintArea = area.Address

        ' Inne ustawienia drukowania (opcjonalne)
        .Orientation = xlPortrait ' Ustawienie orientacji druku (xlPortrait - pionowo, xlLandscape - poziomo)
        .FitToPagesWide = 1 ' Dopasowanie do jednej strony szerokości
        .FitToPagesTall = False ' Nie dopasowywać do wysokości strony
    End With

    ' Drukowanie do pliku PDF
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFileName, Quality:=xlQualityStandard

    ' Przywrócenie ustawień drukowania domyślnych
    ActiveSheet.PageSetup.PrintArea = ""
    ActiveSheet.PageSetup.Orientation = xlPortrait
    ActiveSheet.PageSetup.FitToPagesWide = 1
    ActiveSheet.PageSetup.FitToPagesTall = False
End Sub
