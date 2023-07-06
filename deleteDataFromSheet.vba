Sub resetujFormularz()
    
    Dim resetujFormularz As Worksheet

    ' Ustawienie wskaźnika na odpowiedni arkusz
    Set resetujFormularz = ThisWorkbook.Sheets("formularz_zgloszeniowy")

    ' Wyczyszczenie zawartości komórek
    resetujFormularz.Range("E4:E31").ClearContents
   
End Sub
