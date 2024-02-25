Attribute VB_Name = "Module2"
Sub UsunRekord()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim idToDelete As String

    ' Ustaw arkusz, na którym chcesz dzia³aæ
    Set ws = ThisWorkbook.Sheets("Online Shopping")

    ' ZnajdŸ ostatni¹ u¿ywan¹ wiersz
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Pobierz identyfikator do usuniêcia (np. od u¿ytkownika)
    idToDelete = InputBox("Enter the ID of the record you wish to delete:")

    ' Przeszukaj kolumnê A w poszukiwaniu identyfikatora do usuniêcia
    For i = 1 To lastRow
        If ws.Cells(i, 1).Value = idToDelete Then
            ' Znaleziono rekord do usuniêcia, usuñ wiersz
            ws.Rows(i).Delete
            MsgBox "Rekord o ID " & idToDelete & " zosta³ usuniêty."
            Exit Sub
        End If
    Next i

    ' Jeœli nie znaleziono identyfikatora, wyœwietl komunikat
    MsgBox "No record with the specified ID found."

End Sub

