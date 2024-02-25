Attribute VB_Name = "Module2"
Sub UsunRekord()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim idToDelete As String

    ' Ustaw arkusz, na kt�rym chcesz dzia�a�
    Set ws = ThisWorkbook.Sheets("Online Shopping")

    ' Znajd� ostatni� u�ywan� wiersz
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Pobierz identyfikator do usuni�cia (np. od u�ytkownika)
    idToDelete = InputBox("Enter the ID of the record you wish to delete:")

    ' Przeszukaj kolumn� A w poszukiwaniu identyfikatora do usuni�cia
    For i = 1 To lastRow
        If ws.Cells(i, 1).Value = idToDelete Then
            ' Znaleziono rekord do usuni�cia, usu� wiersz
            ws.Rows(i).Delete
            MsgBox "Rekord o ID " & idToDelete & " zosta� usuni�ty."
            Exit Sub
        End If
    Next i

    ' Je�li nie znaleziono identyfikatora, wy�wietl komunikat
    MsgBox "No record with the specified ID found."

End Sub

