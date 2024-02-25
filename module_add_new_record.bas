Attribute VB_Name = "Module1"
Sub DodajRekord()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Ustaw arkusz, na którym chcesz operowaæ
    Set ws = ThisWorkbook.Sheets("Online Shopping")
    
    ' ZnajdŸ ostatni¹ u¿ywan¹ wiersz
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Pobierz wszystkie dane od u¿ytkownika
    Dim dane(1 To 21) As String
    dane(1) = InputBox("Podaj Index")
    dane(2) = InputBox("Podaj Order ID")
    dane(3) = InputBox("Podaj Customer ID")
    dane(4) = InputBox("Podaj Gender")
    dane(5) = InputBox("Podaj Age")
    dane(7) = InputBox("Podaj Date")
    dane(10) = InputBox("Podaj Status")
    dane(11) = InputBox("Podaj Shop")
    dane(12) = InputBox("Podaj SKU")
    dane(13) = InputBox("Podaj Category")
    dane(14) = InputBox("Podaj Size")
    dane(15) = InputBox("Podaj Quantity")
    dane(16) = InputBox("Podaj Currency")
    dane(17) = InputBox("Podaj Amount")
    dane(18) = InputBox("Podaj ship-state")
    dane(19) = InputBox("Podaj ship-postal-code")
    dane(20) = InputBox("Podaj ship-country")
    dane(21) = InputBox("Podaj B2B")
    
    ' Dodaj dane do arkusza Excela
    For i = 1 To 21
        ws.Cells(lastRow, i).Value = dane(i)
    Next i
End Sub
