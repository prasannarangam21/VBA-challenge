Sub Easy()
For Each WS In Worksheets
    WS.Activate
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "TotalStockVolume"
    Dim Volume As Double
    Volume = 0
    Dim j As Integer
    j = 0
    
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               Volume = Cells(i, 7).Value + Volume
               Range("I" & 2 + j).Value = Cells(i, 1).Value
               Range("J" & 2 + j).Value = Volume
               Volume = 0
               j = j + 1
        Else
               Volume = Cells(i, 7).Value + Volume
               
        End If
    Next i
 Next WS
End Sub