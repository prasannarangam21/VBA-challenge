Sub Moderate()
For Each ws In Worksheets
  ws.Activate
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "TotalStockVolume"
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    Volume = 0
    Dim Column As Integer
    Column = 1
    
    Dim j As Integer
    j = 0
    
    OpenPrice = Cells(2, Column + 2).Value
    For i = 2 To LastRow
        If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
               Volume = Cells(i, 7).Value + Volume
               Range("I" & 2 + j).Value = Cells(i, Column).Value
               Range("L" & 2 + j).Value = Volume
               ClosePrice = Cells(i, Column + 5).Value
               YearlyChange = ClosePrice - OpenPrice
               Range("J" & 2 + j).Value = YearlyChange
               If (OpenPrice = 0) Then
                    PercentChange = 0
               Else
                    PercentChange = YearlyChange / OpenPrice
                    Range("K" & 2 + j).Value = PercentChange
                    Range("K" & 2 + j).NumberFormat = "0.00%"
               End If
               If Range("J" & 2 + j).Value >= 0 Then
                   Range("J" & 2 + j).Interior.ColorIndex = 4
               Else
                   Range("J" & 2 + j).Interior.ColorIndex = 3
               End If
                OpenPrice = Cells(i + 1, Column + 2)
               Volume = 0
               j = j + 1
        Else
               Volume = Cells(i, 7).Value + Volume
               
        End If
    Next i
 Next ws
End Sub

