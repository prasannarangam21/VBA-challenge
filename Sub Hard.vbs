Sub Hard()
For Each ws In Worksheets
  ws.Activate
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "TotalStockVolume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest total volume"
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    Dim Volume As Double
    Volume = 0
    Dim Column As Integer
    Column = 1
    
    Dim j As Integer
    j = 0
    
    OpenPrice = Cells(2, Column + 2).Value
    Range("P2").NumberFormat = "0.00%"
    Range("P3").NumberFormat = "0.00%"
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
                
               If (PercentChange > Greatest_Increase) Then
                   Greatest_Increase = PercentChange
                   Range("O2").Value = Cells(i, Column).Value
                   Range("P2").Value = Greatest_Increase
             
               End If
               If (PercentChange < Greatest_Decrease) Then
                   Greatest_Decrease = PercentChange
                   Range("O3").Value = Cells(i, Column).Value
                   Range("P3").Value = Greatest_Decrease
             
               End If
               If (Volume > Greatest_Total_Volume) Then
                   Greatest_Total_Volume = Volume
                   Range("O4").Value = Cells(i, Column).Value
                   Range("P4").Value = Greatest_Total_Volume
               End If
                            
               
               Volume = 0
               j = j + 1
        Else
               Volume = Cells(i, 7).Value + Volume
               
        End If
    Next i
 Next ws
End Sub


