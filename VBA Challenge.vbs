Attribute VB_Name = "Module1"
Sub StockData():

For Each WS In Worksheets
    WS.Cells(1, 9).Value = "Stock Ticker"
    WS.Cells(1, 16).Value = "Stock Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"
    WS.Cells(1, 17).Value = "Value"
    WS.Cells(2, 15).Value = "Greatest % Increase"
    WS.Cells(3, 15).Value = "Greatest % Decrease"
    WS.Cells(4, 15).Value = "Greatest Total Volume"
    
Dim i As Long
Dim tickerName As String
Dim openYearly As Double
Dim totalVolume As Double
totalVolume = 0
Dim totalYearly As Double
totalYearly = 0
Dim percentChange As Double
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Long
lastRow = WS.Cells(Rows.Count, 1).End(xlUp).row

For i = 2 To lastRow
openYearly = WS.Cells(tickerRow, 3).Value

    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
        tickerName = WS.Cells(i, 1).Value
        WS.Range("I" & tickerRow).Value = tickerName
    
        totalYearly = totalYearly + (WS.Cells(i, 6).Value - openYearly)
        WS.Range("J" & tickerRow).Value = totalYearly
    
        percentChange = (totalYearly / openYearly)
        WS.Range("K" & tickerRow).Value = percentChange
        WS.Range("K" & tickerRow).Style = "Percent"
        
        totalVolume = totalVolume + WS.Cells(i, 7).Value
        WS.Range("L" & tickerRow).Value = totalVolume
        
        tickerRow = tickerRow + 1
        totalYearly = 0
        totalVolume = 0
        openYearly = WS.Cells(tickerRow, 3).Value
    Else
        totalVolume = totalVolume + WS.Cells(i, 7).Value
    End If
Next i

Dim yearLastRow As Long
yearLastRow = WS.Cells(Rows.Count, 10).End(xlUp).row

For i = 2 To yearLastRow

    If WS.Cells(i, 10).Value >= 0 Then
        WS.Cells(i, 10).Interior.ColorIndex = 4
    Else
        WS.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
    
Dim percentLastRow As Long
percentLastRow = WS.Cells(Rows.Count, 11).End(xlUp).row
Dim percent_max As Double
percent_max = 0
Dim percent_min As Double
percent_min = 0

For i = 2 To percentLastRow

    If percent_max < WS.Cells(i, 11).Value Then
        percent_max = WS.Cells(i, 11).Value
        WS.Cells(2, 17).Value = percent_max
        WS.Cells(2, 17).Style = "Percent"
        WS.Cells(2, 16).Value = WS.Cells(i, 9).Value
    ElseIf percent_min > WS.Cells(i, 11).Value Then
        percent_min = WS.Cells(i, 11).Value
        WS.Cells(3, 17).Value = percent_min
        WS.Cells(3, 17).Style = "Percent"
        WS.Cells(3, 16).Value = WS.Cells(i, 9).Value
    End If
Next i

Dim totalVolumeRow As Long
totalVolumeRow = WS.Cells(Rows.Count, 12).End(xlUp).row
Dim totalVolumeMax As Double
totalVolumeMax = 0

For i = 2 To totalVolumeRow

    If totalVolumeMax < WS.Cells(i, 12).Value Then
        totalVolumeMax = WS.Cells(i, 12).Value
        WS.Cells(4, 17).Value = totalVolumeMax
        WS.Cells(4, 16).Value = WS.Cells(i, 9).Value
    End If
Next i
    
Next WS

End Sub
