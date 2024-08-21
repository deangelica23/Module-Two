Attribute VB_Name = "Module1"
Sub Multi_year_Stocks()

    Dim quarters As Worksheet
    Dim totalVolume As Double
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim lastRow As Long
    Dim endRow As Long
    Dim startRow As Long
    Dim tickerRow As Long
    Dim i As Long
    
    For Each quarters In Worksheets
    
        ' Headers
        quarters.Cells(1, 9).Value = "Ticker"
        quarters.Cells(1, 10).Value = "Quarterly Change"
        quarters.Cells(1, 11).Value = "Percent Change"
        quarters.Cells(1, 12).Value = "Total Stock Volume"
        quarters.Cells(1, 17).Value = "Ticker"
        quarters.Cells(1, 18).Value = "Value"
        quarters.Cells(2, 16).Value = "Greatest % Increase"
        quarters.Cells(3, 16).Value = "Greatest % Decrease"
        quarters.Cells(4, 16).Value = "Greatest % Total Volume"
        
        lastRow = quarters.Cells(quarters.Rows.Count, 1).End(xlUp).Row
        
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        tickerRow = 2
        
        startRow = 2
        totalVolume = 0
        openPrice = quarters.Cells(startRow, 3).Value
        
        ' Data for ticker, quarterly change, percent change, and total stock value
        For i = 2 To lastRow
            If quarters.Cells(i + 1, 1).Value <> quarters.Cells(i, 1).Value Or i = lastRow Then
            
                endRow = i
                closePrice = quarters.Cells(endRow, 6).Value
                ticker = quarters.Cells(i, 1).Value
                
                quarters.Cells(tickerRow, 9).Value = ticker
                totalVolume = totalVolume + quarters.Cells(i, 7).Value
                quarters.Cells(tickerRow, 12).Value = totalVolume
                
                quarterlyChange = closePrice - openPrice
                quarters.Cells(tickerRow, 10).Value = quarterlyChange
                
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice)
                Else
                    percentChange = 0
                End If
                
                quarters.Cells(tickerRow, 11).Value = percentChange
                
                ' Format quarterly change
                If quarterlyChange > 0 Then
                    quarters.Cells(tickerRow, 10).Interior.ColorIndex = 4 ' green
                ElseIf quarterlyChange < 0 Then
                    quarters.Cells(tickerRow, 10).Interior.ColorIndex = 3 ' red
                End If
                
                ' Format percent change
                If percentChange > 0 Then
                    quarters.Cells(tickerRow, 11).Interior.ColorIndex = 4 ' green
                ElseIf percentChange < 0 Then
                    quarters.Cells(tickerRow, 11).Interior.ColorIndex = 3 ' red
                End If
                
                ' Increment tickerRow
                tickerRow = tickerRow + 1
                
                ' Greatest % increase, decrease, and total volume
                If percentChange > greatestIncrease Then
                    quarters.Cells(2, 17).Value = ticker
                    quarters.Cells(2, 18).Value = percentChange
                    greatestIncrease = percentChange
                End If
                
                If percentChange < greatestDecrease Then
                    quarters.Cells(3, 17).Value = ticker
                    quarters.Cells(3, 18).Value = percentChange
                    greatestDecrease = percentChange
                End If
                
                If totalVolume > greatestVolume Then
                    quarters.Cells(4, 17).Value = ticker
                    quarters.Cells(4, 18).Value = totalVolume
                    greatestVolume = totalVolume
                End If
                
                ' Formatting Percentages
                quarters.Columns("K").NumberFormat = "0.00%"
                quarters.Columns("R").NumberFormat = "0.00%"
                quarters.Cells(4, 18).NumberFormat = "0.00E+00"
                
                ' Reset for the next ticker
                openPrice = quarters.Cells(i + 1, 3).Value
                totalVolume = 0
                
            Else
                totalVolume = totalVolume + quarters.Cells(i, 7).Value
            End If
        Next i
        
    Next quarters

End Sub


