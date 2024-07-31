Attribute VB_Name = "Module1"
Sub Stocks()
    
    Dim quarters As Worksheet
    Dim totalVolume As Double
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTikcer As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim lastRow As Long
    Dim endRow As Long
    Dim startRow As Long
    Dim tickerRow As Long
    Dim i As Long
    
         
For Each quarters In Worksheets

 'headers
 
    quarters.Cells(1, 9).Value = "Ticker"
    quarters.Cells(1, 10).Value = "Quaterly Change"
    quarters.Cells(1, 11).Value = "Percent Change"
    quarters.Cells(1, 12).Value = " Total Stock Volume"
    
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

'data for ticker, quarterly change, percent change and total stock value

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
            percentChange = (quarterlyChange / openPrice)
        
            quarters.Cells(tickerRow, 11).Value = percentChange
            tickerRow = 1 + tickerRow
 
 
 'data for ticker and value - greatest, increase, decrease and total volume
        
        If percentChange > greatestIncrease Then
            quarters.Cells(2, 17).Value = ticker
            quarters.Cells(2, 18).Value = percentChange
            greatestIncrease = percentChange
            greatestIncreaseTicker = ticker
            

        End If
            

            If percentChange < greatestDecrease Then
                quarters.Cells(3, 17).Value = ticker
                quarters.Cells(3, 18).Value = percentChange
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
                
    
        End If


            If totalVolume > greatestVolume Then
                quarters.Cells(4, 17).Value = ticker
                quarters.Cells(4, 18).Value = totalVolume
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
                

        End If

            quarters.Columns("k").NumberFormat = "0.00%"
            quarters.Columns("r").NumberFormat = "0.00%"
            quarters.Cells(4, 18).NumberFormat = "0.00E+00%"
            
            
            openPrice = quarters.Cells(i + 1, 3).Value
            totalVolume = 0
            rowTicker = 2

        Else
    
            totalVolume = totalVolume + quarters.Cells(i, 7).Value
     
        End If
     
    Next i

   Next quarters

End Sub

Sub AddColors()

Dim quarters As Worksheet
Dim lastRow As Long
Dim i As Long
    
     For Each quarters In Worksheets
        lastRow = quarters.Cells(quarters.Rows.Count, 10).End(xlUp).Row
   
   'Quaterly change
   
    For i = 2 To lastRow
    
        If quarters.Cells(i, 10) > 0 Then
           
                quarters.Cells(i, 10).Interior.ColorIndex = 4 'green
            
        ElseIf quarters.Cells(i, 10) < 0 Then
           
             quarters.Cells(i, 10).Interior.ColorIndex = 3 'red
   
   'Percent change
        
        End If
   
    Next i
        
    For i = 2 To lastRow
        lastRow = quarters.Cells(quarters.Rows.Count, 11).End(xlUp).Row

        If quarters.Cells(i, 11) > 0 Then
           
            quarters.Cells(i, 11).Interior.ColorIndex = 4 'green
            
        ElseIf quarters.Cells(i, 11) < 0 Then
           
             quarters.Cells(i, 11).Interior.ColorIndex = 3 'red
        
      
            End If
            
       Next i
       
    Next quarters
    
End Sub

