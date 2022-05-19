Sub StockSummary():
    ' Set up and define column number variables for data
    Dim tickerColumn1 As Integer
    Dim openColumn As Integer
    Dim closeColumn As Integer
    Dim volumeColumn1 As Integer
    
    tickerColumn1 = 1
    openColumn = 3
    closeColumn = 6
    volumeColumn1 = 7
    
    ' Set up and define column number variables for results
    Dim tickerColumn2 As Integer
    Dim changeColumn As Integer
    Dim percentColumn As Integer
    Dim volumeColumn2 As Integer
    
    tickerColumn2 = 9
    changeColumn = 10
    percentColumn = 11
    volumeColumn2 = 12

    ' Set up counter for the total volume of each stock. Needs to be a Long as some volumes may exceed 2 billion
    Dim totalVolume As LongLong
    totalVolume = 0
    
    ' Set up variables to hold open and close prices and changes
    Dim openPrice As Double
    Dim closePrice As Double
    openPrice = 0
    closePrice = 0
    Dim yearChange As Double
    Dim percentChange As Double
    
    'Set up counter for how many unique stocks have been seen
    Dim uniqueStockCount As Integer
    uniqueStockCount = 0
    
    ' Count the total number of rows
    Dim numRows As Long
    numRows = Cells(Rows.Count, tickerColumn1).End(xlUp).Row
    
    ' Iterate over all the data rows
    For i = 2 To numRows
    
        ' If the ticker name is different to the cell above it, assign the open price to openPrice and reset totalVolume to new starting value. Also increase uniqueStockCount by 1:
        If Cells(i, tickerColumn1).Value <> Cells(i - 1, tickerColumn1).Value Then
            
            openPrice = Cells(i, openColumn).Value
            totalVolume = Cells(i, volumeColumn1).Value
            
            uniqueStockCount = uniqueStockCount + 1
            
        ' Else if the ticker name is different to the cell below it, assign the close price to closePrice and increase totalVolume
        ' Then also write the ticker name, calculated changes, and total stock volume to the appropriate cells:
        ElseIf Cells(i, tickerColumn1).Value <> Cells(i + 1, tickerColumn1).Value Then
            
            closePrice = Cells(i, closeColumn).Value
            totalVolume = totalVolume + Cells(i, volumeColumn1).Value
            
            yearChange = closePrice - openPrice
            ' Calculate percentChange, but allow for if openPrice starts at 0
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearChange / openPrice
            End If
            
            ' Write ticker name (note: uniqueStockCount + 1 is needed because of header row)
            Cells(uniqueStockCount + 1, tickerColumn2).Value = Cells(i, tickerColumn1).Value
            
            ' Write changes and total volume (note: uniqueStockCount + 1 is needed because of header row)
            Cells(uniqueStockCount + 1, changeColumn).Value = yearChange
            Cells(uniqueStockCount + 1, percentColumn).Value = percentChange
            Cells(uniqueStockCount + 1, volumeColumn2).Value = totalVolume
            
            ' Set colour of yearly change cells, red for negative, green for positive, yellow in the rare case of 0
            If yearChange < 0 Then
                Cells(uniqueStockCount + 1, changeColumn).Interior.ColorIndex = 3
            ElseIf yearChange > 0 Then
                Cells(uniqueStockCount + 1, changeColumn).Interior.ColorIndex = 4
            Else
                Cells(uniqueStockCount + 1, changeColumn).Interior.ColorIndex = 6
            End If
        ' Otherwise if the ticker name is the same as the cell below it, increase totalVolume:
        Else
            totalVolume = totalVolume + Cells(i, volumeColumn1).Value
            
            
        End If
        
    Next i
    
    
    ' Count the total number of rows of summary data
    Dim numRows2 As Long
    numRows2 = Cells(Rows.Count, tickerColumn2).End(xlUp).Row
    
    ' Set up variables to hold ticker name and value for greatest % increase and decrease and greatest total volume
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerTV As String
    Dim tickerIncreaseValue As Double
    Dim tickerDecreaseValue As Double
    Dim tickerTVValue As LongLong
    
    tickerIncreaseValue = 0
    tickerDecreaseValue = 0
    tickerTVValue = 0
    
    ' Iterate over all the data rows
    For j = 2 To numRows2
        ' If the current row has a higher % increase than any previous rows, update the Increase variables
        If Cells(j, percentColumn).Value > tickerIncreaseValue Then
            tickerIncreaseValue = Cells(j, percentColumn).Value
            tickerIncrease = Cells(j, tickerColumn2).Value
        End If
        
        ' If the current row has a lower % decrease than any previous rows, update the Decrease variables
        If Cells(j, percentColumn).Value < tickerDecreaseValue Then
            tickerDecreaseValue = Cells(j, percentColumn).Value
            tickerDecrease = Cells(j, tickerColumn2).Value
        End If
        
        ' If the current row has a higher total volume than any previous rows, update the TV variables
        If Cells(j, volumeColumn2).Value > tickerTVValue Then
            tickerTVValue = Cells(j, volumeColumn2).Value
            tickerTV = Cells(j, tickerColumn2).Value
        End If
        
    Next j
    
    Range("P2").Value = tickerIncrease
    Range("Q2").Value = tickerIncreaseValue
    Range("P3").Value = tickerDecrease
    Range("Q3").Value = tickerDecreaseValue
    Range("P4").Value = tickerTV
    Range("Q4").Value = tickerTVValue
    
End Sub


