Attribute VB_Name = "Module1"
Sub GetStockDataForOneYear()

'Create a script that loops through all the stocks for one year and outputs the following information:

'The ticker symbol

'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

'The total stock volume of the stock.
    
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim summaryRow As Long
    Dim tickerRow As Long, tickerStartingPrice As Double, tickerEndingPrice As Double
    Dim tickerYearStartDate As Date, tickerYearEndDate As Date
    Dim yearlyChange As Double, percentChange As Double, totalVolume As Double

    ' Set the worksheet containing the source data
    Set ws = ThisWorkbook.Worksheets("A")
    
    ' Set the destination where we want to put calculated data
    summaryRow = 2
    
    ' Output header row for the summary table
    ws.Cells(summaryRow, 9).Value = "Ticker"
    ws.Cells(summaryRow, 10).Value = "Yearly Change"
    ws.Cells(summaryRow, 11).Value = "Percentage Change"
    ws.Cells(summaryRow, 12).Value = "Total Stock Volume"
    ws.Cells(summaryRow, 15).Value = "Ticker"
    ws.Cells(summaryRow, 16).Value = "Value"
    ws.Cells(summaryRow, 17).Value = "Greatest Percent Increase"
    ws.Cells(summaryRow, 18).Value = "Greatest Percent Decrease"
    ws.Cells(summaryRow, 19).Value = "Greatest Total Volume"
    
    ' Find the last row of data in ws
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Loop through each row of data
    For i = 2 To lastRow ' Assuming header row is in row 1

        tickerRow = i

        ' Get the ticker symbol, starting price, and ending price
        tickerSymbol = ws.Cells(i, 1).Value ' ticker is in column A
        startingPrice = ws.Cells(i, 3).Value ' starting price is in column C
        endingPrice = ws.Cells(i, 6).Value ' ending price is in column F
        stockVolume = ws.Cells(i, 7).Value ' ending price is in column G
        
        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
          ' Calculate yearly and percentage change
            yearlyChange = endingPrice - startingPrice
            percentChange = (yearlyChange / startingPrice) * 100
            
            ' Update greatest values if necessary
        If percentChange > greatestPercentIncrease Then
        greatestPercentIncrease = percentChange
        greatestPercentIncreaseTicker = tickerSymbol
    End If
        
        If percentChange < greatestPercentDecrease Then
        greatestPercentDecrease = percentChange
        greatestPercentDecreaseTicker = tickerSymbol
    End If
    
        If totalVolume > greatestTotalVolume Then
        greatestTotalVolume = totalVolume
        greatestTotalVolumeTicker = tickerSymbol
    End If
        
         ' Output to summary table
            summaryRow = summaryRow + 1
            ws.Cells(summaryRow, 9).Value = tickerSymbol
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = totalVolume
        
        ' Reset the Total
            totalVolume = 0
        
        End If
        
         ' Accumulate totalVolume for the current ticker within the loop
           totalVolume = totalVolume + stockVolume
           
           ' Output the greatest values
   ws.Cells(summaryRow + 2, 15).Value = "Greatest Percent Increase"
    ws.Cells(summaryRow + 2, 16).Value = greatestPercentIncreaseTicker
    ws.Cells(summaryRow + 2, 17).Value = greatestPercentIncrease
    ws.Cells(summaryRow + 3, 15).Value = "Greatest Percent Decrease"
    ws.Cells(summaryRow + 3, 16).Value = greatestPercentDecreaseTicker
    ws.Cells(summaryRow + 3, 17).Value = greatestPercentDecrease
    ws.Cells(summaryRow + 4, 15).Value = "Greatest Total Volume"
    ws.Cells(summaryRow + 4, 16).Value = greatestTotalVolumeTicker
    ws.Cells(summaryRow + 4, 17).Value = greatestTotalVolume
    
    Next i

End Sub

     
    
            
                 


