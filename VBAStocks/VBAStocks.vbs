Sub VBAStocks()

    ' Declaring variables
    Dim counter As Double
    Dim lastRow As Double
    Dim tickerSymbol As String
    Dim column As Integer
    Dim row As Integer
    Dim totalVolume As Double
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyPercentageChange As Double
    Dim percentageChange As Double
    Dim colorRed As Integer
    Dim colorGreen As Integer
    
    ' Initializing column and row variables
    column = 1
    row = 2
    
    ' Initializing color variables
    colorRed = 3
    colorGreen = 4
    
    ' Loop through each worksheet
    For Each ws In Worksheets
    
        ' Add Column Names
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
        
        ' Resize Columns to fit data
        Columns("I:L").AutoFit

        ' Find last row in each worksheet
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Loop through all rows in each worksheet
        For i = 2 To lastRow
        
            ' Find ticker symbol change
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                ' Add that year's volume to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ' Count number of rows for current ticker price
                counter = counter + 1
            Else
                'Add that year's volume to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                ' Write ticker symbol to TickerSymbol
                tickerSymbol = ws.Cells(i, 1).Value
                ' Write opening price to openingPrice
                openingPrice = ws.Cells(i - counter, 3).Value
                ' Write closing price to closingPrice
                closingPrice = ws.Cells(i, 6).Value
                ' Calculate Yearly Change
                yearlyChange = closingPrice - openingPrice
                
                ' Fixing dividing by zero
                If openingPrice <> 0 Then
                    ' Calculate yearlyPercentageChange
                    yearlyPercentageChange = (yearlyChange / openingPrice)
                Else
                    yearlyPercentageChange = 0
                End If
                
                ' Write tickerSymbol to cell in 9th column
                ws.Cells(row, 9).Value = tickerSymbol
                ' Write yearlyChange to cell in 10th column
                ws.Cells(row, 10).Value = yearlyChange
                ' Format 10th column to 2 decimal places
                ws.Cells(row, 10).NumberFormat = "0.00"
                
                ' Conditional Formatting for yearlyChange in 10th colum
                If yearlyChange > 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = colorGreen
                ElseIf yearlyChange < 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = colorRed
                End If
                
                ' Write yearlyPercentageChange to cell in 11th column
                ws.Cells(row, 11).Value = yearlyPercentageChange
                
                ' Format 11th column to 2 decimal places and percent
                ws.Cells(row, 11).NumberFormat = "0.00%"
                
                ' Write totalVolume to cell in 12th column
                ws.Cells(row, 12).Value = totalVolume
                ' Set totalVolume equal to zero at end of ticker symbol
                totalVolume = 0
                ' Move to next row for 9th column
                row = row + 1
                counter = 0
            End If
        Next i
        ' Set row equal to 2 at end of worksheet
        row = 2
    Next ws
End Sub