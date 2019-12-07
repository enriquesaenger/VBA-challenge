Sub WallStreet()

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
    Dim greatestPercentIncrease As Double
    Dim greatestPercentDecrease As Double
    Dim greatestTotalVolume As Double
    
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
        
        ' Add Column Names for Challenge 1
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        ' Add Names to cells for Challenge 1
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        ' Reset Greatest variables for each worksheet
        greatestPercentIncrease = 0
        greatestPercentDecrease = 0
        greatestTotalVolume = 0
        
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
                
                ' If totalVolume is greatest yet, write to greatestTotalVolume
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                End If
                
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
                    
                    ' If yearlyPercentageChange is greatest increase or greatest decrease, write to respective variable
                    If yearlyPercentageChange > greatestPercentageIncrease Then
                        greatestPercentageIncrease = yearlyPercentageChange
                    ElseIf yearlyPercentageChange < greatestPercentageDecrease Then
                        greatestPercentageDecrease = yearlyPercentageChange
                    End If
                    
                Else
                    yearlyPercentageChange = 0
                End If
                
                ' Write tickerSymbol to cell in 9th column
                ws.Cells(row, 9).Value = tickerSymbol
                ' Write yearlyChange to cell in 10th column
                ws.Cells(row, 10).Value = yearlyChange
                
                ' Conditional Formatting for yearlyChange in 10th colum
                If yearlyChange > 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = colorGreen
                ElseIf yearlyChange < 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = colorRed
                End If
                
                ' Write yearlyPercentageChange to cell in 11th column
                ws.Cells(row, 11).Value = yearlyPercentageChange
                
                ' Percentage Styling for cell in 11th column
                ws.Cells(row, 11).Style = "Percent"
                
                ' Write totalVolume to cell in 12th column
                ws.Cells(row, 12).Value = totalVolume
                
                ' Write greatestPercentageIncrease to P2
                ws.Range("P2").Value = greatestPercentageIncrease
                
                ' Change style in P2 to Percent
                ws.Range("P2").Style = "Percent"
                
                'Change style in P3 to Percent
                ws.Range("P3").Style = "Percent"
                
                'Write greatestPercentageDecrease to P3
                ws.Range("P3").Value = greatestPercentageDecrease
                
                ' Write greatestTotalVolume to P4
                ws.Range("P4").Value = greatestTotalVolume
                
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
