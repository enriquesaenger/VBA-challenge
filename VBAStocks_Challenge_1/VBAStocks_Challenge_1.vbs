Sub Challenge1()

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
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentDecrease As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestTotalVolume As Double
    Dim greatestTotalVolumeTicker As String
    
    ' Initializing column and row variables
    column = 1
    row = 2
    
    ' Initializing color variables
    colorRed = 3
    colorGreen = 4
    
    ' Loop through each worksheet
    For Each ws In Worksheets
        
        ' Add Column Names
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Add Column Names for Challenge 1
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Add Names to cells for Challenge 1
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
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
                
                ' Write ticker symbol to tickerSymbol
                tickerSymbol = ws.Cells(i, 1).Value
                ' Write opening price to openingPrice
                openingPrice = ws.Cells(i - counter, 3).Value
                ' Write closing price to closingPrice
                closingPrice = ws.Cells(i, 6).Value
                ' Calculate Yearly Change
                yearlyChange = closingPrice - openingPrice
                
                  ' If totalVolume is greatest yet, write to greatestTotalVolume and write Ticker to greatestTotalVolumeTicker
                If totalVolume > greatestTotalVolume Then
                    greatestTotalVolume = totalVolume
                    greatestTotalVolumeTicker = tickerSymbol
                End If
                
                ' Fixing dividing by zero
                If openingPrice <> 0 Then
                    ' Calculate yearlyPercentageChange
                    yearlyPercentageChange = (yearlyChange / openingPrice)
                    
                    ' If yearlyPercentageChange is greatest increase or greatest decrease, write to respective variable and update [respective]Ticker
                    If yearlyPercentageChange > greatestPercentageIncrease Then
                        greatestPercentageIncrease = yearlyPercentageChange
                        greatestPercentageIncreaseTicker = tickerSymbol
                    ElseIf yearlyPercentageChange < greatestPercentageDecrease Then
                        greatestPercentageDecrease = yearlyPercentageChange
                        greatestPercentageDecreaseTicker = tickerSymbol
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
                
                ' Set totalVolume equal to zero at end of ticker symbol
                totalVolume = 0
                ' Move to next row for 9th column
                row = row + 1
                counter = 0
            End If
        Next i
        ' Set row equal to 2 at end of worksheet
        row = 2
        
        ' Write greatestPercentageIncreaseTicker to P2
        ws.Range("P2").Value = greatestPercentageIncreaseTicker
        ' Write greatestPercentageIncrease to Q2
        ws.Range("Q2").Value = greatestPercentageIncrease
        ' Change style in Q2 to Percent
        ws.Range("Q2").Style = "Percent"
        
        ' Write greatestPercentageDecreaseTicker to P3
        ws.Range("P3").Value = greatestPercentageDecreaseTicker
        ' Write greatestPercentageDecrease to Q3
        ws.Range("Q3").Value = greatestPercentageDecrease
        ' Change style in Q3 to Percent
        ws.Range("Q3").Style = "Percent"
        
        ' Write greatestTotalVolumeTicker to P4
        ws.Range("P4").Value = greatestTotalVolumeTicker
        ' Write greatestTotalVolume to Q4
        ws.Range("Q4").Value = greatestTotalVolume
        
    Next ws
End Sub


