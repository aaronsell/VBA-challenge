Attribute VB_Name = "Module1"
Sub stockTracker():

    Dim totalVolume As LongLong
    Dim row As Long
    Dim yearlyChange As Double
    Dim summaryTable As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim stockStart As Long
    Dim stockOpen
    Dim stockClose As Long
    Dim tickerName As String
    Dim lastRow As Long
    
    For Each ws In Worksheets
    
         ' Initialize the values
        SummaryRow = 0
        rowCount = ws.Cells(Rows.Count, 1).End(xlUp).row
        yearlyChange = 0
        totalVolume = 0
        stockOpen = 2
        
        ' Set the column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Loop to the last row
        For row = 2 To rowCount
            ' Check for changes in column 1
            
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
                ' Take the value that is in column 7
                totalVolume = totalVolume + ws.Cells(row, 7).Value
                ' Test to see if the total volume is 0
                
                If totalVolume = 0 Then
                
                    ' Populate the results in the summary rows (I, J, K, L)
                    ws.Range("I" & 2 + SummaryRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + SummaryRow).Value = 0
                    ws.Range("K" & 2 + SummaryRow).Value = 0
                    ws.Range("L" & 2 + SummaryRow).Value = 0
                    
                Else
                    ' Find first stock open that isn't 0
                    If ws.Cells(stockOpen, 3).Value = 0 Then
                    
                        For findValue = stockOpen To row
                            ' Cycle to next value that isn't 0
                            
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                stockOpen = findValue
                                
                                Exit For
                                
                            End If
                            
                        Next findValue
                        
                    End If
                    
                    ' Calculate the yearly change
                    yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(stockOpen, 3).Value)
                    
                    ' Take the yearly change and divide it by the stock open
                    percentChange = yearlyChange / ws.Cells(stockOpen, 3).Value
                    
                    ' Populate the results into the summary rows (I, J, K, L)
                    ws.Range("I" & 2 + SummaryRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + SummaryRow).Value = yearlyChange
                    ws.Range("K" & 2 + SummaryRow).Value = percentChange
                    ws.Range("L" & 2 + SummaryRow).Value = totalVolume
                    
                    ' Format the summary data
                    ws.Range("J" & 2 + SummaryRow).NumberFormat = "$0.00"
                    ws.Range("K" & 2 + SummaryRow).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + SummaryRow).NumberFormat = "#,###"
                    
                    ' Highlight positive or negative yearly change
                    If yearlyChange > 0 Then
                        ws.Range("J" & 2 + SummaryRow).Interior.ColorIndex = 4
                        
                    ElseIf yearlyChange < 0 Then
                        ws.Range("J" & 2 + SummaryRow).Interior.ColorIndex = 3
                        
                    Else ' If there is no change, then no highlight
                        ws.Range("J" & 2 + SummaryRow).Interior.ColorIndex = 0
                        
                    End If
                    
                End If
                
                ' Reset total volume and yearly change
                yearlyChange = 0
                totalVolume = 0
                SummaryRow = SummaryRow + 1
                
            ' If no change, run this
            Else
                ' Take the value that is in column 7
                totalVolume = totalVolume + ws.Cells(row, 7).Value
                
            End If
            
        Next row
        
        
        ' Use Max() and Min() functions to get the largest increase, largest decrease, & largest volume
        ws.Range("$Q$2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("$Q$3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
        ws.Range("Q4").NumberFormat = "##0.0E+0"
        
        'Use Match() function to pair the largest increase, largest decrease, and largest total volume
        greatestIncrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P2").Value = ws.Cells(greatestIncrease + 1, 9)
        
        greatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        ws.Range("P3").Value = ws.Cells(greatestDecrease + 1, 9)
        
        greatestTotal = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount))
        ws.Range("P4").Value = ws.Cells(greatestTotal + 1, 9)
        
        ' Add formatting to columns
        Columns("A:Q").AutoFit
        
    Next ws

End Sub


