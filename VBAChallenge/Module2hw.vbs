Attribute VB_Name = "Module1"
Sub stockChallenge():

Dim total As LongLong  ' total stock volume
Dim row As Long ' loop control variable that will go through the rows of the sheet
Dim rowCount As Long ' variable will hold number of rows
Dim yearlyChange As Double 'variable that holds yearly change
Dim percentChange As Double ' variable that holds the percent change for each stock
Dim summaryTableRow As Long 'variable holds the rows of the summary table
Dim stockStartRow As Long ' variable that holds where a stock starts
'loop through all the worksheets
    For Each ws In Worksheets
    

    'set sheet headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stack Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        
        
    'initialize the values
        summaryTableRow = 0
        total = 0
        yearlyChange = 0
        stockStartRow = 2 'first stock in the sheet is on row 2
        
    'get the value of our last row
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
    'loop until we get to the end of the sheet
        For row = 2 To rowCount
        
    'check for changes in column A
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
            
            total = total + ws.Cells(row, 7).Value
            
    'check to see if the value of the total volume is 0
            If total = 0 Then
                ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                ws.Range("J" & 2 + summaryTableRow).Value = 0
                ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"
                ws.Range("L" & 2 + summaryTableRow).Value = 0
            Else
    'find the first non zero starting value
            If ws.Cells(stockStartRow, 3).Value = 0 Then
                For findValue = stockStartRow To row
                
                If ws.Cells(findValue, 3).Value <> 0 Then
                    stockStartRow = findValue
    'once we find a value, break the loop
                Exit For
            End If
            Next findValue
            End If
            
    'calculate the yearly change (difference in last closed - first open)
            yearlyChange = (ws.Cells(row, 6).Value - ws.Cells(stockStartRow, 3).Value)
    'calculate the percent change (yearly change / first open)
            percentChange = yearlyChange / ws.Cells(stockStartRow, 3).Value
            
            
            
            ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
            ws.Range("J" & 2 + summaryTableRow).Value = yearlyChange
            ws.Range("J" & 2 + summaryTableRow).NumberFormat = "0.00"
            ws.Range("K" & 2 + summaryTableRow).Value = percentChange
            ws.Range("K" & 2 + summaryTableRow).NumberFormat = "0.00%"
            ws.Range("L" & 2 + summaryTableRow).Value = total
            ws.Range("L" & 2 + summaryTableRow).NumberFormat = "#,###"
            
    'formatting for the yearly change column
            If yearlyChange > 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4 'green for positives
            ElseIf yearlyChange < 0 Then
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
            Else
                ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
            End If
                
            
            End If
            total = 0
            yearlyChange = 0
            summaryTableRow = summaryTableRow + 1
            
        Else
            total = total + ws.Cells(row, 7).Value
        
        
        End If
        
    Next row
'After looping through rows, find the max and min and place them accordingly
    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
'value of the greatest max volume
    ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & rowCount)) * 100
    ws.Range("Q4").NumberFormat = "#,###"
    
    
'do matching in order to match the ticker names with the values
'tell the row in the summary table where the ticker matches the greatest increase
    increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    ws.Range("P2").Value = ws.Cells(increaseNumber + 1, 9)
    
    decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    ws.Range("P3").Value = ws.Cells(decreaseNumber + 1, 9)
    
    volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
    ws.Range("P4").Value = ws.Cells(increaseNumber + 1, 9)
    'AutoFit the columns
        ws.Columns("A:Q").AutoFit


Next ws

End Sub

