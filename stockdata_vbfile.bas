'Attribute VB_Name = "Module1"
'
'  Version 3 - added printResults Sub, percentChange function
'            - removed checking for empty worksheet rows, etc
'            - created a structure for keeping output info
'            - moved output to a separate subroutine
'

Public Type stockType
     ticker As String
     open As Double
     close As Double
     volume As Double
End Type


Sub processStockData()

    'Altnernate solution that doesn't look past the active rows
    'Looking past the last row in Excel is not an issue but for
    'many programming languages, looking past the end of a
    'container can result in undefined behavior.
    
    Dim Sheet As Worksheet
    Dim stock As stockType
    
    'looping variables
    Dim currentRow As Long
    Dim nextRow As Long
    Dim lastRow As Long
    Dim ticker As String
    Dim nextTicker As String
    Dim total As Double
    Dim openValue As Double
    Dim closeValue As Double

    'input and output variables/constants
    Dim outputRow As Long
    Dim inTickerCol As String
    Dim inVolumeCol As String
    Dim inOpenCol As String
    Dim inCloseCol As String
    Dim outTickerCol As String
    Dim outVolumeCol As String
    Dim outChangeCol As String
    Dim outPercentCol As String
    Dim outColor As Integer
    
    'Set the input and output constants - minimize constants in code.
    inTickerCol = "A"
    inOpenCol = "C"
    inCloseCol = "F"
    inVolumeCol = "G"
    outTickerCol = "I"
    outChangeCol = "J"
    outPercentCol = "K"
    outVolumeCol = "L"

    'loop over each sheet in worksheets
    For Each Sheet In Worksheets
        'find the last active row in the Sheet
        lastRow = Sheet.Cells(Rows.Count, inTickerCol).End(xlUp).row

        'initialize for the current Sheet and output
        currentRow = 2
        outputRow = 2
        total = 0
        ticker = Sheet.Range(inTickerCol & currentRow)
        openValue = Sheet.Range(inOpenCol & currentRow)
        
        'print column headers for results
        Call summarizeStock(Sheet, True, False, 1, stock)
        
        Do While currentRow < lastRow
        
            total = total + _
                            Sheet.Range(inVolumeCol & currentRow).Value
            nextRow = currentRow + 1
            nextTicker = Sheet.Range(inTickerCol & nextRow).Value
        
            If ticker <> nextTicker Then
                'This is the last row for the current stock - out results
                closeValue = Sheet.Range(inCloseCol & currentRow).Value
                
                stock.ticker = ticker
                stock.close = closeValue
                stock.open = openValue
                stock.volume = total
                
                Call summarizeStock(Sheet, False, False, outputRow, stock)
                
                'update the constants and accumulator for the next stock
                total = 0
                ticker = nextTicker
                openValue = Sheet.Range(inOpenCol & nextRow).Value
                outputRow = outputRow + 1
            End If
        
            currentRow = nextRow
        Loop
    
        'Now take care of the last row of data in the current sheet
        'This is the last Volume for the last stock
        total = total + Sheet.Range(inVolumeCol & currentRow).Value
        closeValue = Sheet.Range(inCloseCol & currentRow).Value
            
        stock.ticker = ticker
        stock.close = closeValue
        stock.open = openValue
        stock.volume = total
                
        Call summarizeStock(Sheet, False, True, outputRow, stock)
                        
    Next Sheet

    'Prepare for finding min/max percent and volume
    'declare looping variables
    Dim currentMinPercent As Double
    Dim currentMinTicker As String
    Dim currentMaxPercent As Double
    Dim currentMaxPTicker As String
    Dim currentMaxVolume As Double
    Dim currentMaxVTicker As String
    Dim J As Long
    Dim currentPercent As Double
    Dim currentVolume As Double

    'declare and initialize output columns
    Dim outMaxTitleCol As String
    Dim outMaxTickerCol As String
    Dim outMaxStatCol As String
    outMaxTitleCol = "O"
    outMaxTickerCol = "P"
    outMaxStatCol = "Q"
 
    'Now loop across all sheets
    'within each sheet loop through the output Cols to look for results
    For Each Sheet In Worksheets
    
        'Initialize variables for finding min and max
        currentMinPercent = Sheet.Range(outPercentCol & 2).Value
        currentMaxPercent = Sheet.Range(outPercentCol & 2).Value
        currentMaxVolume = Sheet.Range(outVolumeCol & 2).Value
        currentMinTicker = Sheet.Range(outTickerCol & 2).Value
        currentMaxPTicker = currentMinTicker
        currentMaxVTicker = currentMinTicker
        
        lastRow = Sheet.Cells(Rows.Count, outVolumeCol).End(xlUp).row
        
        For J = 2 To lastRow
            currentPercent = Sheet.Range(outPercentCol & J).Value
            currentVolume = Sheet.Range(outVolumeCol & J).Value
            ticker = Sheet.Range(outTickerCol & J).Value
            
            'check if this stock is the current min/max etc
            If currentVolume > currentMaxVolume Then
                currentMaxVolume = currentVolume
                currentMaxVTicker = ticker
            End If
            If currentPercent < currentMinPercent Then
                currentMinPercent = currentPercent
                currentMinTicker = ticker
            ElseIf currentPercent > currentMaxPercent Then
                currentMaxPercent = currentPercent
                currentMaxPTicker = ticker
            End If
        Next J
          
        'output the max/min statistics
        Call printResults(Sheet, outMaxTitleCol, 2, "Greatest % Increase")
        Call printResults(Sheet, outMaxTitleCol, 3, "Greatest % Decrease")
        Call printResults(Sheet, outMaxTitleCol, 4, "Greatest Total Volume")
        Call printResults(Sheet, outMaxTickerCol, 1, "Ticker")
        Call printResults(Sheet, outMaxTickerCol, 2, currentMaxPTicker)
        Call printResults(Sheet, outMaxTickerCol, 3, currentMinTicker)
        Call printResults(Sheet, outMaxTickerCol, 4, currentMaxVTicker)
        Call printResults(Sheet, outMaxStatCol, 1, "Value")
        Call printResults(Sheet, outMaxStatCol, 2, currentMaxPercent)
        Call printResults(Sheet, outMaxStatCol, 3, currentMinPercent)
        Call printResults(Sheet, outMaxStatCol, 4, currentMaxVolume)
    
        'clean up the output Columns
        Sheet.Range(outMaxStatCol & 2 & ":" & outMaxStatCol & 3).NumberFormat = "0.00%"
        Sheet.Columns(outMaxTitleCol & ":" & outMaxStatCol).AutoFit
  
    Next Sheet
    
    'all the sheets have been processeed return to the first worksheet
    Worksheets(1).Activate
         
End Sub

Private Sub summarizeStock(ws As Worksheet, firstCall As Boolean, _
                            lastCall As Boolean, row As Long, results As stockType)
'output the results for a given stock
'If firstCall is true - output the headers
'If lastCall is true - clean up the columns of output

    Dim outTickerCol As String
    Dim outVolumeCol As String
    Dim outChangeCol As String
    Dim outPercentCol As String
    Dim outColor As Integer
    Dim yearlyChange As Double
    Dim changePercent As Double
    
    outTickerCol = "I"
    outChangeCol = "J"
    outPercentCol = "K"
    outVolumeCol = "L"


    If firstCall = True Then
        'Prepare the outputColumns
        Call printResults(ws, outTickerCol, 1, "Ticker")
        Call printResults(ws, outChangeCol, 1, "Yearly Change")
        Call printResults(ws, outPercentCol, 1, "Percent Change")
        Call printResults(ws, outVolumeCol, 1, "Total Stock Volume")
    Else
        'Print results for the current stock
        Debug.Print results.ticker
        
        Call printResults(ws, outTickerCol, row, results.ticker)
        Call printResults(ws, outVolumeCol, row, results.volume)
    
        yearlyChange = results.close - results.open
                
        Call printResults(ws, outChangeCol, row, yearlyChange)
    
        changePercent = percentChange(results.open, results.close)
        Call printResults(ws, outPercentCol, row, _
                    FormatPercent(changePercent))
                                    
        'apply conditional formatting
        If yearlyChange > 0 Then
            outColor = 4
        Else
            outColor = 3
        End If
        ws.Range(outChangeCol & row).Interior.ColorIndex = outColor
    End If
    
    If lastCall = True Then
        'cleanup the output columns with formatting
        ws.Range(outPercentCol & 2 & ":" & _
                    outPercentCol & row).NumberFormat = "0.00%"
        ws.Columns(outTickerCol & ":" & outVolumeCol).AutoFit
    End If
    
End Sub

Private Sub printResults(ws As Worksheet, col As String, row As Long, out As Variant)
    'write out to ws.cells(row, col)
    ws.Range(col & row).Value = out

End Sub

Function percentChange(openValue, closingValue)

    'calculate the percent change from start to end
    'if start is 0, result is undefined so return 0

    Dim change As Double
    
    change = closingValue - openValue
    If openValue <> 0 Then
        percentChange = change / openValue
    Else
        percentChange = 0
    End If

End Function

