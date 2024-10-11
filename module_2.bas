Attribute VB_Name = "Module2"
Sub stock_data()

' cycles through each worksheet
Dim ws As Worksheet

For Each ws In Worksheets
    
    ' tracks ending of previous ticker's data
    ' helpful to track each stocks' first and last rows for easy reference
    Dim line_tracker As Long
    line_tracker = 1
    
    ' add labels to columns we plan to add
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' cycles through each stock
    For i = 1 To 10000
    
        ' exits for loop when we run out of rows of data
        If IsEmpty(ws.Cells(line_tracker + 1, 1).Value) Then
            Exit For
        End If
            
        ' current stock
        Dim ticker As String
        ticker = ws.Cells(line_tracker + 1, 1).Value
        
        ' find first and last row of current ticker
        ' only need these for easy reference later
        
        ' first row
        Dim first_row As Long
        first_row = line_tracker + 1
        
        'last row
        For j = 1 To 1000
            If ws.Cells(first_row + j, 1).Value <> ticker Then
                Dim last_row As Long
                last_row = first_row + j - 1
                line_tracker = last_row
                Exit For
            End If
        Next j
        
        ' Inputs ticker of the current stock into the "Ticker" column
        ws.Cells(i + 1, 9).Value = ticker
        
        ' Inputs the quarterly change of the current stock into the "Quarterly Change" column
        Dim change As Double
        change = ws.Cells(last_row, 6).Value - ws.Cells(first_row, 3).Value
        ws.Cells(i + 1, 10).Value = change
        
        ' Formats positive change to green fill and negative change to red fill
        If change > 0 Then
            ws.Cells(i + 1, 10).Interior.ColorIndex = 4
        ElseIf change < 0 Then
            ws.Cells(i + 1, 10).Interior.ColorIndex = 3
        End If
        
        ' Inputs the percent change of the current stock into the "Percent Change" column
        Dim percent As Double
        percent = change / ws.Cells(first_row, 3).Value
        ws.Cells(i + 1, 11).Value = FormatPercent(percent)
        
        ' Inputs the total stock volume of the current stock into the "Total Stock Volume" column
        Dim vol_sum As Variant
        vol_sum = 0
        ' takes the sum of the volume for each row in the current stock
        For K = first_row To last_row
            vol_sum = vol_sum + ws.Cells(K, 7).Value
            Next K
        ws.Cells(i + 1, 12).Value = vol_sum
        
    Next i
    
    ' adds row and column labels for ticker comparison table
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' create percents and volumes arrays that will store the perecents and volumes of all stocks
    Dim percents() As Double
    Dim volumes() As Variant
    
    ' cycle through each stock
    Dim r As Integer
    For r = 2 To 10000
    
        ' exits for loop when we have cycled through all stocks
        If IsEmpty(ws.Cells(r, 11).Value) Then
            Exit For
        End If
        
        ' adds percent change of all stocks into the percents array
        ReDim Preserve percents(r)
        percents(r) = ws.Cells(r, 11).Value
        
        ' adds total volume of all stocks into the volumes array
        ReDim Preserve volumes(r)
        percents(r) = ws.Cells(r, 11).Value
        volumes(r) = ws.Cells(r, 12).Value
        
        Next r
    
    ' calculates the values needed for the table
    Maximum = WorksheetFunction.Max(percents)
    Minimum = WorksheetFunction.Min(percents)
    Volume_Max = WorksheetFunction.Max(volumes)
    
    ' inputs calculated values to the table
    ws.Cells(2, 17).Value = Maximum
    ws.Cells(3, 17).Value = Minimum
    ws.Cells(4, 17).Value = Volume_Max
    
    ' finds and inputs tickers for the table
    ws.Cells(2, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(Maximum, ws.Range("K:K"), 0))
    ws.Cells(3, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(Minimum, ws.Range("K:K"), 0))
    ws.Cells(4, 16).Value = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(Volume_Max, ws.Range("L:L"), 0))
    
Next ws

End Sub

