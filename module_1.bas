Attribute VB_Name = "Module1"
Sub stock_data()

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
    
    ' initializes all variables we will need
    Dim Maximum As Double
    Dim Maximum_ticker As String
    
    Dim Minimum As Double
    Dim Minimum_ticker As String
    
    Dim Volume As Double
    Dim Volume_ticker As String
    
    Maximum = 0
    Minimum = 0
    Volume = 0
    
    ' cycles through each stock
    For Row = 2 To 10000
    
        ' exits the for loop once we have cycled through all stocks
        If IsEmpty(ws.Cells(Row, 11).Value) Then
            Exit For
        End If
        
        ' finds the stock with the maximum percent change
        If ws.Cells(Row, 11).Value > Maximum Then
            Maximum = ws.Cells(Row, 11).Value
            Maximum_ticker = ws.Cells(Row, 9).Value
            
        ' finds the stock with the minimum percent change
        ElseIf ws.Cells(Row, 11).Value < Minimum Then
            Minimum = ws.Cells(Row, 11).Value
            Minimum_ticker = ws.Cells(Row, 9).Value
            
        End If
        
        ' finds the stock with the maximum total stock volume
        If ws.Cells(Row, 12).Value > Volume Then
            Volume = ws.Cells(Row, 12).Value
            Volume_ticker = ws.Cells(Row, 9).Value
            
        End If
        
    Next Row
    
    ' inputs all of these variables into the corresponding entries of our table
    ws.Cells(2, 16).Value = Maximum_ticker
    ws.Cells(2, 17).Value = FormatPercent(Maximum)
    ws.Cells(3, 16).Value = Minimum_ticker
    ws.Cells(3, 17).Value = FormatPercent(Minimum)
    ws.Cells(4, 16).Value = Volume_ticker
    ws.Cells(4, 17).Value = Volume

Next ws

End Sub
