Attribute VB_Name = "StockSummary"
Sub StockSummary()

'Challenge 2: Allows the script to run for every worksheet
Dim ws As Worksheet

'Loops the script for every worksheet in the data file
For Each ws In Worksheets

Dim LastRow As Long
Dim Ticker As String
Dim StkOpen As Double
Dim StkClose As Double
Dim PerChange As Double
Dim TotStk As Double
Dim SumRow As Double

'Set SumRow to 2 to start after header row for Summary Table
SumRow = 2
'Get the opening price of first ticker in year
StkOpen = ws.Cells(2, 3).Value
'Get first ticker name of year
Ticker = ws.Cells(2, 1).Value

' Get the last row in the sheet
LastRow = ws.Cells(Rows.Count, 2).End(xlUp).Row

'Print the Headers of Summary Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'Look through every row in the sheet
For i = 2 To LastRow
    
     If Ticker <> ws.Cells(i + 1, 1).Value Then
        
        'Get closing stock price
        StkClose = ws.Cells(i, 6).Value
        
        'Add Ticker to Summary Row
        ws.Cells(SumRow, 9).Value = Ticker
        
        'Print final change on Summary Row
        ws.Cells(SumRow, 10).Value = StkClose - StkOpen
        
        'Put the percentage change on Summary Row
        If StkOpen <> 0 Then
            ws.Cells(SumRow, 11).Value = ((StkClose / StkOpen) - 1)
        Else
            ws.Cells(SumRow, 11).Value = 0
        End If
        
        'Change style to percentage
        ws.Cells(SumRow, 11).NumberFormat = "0.00%"
        
        'Get final stock volume
        TotStk = TotStk + ws.Cells(i, 7).Value
        
        'Print Total Stock Volume on Summary Row
        ws.Cells(SumRow, 12).Value = TotStk
        
            'Check if the Yearly Change was positive or negative
            If ws.Cells(SumRow, 10).Value >= 0 Then
            
                'Print green
                ws.Cells(SumRow, 10).Interior.ColorIndex = 4
            
            Else
                
                'Print red
                ws.Cells(SumRow, 10).Interior.ColorIndex = 3
            
            End If
        
        'MovesSummary Row down for the next Ticker
        SumRow = SumRow + 1
        
        'Get the opening price of next Ticker
        StkOpen = ws.Cells(i + 1, 3).Value
        
        'Resets Stock Volume for next Ticker
        TotStk = 0
        
        'Gets name of the next Ticker
        Ticker = ws.Cells(i + 1, 1).Value
        
    Else
        
        'Adds to the total of the stock
        TotStk = TotStk + ws.Cells(i, 7).Value
        
      End If
Next i



'Formats the Headers of the Summary Row
ws.Range("I1:L1").Font.FontStyle = "Bold"
ws.Range("I1:L1").HorizontalAlignment = xlCenter
ws.Range("I1:L1").Columns.AutoFit

Next ws

End Sub
