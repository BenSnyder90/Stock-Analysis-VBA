Attribute VB_Name = "Module2"
Sub Challenge()

'Challenge 1: Make the sheet show the Greatest % Increase,
'             Greatest % Decrease, and Greatest Total Volume

Dim ws As Worksheet

For Each ws In Worksheets

Dim SumLastRow As Double
Dim PerInc As Double
Dim PerDec As Double
Dim TotVol As Double
Dim SumTickerInc As String
Dim SumTickerDec As String
Dim SumTickerVol As String

'Print the Headers for the Ticker and Value Comparison
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

'Change style to percentage
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 16).NumberFormat = "0.00%"


'Gets the last row of the summary table
SumLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Sets all the comparison values to the first result on the summary table
PerInc = ws.Cells(2, 11).Value
SumTickerInc = ws.Cells(2, 9).Value
PerDec = ws.Cells(2, 11).Value
SumTickerDec = ws.Cells(2, 9).Value
TotVol = ws.Cells(2, 12).Value
SumTickerInc = ws.Cells(2, 9).Value


'Looks through every row in the Summary Chart to compare values
For j = 3 To SumLastRow

    'Checks if the Percent Change in the loop is higher than the stored value.
    'If it is higher, the stored value becomes the current value
    If PerInc < ws.Cells(j, 11).Value Then
        PerInc = ws.Cells(j, 11).Value
        SumTickerInc = ws.Cells(j, 9).Value
    End If
    
    'Checks if the Percent Change in the loop is lower than the stored value.
    'If it is smaller, the stored becomes the current value
    If PerDec > ws.Cells(j, 11).Value Then
        PerDec = ws.Cells(j, 11).Value
        SumTickerDec = ws.Cells(j, 9).Value
    End If
    
    'Checks if the Total Stock Volume in the loop is higher than the stored value.
    'If it is higher, the stored becomes the current value
    If TotVol < ws.Cells(j, 12).Value Then
        TotVol = ws.Cells(j, 12).Value
        SumTickerVol = ws.Cells(j, 9).Value
    End If
    
Next j

'Prints the results after looping through the summary table
ws.Cells(2, 15).Value = SumTickerInc
ws.Cells(2, 16).Value = PerInc
ws.Cells(3, 15).Value = SumTickerDec
ws.Cells(3, 16).Value = PerDec
ws.Cells(4, 15).Value = SumTickerVol
ws.Cells(4, 16).Value = TotVol

'Formats the Headers of the Comparison Table
ws.Range("O1:P1").Font.FontStyle = "Bold"
ws.Range("O1:P1").HorizontalAlignment = xlCenter
ws.Range("N1:P5").Columns.AutoFit

Next ws
End Sub

