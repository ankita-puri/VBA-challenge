
Sub VBA_Challenge():

For Each ws In Worksheets
        
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


        
Dim Ticker As String
Dim Total_Stock As Double
Dim Table As Long
Dim Pre_Table As Long
Dim Yearly_Change As Double
Dim Yearly_Open As Double
Dim Yearly_Close As Double
Dim Percent As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total As Double
Dim Value As Double
             
Total_Stock = 0
Table = 2
Pre_Table = 2
Greatest_Increase = 0
Greatest_Decrease = 0
Greatest_Total = 0

For i = 2 To LastRow
            
Total_Stock = Total_Stock + ws.Cells(i, 7).Value
            
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
ws.Range("I" & Table).Value = Ticker
ws.Range("L" & Table).Value = Total_Stock
Total_Stock = 0
Yearly_Open = ws.Range("C" & Pre_Table)
Yearly_Close = ws.Range("F" & i)
Yearly_Change = Yearly_Close - Yearly_Open
ws.Range("J" & Table).Value = Yearly_Change
            
If Yearly_Open = 0 Then
Percent = 0

Else
Yearly_Open = ws.Range("C" & Pre_Table)
Percent = Yearly_Change / Yearly_Open

End If
            
ws.Range("K" & Table).NumberFormat = "0.00%"
ws.Range("K" & Table).Value = Percent
            
If ws.Range("J" & Table).Value >= 0 Then
ws.Range("J" & Table).Interior.ColorIndex = 4
                
Else
ws.Range("J" & Table).Interior.ColorIndex = 3

End If
            
Table = Table + 1
Pre_Table = i + 1

End If

Next i

LastRow_Value = ws.Cells(Rows.Count, 11).End(xlUp).Row
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"
        
For j = 2 To LastRow_Value

If ws.Range("K" & j).Value > Greatest_Increase Then
Greatest_Increase = ws.Range("K" & j).Value
ws.Range("Q2").Value = Greatest_Increase
ws.Range("P2").Value = ws.Range("I" & j).Value

End If
            
If ws.Range("K" & j).Value < Greatest_Decrease Then
Greatest_Decrease = ws.Range("K" & j).Value
ws.Range("Q3").Value = Greatest_Decrease
ws.Range("P3").Value = ws.Range("I" & j).Value

End If

If ws.Range("L" & j).Value > Greatest_Total Then
Greatest_Total = ws.Range("L" & j).Value
ws.Range("Q4").Value = Greatest_Total
ws.Range("P4").Value = ws.Range("I" & j).Value

End If

Next j
Next ws
End Sub


