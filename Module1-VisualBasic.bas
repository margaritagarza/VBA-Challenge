Attribute VB_Name = "Module1"
Sub VBA()

Dim Total As Double
Dim YC As Double
Dim YO As Double
Dim YearlyCh As Double
Dim Percentage As Double
Dim counter As Integer
Dim dblMax As Double


For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"


'Determine last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set initial value for the yearly opening
YO = ws.Cells(2, 3).Value

'Set a counter to paste values
counter = 2

'Loop to determine values

 For i = 2 To Lastrow

    If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
             
            Total = ws.Cells(i, 7).Value + Total
    Else
        
            Total = ws.Cells(i, 7).Value + Total
            
            If YO = 0 Then
                
                Percentage = 0
                
            Else
            
                 YC = ws.Cells(i, 6).Value
                 YearlyCh = YC - YO
                 Percentage = CDbl((YearlyCh / YO))
                 
            End If
            

            'Print values on table
            ws.Cells(counter, 9).Value = ws.Cells(i, 1)
            ws.Cells(counter, 10).Value = YearlyCh
            ws.Cells(counter, 11).Value = Percentage
            ws.Cells(counter, 11).NumberFormat = "0.00%"
            ws.Cells(counter, 12).Value = Total
            
            'Format Color of Yearly Change
            If YearlyCh >= 0 Then
                ws.Cells(counter, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(counter, 10).Interior.ColorIndex = 3
            End If
        
            ' Set value of Yearly opening of next ticker
            YO = ws.Cells(i + 1, 3).Value
            
            ' Counts the times we skip row
            counter = counter + 1
        
            ' Reset total
            Total = 0
    
    End If
        
 Next i

Max = WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & Lastrow))
Min = WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & Lastrow))
Maxtotal = Application.WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & Lastrow))
ws.Range("P2").Value = Max
ws.Range("P3").Value = Min
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
ws.Range("P4").Value = Maxtotal

' rownummax is the row number of the maximum value
rownummax = WorksheetFunction.Match(Max, ws.Range("K" & 2 & ":" & "K" & Lastrow), 0)
' use the rownum and highlight the cell
ws.Range("O2").Value = ws.Range("I" & rownummax)
' rownummin is the row number of the min value
rownummin = WorksheetFunction.Match(Min, ws.Range("K" & 2 & ":" & "K" & Lastrow), 0)
' use the rownum and highlight the cell
ws.Range("O3").Value = ws.Range("I" & rownummin)
' rownumaxtotal is the row number of the maxtotal value
rowmaxtotal = WorksheetFunction.Match(Maxtotal, ws.Range("L" & 2 & ":" & "L" & Lastrow), 0)
' use the rownummaxtotal and highlight the cell
ws.Range("O4").Value = ws.Range("I" & rowmaxtotal)


Next


End Sub
 
