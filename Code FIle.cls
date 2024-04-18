VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub wallstreet()
    Dim ws As Worksheet
    Dim total As Double
    Dim j As Integer
    Dim change As Single
    Dim start As Long
    Dim RowCount As Long
    Dim percentchange As Single

    For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Stock Volume"
        


    total = 0
    j = 0
    change = 0
    start = 2
    RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For I = 2 To RowCount
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            total = total + ws.Cells(I, 7).Value

            If total = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(I, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
            Else
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To I
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If

                change = (ws.Cells(I, 6).Value - ws.Cells(start, 3))
                
                percentchange = Round((change / ws.Cells(start, 3).Value * 100), 2)
                
                start = I + 1

                ws.Range("I" & 2 + j).Value = ws.Cells(I, 1).Value
                
                ws.Range("J" & 2 + j).Value = Round(change, 2)
                
                ws.Range("K" & 2 + j).Value = "%" & percentchange
                
                ws.Range("L" & 2 + j).Value = total

                Select Case change
                    
                    Case Is > 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                
                End Select
            
            End If

            total = 0
            
            change = 0
            
            j = j + 1
        Else
            total = total + ws.Cells(I, 7).Value
        End If
    Next I
ws.Range("Q2") = "%" & (WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100)

ws.Range("Q3") = "%" & (WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100)

ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & RowCount))


increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)

ws.Range("P2") = ws.Cells(increase_number + 1, 9)
ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
ws.Range("P4") = ws.Cells(volume_number + 1, 9)
 
Next ws

End Sub

