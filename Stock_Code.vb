Sub stockVolume()

Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate

' Set initial variables for names
Dim Stock_Name As String
Dim Volume_Total As Double
Dim Open_Value As Double
Dim Close_Value As Double
Dim Stock_Counter As Double
Volume_Total = 0
Stock_Counter = 0

' Keep track of the location for each stock in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Find last row
Dim lastRow As Long
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Add chart names
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Volume"

' Loop through all stock volumes
    For i = 2 To lastRow
                        
' Check if we are still within the same stock name, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i - Stock_Counter, 6) = 0 Then
             Stock_Name = Cells(i, 1).Value
             Open_Value = Cells(i - Stock_Counter, 3).Value
             Close_Value = Cells(i, 6).Value
             Volume_Total = Volume_Total + Cells(i, 7).Value
             Range("I" & Summary_Table_Row).Value = Stock_Name
             Range("J" & Summary_Table_Row).Value = "N/A"
             Range("K" & Summary_Table_Row).Value = 0
             Range("L" & Summary_Table_Row).Value = Volume_Total
             Summary_Table_Row = Summary_Table_Row + 1
             Volume_Total = 0
             Stock_Counter = 0
        
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the stock name, open value, close value, and add to volume total
            Stock_Name = Cells(i, 1).Value
            Open_Value = Cells(i - Stock_Counter, 3).Value
            Close_Value = Cells(i, 6).Value
            Volume_Total = Volume_Total + Cells(i, 7).Value

' Print to the Summary Table and add one to the summary table row then reset total
            Range("I" & Summary_Table_Row).Value = Stock_Name
            Range("J" & Summary_Table_Row).Value = Close_Value - Open_Value
            Range("K" & Summary_Table_Row).Value = (Close_Value - Open_Value) / Open_Value
            Range("L" & Summary_Table_Row).Value = Volume_Total
            Summary_Table_Row = Summary_Table_Row + 1
            Volume_Total = 0
            Stock_Counter = 0

' If the cell immediately following a row is the same stock add volume to total
        Else
            Volume_Total = Volume_Total + Cells(i, 7).Value
            Stock_Counter = Stock_Counter + 1
        End If
    Next i

' Add formating to cells
    For i = 2 To lastRow
        If Cells(i, 10) < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        ElseIf Cells(i, 10) > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        End If
    Next i

Range("K2:K" & lastRow).NumberFormat = "0.00%"
Range("J2:J" & lastRow).NumberFormat = "0.00"

'name the cells
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker Symbol"
Range("P1").Value = "Value"

' Find last row
Dim laststatRow As Long
laststatRow = Cells(Rows.Count, 9).End(xlUp).Row

'Find min and max values
Dim Dmax As Double
Dim Dmin As Double
Dim MaxVol As Double

Dmax = WorksheetFunction.Max(Range("k2:k" & laststatRow).Value)
Dmin = WorksheetFunction.Min(Range("k2:k" & laststatRow).Value)
MaxVol = WorksheetFunction.Max(Range("l2:l" & laststatRow).Value)

Range("P2").Value = Dmax
Range("P3").Value = Dmin
Range("P4").Value = MaxVol

'Find max % Change name
For i = 2 To laststatRow
    If Cells(i, 11).Value = Dmax Then
        Range("O2").Value = Cells(i, 9).Value
    End If
Next i

'Find min % Change name
For i = 2 To laststatRow
    If Cells(i, 11).Value = Dmin Then
        Range("O3").Value = Cells(i, 9).Value
    End If
Next i

'Find max volume name
For i = 2 To laststatRow
    If Cells(i, 12).Value = MaxVol Then
        Range("O4").Value = Cells(i, 9).Value
    End If
Next i

Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"

Next

End Sub
