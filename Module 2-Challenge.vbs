
Sub Main()

    Call ColumnTitlesFunc
    Call TickerSymbol
    Call YearlyChange
    Call YearlyChangeColour
    
End Sub
Sub ColumnTitlesFunc()
'column titles

Dim i As Integer

For i = 1 To Worksheets.Count

    Worksheets(i).Select
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

Next i

End Sub

Sub TickerSymbol()

Dim row As Long
Dim i As Integer

For i = 1 To Worksheets.Count

    Worksheets(i).Select
    row = Cells(Rows.Count, "A").End(xlUp).row
    ActiveSheet.Range("A2:A" & row).AdvancedFilter _
    Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("I2"), _
    Unique:=True
    Range("I2").Delete Shift:=xlUp

Next i

End Sub
Sub YearlyChange()

Dim i As Long
Dim j As Integer
Dim beginning As Long
Dim ticker As String
Dim diff As Double
Dim tickerRow As Long

For j = 1 To Worksheets.Count
    
    Worksheets(j).Select
    LastRow = Cells(Rows.Count, "A").End(xlUp).row
    ticker = Range("A2").Value
    beginning = 2
    tickerRow = 2
    totalVolume = 0
    
    For i = 2 To LastRow

        totalVolume = totalVolume + Range("G" & i).Value
        
        If ticker <> Range("A" & i).Value Then
            
            'save and reset total volume
            Range("L" & tickerRow).Value = totalVolume - Range("G" & i).Value
            totalVolume = 0
            
            'yearly change for current ticker
            diff = Round(Range("C" & beginning).Value - Range("F" & i - 1).Value, 4)
            
            'set value for ticker yearly change
            Range("J" & tickerRow).Value = diff
            
            'set yearly percentage change for ticker
            percentChange = FormatPercent(Round(Abs(diff) / Range("C" & beginning), 4), 4)
            Range("K" & tickerRow).Value = percentChange
            
            'move values to next ticker
            beginning = i
            ticker = Range("A" & i).Value
            tickerRow = tickerRow + 1

        ElseIf i = LastRow Then
            
            'set value for ticker yearly change
            diff = Round(Range("C" & beginning).Value - Range("F" & i - 1).Value, 4)
            Range("J" & tickerRow).Value = diff
            
             'set yearly percentage change for ticker
            percentChange = FormatPercent(Round(Abs(diff) / Range("C" & beginning), 4), 4)
            Range("K" & tickerRow).Value = percentChange
            
             'save total volume
            Range("L" & tickerRow).Value = totalVolume
            
            tickerRow = tickerRow + 1
            
        End If

    Next i
        
    Range("I:L").EntireColumn.AutoFit
    
Next j
 
End Sub

Sub YearlyChangeColour()

For j = 1 To Worksheets.Count
Worksheets(j).Select

LastRow = Cells(Rows.Count, "J").End(xlUp).row

    For i = 2 To LastRow
        If Range("J" & i) > 0 Then
        Range("J" & i).Interior.ColorIndex = 4
        Else
        Range("J" & i).Interior.ColorIndex = 3
        
        End If
    
    Next i
    
Next j

End Sub




