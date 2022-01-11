Attribute VB_Name = "Module1"
Sub StockCalculations()

Dim TickerName, PercentChange As String
Dim i, LastRow, SummaryTable As Long
Dim VolumeTotal, Opening, Closeing, YearlyChange As Double

'Loop for all worksheets (https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/)
Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

    'Determine last row #
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of the location for each ticker in the summary table
    SummaryTable = 2
    VolumeTotal = 0
    Opening = Range("C2").Value
    closing = 0
    
    'Headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    
        'Loop through all tickers
        For i = 2 To LastRow
            
        VolumeTotal = VolumeTotal + Cells(i, 7).Value
            
                'Check if we are within the same ticker name
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                    'Set the ticker name
                    TickerName = Cells(i, 1).Value
                    
                    'Set YearlyChange calc
                    closing = Cells(i, 6).Value
                    YearlyChange = closing - Opening
                    
                        'added to address the overflow error when percentchange runs
                        'https://stackoverflow.com/questions/38246478/how-to-solve-runtime-error-11-division-by-0-in-vba
                        If Opening = 0 Or IsEmpty(Opening) Then
                        PercentChange = 0
                        Else
                        PercentChange = FormatPercent((YearlyChange / Opening), 2)
                        End If
                    
                    'Print the Ticker Name/VolTotal in the Summary Table
                    Range("I" & SummaryTable).Value = TickerName
                    Range("L" & SummaryTable).Value = VolumeTotal
                    Range("J" & SummaryTable).Value = YearlyChange
                    Range("K" & SummaryTable).Value = PercentChange
                        
                    'Add one to the summary table row
                    SummaryTable = SummaryTable + 1
                    
                    VolumeTotal = 0
                    Opening = Cells(i + 1, 3).Value
                    closing = 0
                    PercentChange = 0
                    
                End If
         
        Next i
        
Next ws

'Colors yearly change if pos/neg in the cell box
For Each ws In Worksheets
ws.Activate
For i = 2 To ws.Cells(Rows.Count, 10).End(xlUp).Row
            If (Cells(i, 10) > 0) Then
                'If positive, color green
                Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                'If negative, color red
                Cells(i, 10).Interior.ColorIndex = 3
        
            End If
        Next i
Next ws
            

End Sub
