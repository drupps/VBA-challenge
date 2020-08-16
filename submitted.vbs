Sub vba()

Dim i As Long
Dim LastRow As Long
Dim StockName As String
Dim StockOpen As Double
Dim StockClose As Double
Dim StockVolume As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TickerLoop As Long
Dim ws As Worksheet

For Each ws In Worksheets
'Set defaults for each worksheet
StockVolume = 0

'Setup results tabel column headers
ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "Total Stock Volume"

' Keep track of the location for each stock in the summary table
Dim Results_Table As Integer
Results_Table = 2

    
' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'Loop through all stocks
For i = 2 To LastRow
        
    'Check if we are still within the same stock , if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        'Set the initial stock name
        StockName = StockName + Cells(i, 9).Value
        'Add the volume to the current volume
        StockVolume = StockVolume + ws.Cells(i, 7).Value

        'Set stock ticker in the results table
        ws.Cells(Results_Table, 9).Value = ws.Cells(i, 1).Value

        'Set the volume in the results table
        ws.Cells(Results_Table, 10).Value = StockVolume

        'Add one to the results table so that we don't overwrite
        Results_Table = Results_Table + 1

        'Reset counts
        StockVolume = 0
    Else
        'It's a new ticker, add the first volume
        StockVolume = StockVolume + ws.Cells(i, 7).Value

    End If
    
Next i
    
Next ws


End Sub


