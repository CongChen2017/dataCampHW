Sub HWeasy()
    
  'loop through all sheets
    For Each ws In Worksheets

        'print header of summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"

        ' Set initial variables
        Dim i As Double
        Dim SumVol As Double
        SumVol = 0
        Dim ticker As String
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all rows
         For i = 2 To LastRow
        
            'check if ticker is the same, if not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
               'set the ticker
               ticker = ws.Cells(i, 1).Value

               ' Add to the Sum of Volume
               SumVol = SumVol + ws.Cells(i, 7).Value
    
               ' Print the ticker in summary table
                ws.Cells(Summary_Table_Row, 9).Value = ticker

                ' Print the sum of vol in summary table
                ws.Cells(Summary_Table_Row, 10).Value = SumVol

                'Add one to the summary table row
               Summary_Table_Row = Summary_Table_Row + 1

                'Reset the sum of vol
               SumVol = 0
        
            'if the cell immediately following a row is the same ticker...
             Else
          
               'Add to the sum of vol
               SumVol = SumVol + ws.Cells(i, 7).Value

            End If

        ' Call the next iteration
        Next i
    
     Next

End Sub
