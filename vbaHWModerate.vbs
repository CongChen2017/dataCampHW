Sub HWModerate()
    
  'loop through all sheets
    For Each ws In Worksheets

        'print header of summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Set initial variables
        Dim i As Double
        Dim SumVol As Double
        SumVol = 0
        Dim ticker As String
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim YearOpen As Double
        YearOpen = ws.Cells(2, 3).Value
        Dim YearClose As Double
        Dim YearChange As Double

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

               'Set the yearly close price
               YearClose = ws.Cells(i, 6).Value

               'calculate yearly change
               YearChange = YearClose - YearOpen
    
               ' Print the ticker in summary table
                ws.Cells(Summary_Table_Row, 9).Value = ticker

                'Print the yearly change in summary table
                ws.Cells(Summary_Table_Row, 10).Value = YearChange
                    
                    'conditional formatting
                    If YearChange < 0 Then
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                    End IF

                'Print the Percent change
                    
                    If YearOpen = 0 Then
                        ws.Cells(Summary_Table_Row, 11).Value = "NA"
                    Else
                        ws.Cells(Summary_Table_Row, 11).Value = YearChange/YearOpen
                    End IF

                'change style
                ws.Cells(Summary_Table_Row, 11).Style = "Percent"

                ' Print the sum of vol in summary table
                ws.Cells(Summary_Table_Row, 12).Value = SumVol

                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                'Reset the sum of vol
                SumVol = 0

                'Reset the Year open price
                YearOpen = ws.Cells(i+1, 3).Value
        
            'if the cell immediately following a row is the same ticker...
             Else
          
               'Add to the sum of vol
               SumVol = SumVol + ws.Cells(i, 7).Value

            End If

        ' Call the next iteration
        Next i
    
     Next

End Sub
