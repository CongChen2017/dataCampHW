Sub HWHard()
    
  'loop through all sheets
    For Each ws In Worksheets

        'print header of summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

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
        Dim MostIn as Double
        Dim MostDe as Double
        Dim MostVol as Double

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
                        ws.Cells(Summary_Table_Row, 11).Value = 0
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
        
        ' find the greatest increased stock

        ' initial values for variables
        ticker = ws.cells(2, 9).value
        MostIn = ws.cells(2, 11).value

        'loop through all rows in summary table
        For j = 2 to Summary_Table_Row
            if ws.cells(j, 11).value > MostIn Then
                ticker = ws.cells(j, 9).value
                MostIn = ws.cells(j, 11).value
            End if
        next j

        'print result

        ws.cells(2, 16).value = ticker
        ws.cells(2, 17).value = MostIn

        'change style
        ws.Cells(2, 17).Style = "Percent"

        ' initial values for variables
        ticker = ws.cells(2, 9).value
        MostDe = ws.cells(2, 11).value

        'loop through all rows in summary table
        For j = 2 to Summary_Table_Row
            if ws.cells(j, 11).value < MostDe Then
                ticker = ws.cells(j, 9).value
                MostDe = ws.cells(j, 11).value
            End if
        next j

        'print result

        ws.cells(3, 16).value = ticker
        ws.cells(3, 17).value = MostDe

        'change style
        ws.Cells(3, 17).Style = "Percent"

        ' initial values for variables
        ticker = ws.cells(2, 9).value
        MostVol = ws.cells(2, 12).value

        'loop through all rows in summary table
        For j = 2 to Summary_Table_Row
            if ws.cells(j, 12).value > MostVol Then
                ticker = ws.cells(j, 9).value
                MostVol = ws.cells(j, 12).value
            End if
        next j

        'print result

        ws.cells(4, 16).value = ticker
        ws.cells(4, 17).value = MostVol

     Next

End Sub
