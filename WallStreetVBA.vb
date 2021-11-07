Sub WallStreetVBA()

'Turn off screen updating to improve speed of script
Application.ScreenUpdating = False

'Create variables for summary table
Dim SheetCount As Integer
Dim NumRows As Long
Dim Ticker_Name As String
Dim Yearly_Change As Double
Dim Summary_Row As Integer

'Create variables for formatting
Dim Change_Rows As Long

'Count the number of sheets in the file and assign to a variable
SheetCount = Application.Sheets.Count

'Loop through the sheets (s)
For s = 1 To SheetCount

    Sheets(s).Select

    'Get the number of rows
    Cells(1, 1).Select
    NumRows = Selection.CurrentRegion.Rows.Count

    'Create the Headers
    Cells(1, 9) = "Ticker"
    Cells(1, 10) = "Yearly Change"
    Cells(1, 11) = "% Change"
    Cells(1, 12) = "Total Stock Volume"

    'Reset Variables
    First_Open = Cells(2, 3).Value
    Last_Close = 0
    Yearly_Change = 0
    Stock_Vol = 0
    Summary_Row = 2

    'Loop through all rows in sheet to create the summary table
    For i = 2 To NumRows

        'Check if still looking at the same Ticker
        If Cells(i + 1, 1) <> Cells(i, 1) Then
    
        'If not, assign final variables for the calculations
            Ticker_Name = Cells(i, 1).Value
            Last_Close = Cells(i, 6).Value
            Stock_Vol = Stock_Vol + Cells(i, 7).Value
        
            'Fill in the row on the summary table
            Cells(Summary_Row, 9) = Ticker_Name
            Cells(Summary_Row, 10) = Last_Close - First_Open
            Cells(Summary_Row, 11).NumberFormat = "0.00%"
            Cells(Summary_Row, 12) = Stock_Vol
        
            'Return Null for % Change to prevent an overflow error for dividing by zero
            If First_Open = 0 Then
                Cells(Summary_Row, 11) = Null
            Else
                Cells(Summary_Row, 11) = (Last_Close / First_Open) - 1
            End If
        
        
            'Update the variables for the next row
            First_Open = Cells(i + 1, 3).Value
            Summary_Row = Summary_Row + 1
            Stock_Vol = 0
        
        Else
    
            'Increase the stock volume with the new row values
            Stock_Vol = Stock_Vol + Cells(i, 7).Value
        
        End If
    
    Next i
        
    'Format the new summary table
    Cells(1, 10).Select
    Change_Rows = Selection.CurrentRegion.Rows.Count

    'Loop to format the yearly change column
    For i = 2 To Change_Rows

        'Check if positive
        If Cells(i, 10).Value > 0 Then
    
            'Turn Green
            Cells(i, 10).Interior.ColorIndex = 4
        
        'Check if negative
        ElseIf Cells(i, 10).Value < 0 Then
    
            'Turn Red
            Cells(i, 10).Interior.ColorIndex = 3
        
        Else
    
            'Turn Yellow
            Cells(i, 10).Interior.ColorIndex = 27
        
        End If
    
    Next i

    'Develop headers for further insight into the summary table
    Cells(2, 14).Value = "Greatest % increase"
    Cells(3, 14).Value = "Greatest % decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"

    'Use worksheet functions to determine Min/Max values

    'Greatest % increase
    Cells(2, 16).Value = WorksheetFunction.Max(Range("K:K"))
    Cells(2, 16).NumberFormat = "0.00%"
    Cells(2, 15).Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(2, 16), Range("K:K"), 0))

    'Greatest % decrease
    Cells(3, 16).Value = WorksheetFunction.Min(Range("K:K"))
    Cells(3, 16).NumberFormat = "0.00%"
    Cells(3, 15).Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(3, 16), Range("K:K"), 0))

    'Greatest total volume
    Cells(4, 16).Value = WorksheetFunction.Max(Range("L:L"))
    Cells(4, 15).Value = WorksheetFunction.Index(Range("I:I"), WorksheetFunction.Match(Cells(4, 16), Range("L:L"), 0))

    'Autofits the column widths to make the new content more presentable
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select

Next s

'Turn screen updating back on
Application.ScreenUpdating = True

End Sub