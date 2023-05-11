Sub tickerSummary()

'Set variable to loop through WorkSheets
Dim ws As Integer

'Set a cariable to identify number of sheets in the WorkBook
Dim sheet_count As Integer
sheet_count = ActiveWorkbook.Worksheets.Count

'Run script on different worksheets
For ws = 1 To sheet_count
    ActiveWorkbook.Worksheets(ws).Activate

    'Set a variable to identify last row in a table
    Dim num_row As Long
    num_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set an initial variable to hold Ticker Name
    Dim Ticker_name As String
    
    ' Set a variable to hold the date for the Opening Price of the year
    Dim Open_date As Double
    
    ' Set an initial variable to hold the Opening Price of the year
    Dim Open_price As Double
    
    ' Set a variable to hold the date for the Closing Price of the year
    Dim Close_date As Double
    
    ' Set an initial variable to hold the Closing Price of the year
    Dim Close_price As Double
    
    'Set an initial variable to hold the Yearly change
    Dim Yearly_change As Double
    
    ' Set an initial variable to hold the Greatest % Increase
    Dim great_increase As Double
    great_increase = 0
    
    ' Set an initial variable to hold the Greatest % Decrease
    Dim great_decrease As Double
    great_decrease = 0
    
    ' Set an initial variable to hold the Greatest Total Volume
    Dim great_total_vol As Double
    great_total_vol = 0
    
    ' Set a value to keep track of the location for each Ticker name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Set a variable to loop through the original table
    Dim i As Long
    
    'Set a variable to loop through the summary table
    Dim j As Long
    
    'Convert <date> column values to number - solution is taken from https://stackoverflow.com/questions/36771458/vba-convert-text-to-number
    [B:B].Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    
    'Sort <ticker> column in a ascending order
    Range(Cells(1, 1), Cells(num_row, 7)).Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    'Create columns for ticker_name, yearly_change, percent_change, total_stock_volume
    Range("I1").Value = "Ticker name"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
      
    'Populate column ticker_name with unique values
        'Loop through all the ticker records in a table
        For i = 2 To num_row
        
        ' Check if a ticker name is not yet in the summary table
        If Application.WorksheetFunction.CountIf(Range(Cells(1, 9), Cells(Summary_Table_Row, 9)), Cells(i, 1).Value) = 0 Then
    
            ' Set the Ticker name
            Ticker_name = Cells(i, 1).Value
          
            ' Print the Ticker name into the Summary Table
            Cells(Summary_Table_Row, 9).Value = Ticker_name
          
            ' Count and print the Total volume for the printed ticker name
            Cells(Summary_Table_Row, 12).Value = Application.WorksheetFunction.SumIf(Range("A:A"), Ticker_name, Range("G:G"))
          
            ' Set open and close dates
            Open_date = Cells(i, 2).Value
            Close_date = Cells(i, 2).Value
                    
            ' Set initial openning and closing prices
            Open_price = Cells(i, 3).Value
            Close_price = Cells(i, 6).Value
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
                              
        'If the ticker is already in the summary table, check if there is an earlier date with open price
        ElseIf Cells(i, 2).Value < Open_date Then
            Open_price = Cells(i, 3).Value
        
        'Check if there is later date with close price
        ElseIf Cells(i, 2).Value > Close_date Then
            Close_price = Cells(i, 6).Value
                              
        End If
        
        'Calculate yearly change
        Yearly_change = Close_price - Open_price
        
        'Print yearly change
        Cells(Summary_Table_Row - 1, 10).Value = Yearly_change
        
        'Format for the yearly change cells as negative - red, positive - green
        If Yearly_change > 0 Then
            Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 4
        ElseIf Yearly_change < 0 Then
            Cells(Summary_Table_Row - 1, 10).Interior.ColorIndex = 3
        
        End If
        
        'Calculate and print Percent change Value
        Cells(Summary_Table_Row - 1, 11).Value = Close_price / Open_price - 1
        
        'Format Percent change Value
        Cells(Summary_Table_Row - 1, 11).NumberFormat = "0.00%"
    
      Next i
        
    'Set a value to identify last row in a summary table
    Dim num_row_new As Long
    num_row_new = Cells(Rows.Count, 9).End(xlUp).Row
    
    'Summarize greatest values
        'Create column headers and row headers.
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Fill in the "greatest" summary table
        For j = 2 To num_row_new
            
        'Identify Greatest % increase
        If Cells(j, 11).Value > great_increase Then
            great_increase = Cells(j, 11).Value
            
            'Print greatest % increase in the table
            Range("Q2").Value = great_increase
            Range("Q2").NumberFormat = "0.00%"
            
            'Print corresponding Ticker name
            Range("P2").Value = Cells(j, 9).Value
        End If
        
        'Identify Greatest % decrease
        If Cells(j, 11).Value < great_decrease Then
            great_decrease = Cells(j, 11).Value
            
            'Print greatest % decrease in the table
            Range("Q3").Value = great_decrease
            Range("Q3").NumberFormat = "0.00%"
            
            'Print corresponding Ticker name
            Range("P3").Value = Cells(j, 9).Value
        End If
        
        'Identify Greatest Total Volume
        If Cells(j, 12).Value > great_total_vol Then
            great_total_vol = Cells(j, 12).Value
            
            'Print greatest total volume in the table
            Range("Q4").Value = great_total_vol
            
            'Print corresponding Ticker name
            Range("P4").Value = Cells(j, 9).Value
        End If
        
        Next j
        
    'Adjust Columns width
    Columns("I:Q").AutoFit
    
    'Reset number of rows in a table to 0
    num_row = 0
        
Next ws
    

End Sub

