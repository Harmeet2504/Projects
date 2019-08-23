  'Analysing stock data for single sheet

Sub stockHard()
    
        'Creating a summary table
        'Declare variables to store ticker name, total volume, yearly_change, percent_change
        Dim tickername As String
        Dim totalvolume, yearly_change, percent_change As Double
        
         'Labelling Summary Table headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Initialize totalvolume
        totalvolume = 0

        'Declaring variable to store row number in the summary table, setting it  to 2 as 1st is for headers
        Dim summary_row As Integer
        summary_row = 2
        
        'Declare variables to store open and close price to calculate yearly change 
        Dim open_price, close_price As Double
        
        'Set initial open_price
        open_price = Cells(2, 3).Value
        
        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows of the ticker names and checks the condition if the value of previous cell is different from current cell
        For i = 2 To lastrow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              'Store the values of ticker name and totalvolume from respective cells
              tickername = Cells(i, 1).Value
              totalvolume = totalvolume + Cells(i, 7).Value

              'Print the ticker name and total volume in the summary table
              Range("I" & summary_row).Value = tickername
              Range("L" & summary_row).Value = totalvolume

              'Store closing price
              close_price = Cells(i, 6).Value

              'Calculate yearly change
               yearly_change = (close_price - open_price)
              
              'Print yearly change in the summary table
              Range("J" & summary_row).Value = yearly_change

              'Calculate percent change, check if a condition renders it zero
                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = yearly_change / open_price
                End If

              'Print the value in the required number format in the summary table
              Range("K" & summary_row).Value = percent_change
              Range("K" & summary_row).NumberFormat = "0.00%"
   
              'Reset the summary row counter, total volume, open price
              summary_row = summary_row + 1
              totalvolume = 0
              open_price = Cells(i + 1, 3)

            Else
              
               'Add the totalvolume to return the last value
              totalvolume = totalvolume + Cells(i, 7).Value
            
            End If
        
        Next i

    'Conditional formatting to show +ve and -ve yearly change
    
    'Find the last row of the summary table and loop through
    lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To lastrow_summary
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

    'Overall summary of the analysis of stock price changes, labelling the cells
    
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    'Loop through the summary table
        For i = 2 To lastrow_summary
        
            'Find maximum percent change using Application.WorksheetFunction.Max
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Find minimum percent change using Application.WorksheetFunction.Min
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the greatest total volume
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub
  