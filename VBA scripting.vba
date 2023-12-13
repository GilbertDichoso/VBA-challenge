Attribute VB_Name = "Module1"
Sub VBA_Scripting_Assignment()

Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()


'New repository created

Dim tickername As String
Dim tickervolume As Double
        tickervolume = 0
Dim summary_table_row As Integer
        summary_table_row = 2
        
 ' Yearly change = closeprice at the year end - open price at beginning of the year
Dim open_price As Integer
        open_price = Cells(2, 3).Value
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'labels of description
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly change"
Cells(1, 11).Value = "Percent change"
Cells(1, 12).Value = "Total Stock Volume"

'count the number of rows
lastrow = Cells(Rows.count, 1).End(xlUp).Row

For i = 2 To lastrow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'put ticker name
            tickername = Cells(i, 1).Value
            
            'add volume of stocks
            tickervolume = tickervolume + Cells(i, 7).Value
            
            'put the ticker name in the summary table
            Range("I" & summary_table_row).Value = tickername
            
            'put the stock volume for each ticker in the summary table
            Range("L" & summary_table_row).Value = tickervolume
            
            'open_price
            open_price = Cells(i, 3).Value
            
            'closing price
            close_price = Cells(i, 6).Value
            
            'Yearly change computation
            yearly_change = ((close_price) - (open_price))
            
            'put the yearly change on summary table
            Range("J" & summary_table_row).Value = yearly_change
            
            'confirm for Mod
                If (open_price = 0) Then
                
                percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                    
                End If
                
             'input yearly change on the summary table
             Range("K" & summary_table_row).Value = percent_change
             Range("K" & summary_table_row).NumberFormat = "0.00%"
             
             
             'reset row counter
             summary_table_row = summary_table_row + 1
             
             'reset stocks to 0
             tickervolume = 0
             
             'reset opening price
             open_price = Cells(i + 1, 3)
             
       Else
       
            'add volume of stocks
            tickervolume = tickervolume + Cells(i, 7).Value
            
       End If
       
       
   Next i
   
   
   
   'Conditionalformatting
   
   lastrow_summary_table = Cells(Rows.count, 9).End(xlUp).Row
   
   'Color code yearly change
        For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i


   
    'Highlight the stock price changes
    'First label the cells according to the sample .png provided in the assignment

        Cells(2, 15).Value = "Greatest percentage increase"
        Cells(3, 15).Value = "Greatest percentage decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    'Determine the max and min values in column "Percent Change" and just max in column "Total Stock Volume"
    'Then collect the ticker name, and the corresponding values for the percent change and total volume of trade for that ticker
    '
        For i = 2 To lastrow_summary_table
            'Find the maximum percent change
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
                Cells(2, 16).Value = Cells(i, 9).Value
                Cells(2, 17).Value = Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
                Cells(3, 16).Value = Cells(i, 9).Value
                Cells(3, 17).Value = Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
                Cells(4, 16).Value = Cells(i, 9).Value
                Cells(4, 17).Value = Cells(i, 12).Value
            
            End If
        
        Next i
        
End Sub
            
            
            
            
            
            


