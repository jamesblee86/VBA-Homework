Attribute VB_Name = "Module1"
Sub StockTesting()

    ' Create initial variable to hold summary table stock tickers
    Dim Summary_Ticker As String

    ' Create initial variable to hold total volume per summary table ticker
    Dim Ticker_Total_Volume As Long
  
    ' Create initial variables to hold opening and closing stock prices per ticker in the summary table
    Dim Open_Price As Double
    Dim Close_Price As Double
  
    ' Create initial variables for yearly change in stock price and the
    ' corresponding percent change, both in the summary table
    Dim Yearly_Change As Double
    Dim Percent_Change As Double

    ' Create initial variable to track location in summary table
    Dim Summary_Table_Row As Long
  
    ' Create initial variables to track the last row and column for each worksheet
    Dim Last_Row As Long
    Dim Last_Column As Integer
    
    ' Create initial variable to track last row of the summary table
    Dim Last_Row_Summary As Long
    
    ' Create initial variables to track the greatest increase, decrease, and volume data
    Dim Greatest_Increase_Count As Double
    Dim Greatest_Decrease_Count As Double
    Dim Greatest_Volume_Count As LongLong
    
    ' Create initial variables to store the tickers for greatest increase, decrease, and volume data
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Derease_Ticker As String
    Dim Greatest_Volume_Ticker As String
    
    ' Create ws variable to cycle through worksheets
    Dim ws As Worksheet
  
  ' For loop to cycle through each worksheet
  For Each ws In Worksheets

        ' Set intitial summary table row
        Summary_Table_Row = 2
        
        ' Set initial summary table volume counter
        Ticker_Total_Volume = 0
        
        ' Establish last row for each worksheet
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Establish last column for each worksheet
        Last_Column = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        ' Loop through all stock data
        For i = 2 To Last_Row
        
            ' Conditional to test for start of new stock ticker row and test for a zero open price
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 3).Value = 0 Then
            
            ' If open price is zero, then the cell value is Null
            ws.Cells(i, 3).Value = Null
            
                ' If row is a new stock and open price is non-zero value
                ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 3).Value > 0 Then
                
                ' Set open price value
                Open_Price = ws.Cells(i, 3).Value
                        
                ' Check for last row for a given ticker
                ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                ' Establish ticker for each row in the summary table
                Summary_Ticker = ws.Cells(i, 1).Value
      
                ' Set close price from last row of each ticker
                Close_Price = ws.Cells(i, 6).Value
      
                ' Calc yearly change for each stock
                Yearly_Change = Close_Price - Open_Price
      
                'Calc percent change for each stock
                Percent_Change = ((Close_Price / Open_Price) - 1)
      
                ' Set volume total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value

                ' Print stock ticker in summary table
                ws.Range("I" & Summary_Table_Row).Value = Summary_Ticker

                ' Print yearly change in summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
      
                ' Print percent change in summary table
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
      
                ' Print stock volume in summary table
                ws.Range("L" & Summary_Table_Row).Value = Volume_Total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset volume total variable
                Volume_Total = 0
      
                'Reset yearly change variable
                Yearly_Change = 0
      
                ' Reset percent change variable
                Percent_Change = 0

            ' If the cell immediately following a row is the same ticker...
            Else

            ' Add to the volume total
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
        
        ' End summary table conditionals
        End If
        
    ' End summary table loop
    Next i
    
    ' Search for last row of summary table
    Last_Row_Summary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' Set initial value for test of greatest percentage increase in stock price
    Greatest_Increase_Count = 0
    
    ' Set initial value for test of greatest percentage decrease in stock price
    Greatest_Decrease_Count = 0
    
    ' Set initial value for test of greatest volume in a stock
    Greatest_Volume_Count = 0
    
    ' Loop through summary table data
    For j = 2 To Last_Row_Summary
    
        ' Conditional to test for greatest increase in stock prices
        If ws.Cells(j, 11).Value > Greatest_Increase_Count Then
        
        ' Set value for greatest increase counter
        Greatest_Increase_Count = ws.Cells(j, 11).Value
        
        ' Set ticker for stock with greatest increase in price
        Greatest_Increase_Ticker = ws.Cells(j, 9).Value
        
        ' Print greatest increase in a stock and the corresponding ticker in the "Greatest" table
        ws.Cells(2, 17).Value = Greatest_Increase_Count
        ws.Cells(2, 16).Value = Greatest_Increase_Ticker
            
            ' Conditional to test for greatest decrease in stock prices
            ElseIf ws.Cells(j, 11).Value < Greatest_Decrease_Count Then
            
            ' Set value for greatest decrease counter
            Greatest_Decrease_Count = ws.Cells(j, 11).Value
            
            ' Set ticker for stock with greatest decrease in price
            Greatest_Decrease_Ticker = ws.Cells(j, 9).Value
            
            ' Print greatest decrease in a stock and the corresponding ticker in the "Greatest" table
            ws.Cells(3, 17).Value = Greatest_Decrease_Count
            ws.Cells(3, 16).Value = Greatest_Decrease_Ticker
        
        ' End conditional to test for greatest increase or decrease in a stock prices
        End If
        
        ' Conditional to test for greatest volume in the summary table
        If ws.Cells(j, 12).Value > Greatest_Volume_Count Then
        
        ' Set value for greatest volume counter
        Greatest_Volume_Count = ws.Cells(j, 12).Value
        
        ' Set ticker for stock with greatest volume
        Greatest_Volume_Ticker = ws.Cells(j, 9).Value
        
        ' Print greatest volume in a stock and the corresponding ticker in the "Greatest" table
        ws.Cells(4, 17).Value = Greatest_Volume_Count
        ws.Cells(4, 16).Value = Greatest_Volume_Ticker
        
        ' End conditional for greatest volume test
        End If
        
    ' End loop through summary table data
    Next j
    
    ' Resize each column in a worksheet
    ws.Columns("A:Q").AutoFit

' Loop through next worksheet
Next ws

End Sub


