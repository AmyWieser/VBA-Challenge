Attribute VB_Name = "Module1"
Sub Stock_Market()

''Loop through all worksheets in file
For Each ws In Worksheets

'Define all of the variables
    'Define variable for ticker
            Dim Ticker As String
    'Define variable for Total Stock Volume
            Dim StockTotal As Single
            StockTotal = 0
    'Define variable for Yearly Change
            Dim Yrly_Chg As Double
    'Define variable for % Change
            Dim Pct_Chg As Double
    'Define Summary Table Row
            Dim Summary_Table_Row As Integer
    'Define Open Value
            Dim Open_Val As Double
    'Define Close Value
            Dim Close_Val As Double
    'Define Volume
            Dim Volume As Single
    'Define Last Row
            Dim lastrow As Double
            
                        
'Determine the Last Row in Worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Keep track of the location for each ticker summary table
    Summary_Table_Row = 2
    
'Set beginning open value
    Open_Val = ws.Cells(2, 3).Value

'Create column headings for summary data
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Formatting for Percent Change Column
    ws.Columns("K").NumberFormat = "0.00%"
    
    
'Create loop to summarize data for Total Volume, Yearly Change and % Change by Ticker symbol for year
       
   
    'Loop through all tickers for year
For i = 2 To lastrow
    
    'Check if we are still within the same ticker
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set Ticker
        Ticker = ws.Cells(i, 1).Value
    
    'Add to the Volume Total
        StockTotal = StockTotal + ws.Cells(i, 7).Value
    
    'Print ticker symbol in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
            
    'Print the Total Volume to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = StockTotal
        
    'Get Yearly Change
        Close_Val = ws.Cells(i, 6).Value
        Yrly_Chg = Close_Val - Open_Val
            
    'Print the Yearly Change to the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = Yrly_Chg
        
    'Calculate Percent Change
        If Open_Val <> 0 Then
             
        Pct_Chg = Yrly_Chg / Open_Val

                
    'Print Percent Change to Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Pct_Chg

        Else
        Pct_Chg = 0

    'Print Percent Change to Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Pct_Chg

        End If
        
    'Conditional Formatting for the Yearly Change
    'If greater than zero, make the cell green
    'If less than zero, make the cell red
        If Yrly_Chg > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
              
           
    'Reset Open Value
        Open_Val = ws.Cells(i + 1, 3).Value
        
    
    'Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset the Ticker Volume Total
        StockTotal = 0
        
    'If the cell immediately following a row is the ticker
        Else
    
    'Add to the Stock Volume
        StockTotal = StockTotal + ws.Cells(i, 7).Value
        

        End If


Next i


Next ws
    
End Sub
