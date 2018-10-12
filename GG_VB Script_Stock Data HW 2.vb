
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.

'Create a script that will loop through all the stocks and take the following info.
'Yearly change from what the stock opened the year at to what the closing price was.
'The percent change from the what it opened the year at to what it closed.
'The total Volume of the stock
'Ticker Symbol

'You should also have conditional formatting that will highlight positive change in green and negative change in red.
'-----------------------------------------------------------------------------------------------------
'Homework 2:VBA Multiple_year_stock_data
'-----------------------------------------------------------------------------------------------------

Sub Stock_Data()

 ' Loop through all the sheets
    
    For Each ws In Worksheets
 
        ' Insert column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        'format columns
        ws.Columns("J").NumberFormat = "0.000000000"
        ws.Columns("K").NumberFormat = "0.00%"
        
        ' Find the last row of the each (year) worksheet and find each stock symbol in column A and put it in column I
        ' Find the total volume of the each stock symbol and put it in column L so that it coincides with the its stock symbol.

        Dim LastRow As Long
        Dim i As Long
        Dim ticker_location As Integer
        Dim Total_volume As Double
        Dim stock_open As Double
        Dim stock_close As Double
        Dim stock_start As Long

        ticker_location = 2
        stock_start = 2
        Total_volume = 0
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'loop through tickers
        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ' Set the ticker 
                ws.Range("I" & ticker_location).Value = ws.Range("A" & i).Value

                'Stock opening price and increment stock_start each loop
                stock_open = ws.Cells(stock_start, 3).Value
                stock_start = i + 1 
                
             'Add the last row to volume total
                Total_volume = Total_volume + ws.Range("G" & i).Value
            
            ' Place the total volume per ticker in column L 
                ws.Range("L" & ticker_location).Value = Total_volume

            ' Get the closing value for the year 

                stock_close = ws.Cells(i, 6).Value

            ' Place the yearly change value 
           
                ws.Range("J" & ticker_location).Value = (stock_close - stock_open)

             'Percent change
                If stock_open = 0 Then
                    ws.Range("K" & ticker_location).Value = 0
                Else
                    ws.Range("K" & ticker_location).Value = ((stock_close - stock_open) / stock_open)
                End If

            ' Add 1
               ticker_location = ticker_location + 1

             ' Reset the total volume

                Total_volume = 0

            Else
            ' If the cell immediately following a row is the same stock then add to the volume for the stock 
            
               Total_volume = Total_volume + ws.Range("G" & i).Value


            End If
        
        Next i

        Dim Yearly_Change As Double
  
        ' Conditional formatting 
        Yearly_Change = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
        For i = 2 To Yearly_Change
            If ws.Range("J" & i).Value > 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & i).Value < 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 3
            End If
             
        Next i

        ws.Columns("A:L").AutoFit

    Next ws
         
'message box when complete
MsgBox ("Done")

End Sub
