Sub Stock_Exchange()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        'Set the Header for Ticker and Total stock Volume
        ws.Cells(1, 11).Value = "Ticker"
        ws.Cells(1, 12).Value = "Total Stock volume"

        ' Created a Variable to Hold Worksheet Name, Last Row
        Dim WorksheetName As String

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'MsgBox lastRow

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
                
        'MsgBox WorksheetName
        
        ' Set an initial variable for holding the Ticker name
  
        Dim Ticker_Name As String

        ' Set an initial variable for holding the total stock volume per Ticker
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0

        ' Keep track of the location for each Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Loop through all Tickers
        For i = 2 To lastRow

            ' Check if we are still within the same Ticker, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker
                Ticker_Name = ws.Cells(i, 1).Value
                
                ' Print the Ticker in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Ticker_Name
                

                ' Print the Ticker Total Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume + ws.Cells(i, 7).Value

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the Total_Stock_Volume
                Total_Stock_Volume = 0

            ' If the cell immediately following a row is the same Ticker...
            Else

                ' Add to the Total_Stock_Volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                

            End If

        Next i
 
    Next ws
    
    

End Sub





