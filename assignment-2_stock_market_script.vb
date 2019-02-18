Sub stock_market()

      ' Declare Current as a worksheet object variable
      Dim ws As Worksheet
      
      ' Loop through all of the worksheets in the active workbook.
       For Each ws In Sheets
           ws.Activate
            
            ' Set an initial variable for holding the ticker
            Dim ticker As String
      
            ' Set an initial variable for holding the total volum per ticker
            Dim TickerVol_Total As Double
            Ticker_TotalVolum = 0

            ' Keep track of the location for each ticker in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2

            ' Dynamically detect the number of rows in a column on a sheet
            Dim area As Range
            Dim MaxRow As Long
            Dim iAreaCount As Long
  
            MaxRow = Range("A" & Rows.Count).End(xlUp).Row
            Set area = Range("A1:A" & MaxRow)
            iAreaCount = area.Count

            ' Loop through all stock
            For i = 2 To MaxRow

            ' Check if we are still within the same ticker, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the ticker
            ticker = Cells(i, 1).Value

            ' Add to the Ticker Total
            Ticker_TotalVolum = Ticker_TotalVolum + Cells(i, 7).Value

            ' Print the ticker type in the Summary Table
            Range("I1").Value = "Ticker" 

            ' Print the total stock volume in the Summary Table
            Range("J1").Value = "Total Stock Volume"

            ' Print the ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = ticker

            ' Print the ticker Amount to the Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker_TotalVolum

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the ticker Total
            Ticker_TotalVolum = 0

            ' If the cell immediately following a row is the same ticker...
            Else

            ' Add to the ticker Total
            Ticker_TotalVolum = Ticker_TotalVolum + Cells(i, 7).Value

            End If

            Next i

        Next ws
            
    End Sub

   

   