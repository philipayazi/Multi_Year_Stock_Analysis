Option Explicit

Sub stock_summary()

  ' Set an initial variable for holding the ticker name
  Dim Unique_Ticker As String

  ' Set an initial variable for holding the total volume per ticker symbol
  Dim Ticker_Total As Double
  Ticker_Total = 0

  Dim year_change As Double
  Dim year_open As Double
  Dim year_close As Double
  Dim year_percentChange As Double
  Dim i As Long
  Dim j As Long
  Dim k As LongPtr
  Dim l As Long
  Dim LastRowFromA As Long
  Dim LastRowFromJ As Long
  Dim ws As Worksheet
  
  
  For Each ws In Worksheets
  
    ws.Cells(1, "J") = "ticker"
    ws.Cells(1, "K") = "year change"
    ws.Cells(1, "L") = "percent change"
    ws.Cells(1, "M") = "total volume"
    ws.Cells(1, "O") = "Open"
    ws.Cells(1, "P") = "Close"
  
      Dim Summary_SubTable_Row As Long
      Summary_SubTable_Row = 2
      
      
      ' Keep track of the location for each ticker symbol in the summary table
      Dim Summary_Table_Row As Long
      Summary_Table_Row = 2
    
      ' Loop through all ticker symbols
    
      LastRowFromA = Range("A" & Rows.Count).End(xlUp).Row
      
      For i = 2 To LastRowFromA
    
        ' Check if we are still within the same ticker symbol, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
          ' Set the Ticker name
          Unique_Ticker = ws.Cells(i, 1).Value
    
          ' Add to the Ticker Total
          Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
          year_close = ws.Cells(i, 6).Value
    
    
          ' Print the Credit Card Brand in the Summary Table
          ws.Range("J" & Summary_Table_Row).Value = Unique_Ticker
    
    
          ' Print the Brand Amount to the Summary Table
          ws.Range("M" & Summary_Table_Row).Value = Ticker_Total
          
          ws.Range("P" & Summary_Table_Row).Value = year_close
          
          
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Brand Total
          Ticker_Total = 0
          
    
        ' If the cell immediately following a row is the same brand...
        Else
    
          ' Add to the Ticker Total
          Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
          
          'year_change = (year_open - year_close)
        End If
        
            
        Next i
        
          For j = 2 To LastRowFromA

        If ws.Cells(j, 1).Value <> ws.Cells(j - 1, 1).Value Then


            year_open = ws.Cells(j, 3)

            ws.Range("O" & Summary_SubTable_Row).Value = year_open

            Summary_SubTable_Row = Summary_SubTable_Row + 1

            End If

        Next j
        


        LastRowFromJ = ws.Range("J" & Rows.Count).End(xlUp).Row

        For k = 2 To LastRowFromJ

            year_change = ws.Cells(k, "P") - ws.Cells(k, "O")
            
            If ws.Cells(k, "O") = 0 Then
            
                year_percentChange = 0
            
            Else
            
                year_percentChange = (ws.Cells(k, "P") - ws.Cells(k, "O")) / ws.Cells(k, "O")
            
            End If
            
            ws.Range("K" & k).Value = year_change

            ws.Range("L" & k).Value = year_percentChange

            ws.Range("L2:L" & LastRowFromJ).NumberFormat = "0.00%"

            If year_change > 0 Then

                ws.Range("K" & k).Interior.Color = vbGreen

            Else

                ws.Range("K" & k).Interior.Color = vbRed
            End If
            

        Next k
        

    Next
    
End Sub
