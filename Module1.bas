Attribute VB_Name = "Module1"
Sub stockAnalyzer()

' Iterate through all worksheets in the workbook
' I got how to iterate through the worksheets and access the cells of
' each worksheet from Class Exercise 7 of the 2nd class of the week we went over VBA
For Each ws In Worksheets

    ' VARIABLES
    ' Quarterly change of each stock
    Dim qrtly_change As Double
    qrtly_change = 0
    ' Total volume of a stock during a quarter
    Dim stock_total As LongLong
    stock_total = 0
    ' Opening week of each stock
    Dim opening_wk As Long
    opening_wk = 2
    ' Stock ticker
    Dim ticker As String
    
    
    ' Keep track of the location for each stock in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  
    'Column labels for summary able
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"




  ' Loop through all the stocks on all days of the quarter
  For i = 2 To 93001
  
    ' Add to total volume of stock
    stock_total = stock_total + ws.Cells(i, 7)


    ' Check if we are still within the same stock, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


        ' TICKER
        ' Grab ticker for stock in summary table
        ticker = ws.Cells(i, 1).Value
        ' Print the stock ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker


        ' QUARTERLY CHANGE
        ' Calculate the quarterly change for stock
        qrtly_change = ws.Cells(i, 6).Value - ws.Cells(opening_wk, 3).Value
        ' Put quarterly change in the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = qrtly_change
        ' Color the quarterly changes greater than 0 green and the quarterly changes less than 0 red
        If qrtly_change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        
        ElseIf qrtly_change < 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
        End If

        
        
        ' PERCENTAGE CHANGE
        ' Calculate percentage change
        Dim percentage_change As Double
        percentage_change = qrtly_change / ws.Cells(opening_wk, 3)
        ' Print percentage change in the Summary Table
        ' I found the Format function from https://excelvbatutor.com/vba_lesson9.htm
        ws.Range("K" & Summary_Table_Row).Value = Format(percentage_change, "Percent")
    
        
        ' TOTAL STOCK VOLUME
        ' Print total stock volume in the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = stock_total



        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
    
        ' Reset the Brand Total
        qrtly_change = 0
        
        ' Reset the total volume of stock
        stock_total = 0
      
        ' Reset the opening week
        opening_wk = i + 1
        
        
    End If
    
    
  Next i
  
  
Next ws


End Sub
