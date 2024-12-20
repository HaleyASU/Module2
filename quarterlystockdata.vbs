VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub quarterlystockdata()

    'declare ws as a worksheet objective variable
    Dim ws As Worksheet
    
    'create variables for the data
    Dim i As Long
    Dim ticker As String
    Dim quarterlychange As Double
    Dim percentchange As Double
    Dim stocktotal As Double
    
    'variables for tracking first open and last close prices
    Dim openingprice As Double
    Dim closingprice As Double
         
    'variables for greatest increase, decrease, and total volume
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim greatesttotalvolume As LongLong
    Dim maxticker As String
    Dim increaseticker As String
    Dim decreaseticker As String
    Dim totalvolumeticker As String
    
    'loop through all worksheets in the workbook
    For Each ws In Worksheets
    
        'tracking variables
        quarterlychange = 0
        stocktotal = 0
        greatestincrease = 0
        greatestdecrease = 0
        greatesttotalvolume = 0
    
        'track location of each ticker and changes
        Dim summary_table_row As Long
        summary_table_row = 2
    
        'calculate last row before loop
        Dim lastrow As Long
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
     
        'set headers for summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'reset first open price for first row
        openingprice = ws.Cells(2, 3).Value
        
        'loop through rows in the column
        For i = 2 To lastrow
      
            'check if i reached last row in ticker group
            If i = lastrow Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                  
            'set the ticker name and last close price
            ticker = ws.Cells(i, 1).Value
            closingprice = ws.Cells(i, 6).Value
            
            'calculate the quarterly change for the ticker
            quarterlychange = closingprice - openingprice
            
             'calculate total stock volume for the ticker
             stocktotal = stocktotal + ws.Cells(i, 7).Value
             
             'calcuate the percentage change for the ticker
                 If openingprice <> 0 Then
                     percentchange = ((closingprice - openingprice) / openingprice)
                     
                 Else
                 percentchange = 0
                 
                 End If
                
            'print data in summary table
            ws.Range("I" & summary_table_row).Value = ticker
            ws.Range("J" & summary_table_row).Value = quarterlychange
            ws.Range("K" & summary_table_row).Value = percentchange
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            ws.Range("L" & summary_table_row).Value = stocktotal
         
                 'color formatting for quarterlychange
                 If quarterlychange > 0 Then
                     ws.Range("J" & summary_table_row).Interior.ColorIndex = 4 'green formating
                ElseIf quarterlychange < 0 Then
                     ws.Range("J" & summary_table_row).Interior.ColorIndex = 3 'red formatting
                Else
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = xlNone
                End If
                    
                'find greatest percent increase, decrease and total volume
                If percentchange > greatestincrease Then
                    greatestincrease = percentchange
                    increaseticker = ticker
                End If
                        
                If percentchange < greatestdecrease Then
                    greatestdecrease = percentchange
                    decreaseticker = ticker
                End If
                
                If stocktotal > greatesttotalvolume Then
                    greatesttotalvolume = stocktotal
                    totalvolumeticker = ticker
                End If
        
        'increment summary table
        summary_table_row = summary_table_row + 1
        
        'reset values for openingprice, quarterlychange, and total stock volume
        If i < lastrow Then
            openingprice = ws.Cells(i + 1, 3).Value
        quarterlychange = 0
        stocktotal = 0
        End If
        
        'if the cell immediately following a row is the same ticker...
        
        Else
        
            'if the cell immediately follwing a row is the same ticker
            'add to the total stock volume
            stocktotal = stocktotal + ws.Cells(i, 7).Value
        End If
        
        Next i
        
        'print results for the greatest increase, decrease, and total volume
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = increaseticker
        ws.Range("Q2").Value = greatestincrease
        ws.Range("Q2").NumberFormat = "0.00%"
        
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = decreaseticker
        ws.Range("Q3").Value = greatestdecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = totalvolumeticker
        ws.Range("Q4").Value = greatesttotalvolume
        
    Next ws
    
      
End Sub
