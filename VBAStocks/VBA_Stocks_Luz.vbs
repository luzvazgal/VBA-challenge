Option Explicit
'Module variables to store stock names having the greatest %decrease, %increase and total stock values
Dim greatestIncKey, greatestDecKey, maxTotalStockValKey As String
'Module variables to store values greatest %decrease, %increase and total stock values
Dim greatestInc, greatestDec, maxTotalStockVal As Double
Sub Main()
    
    'Grouping stocks and their values
    Call stockChanges
    
    'Printing greatest Increase, greatest Decrease and totalVolumeDictionary
    'In case the instructions refer to an only table for all sheets, the printChallenge sub should called from here
    'Call printChallenge
    
End Sub

Sub stockChanges()

'ITERATE ON DIFFERENT SHEETS
Dim sheetsNum, i As Integer

sheetsNum = Application.Sheets.Count

For i = 1 To sheetsNum

    'Make current Sheet active
    Worksheets(i).Activate
    
    'Variables Declaration
    Dim stockName As String
    Dim prevStockName As String
        
    'Number of different stocks we have in each sheet
    Dim numStocks As Integer
    Dim j, rowsNum As Long
    
    
    Dim yearlyPriceChange As Double
    Dim yearlyPercentChange As Double
    Dim totalStockVolume As Double
    
    'Variables initialization each time it enters to a new sheet
    stockName = ""
    prevStockName = ""
    numStocks = 0
    yearlyPriceChange = 0
    yearlyPercentChange = 0
    totalStockVolume = 0
               
    greatestIncKey = ""
    greatestDecKey = ""
    maxTotalStockValKey = ""
    greatestInc = 0
    greatestDec = 0
    maxTotalStockVal = 0
    
    'Print Results Tables' Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Value"
   
    
    'ITERATE ON CURRENT SHEET ROWS
    rowsNum = Cells(Rows.Count, 1).End(xlUp).Row
    
    For j = 2 To rowsNum
        
        'Local Variables
        Dim openPrice, closePrice, priceChange, percentChange As Double
        Dim stockVolume As Double
        
    
        stockName = Cells(j, 1).Value
        openPrice = Cells(j, 3).Value
        closePrice = Cells(j, 6).Value
        stockVolume = Cells(j, 7).Value
        
        priceChange = closePrice - openPrice
        
        'Validating openPrice is not zero
        If openPrice <> 0 Then
            percentChange = priceChange / openPrice
        Else
            percentChange = 0
        End If
        
        'The same stock values
        If stockName = prevStockName Then
            yearlyPriceChange = yearlyPriceChange + priceChange
            yearlyPercentChange = yearlyPercentChange + percentChange
            totalStockVolume = totalStockVolume + stockVolume
        
            'In case it's the last row of Sheet, and next value doesn't exist
            If j = rowsNum Then
                'Calling PrintResults to print last sheet's stock info starting on column 10
                printResults numStocks, stockName, yearlyPriceChange, yearlyPercentChange, totalStockVolume
                
                'Calling the sub to select the %Greatest Increase and Decrease as TotalVolume
                giveTheGreatest stockName, yearlyPercentChange, totalStockVolume
            End If
            
            
        'Validating if it's not the first row of first stock
        ElseIf prevStockName <> "" Then
        
            'Calling PrintResults to print previous stock info starting on column 10
            printResults numStocks, prevStockName, yearlyPriceChange, yearlyPercentChange, totalStockVolume
            
            'Calling the sub to select the %Greatest Increase and Decrease as TotalVolume
            giveTheGreatest prevStockName, yearlyPercentChange, totalStockVolume
                                             
            'Resetting variables
            yearlyPriceChange = priceChange
            yearlyPercentChange = percentChange
            totalStockVolume = stockVolume
            numStocks = numStocks + 1
                                   
        'The first row of first stock in Sheet
        Else
            'Assigning values for the first time
            yearlyPriceChange = priceChange
            yearlyPercentChange = percentChange
            totalStockVolume = stockVolume
            numStocks = 1
          End If
    
        prevStockName = stockName
    
    Next j
    
    printChallenge i
'MsgBox ("Termine hoja " + Str(i))
Next i


End Sub

'Printing results in current sheet
Sub printResults(ByVal numStocks As Integer, prevStockName As String, yearlyPriceChange As Double, yearlyPercentChange As Double, totalStockVolume As Double)
    'Print previous stock info starting on column 10
            Cells(numStocks + 1, 9).Value = prevStockName
            Cells(numStocks + 1, 10).Value = yearlyPriceChange
            Cells(numStocks + 1, 11).Value = FormatPercent(yearlyPercentChange, 2)
            Cells(numStocks + 1, 12).Value = totalStockVolume
            
         'Adding color
            If yearlyPriceChange < 0 Then
                Cells(numStocks + 1, 11).Interior.ColorIndex = 3
            Else
                Cells(numStocks + 1, 11).Interior.ColorIndex = 4
            End If
      
End Sub
'Adding to different dictionaries the greatest increase, decrease and totalvolume
Sub giveTheGreatest(ByRef prevStockName As String, yearlyPercentChange As Double, totalStockVolume As Double)
              
        'TO GET THE GREATEST INCREASE
        'If this is the first value
        If (greatestIncKey = "" And greatestInc = 0) Then
           greatestIncKey = prevStockName
           greatestInc = yearlyPercentChange
        
        'Validating if previous item is smaller than previous one
        'If true, it's removed and replaced by the new one; otherwise, nothing's done.
        ElseIf greatestInc < yearlyPercentChange Then
           greatestIncKey = prevStockName
           greatestInc = yearlyPercentChange
        End If
               
        
        'TO GET THE GREATEST DECREASE
        'If this is the first value
        If (greatestDecKey = "" And greatestDec = 0) Then
           greatestDecKey = prevStockName
           greatestDec = yearlyPercentChange
        
        'Validating if previous item is bigger than previous one
        'If true, it's removed and replaced by the new one; otherwise, nothing's done.
        ElseIf greatestDec > yearlyPercentChange Then
           greatestDecKey = prevStockName
           greatestDec = yearlyPercentChange
        End If
        
            
        'TO GET THE GREATEST TOTAL STOCK VOLUME
        'If this is the first value
        If (maxTotalStockValKey = "" And maxTotalStockVal = 0) Then
           maxTotalStockValKey = prevStockName
           maxTotalStockVal = totalStockVolume
        
        'Validating if previous item is smaller than previous one
        'If true, it's removed and replaced by the new one; otherwise, nothing's done.
        ElseIf maxTotalStockVal < totalStockVolume Then
           maxTotalStockValKey = prevStockName
           maxTotalStockVal = totalStockVolume
        End If

End Sub

'Printing greatest decrease and Increase percentages as well as highest Total Stock value
'In case it's only a table for all challenge, the arguments would be void instead the number of active Worksheet
Sub printChallenge(ByRef i As Integer)

MsgBox ("enter printChallenge")
    'Make first Sheet active
    Worksheets(i).Activate  'In case it's only a table per workbook, i=1
    
    'Printing Headers
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 16).Value = greatestIncKey
    Cells(2, 17).Value = FormatPercent(greatestInc, 2)
    
    
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 16).Value = greatestDecKey
    Cells(3, 17).Value = FormatPercent(greatestDec, 2)
    
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 16).Value = maxTotalStockValKey
    Cells(4, 17).Value = maxTotalStockVal
        
        
End Sub

