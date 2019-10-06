Sub StockMarket()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
    For Each ws In Worksheets


        ' Initialize Variables
        Dim TSymbol, PrevSymbol As String
        Dim i, j, k, LastRow, LastColumn, GISymbol, GDSymbol, GTVSymbol As Integer
        Dim TotalVolume As LongLong
        Dim OpenPrice, ClosePrice, YearlyChange, PercentChange, GPIncrease, GPDecrease, GTVolume As Double
        Dim FirstInd As Boolean
        
                       
        'Sort the data by Ticker Symbol and Date in ascending order
        Columns("A:G").Sort key1:=Range("A2"), order1:=xlAscending, key2:=Range("B2"), order2:=xlAscending, Header:=xlYes
       
        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Determine the Last Column Number
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        k = 1
        
        ' --------------------------------------------
        ' Column Header
        ' --------------------------------------------
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' --------------------------------------------
        ' Initialize
        ' --------------------------------------------
        
        PrevSymbol = ""
        FirstInd = True
        YearlyChange = 0
        PercentChange = 0
        GPIncrease = 0
        GPDecrease = 0
        GTVolume = 0
        GISymbol = 0
        GDSymbol = 0
        GTVSymbol = 0


        For i = 2 To LastRow + 1
            j = 1
                
            TSymbol = ws.Cells(i, j).Value
             
             
           'This section handles the change to a new symbol and
           'Stores/Write the Yearly Change, Percent Change and Total Stock Volume by Ticker Symbol
           
            If PrevSymbol <> TSymbol And FirstInd = False Then            ' Compares to check if symbol changed
                
                
                k = k + 1
                
                'Summary by Ticker Symbol
                'Write Yearly Change, Percent Change and Total Stock Volume
                
                YearlyChange = ClosePrice - OpenPrice
                
                'Zero out if Ope Price=0. Undefined since we cannot divide by 0
                If OpenPrice = 0 Then
                    PercentChange = 0
                Else
                    PercentChange = YearlyChange / OpenPrice
                End If
                
               'Write TickerSymbol, Yearly Change, Percent Change and Total Stock Volume
                ws.Cells(k, LastColumn + 2).Value = PrevSymbol
                ws.Cells(k, LastColumn + 3).Value = YearlyChange
                ws.Cells(k, LastColumn + 3).NumberFormat = "0.000000000"
                ws.Cells(k, LastColumn + 4).Value = Format(PercentChange, "0.00%")
                ws.Cells(k, LastColumn + 5).Value = TotalVolume
                
                If YearlyChange >= 0 Then
                    ws.Cells(k, LastColumn + 3).Interior.ColorIndex = 4
                Else
                    ws.Cells(k, LastColumn + 3).Interior.ColorIndex = 3
                End If
                
                TotalVolume = 0
                FirstInd = True
                               
            End If
            
            'This only stores the OpenPrice for the very first record in the table and for each change in symbol
            If FirstInd Then
               OpenPrice = ws.Cells(i, j + 2).Value
               FirstInd = False
            End If
               
            ClosePrice = ws.Cells(i, j + 5).Value
            TotalVolume = TotalVolume + ws.Cells(i, j + 6).Value
            PrevSymbol = TSymbol
        
        Next i
             
          
       ' Include Summary Statistics
       ' Search for the Highest and Lowest % Using Match function


        Dim LookupRangeP, LookupRangeV As Range
   
        Set LookupRangeP = ws.Range("K2:K" & k)
        GPIncrease = WorksheetFunction.Max(LookupRangeP)
        GPDecrease = WorksheetFunction.Min(LookupRangeP)
        
        Set LookupRangeV = ws.Range("L2:L" & k)
        GTVolume = WorksheetFunction.Max(LookupRangeV)
        

        GISymbol = (WorksheetFunction.Match(GPIncrease, LookupRangeP, 0)) + 1
        GDSymbol = (WorksheetFunction.Match(GPDecrease, LookupRangeP, 0)) + 1
        GTVSymbol = (WorksheetFunction.Match(GTVolume, LookupRangeV, 0)) + 1
        
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = ws.Range("I" & GISymbol)
        ws.Range("Q2").Value = Format(GPIncrease, "0.00%")
        
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = ws.Range("I" & GDSymbol)
        ws.Range("Q3").Value = Format(GPDecrease, "0.00%")

    
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = ws.Range("I" & GTVSymbol)
        ws.Range("Q4").Value = GTVolume
        
        ws.Columns("A:Q").AutoFit
        
    Next ws
    
    MsgBox ("Stock Market Summary Completed.")

End Sub
