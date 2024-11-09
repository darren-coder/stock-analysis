Attribute VB_Name = "Module1"
Sub Stock()
    
    Dim Ticker As String
    Dim QC As String
    Dim PC As String
    Dim TSV As String
    
    Ticker = "Ticker"
    QC = "Quarterly Change"
    PC = "Percent Change"
    TSV = "Total Stock Volume"
    
    Dim GI As String
    Dim GD As String
    Dim GTV As String
    Dim V As String
    
    GI = "Greatest % Increase"
    GD = "Greatest % Decrease"
    GTV = "Greatest Total Volume"
    V = "Value"
    
    Dim Symbol As String
    Dim LastRow As Long
    Dim Stock_Volume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Summary_Table_Row As Integer
    Dim Percent_Change As Double
    Dim WS_Count As Integer
    Dim i As Long
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For Each WS In ThisWorkbook.Worksheets
        
        Stock_Volume = 0
        Open_Price = WS.Range("C2").Value
        Summary_Table_Row = 2
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        WS.Range("I1").Value = Ticker
        WS.Range("J1").Value = QC
        WS.Range("K1").Value = PC
        WS.Range("L1").Value = TSV
        WS.Columns("I:L").AutoFit
        WS.Columns("K").NumberFormat = "0.00%"
        WS.Columns("J").NumberFormat = "0.00"
        WS.Range("O2").Value = GI
        WS.Range("O3").Value = GD
        WS.Range("O4").Value = GTV
        WS.Columns("O").Columns.AutoFit
        WS.Range("P1").Value = Ticker
        WS.Range("Q1").Value = V
        
        For i = 2 To LastRow
            
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1) Then
            
                Symbol = WS.Cells(i, 1).Value
            
                WS.Cells(Summary_Table_Row, 9).Value = Symbol
            
                Close_Price = WS.Cells(i, 6).Value
                
                Quarterly_Change = Close_Price - Open_Price
                WS.Cells(Summary_Table_Row, 10).Value = Quarterly_Change
                    
                    If Quarterly_Change > 0 Then
                        WS.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf Quarterly_Change < 0 Then
                        WS.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0)
                    End If
                    
                    If Open_Price <> 0 Then
                        Percent_Change = Quarterly_Change / Open_Price
                    Else
                        Percent_Change = 0
                    End If
                    
                        If Percent_Change > 0 Then
                            WS.Cells(Summary_Table_Row, 11).Interior.Color = RGB(0, 255, 0)
                        ElseIf Percent_Change < 0 Then
                            WS.Cells(Summary_Table_Row, 11).Interior.Color = RGB(255, 0, 0)
                        End If
                    
                        WS.Cells(Summary_Table_Row, 11).Value = Percent_Change
                
                Stock_Volume = Stock_Volume + WS.Cells(i, 7)
               
                WS.Cells(Summary_Table_Row, 12).Value = Stock_Volume
                
                Open_Price = WS.Cells(i + 1, 3).Value
               
                Quarterly_Change = 0
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Close_Price = 0
                
            Else: WS.Cells(i + 1, 1).Value = WS.Cells(i, 1).Value
                Stock_Volume = Stock_Volume + WS.Cells(i, 7).Value
                
            End If
            
        Next i
        
            Dim Symbol1 As String
            Dim Symbol2 As String
            Dim Symbol3 As String
            Dim Increase As Double
            Dim Decrease As Double
            Dim Vol As Double
            Dim g As Long
            Increase = -1
            Decrease = 1
            Vol = -1
            
            For g = 2 To LastRow
            
                If WS.Cells(g, 11).Value > Increase Then
                    Increase = WS.Cells(g, 11).Value
                    Symbol1 = WS.Cells(g, 9).Value
                End If
                
                If WS.Cells(g, 11).Value < Decrease Then
                    Decrease = WS.Cells(g, 11).Value
                    Symbol2 = WS.Cells(g, 9).Value
                End If
                
                If WS.Cells(g, 12).Value > Vol Then
                    Vol = WS.Cells(g, 12).Value
                    Symbol3 = WS.Cells(g, 9).Value
                End If
                
            Next g
            
        WS.Range("Q2").Value = Increase
        WS.Range("P2").Value = Symbol1
        WS.Range("Q3").Value = Decrease
        WS.Range("P3").Value = Symbol2
        WS.Range("Q4").Value = Vol
        WS.Range("P4").Value = Symbol3
                
        WS.Range("Q4").EntireColumn.AutoFit
        WS.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next WS
              
End Sub



