Attribute VB_Name = "Module1"
Sub StockParse():
    
    'Look through all sheets
    For Each ws In Worksheets
    
    'Keeping track of summary table row instead of the row the data is originally from
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'Initializing some variables
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim YearClose As Double
    Dim YearOpen As Double
    
        'Find length of worksheet we are pulling data from
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Total_Volume = 0
        
        'Add new headers to the worksheets
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Set initial yearOpen
        YearOpen = Cells(2, 3).Value
        
            'Set loop length in each sheet
            For i = 2 To LastRow
            
            Total_Volume = Total_Volume + Cells(i, 7).Value
            
                If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                
                'Return ticker
                Cells(Summary_Table_Row, 9).Value = Cells(i, 1).Value
                
                'Return Total Stock Volume
                Cells(Summary_Table_Row, 12).Value = Total_Volume
                
                'Find yearClose
                YearClose = Cells(i, 6).Value
                
                'Find yearly change, open value on first day - close value on last
                Yearly_Change = YearClose - YearOpen
                
                'Set yearly change into new table
                Cells(Summary_Table_Row, 10).Value = Yearly_Change
                    
                    'Color code yearly change. Green if positive, red if negative
                    If Yearly_Change > 0 Then
                    Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
                    Else: Cells(Summary_Table_Row, 10).Interior.Color = vbRed
                    End If
                    
                    'Percent Change if statement, in case yearOpen is 0
                    If YearOpen <> 0 Then
                    'Set percent change
                    Percent_Change = (YearClose - YearOpen) / YearOpen
                
                    'Print percent change to summary table & format to percent
                    
                    Cells(Summary_Table_Row, 11).Value = Percent_Change
                    Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                    
                    Else: Cells(Summary_Table_Row, 11).Value = "N/A"
                    End If
                
                'Set yearOpen to new value
                YearOpen = Cells(i + 1, 3).Value
                 
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                End If
            Next i
    
    Next ws

End Sub

