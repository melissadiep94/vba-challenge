# vba-challenge


Sub Multi_Yr_Stock_Data()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

    'Set headers
        Cells(1, 9) = "Ticker"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
    
    'Set variables
    
        Dim ticker As String
        Dim i As Long
        Dim Total_sv As Double
        Total_sv = 0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Open_Price = 0
        Close_Price = 0
        Dim LastRow As Long
        
    'To determine last row of each WS
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set Initial Open Price
            Open_Price = Cells(2, 3).Value
    
         
    'Set loop to create summary table
    
    For i = 2 To LastRow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set ticker price
            ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = ticker
            
            
            ' Set Close Price
            Close_Price = Cells(i, 6).Value
            ' Add Yearly Change
            Yearly_Change = Close_Price - Open_Price
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Red if Yearly change is < 0; Green if Yearly Change > 0;
                    If Yearly_Change < 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    
                    ElseIf Yearly_Change > 0 Then
                        Range("J" & Summary_Table_Row).Interior.ColorIndex = 10
                    End If
    
            ' Add Percent Change
                    If (Open_Price = 0 And Close_Price = 0) Then
                        Percent_Change = 0
                    ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                        Percent_Change = 1
                    Else
                        Percent_Change = Yearly_Change / Open_Price
                        Range("K" & Summary_Table_Row).Value = Percent_Change
                        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                    
                    ' reset the Open Price ; add 1 row
                       Open_Price = Cells(i + 1, 3).Value
                           
                    End If
        
            'Set Stock Volume
            Total_sv = Total_sv + Cells(i, 7).Value
            Range("L" & Summary_Table_Row).Value = Total_sv
            Range("L" & Summary_Table_Row).NumberFormat = "#,##0"
           'reset the volume
             Total_sv = 0
             
             
            'Add one to the summary table row
             
             Summary_Table_Row = Summary_Table_Row + 1
             
        
        Else
            
          'Add to the Total
          
          Total_sv = Total_sv + Cells(i, 7).Value
          
        End If
    
        
    Next i

    
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest Total Stock Volume"
    Range("O1").Value = "Value"
    Range("N2:N4,O1").Font.Bold = True
    
    Range("O2").Value = WorksheetFunction.Max(Range("K:K"))
    Range("O2").NumberFormat = "0.00%"
    
    Range("O3").Value = WorksheetFunction.Min(Range("K:K"))
    Range("O3").NumberFormat = "0.00%"
    
    Range("O4").Value = WorksheetFunction.Max(Range("L:L"))
    Range("O4").NumberFormat = "#,##0"
    
    
Next WS

End Sub

