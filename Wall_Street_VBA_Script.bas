Attribute VB_Name = "Module1"
'option explicit

Sub Stock_Data()

'Loop Through Each Worksheet
'---------------------------------

    For Each ws In Worksheets
   
   'Add Heading for summary
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(1, "P").Value = "Ticker"
    ws.Cells(1, "Q").Value = "Value"
    
    'Determine Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Dim Open_Price As Double
    Dim Close_Price As Double
    
    'Define Open_Price
    Open_Price = ws.Cells(2, 3).Value
    
    Dim volume As Variant
    volume = 0
        
    'Loop Through Every Cell
    For i = 2 To lastrow
            
    'Declare Varaibles
     Dim Ticker_Name As String
     Dim Yearly_Change As Double
     Dim Percent_Change As Double
            
        'Calculate Total Volume
            volume = volume + ws.Cells(i, 7).Value
                 
        ' Check if we are still within the same Ticker Label, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                             
        'Set Ticker Label
             Ticker_Name = ws.Cells(i, 1).Value
                           
              'Set Close_Price
              Close_Price = ws.Cells(i, 6).Value
    
         'Calculate The Yearly_Change
                Yearly_Change = Close_Price - Open_Price
          
          'Calculate Percent_Change
            If Open_Price = 0 Then
                     Percent_Change = (Close_Price - Open_Price)
              Else
                    Percent_Change = (Close_Price - Open_Price) / Open_Price
              End If
           
           'Print To Summary_Table_Rows
                'Ticker Name
                 ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                      
           'Yearly_Change
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
           'Change color Index
            If (Yearly_Change < 0) Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
              Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If

            'Print Percent_Change To Summary_Table_Row
                 ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                 ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                 

            'Print volume  To Summary_Table_Row
                 ws.Range("L" & Summary_Table_Row).Value = volume
                 
                 'Reset Volume
                 volume = 0
                 
                 'Add 1 to Summary Table
                 Summary_Table_Row = Summary_Table_Row + 1
                 
                 'Reset Open Price
                 Open_Price = ws.Cells(i + 1, 3).Value
    
           End If

         Next i
            
            lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ' Set Greatest % Increase, % Decrease, and Total Volume
        
        'Identify Variables
        
        Dim Maximum_Percentage As Double
        Dim Minimum_Percentage As Double
        Dim Maximum_Total_Volume As Double
        
        'Calculate the Maximum_Percentage
        Maximum_Percentage = Application.WorksheetFunction.Max(ws.Range("K2: K" & lastrow))
        For j = 2 To lastrow
            If ws.Cells(j, 11) = Maximum_Percentage Then
                ws.Cells(2, 17).Value = Maximum_Percentage
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(2, 16).Value = ws.Cells(j, 9).Value
        Exit For
    End If
Next j
        
        
        'Calculate Minimum_Percentage
        Minimum_Percentage = Application.WorksheetFunction.Min(ws.Range("K2: K" & lastrow))
        For j = 2 To lastrow
            If ws.Cells(j, 11) = Minimum_Percentage Then
                ws.Cells(3, 17).Value = Minimum_Percentage
                ws.Cells(3, 17).NumberFormat = "0.00%"
                ws.Cells(3, 16).Value = ws.Cells(j, 9).Value
    
        Exit For
    End If
Next j
        
        
        'Calculate Maximum_Total_Volume
        Maximum_Total_Volume = Application.WorksheetFunction.Max(ws.Range("L2: L" & lastrow))
        For j = 2 To lastrow
            If ws.Cells(j, 12) = Maximum_Total_Volume Then
                ws.Cells(4, 17).Value = Maximum_Total_Volume
                ws.Cells(4, 16).Value = ws.Cells(j, 9).Value
        Exit For
    End If
Next j


 Next ws

End Sub





