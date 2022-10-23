VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub module2()

For Each ws In Worksheets
            ' Set an initial variable for holding the brand name
                                          Dim Ticker As String
                                        
                                        ' Set yearly change as double
                                          Dim yearly_change As Double
                                          yearly_change = 0
                                          
                                          Dim percentage_change As Double
                                          percentage_change = 0
                                          
                                          ' Set total stock volum
                                          Dim total_stock_volume As Double
                                          total_stock_volume = 0
                                        
                                          ' Keep track of the location for ticker name in the summary table
                                          Dim Summary_Table_Row As Integer
                                          Summary_Table_Row = 2
                                          
                                          Dim row_count As Integer
                                          row_count = 0
                                          
                                        ws.Cells(1, 9).Value = "Ticker"
                                        ws.Cells(1, 10).Value = "Yearly Change"
                                        ws.Cells(1, 11).Value = "Percent Change"
                                        ws.Cells(1, 12).Value = "Total Stock Volume"
                                        ws.Cells(1, 16).Value = "Ticker"
                                        ws.Cells(1, 17).Value = "Value"
                                        ws.Cells(2, 15).Value = "Greatest % increase"
                                        ws.Cells(3, 15).Value = "Greatest % decrease"
                                        ws.Cells(4, 15).Value = "Greatest total volume"
                                        
                                         ' Determine the Last Row
                                        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                                    
                                          ' Loop through
                                            For j = 2 To Lastrow
                                        
                                            ' if we are not in the same ticker
                                                If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                                        
                                              ' Set the ticker name
                                                Ticker = ws.Cells(j, 1).Value
                                              
                                              ' Set the yearly change name
                                                yearly_change = ws.Cells(j, 6).Value - ws.Cells((j - row_count), 3).Value
                                                percent_change = (yearly_change / ws.Cells((j - row_count), 3).Value)
                                              
                                              ' Add to the stock Total
                                                total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
                                              
                                              ' Print the ticker name in the Summary Table
                                                ws.Range("I" & Summary_Table_Row).Value = Ticker
                                              
                                              ' Print the yearly change in the Summary Table
                                                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                                                              
                                              ' conditional formatting to show color for positive and negative values
                                              
                                                                    If yearly_change < 0 Then
                                                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                                                                    Else
                                                                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                                                                    End If
                                            
                                              ' Print the percent change to the Summary Table
                                                ws.Range("k" & Summary_Table_Row).Value = percent_change
                                                ws.Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
                                              
                                              ' Print the Brand Amount to the Summary Table
                                                ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
                                        
                                              ' Add one to the summary table row
                                                Summary_Table_Row = Summary_Table_Row + 1
                                              
                                              ' Reset the stock Total
                                                total_stock_volume = 0
                                                yearly_change = 0
                                                row_count = 0
                                                'If the cell immediately following a row is the same brand...
                                                Else
                                        
                                                total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
                                            
                                                row_count = row_count + 1
                                            
                                                End If
                                         
                                            Next j
                                            
                                            ' to set greatest increase
                                             Dim greatest As Double
                                             Dim great_ticker As String
                                             greatest = 0
                                             
                                             For a = 2 To Lastrow
                                            
                                                If greatest >= ws.Cells(a, 11).Value Then
                                                         
                                                    greatest = greatest
                                                
                                                ElseIf ws.Cells(a, 11).Value > greatest Then
                                                     
                                                    greatest = ws.Cells(a, 11).Value
                                                    great_ticker = ws.Cells(a, 9).Value
                                                End If
                                                     ws.Cells(2, 17).Value = greatest
                                                     
                                                     ws.Cells(2, 16).Value = great_ticker
                                                
                                             Next a
                                             
                                                ' to set greatest decrease
                            Dim lowest As Double
                            Dim low_ticker As String
                            lowest = 0
                            
                            For b = 2 To Lastrow
                           
                                    If lowest <= ws.Cells(b, 11).Value Then
                                        
                                     lowest = lowest
                                    ElseIf lowest > ws.Cells(b, 11).Value Then
                                    
                                     lowest = ws.Cells(b, 11).Value
                                     low_ticker = ws.Cells(b, 9).Value
                                    End If
                                    
                                    ws.Cells(3, 17).Value = lowest
                                    
                                    ws.Cells(3, 16).Value = low_ticker
                                                
                            Next b
                            
                                              ' to set greatest stock
                            Dim high_stock As Double
                            Dim stock_name As String
                            high_stock = 0
                            
                            For c = 2 To Lastrow
                           
                                    If high_stock >= ws.Cells(c, 12).Value Then
                                        
                                     high_stock = high_stock
                                    ElseIf high_stock < ws.Cells(c, 12).Value Then
                                    
                                     high_stock = ws.Cells(c, 12).Value
                                     stock_name = ws.Cells(c, 9).Value
                                    End If
                                    
                                    ws.Cells(4, 17).Value = high_stock
                                    
                                    ws.Cells(4, 16).Value = stock_name
                                                
                            Next c
        
                     
Next ws
              
End Sub
