Attribute VB_Name = "Module12"
Sub vbachallenge():

  
    
Dim ws As Worksheet

    
For Each ws In Worksheets
    

    
    
    
    Dim TickerName As String
    
    Dim YrlyChange As Double
    YrlyChange = 2
    
    Dim PercentChange As Double
    PercentChange = 2
    
    Dim TotalVol As Double
    
    Dim TickerOpen As Double
    
    Dim TickerClose As Double
    
    Dim TickerRows As Double
    TickerRows = 2
    
    Dim TickerRowOpen As Double
    TickerRowOpen = 2
    

    
    
    Dim MinPercent As Double
    Dim MaxPercent As Double
    Dim MaxVol As Double
    
    

    Dim MaxTickerName As String
    Dim MinTickerName As String
    Dim MaxVolName As String

    

    Dim TickerStart As Double
    TickerStart = 2


    Dim row As Long
    

    
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("A:Q").Columns.AutoFit
    
    
    
    
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    
    

        For row = 2 To lastrow
        
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then  'finding info from row above -> "Checking" for info on line above
                
                TotalVol = TotalVol + ws.Cells(row, 7).Value
            
                TickerName = ws.Cells(row, 1).Value
            
                ws.Cells(TickerRows, 12).Value = TotalVol
            
                ws.Cells(TickerRows, 9).Value = TickerName
                
                TickerRows = TickerRows + 1 'tells it to keep going
                
                TotalVol = 0
            
      
            
                If ws.Cells(TickerRowOpen, 3).Value = 0 Then
                
                
                    For TickerRowOpen = TickerRowOpen To lastrow
                    
                        If ws.Cells(TickerRowOpen, 3).Value <> 0 Then
                            TickerRowOpen = TickerRowOpen + 1
                            
                            Exit For
                
                     End If
                
                  
                   Next TickerRowOpen 'looping until we find our value..
                    
      
                    End If
                
                    TickerOpen = ws.Cells(TickerRowOpen, 3).Value 'grab value of first open
                    TickerClose = ws.Cells(row, 6).Value
                    
                    
                YrlyChange = TickerClose - TickerOpen
                PercentChange = YrlyChange / TickerOpen
                
                ws.Cells(TickerStart, 10).Value = YrlyChange
                ws.Cells(TickerStart, 11).Value = FormatPercent(PercentChange)
                    
                TickerRowOpen = row + 1
              
                TickerOpen = ws.Cells(row, 3).Value
                
                TickerClose = ws.Cells(row, 6).Value
                
       
                TickerStart = TickerStart + 1
                YrlyChange = 0
                PercentChange = 0
              
            
               
             Else
                
                 TotalVol = TotalVol + ws.Cells(row, 7).Value
           
        End If
        
      
    Next row

  
         
            

          
            Totallastrow = ws.Cells(Rows.Count, 10).End(xlUp).row
            
            
            ws.Cells(2, 17).Formula = "=max(K2:K" & Totallastrow & ")"
            ws.Cells(3, 17).Formula = "=min(K2:K" & Totallastrow & ")"
            ws.Cells(4, 17).Formula = "=max(L2:L" & Totallastrow & ")"
            
            ws.Range("K2:K" & Totallastrow).NumberFormat = "0.00%"
            
            
            
           
                                          
              
              
            ws.Cells(2, 16).Formula = "=INDEX(I2:I" & Totallastrow & ",MATCH(MAX(K2:K" & Totallastrow & "),K2:K" & Totallastrow & ",0))"
              
            ws.Cells(3, 16).Formula = "=INDEX(I2:I" & Totallastrow & ",MATCH(MIN(K2:K" & Totallastrow & "),K2:K" & Totallastrow & ",0))"
            ws.Cells(4, 16).Formula = "=INDEX(I2:I" & Totallastrow & ",MATCH(MAX(L2:L" & Totallastrow & "),L2:L" & Totallastrow & ",0))"
            
            
            
      
    
    
                                       ' FIND MAX/MIN CORRESPONDING VALUES SOURCE:
                                       '   https://stackoverflow.com/questions/53261956/how-to-find-get-the-variable-name-having-largest-value-in-excel-vba/53262316#53262316

              
        
    For row = 2 To Totallastrow
        
        If ws.Cells(row, 10).Value > 0 Then
            ws.Cells(row, 10).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(row, 10).Value < 0 Then
            ws.Cells(row, 10).Interior.ColorIndex = 3
                
        ElseIf ws.Cells(row, 10).Value = 0 Then
            ws.Cells(row, 10).Interior.ColorIndex = 6
            
             
        End If
        
    Next row
    
 Next ws
 

End Sub



