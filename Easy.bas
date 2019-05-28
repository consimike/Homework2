Attribute VB_Name = "Easy"
Sub easyFinal()

    counter = 1
    
    volume = 0
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("H1").Value = "Ticker"
        
    Range("I1").Value = "Volume"
        
    Range("J1").Value = "% change open/close"
        
    For i = 2 To lrow
    
       temp = Range("A" & i + 1).Value
        
       tickers = Range("A" & i).Value
    
        If tickers <> temp Then
            
            counter = counter + 1
            
            ticker = Range("A" & i).Value
            
            volume = volume + Range("G" & i).Value
            
            Range("H" & counter).Value = ticker
            
            Range("I" & counter).Value = volume
            
            volume = 0
            
        Else
            
            volume = volume + Range("G" & i).Value
        
        End If
        
      Next i

End Sub

