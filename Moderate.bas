Attribute VB_Name = "Moderate1"
Sub Moderate()

    counter = 1
    
    countopen = 0
    
    volume = 0
    
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    Range("H1").Value = "Ticker"
        
    Range("I1").Value = "Volume"
        
    Range("J1").Value = ("Yearly Change")
    
    Range("K1").Value = "% Change of Yearly Change"
        
    
           
    For i = 2 To lrow
    
       temp = Range("A" & i + 1).Value
        
       tickers = Range("A" & i).Value
    
        If tickers <> temp Then
            
            opener = Range("C" & i - countopen).Value
            
            closer = Range("f" & i).Value
            
            counter = counter + 1
            
            ticker = Range("A" & i).Value
            
            volume = volume + Range("G" & i).Value
            
            Range("H" & counter).Value = ticker
            
            Range("I" & counter).Value = volume
            
            'Range("n" & counter).Value = closer
            
            'Range("m" & counter).Value = opener
            
            OPCL = closer - opener
                        
            Range("J" & counter).Value = OPCL
            
            If opener <> 0 Then
            
                Percent = (OPCL / opener)
            
                Range("K" & counter).NumberFormat = "0.00%"
            
                Range("K" & counter).Value = Percent
            
            Else
                
                Range("K" & counter).Value = 0
                
            End If
                
            volume = 0
            
            countopen = 0
            
            If OPCL >= 0 Then
                
                Range("J" & counter).Interior.ColorIndex = 4
            Else
            
                Range("J" & counter).Interior.ColorIndex = 3
                
            End If
            
        Else
            
            volume = volume + Range("G" & i).Value
            
            countopen = countopen + 1
            
        End If
        
      Next i

End Sub


