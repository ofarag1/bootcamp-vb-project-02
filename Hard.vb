Sub StocksModerate()

' Set Worksheet variable

Dim ws As Worksheet

' Loop through each WorkSheet

For Each ws In ThisWorkbook.Sheets

    ws.Activate
    MsgBox (ws.Name)

  ' WorkSheet Variables
  Dim Ticker As String
  Dim TotalStockVolume As Variant
  Dim Summary_Table_Row As Integer
  Dim LastRow As Long
  Dim Closing As Double
  Dim Opening As Double
  Dim max As Long
  Dim min As Long
  Dim GTV As Variant

  ' Keep track of the location for each Ticker
  
  Summary_Table_Row = 2
  
  ' Set WorkSheet Headers
  
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Yearly Change"
  Cells(1, 11).Value = "Percentage Change"
  Cells(1, 12).Value = "Total Stock Volume"
  Cells(2, 14).Value = "Greatest Inc"
  Cells(3, 14).Value = "Greatest Dec"
  Cells(4, 14).Value = "Greatest Total Volume"
  Cells(1, 15).Value = "Ticker"
  Cells(1, 16).Value = "Value"

' Determine the LastRow
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through each Row in the Sheet 
  For i = 2 To LastRow
    If i = 2 then
		Opening = Cells(i,3).Value
	End if	  
    
     ' Add to the Total Stock Volume
     
      TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
           
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Closing = Cells(i, 6).Value
        
        Cells(Summary_Table_Row, 9).Value = Cells(i, 1).Value
	
	'Yearly Change, Subtract Closing - Opening
        
        Cells(Summary_Table_Row, 10).Value = (Closing - Opening) 
      
      'Calculate Percentage Change
        
            If Opening > 0 Then

                Cells(Summary_Table_Row, 11).Value = (FormatPercent(Cells(Summary_Table_Row, 10).Value / Opening))

            ElseIf Opening = 0 And Closing <> 0 Then

                Cells(Summary_Table_Row, 11).Value = (FormatPercent(Cells(Summary_Table_Row, 10).Value))

            Else
                Cells(Summary_Table_Row, 11).Value = 0

            End If
	    
	        Cells(Summary_Table_Row, 12).Value = TotalStockVolume

      	    'Highlight positive change in green and negative change in red
          
            If Cells(Summary_Table_Row, 10).Value < 0 Then
	    
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3

            ElseIf Cells(Summary_Table_Row, 10).Value > 0 Then
	    
                Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
            End If

    
	        TotalStockVolume = 0
	
	        Opening = Cells(i + 1,3).Value
    
	        Summary_Table_Row = Summary_Table_Row + 1
    
    End If
  
  Next i

    For x = 2 To LastRow
    
        If Cells(x, 11).Value > max Then
       
        max = Cells(x, 11).Value
        ticker = Cells(x, 1).Value
    
        End If

    Next x

        Cells(2, 16).Value = max
        Cells(2, 15).Value = ticker

    For y = 2 To LastRow
    
        If Cells(y, 3).Value < min Then
       
        min = Cells(y, 11).Value
        ticker = Cells(i, 1).Value
    
        End If

    Next y

        Cells(3, 16).Value = min
        Cells(3, 15).Value = ticke
  
  For w = 2 To LastRow
    
        If Cells(w, 4).Value > GTV Then
       
        GTV = Cells(w, 4).Value
        ticker = Cells(w, 1).Value
    
        End If

  Next w

        Cells(4, 16).Value = GTV
        Cells(4, 15).Value = ticker

Next ws

End Sub





