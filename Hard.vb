Sub StocksHard()

' Set Worksheet variable

Dim ws As Worksheet

' Loop through each WorkSheet

For Each ws In ThisWorkbook.Sheets

    ws.Activate
    Debug.Print ws.Name

  ' WorkSheet Variables
        Dim Ticker As String
        Dim TotalStockVolume As Variant
        Dim Summary_Table_Row As Integer
        Dim LastRow As Long
        Dim Closing As Double
        Dim Opening As Double
        Dim PercentChange As Double
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalVolume As Double

  ' Keep track of the location for each Ticker
  
  Summary_Table_Row = 2
  
  ' Set WorkSheet Headers
  
        Cells(1,9).Value = "Ticker"
        Cells(1,10).Value = "Yearly Change"
        Cells(1,11).Value = "Percent Change"
        Cells(1,12).Value = "Total Stock Volume"
        Cells(2,15).Value = "Greatest % Increase"
        Cells(3,15).Value = "Greatest % Decrease"
        Cells(4,15).Value = "Greatest Total Volume"
        Cells(1,16).Value = "Ticker"
        Cells(1,17).Value = "Value"

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

                Cells(Summary_Table_Row, 11).Value = (Cells(Summary_Table_Row, 10).Value / Opening) * 100

            ElseIf Opening = 0 And Closing <> 0 Then

                Cells(Summary_Table_Row, 11).Value = FormatPercent(Cells(Summary_Table_Row, 10).Value)

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
    
  ' Loop to calculate Greatest % Increase, Greatest % Decrease & Greatest Total Volume

            For x = 2 To LastRow
                If ws.Cells(x,11).Value > ws.Cells(2,17).Value Then
                    ws.Cells(2,17).Value = ws.Cells(x,11).Value
                    ws.Cells(2,16).Value = ws.Cells(x,9).Value
                End If

                If ws.Cells(x,11).Value < ws.Cells(3,17).Value Then 
                    ws.Cells(3,17).Value = ws.Cells(x,11).Value
                    ws.Cells(3,16).Value = ws.Cells(x,9).Value
                End If

                If ws.Cells(x,12).Value > ws.Cells(4,17).Value Then 
                    ws.Cells(4,17).Value = ws.Cells(x,12).Value
                    ws.Cells(4,16).Value = ws.Cells(x,9).Value
                End If
                

            Next x


Next ws

End Sub
