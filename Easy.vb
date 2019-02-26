Sub StocksEasy()

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

  ' Set an initial variable for holding the TotalStockVolume
  
  TotalStockVolume = 0

  ' Keep track of the location for each Ticker
  
  Summary_Table_Row = 2
  
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Total Stock Volume"

' Determine the LastRow
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through each Sheet
  For i = 2 To LastRow

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker 
      Ticker = Cells(i, 1).Value

      ' Add to the Total Stock Volume
      TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

      ' Print the Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Total Stock Volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = TotalStockVolume

      ' Add one to the Summary Table Row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Total Stock Volume
      TotalStockVolume = 0

    
    Else

      ' Add to the Total Stock Volume
      TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

    End If

  Next i

Next ws

End Sub





