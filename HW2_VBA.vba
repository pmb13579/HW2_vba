Sub stocks()
        
  'loop through worksheets in workbook

  For Each ws In Worksheets

    'Declare and initialize some worksheet level variables

    Dim worksheetName As String
    worksheetName = ws.Name

    MsgBox ("processing " & worksheetName)

    Dim numSym As Integer
    numSym = 0

    Dim totVol As Double
    totVol = 0

    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    Dim newSym As Boolean
    newSym = True

    Dim begPrice As Double

    Dim symInc, symDec, symVol As String
    Dim valInc, valDec, valVol As Double

    ' headings
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'loop through rows
    For r = 2 To lastRow

      'set current symbol
      Dim curSym As String
      curSym = ws.Cells(r, 1).Value

      'if starting new symbol
      If newSym Then
        numSym = numSym + 1
        newSym = False
        begPrice = ws.Cells(r, 3).Value
      End If

      'aAdd to total volume
      totVol = totVol + ws.Cells(r, 7).Value

      'if last entry for symbol
      If curSym <> ws.Cells(r + 1, 1).Value Then

        'write symbol, change, percent change, and total volume to columns 9, 10, 11, and 12

        ws.Cells(numSym + 1, 9).Value = curSym
        Dim change As Double
        change = ws.Cells(r, 6).Value - begPrice
        ws.Cells(numSym + 1, 10).Value = change
        Dim percentChange As Double
        If begPrice <> 0 Then
          percentChange = change / begPrice
        Else
          percentChange = 0
        End If
        ws.Cells(numSym + 1, 11).Value = Format(percentChange, "Percent")
        ws.Cells(numSym + 1, 12).Value = totVol

        'Color red(3) Or green(4)
        If change < 0 Then
          ws.Cells(numSym + 1, 10).Interior.ColorIndex = 3
        Else
          ws.Cells(numSym + 1, 10).Interior.ColorIndex = 4
        End If

        'iif first symbol on sheet
        If numSym = 1 Then
          symInc = curSym
          valInc = percentChange
          symDec = curSym
          valDec = percentChange
          symVol = curSym
          valVol = totVol
        Else
          If percentChange > valInc Then
            symInc = curSym
            valInc = percentChange
          End If
          If percentChange < valDec Then
            symDec = curSym
            valDec = percentChange
          End If
          If totVol > valVol Then
            symVol = curSym
            valVol = totVol
          End If

        End If

      'rReset two variables
      totVol = 0
      newSym = True

    End If

  Next r

  'post largest, smallest, volume with symbol
  ws.Cells(2, 15).Value = "Greatest % Increaser"
  ws.Cells(3, 15).Value = "Greatest % Decreaser"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"
  ws.Cells(2, 16).Value = symInc
  ws.Cells(3, 16).Value = symDec
  ws.Cells(4, 16).Value = symVol
  ws.Cells(2, 17).Value = Format(valInc, "Percent")
  ws.Cells(3, 17).Value = Format(valDec, "Percent")
  ws.Cells(4, 17).Value = valVol

  'Make sure columns wide enough to fit data
  ws.Columns("A:Q").AutoFit
   
  Next ws

End Sub


