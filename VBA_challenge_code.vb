Option Explicit


Public Sub stockOutputs()

Dim ws As Worksheet

Dim i, j, x As Integer

Dim percentchange, yearlychange As Double
Dim openprice, closeprice As Double
Dim volume, greatest_volume As LongLong

Dim lastrow, tickerrow, yearchangerow, percentchangerow, volumechangerow As Integer

Dim greatest_increase, greatest_decrease As Double
Dim increase_ticker, decrease_ticker, volume_ticker As String

For Each ws In Worksheets

'these variable are used to set the next available row foreach of the summary columns
lastrow = Range("A2").End(xlDown).Row
tickerrow = Range("I2").End(xlDown).Count
yearchangerow = Range("J2").End(xlDown).Count
percentchangerow = Range("K2").End(xlDown).Count
volumechangerow = Range("L2").End(xlDown).Count
  
  'this sets the Column titles for each ws
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
   
 'this format paints the column titles for consistency across workbook
  Worksheets("2018").Columns("I:L").Copy
  ws.Columns("I:L").PasteSpecial Paste:=xlPasteFormats
  Application.CutCopyMode = False
   
   For i = 2 To lastrow
       
        volume = volume + Cells(i, 7).Value
    
        If ws.Cells(i, 1).Offset(1).Value <> ws.Cells(i, 1).Value Then
            
            'This generates the each ticker in column I.
            ws.Range("I" & tickerrow + 1).Value = ws.Cells(i, 1).Value
            tickerrow = tickerrow + 1
                        
            'This generates the starting price.
            openprice = ws.Cells(i - 250, 3).Value

            'this generates the closing price
            closeprice = ws.Cells(i, 6).Value
                                  
            'this calculates the yearly change and prints it to column J.
            yearlychange = openprice - closeprice
                    
                    '..and applies conditional formatting to the yearly change column
                    If yearlychange > 0 Then
                        ws.Range("J" & yearchangerow + 1).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & yearchangerow + 1).Interior.ColorIndex = 3
                    End If
                    
            ws.Range("J" & yearchangerow + 1).Value = yearlychange
            yearchangerow = yearchangerow + 1
            
            'this calculates the percentage change and prints it to column K
            percentchange = yearlychange / openprice
            
                '..and applies conditional formatting to the percent change column
                If percentchange > 0 Then
                    ws.Range("K" & percentchangerow + 1).Interior.ColorIndex = 4
                Else
                    ws.Range("K" & percentchangerow + 1).Interior.ColorIndex = 3
                End If
            
            ws.Range("K" & percentchangerow + 1).Value = percentchange
            percentchangerow = percentchangerow + 1
            
            'this generates and prints the trade volume per stock.
            ws.Range("L" & volumechangerow + 1).Value = volume
            volumechangerow = volumechangerow + 1
            volume = 0
                  
        End If
                   
Next i

ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("R3:R4").NumberFormat = "0.00%"

greatest_increase = -1
greatest_decrease = 0
greatest_volume = 0

'This loops through the stock summaries to pick out highest increases, decreases and volume
For x = 2 To percentchangerow

    If ws.Cells(x, 11).Value > greatest_increase Then
        greatest_increase = ws.Cells(x, 11).Value
        increase_ticker = ws.Cells(x, 9).Value
    End If
    If ws.Cells(x, 11).Value < greatest_decrease Then
        greatest_decrease = ws.Cells(x, 11).Value
        decrease_ticker = ws.Cells(x, 9).Value
    End If
    If ws.Cells(x, 12).Value > greatest_volume Then
        greatest_volume = ws.Cells(x, 12).Value
        volume_ticker = ws.Cells(x, 9).Value
    End If
Next x
'This constructs and formats the Bonus summary section
ws.Range("P3").Value = "Greatest % Increase"
ws.Range("Q3").Value = increase_ticker
ws.Range("R3").Value = greatest_increase
ws.Range("P4").Value = "Greatest % Decrease"
ws.Range("Q4").Value = decrease_ticker
ws.Range("R4").Value = greatest_decrease
ws.Range("P5").Value = "Greatest Total Volume"
ws.Range("Q5").Value = volume_ticker
ws.Range("R5").Value = greatest_volume
Worksheets("2018").Columns("M:M").Copy
ws.Columns("P:P").PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

Next ws

End Sub

Sub clear()
'this erases content and formatting from summary area when button is clicked
Dim wsheet As Worksheet
For Each wsheet In Worksheets
wsheet.Range("I:L").ClearContents
wsheet.Range("P3:R5").ClearContents


wsheet.Columns("M:M").Copy
wsheet.Range("I:L").PasteSpecial Paste:=xlPasteFormats
Application.CutCopyMode = False

Next wsheet

End Sub