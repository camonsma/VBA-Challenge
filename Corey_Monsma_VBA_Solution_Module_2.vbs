'sources https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures
'Source Class material
Sub Main()
' The main subroutine that does the work. I know this is very C++ and 1990s like me
   Call processWorkSheet
   Call getHighestLowest
End Sub

Sub processWorkSheet()

For Each ws In Worksheets

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Declair variables.
' code snipets were used by class notes and the Microsoft help site
' https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays
'End variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RowCount As LongLong
Dim slotCounter As LongLong
Dim CountStock As LongLong
Dim CounterOpen As LongLong

Dim closeAmount As Double
Dim openAmount As Double

Dim Volume As LongLong
Dim lastRow As LongLong
Dim TotalStocks As LongLong
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''program excution logic
' Set the headers on the columns for the output
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

 ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total value"

'Bold the titles and make them yellow to stand out
ws.Range("I1:L1").Font.Bold = True
ws.Range("I1:L1").Interior.ColorIndex = 6

'Assign initial values to variables.
slotCounter = 2 'The values start at two so we set the counter to 2 as the starting poistion
CountStock = 0
Volume = 0
TotalStocks = 0
CounterOpen = 0
openAmount = 0

'Like the census exercise in class get the amount of records in the dataset
lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Increment through the list and count the total volume individual stocks
For i = 2 To lastRow
  CounterOpen = CounterOpen + 1
'''''''sum and get the percentage differences'''''''''''''''''''''''''''''
' this is the record that starts the new set so we have the year close
 If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
   closeAmount = ws.Cells(i, 3).Value

'Assign the volume of the ticker sig
    ws.Cells(slotCounter, 9).Value = ws.Cells(i, 1).Value

'In case of one value we have to make sure we do not assign the volume to zero
'Sum the volume
     If (Volume = 0) Then
      Volume = ws.Cells(i, 7)
     Else
      Volume = Volume + ws.Cells(i, 7).Value
     End If ' only do this when there is one instance of the stock ticker
     
     ws.Cells(slotCounter, 12).NumberFormat = "000000000000"
     ws.Cells(slotCounter, 12).Value = Volume
     ws.Cells(slotCounter, 10).Value = (closeAmount - openAmount)
     
 'Assign the percent change
      If (openAmount = closeAmount) Then
       ws.Cells(slotCounter, 11).Value = 0
      ElseIf (openAmount <> 0) Then
       ws.Cells(slotCounter, 11).Value = ((closeAmount - openAmount) / openAmount)
      Else
' do nothing
     End If
     
     ws.Cells(slotCounter, 11).NumberFormat = "000.000%"
     
        If (ws.Cells(slotCounter, 10).Value) > 0 Then
' mark it green see, activity 03 student grade book
         ws.Cells(slotCounter, 10).Interior.ColorIndex = 4
        Else
' mark it red
         ws.Cells(slotCounter, 10).Interior.ColorIndex = 3
        End If

' reseting counters
    slotCounter = slotCounter + 1
    closeAmount = 0
    openAmount = 0
    Volume = 0 'reset the volume counter
    CounterOpen = 0 'reset the counter for the open value
  Else
' if this is the first instance record the opening value
    If (CounterOpen = 1) Then
     openAmount = ws.Cells(i, 3).Value
     End If
   Volume = Volume + ws.Cells(i, 7).Value
  End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Next i

Next ws
  
End Sub

Sub getHighestLowest()

For Each ws In Worksheets

Dim VTicker As String
Dim HTicker As String
Dim LTicker As String
Dim TotalStocks As LongLong

TotalStocks = 0

Dim LargestV As LongLong
LargestV = 0
VTicker = ""

Dim lPercent As Double
lPercent = 0
Dim hPercent As Double
hPercent = 0

VTicker = ""
LTicker = ""
HTicker = ""

TotalStocks = ws.Cells(Rows.Count, "L").End(xlUp).Row

LargestV = 0

'Go through each summed set and get the highest and lowest
For i = 2 To TotalStocks + 1
  If (i = 2) Then
   LargestV = ws.Cells(i, "L").Value
   VTicker = ws.Cells(i, "I").Value
   lPercent = ws.Cells(i, "K").Value
   LTicker = ws.Cells(i, "I").Value
   hPercent = ws.Cells(i, "K").Value
   HTicker = ws.Cells(i, "I").Value
  End If
    
  If (LargestV < ws.Cells(i + 1, "L").Value) Then
   LargestV = ws.Cells(i + 1, "L").Value
   VTicker = ws.Cells(i + 1, "I").Value
  End If
  If (lPercent > ws.Cells(i + 1, "K").Value) Then
   lPercent = ws.Cells(i + 1, "K").Value
   LTicker = ws.Cells(i + 1, "I").Value
  End If
  If (hPercent < ws.Cells(i + 1, "K").Value) Then
   hPercent = ws.Cells(i + 1, "K").Value
   HTicker = ws.Cells(i + 1, "I").Value
  End If
Next i

'write out the largest volue
ws.Range("O2").Cells.Value = "Greatest % Increase"
ws.Range("O3").Cells.Value = "Greatest % Decrease"
ws.Range("O4").Cells.Value = "Greatest tota Volume"

ws.Range("Q1").Cells.Value = "Volume"
ws.Range("P1").Cells.Value = "Ticker"
ws.Range("Q4").Cells.NumberFormat = "000000000000"
ws.Range("Q4").Cells.Value = LargestV
ws.Range("P4").Cells.Value = VTicker
ws.Range("Q3").Cells.NumberFormat = "000.000%"
ws.Range("Q3").Cells.Value = lPercent
ws.Range("P3").Cells.Value = LTicker
ws.Range("Q2").Cells.NumberFormat = "000.000%"
ws.Range("Q2").Cells.Value = hPercent
ws.Range("P2").Cells.Value = HTicker

ws.Columns("O:Q").AutoFit

Next ws

End Sub
