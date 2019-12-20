Sub stock_analysis()

Dim currenttickerval As String
Dim nexttickerval As String
Dim summaryrow As Long
Dim counter As Long
Dim start As Integer
Dim ending As Long
Dim openingprice As Double
Dim endingprice As Double
Dim totalsheets As Integer
Dim sheetcounter As Long
Dim adjustedtotalvolume As Double
Dim yearlypricechange As Double
Dim adjustedpercentchange As Double
Dim percentincreaseposition As Long
Dim percentdecreaseposition As Long
Dim greatestvolumeposition As Long

totalsheets = Worksheets.Count

For sheetcounter = 1 To totalsheets

Worksheets(sheetcounter).Activate

ending = ActiveSheet.UsedRange.Rows.Count
start = 2

With ActiveSheet.Sort
.SortFields.Add Key:=Range("A1"), Order:=xlAscending
.SortFields.Add Key:=Range("B1"), Order:=xlAscending
.SetRange Range("A1:G" & ending)
.Header = xlYes
.Apply
End With

openingprice = Cells(start, 3).Value
adjustedtotalvolume = 0
summaryrow = 1

Cells(summaryrow, 9).Value = "Ticker"
Cells(summaryrow, 10).Value = "Yearly Change"
Cells(summaryrow, 11).Value = "Percent Change"
Cells(summaryrow, 12).Value = "Total Stock Volume"

Cells(summaryrow, 15).Value = "Ticker"
Cells(summaryrow, 16).Value = "Value"
Cells(summaryrow + 1, 14).Value = "Greatest % Increase"
Cells(summaryrow + 2, 14).Value = "Greatest % Decrease"
Cells(summaryrow + 3, 14).Value = "Greatest Total Volume"

For counter = start To ending

currenttickerval = Cells(counter, 1).Value
nexttickerval = Cells(counter + 1, 1).Value

    If currenttickerval = nexttickerval Then
    adjustedtotalvolume = adjustedtotalvolume + Cells(counter, 7).Value / 100
    
    Else
    
    endingprice = Cells(counter, 6).Value
    yearlypricechange = endingprice - openingprice
    adjustedtotalvolume = adjustedtotalvolume + Cells(counter, 7).Value / 100
    summaryrow = summaryrow + 1
    Cells(summaryrow, 9).Value = currenttickerval
    Cells(summaryrow, 10).Value = yearlypricechange
    
        If yearlypricechange > 0 Then
            Cells(summaryrow, 10).Interior.ColorIndex = 4
            ElseIf yearlypricechange < 0 Then
                Cells(summaryrow, 10).Interior.ColorIndex = 3
            End If
        
    If openingprice <> 0 Then
        Cells(summaryrow, 11).Value = Cells(summaryrow, 10).Value / openingprice
        Else: Cells(summaryrow, 11).Value = 0
        End If
    
    Cells(summaryrow, 11).Value = Cells(summaryrow, 11).Value
    Cells(summaryrow, 11).NumberFormat = "0.00%"
    Cells(summaryrow, 12).Value = adjustedtotalvolume
    openingprice = Cells(counter + 1, 3).Value
    adjustedtotalvolume = 0
    
    End If
    
Next counter

For counter = 2 To summaryrow

Cells(counter, 12).Value = Cells(counter, 12).Value * 100

Next counter

percentdecreaseposition = 2
greatestvolumeposition = 2
percentincreaseposition = 2

For counter = 2 To summaryrow

If Cells(counter, 11) > Cells(percentincreaseposition, 11) Then
    percentincreaseposition = counter
    End If
    
If Cells(counter, 11) < Cells(percentdecreaseposition, 11) Then
    percentdecreaseposition = counter
    End If
    
If Cells(counter, 12) > Cells(greatestvolumeposition, 12) Then
    greatestvolumeposition = counter
    End If
    
Next counter

Cells(2, 15).Value = Cells(percentincreaseposition, 9).Value
Cells(2, 16).Value = Cells(percentincreaseposition, 11).Value
Cells(2, 16).NumberFormat = "0.00%"
Cells(3, 15).Value = Cells(percentdecreaseposition, 9).Value
Cells(3, 16).Value = Cells(percentdecreaseposition, 11).Value
Cells(3, 16).NumberFormat = "0.00%"
Cells(4, 15).Value = Cells(greatestvolumeposition, 9).Value
Cells(4, 16).Value = Cells(greatestvolumeposition, 12).Value

ActiveSheet.Columns("N:P").AutoFit

Next sheetcounter

End Sub



















