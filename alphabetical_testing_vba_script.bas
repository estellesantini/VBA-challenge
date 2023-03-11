Attribute VB_Name = "Module1"
Sub alphabetical()

'Declaring Variable with Datatypes
Dim i As Long
Dim output As Long
Dim volume As Double
Dim opened As Double
Dim closed As Double
Dim least As Double
Dim greatest As Double
Dim g_volume As Double
Dim Current As Worksheet
For Each Current In Worksheets

'Initialize or assigning a variable
Current.Cells(2, 9).Value = Current.Cells(2, 1).Value
Current.Cells(1, 9).Value = "Ticker"
Current.Cells(1, 10).Value = "Yearly Change"
Current.Cells(1, 11).Value = "Percent Change"
Current.Cells(1, 12).Value = "Total Stock Volume"
Current.Cells(1, 16).Value = "Ticker"
Current.Cells(1, 17).Value = "Value"
Current.Cells(2, 15).Value = "Greatest % Increase"
Current.Cells(3, 15).Value = "Greatest % Decrease"
Current.Cells(4, 15).Value = "Greatest Total Volume"
opened = Current.Cells(2, 3).Value
output = 2
volume = 0
greatest = Current.Cells(2, 11).Value
least = Current.Cells(2, 11).Value
g_volume = Current.Cells(2, 12).Value

'Create a for loop from 3 to 22771
For i = 3 To 22771

'Conditional when there is no ticker change
If Current.Cells(i - 1, 1).Value = Current.Cells(i, 1).Value Then
    volume = volume + Current.Cells(i - 1, 7).Value
' Close the If/Else Statement
End If

'Conditional when a ticker change occurs
If Current.Cells(i - 1, 1).Value <> Current.Cells(i, 1).Value Then
    closed = Current.Cells(i - 1, 6).Value
    Current.Cells(output, 9).Value = Current.Cells(i - 1, 1).Value
    Current.Cells(output, 10).Value = closed - opened
    Current.Cells(output, 11).Value = FormatPercent((closed - opened) / (opened), 2)
    Current.Cells(output, 12).Value = volume + Current.Cells(i - 1, 7).Value
    output = output + 1 'This makes the next ticker go onto the next row
    opened = Current.Cells(i, 3).Value
    volume = 0 'Resets volume to zero for next ticker's volume
' Close the If/Else Statement
End If

Next i

'Create a second loop from 2 to output
For i = 2 To (output - 1)

'Summary Statistics Loop
If Current.Cells(i, 11).Value > greatest Then
    greatest = Current.Cells(i, 11).Value
    Current.Cells(2, 17).Value = FormatPercent(greatest, 2)
    Current.Cells(2, 16).Value = Current.Cells(i, 9).Value
End If

If Current.Cells(i, 11).Value < least Then
least = Current.Cells(i, 11).Value
    Current.Cells(3, 17).Value = FormatPercent(least, 2)
    Current.Cells(3, 16).Value = Current.Cells(i, 9).Value
End If

If Current.Cells(i, 12).Value > g_volume Then
    g_volume = Current.Cells(i, 12).Value
    Current.Cells(4, 17).Value = g_volume
    Current.Cells(4, 16).Value = Current.Cells(i, 9).Value
End If

If Current.Cells(i, 10).Value > 0 Then
    Current.Cells(i, 10).Interior.Color = vbGreen
    Else
    Current.Cells(i, 10).Interior.Color = vbRed
End If

Next i

Next

End Sub

