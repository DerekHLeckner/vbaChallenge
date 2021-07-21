Attribute VB_Name = "Module1"
Public Sub startFresh()

Dim ticker1 As String
Dim ticker2 As String
Dim beginning As Double
Dim ending As Double
Dim total As Long
Dim n As Integer
Dim o As Integer



Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Open Price"
Cells(1, "K").Value = "Closing Price"
Cells(1, "L").Value = "Difference"
Cells(1, "M").Value = "Percent Change"
Cells(1, "N").Value = "Total Volume (hundreds)"
Cells(1, "Q").Value = "Ticker"
Cells(1, "R").Value = "Value"
Cells(2, "P").Value = "Greatest % Increase"
Cells(3, "P").Value = "Greatest % Decrease"
Cells(4, "P").Value = "Greatest Total Volume"




o = 2
ticker1 = "A"


For l = 2 To 705714
If ticker1 <> Cells(l - 1, "A") Then
If l = 2 Then
beginning = Cells(l, "C")
Cells(o, "J").Value = beginning
ticker1 = Cells(l, "A")
o = o + 1

ElseIf l > 2 Then
beginning = Cells(l - 1, "C")
Cells(o, "J").Value = beginning
ticker1 = Cells(l, "A")
o = o + 1




End If
End If
Next l

n = 2
ticker2 = "A"

For i = 2 To 705714
If ticker2 = Cells(i, "A") Then
total = total + Cells(i, "G") / 100



ElseIf ticker2 <> Cells(i, "A") Then
closing = Cells(i - 1, "F")

Cells(n, "I").Value = ticker2
Cells(n, "N").Value = total
total = 0
Cells(n, "K").Value = closing
Cells(n, "L").Value = Cells(n, "K") - Cells(n, "J")

If Cells(n, "J") = 0 Then
Cells(n, "M").Value = 0
Else:
Cells(n, "M").Value = ((Cells(n, "K") - Cells(n, "J")) / Cells(n, "J")) * 100
End If
If Cells(n, "L") >= 0 Then
Cells(n, "L").Interior.Color = vbGreen
ElseIf Cells(n, "L") < 0 Then
Cells(n, "L").Interior.Color = vbRed
End If


ticker2 = Cells(i, "A")
n = n + 1

End If
Next i

' Min
Dim dblMin As Double
Set rng = Range("M:M")
dblMin = Application.Min(rng)
Cells(3, "R").Value = dblMin
For i = 2 To 2836
If Cells(i, "M") = dblMin Then
Cells(3, "Q").Value = Cells(i, "I")
End If
Next i

' Max Percent
Dim dblMax As Double
Set rng = Range("M:M")
dblMax = Application.Max(rng)
Cells(2, "R").Value = dblMax
For i = 2 To 2836
If Cells(i, "M") = dblMax Then
Cells(2, "Q").Value = Cells(i, "I")
End If
Next i

' Max Volume
Dim dblVol As Double
Set rng = Range("N:N")
dblVol = Application.Max(rng)
Cells(4, "R").Value = dblVol
For i = 2 To 2836
If Cells(i, "N") = dblVol Then
Cells(4, "Q").Value = Cells(i, "I")
End If
Next i


End Sub


Sub applyAll()


Dim ws As Worksheet
For Each ws In Worksheets
    ws.Activate
    Call startFresh

Next

End Sub

