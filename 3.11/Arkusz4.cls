VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Range("B3").Value = 2
Range("A1:A4").Value = 5
Range("A1:A2,B3:C4").Value = 10
Range("Prices").Value = 15
Cells(3, 2).Value = 2
Range(Cells(1, 1), Cells(4, 1)).Value = 5
Dim example As Range
Set example = Range("A1:C4")

example.Value = 8
Dim example As Range
Set example = Range("A1:C4")

example.Select
Worksheets(3).Activate
Worksheets(3).Range("B7").Select
Dim example As Range
Set example = Range("A1:C4")

example.Rows(3).Select
Dim example As Range
Set example = Range("A1:C4")

example.Columns(2).Select
Range("A1:A2").Select
Selection.Copy

Range("C3").Select
ActiveSheet.Paste
Range("C3:C4").Value = Range("A1:A2").Value
Range("A1").ClearContents
Range("A1").Value = ""
Dim example As Range
Set example = Range("A1:C4")

MsgBox example.Count
Dim example As Range
Set example = Range("A1:C4")

MsgBox example.Rows.Count


End Sub

Private Sub CommandButton10_Click()
Range("A1").Font.Color = -16776961
Range("A1").Font.Color = vbRed
Range("A1").Font.Color = RGB(255, 0, 0)
Range("A1").Font.Bold = True

End Sub

Private Sub CommandButton11_Click()
Range("A1").Interior.ColorIndex = 37
Range("A1").Interior.ColorIndex = 0
MsgBox Selection.Interior.ColorIndex
Range("A1").Interior.Color = RGB(255, 0, 0)

End Sub

Private Sub CommandButton12_Click()
Range("A1:A8").Sort Key1:=Range("A1"), Order1:=xlAscending, Header:=xlNo
Range("A1:B8").Sort Key1:=Range("B2"), Order1:=xlDescending, Header:=xlYes
Range("A1:C8").Sort Key1:=Range("B2"), Order1:=xlAscending, Key2:=Range("C2"), Order2:=xlDescending, Header:=xlYes
Range("A1").CurrentRegion.Sort Key1:=Range("B2"), Order1:=xlDescending, Header:=xlYes
Cells.Sort Key1:=Columns("B"), Order1:=xlDescending, Header:=xlYes

End Sub

Private Sub CommandButton13_Click()
Dim rangeToUse As Range, singleArea As Range
Set rangeToUse = Range("B2:C3,C5:E5")
MsgBox rangeToUse.Areas.Count
MsgBox rangeToUse.Areas(1).Count
For Each singleArea In rangeToUse.Areas
    MsgBox singleArea.Count
Next singleArea

End Sub

Private Sub CommandButton14_Click()
Dim rangeToUse As Range, singleArea As Range, cell1 As Range, cell2 As Range, i As Integer, j As Integer
Set rangeToUse = Selection
Cells.Interior.ColorIndex = 0
Cells.Borders.LineStyle = xlNone
If Selection.Areas.Count <= 1 Then
      MsgBox "Please select more than one area."
Else


End If
rangeToUse.Interior.ColorIndex = 38
For Each singleArea In rangeToUse.Areas
    singleArea.BorderAround ColorIndex:=1, Weight:=xlThin
Next singleArea
For i = 1 To rangeToUse.Areas.Count
    For j = i + 1 To rangeToUse.Areas.Count
        For Each cell1 In rangeToUse.Areas(i)
            For Each cell2 In rangeToUse.Areas(j)
                If cell1.Value = cell2.Value Then
                    cell1.Interior.ColorIndex = 0
                    cell2.Interior.ColorIndex = 0
                End If
            Next cell2
        Next cell1
    Next j
Next i

End Sub

Private Sub CommandButton2_Click()
Range("A1").CurrentRegion.Select
Range("B3").CurrentRegion.Select

End Sub

Private Sub CommandButton3_Click()
Dim maximum As Double, rng As Range, cell As Range
Cells.Interior.ColorIndex = 0
Set rng = Range("A1").CurrentRegion
maximum = WorksheetFunction.Max(rng)
For Each cell In rng
    If cell.Value = maximum Then cell.Interior.ColorIndex = 22
Next cell

End Sub

Private Sub CommandButton4_Click()
Range("A1:C4").Resize(3, 2).Select
Range("A1:C4").Resize(, 1).Select
Range("A1:C4").Resize(, 5).Select
Range("A1:C4").Resize(10) = 99

End Sub

Private Sub CommandButton5_Click()
Worksheets(2).Activate
Worksheets(2).Cells.Select
Columns(2).Select
Rows(7).Select
Rows("5:7").Select
Columns("B:E").Select
MsgBox Cells(5, 2).Row
ActiveCell.EntireRow.Select
ActiveCell.EntireColumn.Cells(1).Value = 2
ActiveCell.EntireRow.Offset(1, 0).Cells(1).Value = 3

End Sub

Private Sub CommandButton6_Click()
Range("A1:A2").Offset(3, 2).Select
Range("F5:H5").Offset(-1, -4).Select
Range("B2:D3").Offset(5).Select
Range("G8").Offset(0, -2).Select

End Sub

Private Sub CommandButton7_Click()
Range("A5").End(xlDown).Select
Range(Range("A5"), Range("A5").End(xlDown)).Select
Range(ActiveCell, ActiveCell.End(xlDown)).Select

End Sub

Private Sub CommandButton8_Click()
Union(Range("B2:C7"), Range("C6:F8")).Select
Intersect(Range("B2:C7"), Range("C6:F8")).Select

End Sub

Private Sub CommandButton9_Click()
Dim rng As Range, cell As Range
Set rng = Range("C4:D10")
Set cell = Selection
If cell.Count = 1 And IsNumeric(cell) And Not IsEmpty(cell) And Not Intersect(rng, cell) Is Nothing Then
    MsgBox "You win"
Else
    MsgBox "You lose"
End If

End Sub
