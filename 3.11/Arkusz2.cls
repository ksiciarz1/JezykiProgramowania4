VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
MsgBox "This is fun"
MsgBox "Entered value is " & Range("A1").Value
MsgBox "Line 1" & vbNewLine & "Line 2"

End Sub

Private Sub CommandButton2_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to empty the sheet?", vbYesNo + vbQuestion, "Empty Sheet")
If answer = vbYes Then
    Cells.ClearContents
Else
    'do nothing
End If

End Sub

Private Sub CommandButton3_Click()
Dim myValue As Variant
myValue = InputBox("Give me some input")
Range("A1").Value = myValue
myValue = InputBox("Give me some input", "Hi", 1)

End Sub
