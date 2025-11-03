VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim temp As Double
temp = Range("A1").Value
Range("A1").Value = Range("B1").Value
Range("B1").Value = temp
End Sub

Private Sub CommandButton2_Click()
Cells.Interior.ColorIndex = 28
End Sub

Private Sub CommandButton3_Click()
Range("D4").Formula = "=B3*10"
Range("D4").FormulaR1C1 = "=R3C2*10"
Range("D4").FormulaR1C1 = "=R[-1]C[-2]*10"

End Sub
