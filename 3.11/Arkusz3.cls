VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Worksheets("Sales").Range("A1").Value = "Hello"
Worksheets(1).Range("A1").Value = "Hello"
Sheet1.Range("A1").Value = "Hello"

End Sub

Private Sub CommandButton2_Click()
Workbooks.Add
MsgBox Worksheets.Count
End Sub

Private Sub CommandButton3_Click()
MsgBox Workbooks("path-fullname.xlsm").Path
MsgBox ActiveWorkbook.FullName

End Sub

Private Sub CommandButton4_Click()
Workbooks("close-open.xlsm").Close
Workbooks(1).Close
ActiveWorkbook.Close
Workbooks.Close
Workbooks.Open ("sales.xlsx")
Dim MyFile As String
MyFile = Application.GetOpenFilename()
Workbooks.Open (MyFile)
End Sub

Private Sub CommandButton5_Click()
Dim book As Workbook, sheet As Worksheet, text As String
For Each book In Workbooks
text = text & "Workbook: " & book.Name & vbNewLine & "Worksheets: " & vbNewLine
For Each sheet In book.Worksheets
text = text & sheet.Name & vbNewLine
Next sheet
text = text & vbNewLine
Next book
MsgBox text

End Sub

Private Sub CommandButton6_Click()
Dim employee As String, total As Integer, sheet As Worksheet, i As Integer
total = 0
employee = InputBox("Enter the employee name (case sensitive)")
For Each sheet In Worksheets
For i = 2 To 13
If sheet.Cells(i, 2).Value = employee Then
    total = total + sheet.Cells(i, 3).Value
End If
    Next i
Next sheet
MsgBox "Total sales of " & employee & " is " & total

End Sub

Private Sub CommandButton7_Click()
Dim directory As String, fileName As String, sheet As Worksheet, i As Integer, j As Integer
Application.ScreenUpdating = False
directory = "c:\test\"
fileName = Dir(directory & "*.xl??")
Do While fileName <> ""

Loop
i = i + 1
j = 2
Cells(i, 1) = fileName
For Each sheet In Workbooks(fileName).Worksheets
    Workbooks("files-in-a-directory.xlsm").Worksheets(1).Cells(i, j).Value = sheet.Name
    j = j + 1
Next sheet
Workbooks(fileName).Close
fileName = Dir()
Application.ScreenUpdating = True

End Sub

Private Sub NieKlikac_Click()
Range("A1").Value = "Hello"
Application.Workbooks("create-a-macro").Worksheets(1).Range("A1").Value = "Hello"

End Sub
