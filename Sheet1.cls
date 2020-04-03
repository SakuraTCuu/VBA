VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub Test()
  Dim d, ar
    Set d = CreateObject("Scripting.Dictionary")
    ar = Sheet2.Range("A1").CurrentRegion
    For i = 1 To UBound(ar)
        d(ar(i, 1)) = Sheet2.Range("A" & i).Value
    Next
    With Sheet1
    For i = 1 To UBound(ar)
          Sheet1.Range("A" & (i + 10)).Value = d(ar(i, 1))
        Next
    End With
End Sub

Private Sub modify_Click()

End Sub

Private Sub query_Click()

Sheet1.Range("A1") = TextBox1.Text
Dim count As Long
Dim index As Long
count = Sheet2.UsedRange.Rows.count

Do While count > 0
    Sheet2.Rows(index).Copy Destination:=Sheet1.Rows(index + 10)
    count = count - 1
    Loop


MsgBox count

End Sub

Sub 行()
MsgBox "sheet1共有" & Sheet2.UsedRange.Rows.count & "行记录。"
End Sub
Private Sub TextBox1_Change()

End Sub


