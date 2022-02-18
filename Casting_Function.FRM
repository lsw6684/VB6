Option Explicit
Public Function Format4(temp As Integer) As String

Dim strTest As String
Dim intTest As Integer

'casting & length
strTest = Str(temp)
intTest = Len(strTest)

Format4 = temp & Space(intTest)

End Function
Private Sub Form_Click()

Dim i As Integer
Dim n As Integer

n = InputBox("Please enter number:")

For i = 1 To n
Print Format4(i);
Next i
End Sub