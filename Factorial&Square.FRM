Option Explicit


Public Function factorial(temp As Integer) As Integer

   Dim i As Integer
   Dim fact As Integer
   
   fact = 1
   
   For i = 2 To temp
      fact = fact * i
   Next i
   
   factorial = fact
   
End Function

Public Function square(temp As Integer) As Integer

   square = temp * temp
   
End Function


Private Sub Form_Click()
   
   Dim n As Integer
   
   n = InputBox("Please enter number(1~10) : ")
   
   Print "factorial result ===>" & factorial(n)
   Print "square result ===>" & square(n)
   
End Sub