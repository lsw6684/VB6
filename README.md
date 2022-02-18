# VB6
Visual Basic 6

- [Fundamental Functions](#fundamental-functions)


## Fundamental functions
- 출력, 선언, 입력, 반복문
    ```vb
    Option Explicit
    Private Sub Form_Click()

    '출력
    Print "This is my first VB"
    
    '선언
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
        
    a = 1
    b = 2
    c = a + b
    Print "a = " & a & " b = " & b & " c = " & c
    
    '입력
    Dim inputTest As Integer
    Dim printTest As Integer
    
    inputTest = InputBox("Enter numper : ")
    Print inputTest
    
    '반복문
    Dim i As Integer
    Dim n As Integer
    n = 5
    
    For i = 0 To n
        Print i
    Next i
    End Sub
    ```
- Casting
    ```vb
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
    ```