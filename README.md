# VB6
Visual Basic 6 이건뭐 객체지향에 절차지향 담갔다 뺀건가ㅏㅏㅏㅏㅏ

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
    intTest = Len(strTest) '변수 최소 할당 공간은 2이며, 1만큼 비워두는 것으로 추정.
    '공백과 1의 길이는 2
    '11의 길이는 3

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