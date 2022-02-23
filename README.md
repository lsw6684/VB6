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
    MsgBox("내용", 아이콘 및 버튼 Type, "메시지 박스 제목")

    '선언 - Dim 변수명 As 데이터 형식
    'Dim - Declare In Memory
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
- 세미콜론 ;
    - **&, +** 와 동일하게 문자열을 합칩니다. 
- 배열
```vb
'Dim 배열 이름(사이즈)
Dim a(10) As 
a = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)

Dim b(30, 30) As Integer
```