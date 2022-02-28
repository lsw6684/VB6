# VB6
Visual Basic 6 이건뭐 객체지향에 절차지향 담갔다 뺀건가ㅏㅏㅏㅏㅏ

- [기본 컨트롤](#기본-컨트롤)
    - [Form](#form)
    - [Button](#button)
    - [Text Box](#text-box)
    - [Label](#label)
    - [GroupBox](#groupbox)
    
- [Fundamental Functions](#fundamental-functions)

## 기본 컨트롤

### Form
프로그램을 실행시켰을 때 나타나는 윈도우 객체입니다. 모양과 행동을 제어할 수 있는 **속성, 메서드, 이벤트**가 있습니다.

### Button
명령을 내릴 수 있도록 하는 버튼으로, 모양을 나타내는 **속성과 클릭 이벤트**가 있습니다.

### Text Box
입력, 텍스트 출력에 사용되는 컨트롤로 **글자 수를 제한하는 속성과 여러 줄로 표시하는 속성** 등이 있습니다.

### Label
사용자가 변경할 수 없는 텍스트를 표시할 때 사용하는 컨트롤이며, 레이블의 모양을 결정하는 속성들을 가지고 있습니다.

### GroupBox
그룹핑 하여 사용하는 컨트롤로 모양을 결정하는 속성을 가집니다.



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
    '또는   For 변수 = 초기 값 To 최종 값 Step 증감값
    '           내용
    '       Next 변수
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