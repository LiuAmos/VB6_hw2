VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   7440
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "樂透開獎"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num(7), i, j, m, k, tmp As Variant
Private Sub Command1_Click()
num(7) = Int(Rnd() * 49) + 1
'List1.AddItem num(7)
For i = 1 To 6

    Randomize (Timer)
    num(i) = Int(Rnd() * 49) + 1
    For j = 1 To i - 1 '讓第i個產生的號碼,跟之前已經產生過的號碼比較, 如果重複，便重新選取,因此要讓i=i-1(倒退一位)
        If num(i) = num(j) Then
            i = i - 1
            Exit For
        End If
    Next
    Next
'List1.AddItem num(i)
List1.AddItem "----排序前----"
For i = 1 To 6
    List1.AddItem num(i)
Next
List1.AddItem "特別號:"
List1.AddItem num(7)

List1.AddItem "----排序過程----"
For k = 1 To 6
    For m = k To 6
        If num(k) > num(m) Then ' 當 第 i 位 > 第 j 位
        '(由小排到大 用 > 由大排到小 用 < )
            tmp = num(k) ' 先將 第 i 位 存起來
            num(k) = num(m) ' 將 第 i 位 取代為 第 j 位
            num(m) = tmp ' 將 第 j 位 取代為 原本存起來的 Tag (第 i 位)
        End If
    Next m
    For i = 1 To 6
        List1.AddItem num(i)
    Next
    List1.AddItem "--------"
Next k
List1.AddItem "----排序後----"
For i = 1 To 6
    List1.AddItem num(i)
Next

List1.AddItem "特別號:"
List1.AddItem num(7)

End Sub

