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
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "�ֳz�}��"
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
    For j = 1 To i - 1 '����i�Ӳ��ͪ����X,�򤧫e�w�g���͹L�����X���, �p�G���ơA�K���s���,�]���n��i=i-1(�˰h�@��)
        If num(i) = num(j) Then
            i = i - 1
            Exit For
        End If
    Next
    Next
'List1.AddItem num(i)
List1.AddItem "----�Ƨǫe----"
For i = 1 To 6
    List1.AddItem num(i)
Next
List1.AddItem "�S�O��:"
List1.AddItem num(7)

List1.AddItem "----�ƧǹL�{----"
For k = 1 To 6
    For m = k To 6
        If num(k) > num(m) Then ' �� �� i �� > �� j ��
        '(�Ѥp�ƨ�j �� > �Ѥj�ƨ�p �� < )
            tmp = num(k) ' ���N �� i �� �s�_��
            num(k) = num(m) ' �N �� i �� ���N�� �� j ��
            num(m) = tmp ' �N �� j �� ���N�� �쥻�s�_�Ӫ� Tag (�� i ��)
        End If
    Next m
    For i = 1 To 6
        List1.AddItem num(i)
    Next
    List1.AddItem "--------"
Next k
List1.AddItem "----�Ƨǫ�----"
For i = 1 To 6
    List1.AddItem num(i)
Next

List1.AddItem "�S�O��:"
List1.AddItem num(7)

End Sub

