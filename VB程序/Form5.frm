VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0080FFFF&
   Caption         =   "用户注册"
   ClientHeight    =   7860
   ClientLeft      =   7380
   ClientTop       =   2640
   ClientWidth     =   6690
   LinkTopic       =   "Form5"
   ScaleHeight     =   7860
   ScaleWidth      =   6690
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   3480
      TabIndex        =   5
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   6480
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "权限"
      Height          =   180
      Left            =   2040
      TabIndex        =   9
      Top             =   4560
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "确认密码"
      Height          =   180
      Left            =   2040
      TabIndex        =   8
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   2040
      TabIndex        =   7
      Top             =   1680
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   2040
      TabIndex        =   6
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text4.Text = 1 Or Text4.Text = 2 Or Text4.Text = 3 Then
    If Text2.Text = Text3.Text Then
      usnlen = usnlen + 1
      ReDim Preserve uslist(usnlen)
      uslist(usnlen).yhm = Trim(Text1.Text)
      uslist(usnlen).mi = Trim(Text3.Text)
      uslist(usnlen).qx = Trim(Text4.Text)
      Open App.Path & "\user.dat" For Output As #1
      For i = 1 To usnlen
        Write #1, uslist(i).yhm, uslist(i).mi, uslist(i).qx
      Next i
      Close #1
      MsgBox ("已添加")
    Else
      MsgBox ("密码不一致")
    End If
  Else
    MsgBox ("权限错误")
  End If
End Sub

Private Sub Command2_Click()
  Form2.Show
  Unload Me
End Sub

Private Sub Form_Load()
  Open App.Path & "\user.dat" For Input As #1
  usnlen = 0
  Do While Not EOF(1)
    Input #1, us.yhm, us.mi, us.qx
    usnlen = usnlen + 1
  Loop
  Close #1
  Print usnlen
  ReDim uslist(usnlen)
  Open App.Path & "\user.dat" For Input As #1
  For i = 1 To usnlen
    Input #1, uslist(i).yhm, uslist(i).mi, uslist(i).qx
  Next i
  Close #1
End Sub
