VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FFC0FF&
   Caption         =   "用户删除"
   ClientHeight    =   9030
   ClientLeft      =   8130
   ClientTop       =   2085
   ClientWidth     =   6855
   LinkTopic       =   "Form7"
   ScaleHeight     =   9030
   ScaleWidth      =   6855
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   855
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   855
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "权限"
      Height          =   180
      Left            =   1680
      TabIndex        =   9
      Top             =   5760
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "确认密码"
      Height          =   180
      Left            =   1680
      TabIndex        =   8
      Top             =   4560
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1680
      TabIndex        =   7
      Top             =   3360
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   540
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gg As Integer
Private Sub Command1_Click()
  If Text2.Text = Text3.Text Then
    n = usnlen - 1
    For i = gg To n
      uslist(i).yhm = uslist(i + 1).yhm
      uslist(i).mi = uslist(i + 1).mi
      uslist(i).qx = uslist(i + 1).qx
    Next
    ReDim Preserve uslist(n)
    Open App.Path & "\user.dat" For Output As #1
    For i = 1 To n
      Write #1, uslist(i).yhm, uslist(i).mi, uslist(i).qx
    Next i
    Close #1
    MsgBox ("已删除")
    usnlen = n
    gg = 1
  Else
    MsgBox ("密码不一致")
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
  Form2.Show
End Sub

Private Sub Form_Load()
  gg = 1
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
Private Sub Command3_Click()
  gg = gg - 1
  If gg = 0 Then gg = usnlen
  If gg = usnlen + 1 Then gg = 1
  Text1.Text = uslist(gg).yhm
  Text2.Text = uslist(gg).mi
  Text3.Text = uslist(gg).mi
  Text4.Text = uslist(gg).qx
End Sub

Private Sub Command4_Click()
  gg = gg + 1
  If gg = 0 Then gg = usnlen
  If gg = usnlen + 1 Then gg = 1
  Text1.Text = uslist(gg).yhm
  Text2.Text = uslist(gg).mi
  Text3.Text = uslist(gg).mi
  Text4.Text = uslist(gg).qx
End Sub
