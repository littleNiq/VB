VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFC0FF&
   Caption         =   "用户修改"
   ClientHeight    =   7905
   ClientLeft      =   7935
   ClientTop       =   2640
   ClientWidth     =   5925
   LinkTopic       =   "Form6"
   ScaleHeight     =   7905
   ScaleWidth      =   5925
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   495
      Left            =   3360
      TabIndex        =   11
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   495
      Left            =   720
      TabIndex        =   10
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   735
      Left            =   3360
      TabIndex        =   5
      Top             =   6600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   735
      Left            =   840
      TabIndex        =   4
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "权限"
      Height          =   180
      Left            =   1440
      TabIndex        =   9
      Top             =   5280
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "确认密码"
      Height          =   180
      Left            =   1440
      TabIndex        =   8
      Top             =   4080
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   1440
      TabIndex        =   7
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   540
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gg As Integer
Private Sub Command1_Click()
  If Text2.Text = Text2.Text Then
    uslist(gg).yhm = Trim(Text1.Text)
    uslist(gg).mi = Trim(Text3.Text)
    uslist(gg).qx = Trim(Text4.Text)
    Open App.Path & "\user.dat" For Output As #1
    For i = 1 To usnlen
      Write #1, uslist(i).yhm, uslist(i).mi, uslist(i).qx
    Next i
    Close #1
    MsgBox ("已修改")
  Else
    MsgBox ("密码不一致")
  End If
End Sub

Private Sub Command2_Click()
  Unload Me
  Form2.Show
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
