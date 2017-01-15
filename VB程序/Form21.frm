VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H00FFFFC0&
   Caption         =   "学生名单"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form21"
   ScaleHeight     =   7965
   ScaleWidth      =   8925
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   975
      Left            =   6240
      TabIndex        =   15
      Top             =   6600
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1800
      TabIndex        =   13
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   0
      Left            =   960
      TabIndex        =   12
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   1
      Left            =   960
      TabIndex        =   11
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   4
      Left            =   960
      TabIndex        =   8
      Top             =   4080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   5
      Left            =   4200
      TabIndex        =   7
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   6
      Left            =   4200
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   7
      Left            =   4200
      TabIndex        =   5
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   8
      Left            =   4200
      TabIndex        =   4
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   9
      Left            =   4200
      TabIndex        =   3
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全部显示"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程"
      Height          =   180
      Left            =   840
      TabIndex        =   14
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Private Type stucoutype
  cno As String * 2
  cname As String * 6
  js As String * 6
  bj As String * 2
  lx As String * 1
  dt As String * 3
End Type
Private Type stuxktype
  cno As String * 2
  sno As String * 6
  cj As String * 3
End Type
Private Type stutype
  sno As String * 6
  sname As String * 6
  bj As String * 2
End Type
Private Type xkstutype
  sno As String * 6
  sname As String * 6
  cj As String * 3
End Type
Dim s1 As stuxktype
Dim stuxklen As Integer
Dim stuxklist() As stuxktype
Dim c1 As stucoutype
Dim stucoulist() As stucoutype
Dim curpge As Integer
Dim pgecount As Integer
Dim pgemod As Integer
Private Sub Combo1_Click()
  Call stuxkclea
  If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    For i = 1 To stuxklen
      If Trim(Combo1.List(Combo1.ListIndex)) = Trim(stucoulist(i).cname) Then
        Exit For
      End If
    Next
  End If
  Text1(0).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
End Sub

Private Sub Command1_Click()
  Call stuxkclea
  pgecount = (stuxklen - 1) \ 10 + 1
  pgemod = stuxklen Mod 10
  curpge = 1
  If pgecount <= 1 Then
      For i = 1 To pgemod
        Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
      Next
  Else
      For i = 1 To 10
        Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
      Next
    End If
  Command3.Enabled = True
End Sub

Private Sub Command2_Click()
  If curpge > 1 Then
    Call stuxkclea
    curpge = curpge - 1
    For i = 1 To 10
      Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
    Next
    Command3.Enabled = True
  Else
    Command2.Enabled = False
  End If
End Sub

Private Sub Command3_Click()
  If curpge < pgecount - 1 Then
    Call stuxkclea
    curpge = curpge + 1
    For i = 1 To 10
      Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
    Next
    Command2.Enabled = True
  ElseIf curpge = pgecount - 1 Then
    Call stuxkclea
    curpge = curpge + 1
    For i = 1 To pgemod
      Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
    Next
    Command2.Enabled = True
  Else
    Command3.Enabled = False
  End If
End Sub

Private Sub Command4_Click()
  Unload Me
  Form4.Show
End Sub

Private Sub Form_Load()
    Open App.Path & "\xk.dat" For Input As #1
    stuxklen = 0
    Do While Not EOF(1)
      Input #1, s1.cno, s1.sno, s1.cj
      If Trim(s1.sno) = Trim(student) Then
        stuxklen = stuxklen + 1
      End If
    Loop
    Close #1
    If stuxklen = 0 Then
      a = MsgBox("课程信息为空", 49, "信息确认")
    Else
      ReDim stuxklist(stuxklen)
      Open App.Path & "\xk.dat" For Input As #1
      i = 0
      Do While Not EOF(1)
        Input #1, s1.cno, s1.sno, s1.cj
        If Trim(s1.sno) = Trim(student) Then
          i = i + 1
          stuxklist(i).cno = s1.cno
          stuxklist(i).sno = s1.sno
          stuxklist(i).cj = s1.cj
        End If
      Loop
      Close #1
    End If
    ReDim stucoulist(stuxklen)
    For i = 1 To stuxklen
      Open App.Path & "\cou.dat" For Input As #1
      Do While Not EOF(1)
        Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
        If Trim(c1.cno) = Trim(stuxklist(i).cno) Then
          stucoulist(i).cno = c1.cno
          stucoulist(i).cname = c1.cname
          stucoulist(i).js = c1.js
          stucoulist(i).bj = c1.bj
          stucoulist(i).lx = c1.lx
          stucoulist(i).dt = c1.dt
        End If
      Loop
      Close #1
    Next
    For i = 1 To stuxklen
      Combo1.AddItem stucoulist(i).cname
    Next
    pgecount = (stuxklen - 1) \ 10 + 1
    pgemod = stuxklen Mod 10
    curpge = 1
    If pgecount <= 1 Then
      For i = 1 To pgemod
        Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
      Next
    Else
      For i = 1 To 10
        Text1(i - 1).Text = stucoulist(i).cname + "--" + stuxklist(i).cj
      Next
    End If
    Command3.Enabled = True
End Sub

Private Sub stuxkclea()
  For i = 0 To 9
    Text1(i).Text = ""
  Next
End Sub
