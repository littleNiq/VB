VERSION 5.00
Begin VB.Form Form16 
   BackColor       =   &H00FFFFC0&
   Caption         =   "开课信息"
   ClientHeight    =   7020
   ClientLeft      =   6030
   ClientTop       =   3045
   ClientWidth     =   9720
   LinkTopic       =   "Form16"
   ScaleHeight     =   7020
   ScaleWidth      =   9720
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   735
      Left            =   7560
      TabIndex        =   19
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   615
      Left            =   7560
      TabIndex        =   12
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   615
      Left            =   7560
      TabIndex        =   11
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1680
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   1680
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   1680
      TabIndex        =   8
      Top             =   5760
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   4920
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   4920
      TabIndex        =   5
      Top             =   5760
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   8415
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "课号"
      Height          =   180
      Left            =   1200
      TabIndex        =   18
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "课名"
      Height          =   180
      Left            =   1200
      TabIndex        =   17
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "任课老师"
      Height          =   180
      Left            =   840
      TabIndex        =   16
      Top             =   5880
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   4320
      TabIndex        =   15
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "类型"
      Height          =   180
      Left            =   4320
      TabIndex        =   14
      Top             =   4800
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "时间"
      Height          =   180
      Left            =   4320
      TabIndex        =   13
      Top             =   6000
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "请选择班级"
      Height          =   180
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程名"
      Height          =   180
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Private Type teacoutype
  cno As String * 2
  cname As String * 6
  js As String * 6
  bj As String * 2
  lx As String * 1
  dt As String * 3
End Type
Dim c1 As teacoutype
Dim teacoulist() As teacoutype
Dim teacoucurno As Integer
Dim teacounlen As Integer
Private Sub Combo1_Click()
  Command1.Enabled = True
  Command2.Enabled = True
  List1.Clear
  If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    For i = 1 To teacounlen
      If Trim(Combo1.List(Combo1.ListIndex)) = Trim(teacoulist(i).cname) Then
        Text1.Text = teacoulist(i).cno
        Text2.Text = teacoulist(i).cname
        Text3.Text = teacoulist(i).js
        Text4.Text = teacoulist(i).bj
        Text5.Text = teacoulist(i).lx
        Text6.Text = teacoulist(i).dt
        Exit For
      End If
    Next
  End If
  Combo1.Text = ""
  Combo2.Text = ""
  If teacoucurno = teacounlen Then Command2.Enabled = False
  If teacoucurno = 1 Then Command1.Enabled = False
End Sub

Private Sub Combo2_click()
  Dim lxx As String
  Dim n As Integer
  n = 0
  List1.Clear
  If Trim(Combo2.List(Combo2.ListIndex)) <> "" Then
    For i = 1 To teacounlen
      If Trim(Combo2.List(Combo2.ListIndex)) = Trim(teacoulist(i).bj) Then
        lxx = "课号：" + teacoulist(i).cno + "--课程名：" + teacoulist(i).cname + "--班级" + teacoulist(i).bj
        List1.List(n) = lxx
        n = n + 1
      End If
    Next i
  End If
  Combo1.Text = ""
  Combo2.Text = ""
End Sub

Private Sub Command1_Click()
  teacoucurno = teacoucurno - 1
  Command2.Enabled = True
  If teacoucurno = 1 Then Command1.Enabled = False
  Text1.Text = teacoulist(teacoucurno).cno
  Text2.Text = teacoulist(teacoucurno).cname
  Text3.Text = teacoulist(teacoucurno).js
  Text4.Text = teacoulist(teacoucurno).bj
  Text5.Text = teacoulist(teacoucurno).lx
  Text6.Text = teacoulist(teacoucurno).dt
End Sub

Private Sub Command2_Click()
  teacoucurno = teacoucurno + 1
  Command1.Enabled = True
  If teacoucurno = teacounlen Then Command2.Enabled = False
  Text1.Text = teacoulist(teacoucurno).cno
  Text2.Text = teacoulist(teacoucurno).cname
  Text3.Text = teacoulist(teacoucurno).js
  Text4.Text = teacoulist(teacoucurno).bj
  Text5.Text = teacoulist(teacoucurno).lx
  Text6.Text = teacoulist(teacoucurno).dt
End Sub

Private Sub Command3_Click()
  Unload Me
  Form3.Show
End Sub

Private Sub Form_Load()
  Dim strfilename As String
  Dim n As Integer
  strfilename = App.Path & "\cou.dat"
  If Dir(strfilename) <> "" Then
    Open App.Path & "\cou.dat" For Input As #1
    teacounlen = 0
    teacoucurno = 0
    Do While Not EOF(1)
    Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
    If Trim(c1.js) = Trim(teacher) Then
      teacounlen = teacounlen + 1
    End If
    Loop
    Close #1
    If teacounlen = 0 Then
      Command2.Enabled = False
      Command3.Enabled = False
      a = MsgBox("课程信息为空", 49, "信息确认")
    Else
      ReDim teacoulist(teacounlen)
      Open App.Path & "\cou.dat" For Input As #1
      i = 0
      Do While Not EOF(1)
        Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
        If Trim(c1.js) = Trim(teacher) Then
          i = i + 1
          teacoulist(i).cno = c1.cno
          teacoulist(i).cname = c1.cname
          teacoulist(i).js = c1.js
          teacoulist(i).bj = c1.bj
          teacoulist(i).lx = c1.lx
          teacoulist(i).dt = c1.dt
        End If
      Loop
      Close #1
      teacoucurno = 1
      Text1.Text = teacoulist(teacoucurno).cno
      Text2.Text = teacoulist(teacoucurno).cname
      Text3.Text = teacoulist(teacoucurno).js
      Text4.Text = teacoulist(teacoucurno).bj
      Text5.Text = teacoulist(teacoucurno).lx
      Text6.Text = teacoulist(teacoucurno).dt
      Command1.Enabled = False
      For i = 1 To teacounlen
        Combo1.AddItem teacoulist(i).cname
      Next
      For i = 1 To teacounlen
        n = 0
        For j = 0 To Combo2.ListCount - 1
          If Trim(Combo2.List(j)) = Trim(teacoulist(i).bj) Then
            n = n + 1
            Exit For
          End If
        Next
        If n = 0 Then
          Combo2.AddItem teacoulist(i).bj
        End If
      Next
    End If
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
    Command2.Enabled = False
    Command3.Enabled = False
  End If
  teacoucurno = 1
End Sub

Private Sub List1_Click()
  Command1.Enabled = True
  Command2.Enabled = True
  Dim lxx As String
  lxx = List1.List(List1.ListIndex)
  For i = 1 To teacounlen
    If teacoulist(i).cno = Mid(lxx, 4, 2) Then
      Text1.Text = teacoulist(i).cno
      Text2.Text = teacoulist(i).cname
      Text3.Text = teacoulist(i).js
      Text4.Text = teacoulist(i).bj
      Text5.Text = teacoulist(i).lx
      Text6.Text = teacoulist(i).dt
      teacoucurno = i
      Exit For
    End If
  Next
  If teacoucurno = teacounlen Then Command2.Enabled = False
  If teacoucurno = 1 Then Command1.Enabled = False
End Sub
