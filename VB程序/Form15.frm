VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00FFFFC0&
   Caption         =   "删除课程信息"
   ClientHeight    =   8610
   ClientLeft      =   5895
   ClientTop       =   2280
   ClientWidth     =   10395
   LinkTopic       =   "Form15"
   ScaleHeight     =   8610
   ScaleWidth      =   10395
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   5160
      TabIndex        =   16
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   5160
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   5160
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   1080
      TabIndex        =   13
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   1080
      TabIndex        =   12
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   1080
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<<"
      Height          =   615
      Left            =   8280
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   615
      Left            =   8280
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "查询"
      Height          =   615
      Left            =   3600
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   9975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "查询"
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "查询"
      Height          =   615
      Left            =   3600
      TabIndex        =   5
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "查询"
      Height          =   615
      Left            =   7320
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command9 
      Caption         =   "查询"
      Height          =   615
      Left            =   7320
      TabIndex        =   3
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command10 
      Caption         =   "查询"
      Height          =   615
      Left            =   7320
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除"
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "时间"
      Height          =   180
      Left            =   4560
      TabIndex        =   22
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "类型"
      Height          =   180
      Left            =   4560
      TabIndex        =   21
      Top             =   2400
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   4560
      TabIndex        =   20
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "任课老师"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "课名"
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   2280
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "课号"
      Height          =   180
      Left            =   480
      TabIndex        =   17
      Top             =   1320
      Width           =   360
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type coutype
  sno As String * 2
  sname As String * 6
  tc As String * 6
  bj As String * 2
  lx As String * 1
  sj As String * 3
End Type
Dim c1 As coutype
Dim coulist() As coutype
Dim counlen As Integer
Dim gg As Integer

Private Sub Command1_Click()
    n = counlen - 1
    For i = gg To n
      coulist(i).sno = coulist(i + 1).sno
      coulist(i).sname = coulist(i + 1).sname
      coulist(i).tc = coulist(i + 1).tc
      coulist(i).bj = coulist(i + 1).bj
      coulist(i).lx = coulist(i + 1).lx
      coulist(i).sj = coulist(i + 1).sj
    Next
    ReDim Preserve coulist(n)
    Open App.Path & "\cou.dat" For Output As #1
    For i = 1 To counlen - 1
      Write #1, coulist(i).sno, coulist(i).sname, coulist(i).tc, coulist(i).bj, coulist(i).lx, coulist(i).sj
    Next i
    Close #1
    List1.Clear
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    MsgBox ("已删除")
    counlen = n
    gg = 1
End Sub
Private Sub Command5_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text1.Text) = Trim(coulist(i).sno) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub Command6_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text2.Text) = Trim(coulist(i).sname) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub Command7_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text3.Text) = Trim(coulist(i).tc) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub Command8_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text4.Text) = Trim(coulist(i).bj) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub Command9_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text5.Text) = Trim(coulist(i).lx) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub Command10_Click()
  List1.Clear
  For i = 1 To counlen
    If Trim(Text6.Text) = Trim(coulist(i).sj) Then
      lxx = "课号：" + coulist(i).sno + "课名：" + coulist(i).sname + "--任课老师：" + coulist(i).tc + "--班级：" + coulist(i).bj + "--类型：" + coulist(i).lx + "--时间：" + coulist(i).sj
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub
Private Sub List1_Click()
  lxx = List1.List(List1.ListIndex)
  For i = 1 To counlen
    If coulist(i).sno = Mid(lxx, 4, 2) Then
      Text1.Text = coulist(i).sno
      Text2.Text = coulist(i).sname
      Text3.Text = coulist(i).tc
      Text4.Text = coulist(i).bj
      Text5.Text = coulist(i).lx
      Text6.Text = coulist(i).sj
      gg = i
    End If
  Next
End Sub

Private Sub Form_Load()
  Open App.Path & "\cou.dat" For Input As #1
  counlen = 0
  Do While Not EOF(1)
    Input #1, c1.sno, c1.sname, c1.tc, c1.bj, c1.lx, c1.sj
    counlen = counlen + 1
  Loop
  Close #1
  Print counlen
  ReDim coulist(counlen)
  Open App.Path & "\cou.dat" For Input As #1
  For i = 1 To counlen
    Input #1, coulist(i).sno, coulist(i).sname, coulist(i).tc, coulist(i).bj, coulist(i).lx, coulist(i).sj
  Next i
  Close #1
  gg = 1
End Sub
Private Sub Command3_Click()
  gg = gg - 1
  If gg = 0 Then gg = counlen
  If gg = counlen + 1 Then gg = 1
  Text1.Text = coulist(gg).sno
  Text2.Text = coulist(gg).sname
  Text3.Text = coulist(gg).tc
  Text4.Text = coulist(gg).bj
  Text5.Text = coulist(gg).lx
  Text6.Text = coulist(gg).sj
End Sub

Private Sub Command4_Click()
  gg = gg + 1
  If gg = 0 Then gg = counlen
  If gg = counlen + 1 Then gg = 1
  Text1.Text = coulist(gg).sno
  Text2.Text = coulist(gg).sname
  Text3.Text = coulist(gg).tc
  Text4.Text = coulist(gg).bj
  Text5.Text = coulist(gg).lx
  Text6.Text = coulist(gg).sj
End Sub
Private Sub Command2_Click()
  Unload Me
  Form2.Show
End Sub

