VERSION 5.00
Begin VB.Form Form19 
   BackColor       =   &H00C0E0FF&
   Caption         =   "选修课程"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form19"
   ScaleHeight     =   7575
   ScaleWidth      =   8985
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   735
      Left            =   3360
      TabIndex        =   19
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "提交"
      Height          =   735
      Left            =   600
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消选修"
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选修"
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   7080
      TabIndex        =   7
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox Text5 
      Height          =   735
      Left            =   7080
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   735
      Left            =   7080
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   735
      Left            =   7080
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   7080
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7080
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2760
      Left            =   3480
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "时间"
      Height          =   180
      Left            =   6480
      TabIndex        =   18
      Top             =   6480
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "类型"
      Height          =   180
      Left            =   6480
      TabIndex        =   17
      Top             =   5280
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   6480
      TabIndex        =   16
      Top             =   4080
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "教师"
      Height          =   180
      Left            =   6480
      TabIndex        =   15
      Top             =   2880
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "课程名"
      Height          =   180
      Left            =   6360
      TabIndex        =   14
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "课号"
      Height          =   180
      Left            =   6480
      TabIndex        =   13
      Top             =   720
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "已选课程"
      Height          =   180
      Left            =   3840
      TabIndex        =   12
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "未选课程"
      Height          =   180
      Left            =   960
      TabIndex        =   11
      Top             =   960
      Width           =   720
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Private Type bjcoutype
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
Dim c1 As bjcoutype
Dim xk1 As stuxktype
Dim stucoulist() As stuxktype
Dim stuxkcoulist() As stuxktype
Dim bjcoulist() As bjcoutype
Dim bjcoulen As Integer
Private Sub Command1_Click()
  If List1.ListCount > 0 Then
    List2.AddItem List1.List(List1.ListIndex)
    List1.RemoveItem List1.ListIndex
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
  End If
End Sub

Private Sub Command2_Click()
  If List2.ListCount > 0 Then
    List1.AddItem List2.List(List2.ListIndex)
    List2.RemoveItem List2.ListIndex
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
  End If
End Sub
Private Sub Command3_Click()
  Dim cno As String
  Dim flag As Boolean
  Dim flag1 As Boolean
  Dim k As Integer
  Dim h As Integer
  Dim xkcoulen As Integer
  Open App.Path & "\xk.dat" For Input As #1
  xkcounlen = 0
  Do While Not EOF(1)
    Input #1, xk1.cno, xk1.sno, xk1.cj
    xkcoulen = xkcoulen + 1
  Loop
  Close #1
  If xkcoulen = 0 Then
    a = MsgBox("课程信息为空", 49, "信息确认")
  Else
    ReDim stuxkcoulist(xkcoulen)
    Open App.Path & "\xk.dat" For Input As #1
    i = 0
    Do While Not EOF(1)
      Input #1, xk1.cno, xk1.sno, xk1.cj
      i = i + 1
      stuxkcoulist(i).cno = xk1.cno
      stuxkcoulist(i).sno = xk1.sno
      stuxkcoulist(i).cj = xk1.cj
    Loop
    Close #1
  End If
  If List1.ListCount > 0 Then
    For i = 0 To List1.ListCount - 1
      flag1 = False
      For j = 1 To bjcoulen
        If Trim(List1.List(i)) = Trim(bjcoulist(j).cname) Then
          cno = bjcoulist(j).cno
        End If
      Next
      For k = 1 To xkcoulen
        If Trim(cno) = Trim(stuxkcoulist(k).cno) And Trim(student) = Trim(stuxkcoulist(k).sno) Then
          For h = k To xkcoulen - 1
            stuxkcoulist(h) = stuxkcoulist(h + 1)
          Next
          flag1 = True
        End If
      Next
      If flag1 = True Then
        xkcoulen = xkcoulen - 1
        ReDim Preserve stuxkcoulist(xkcoulen)
      End If
    Next
  End If
  If List2.ListCount > 0 Then
    For i = 0 To List2.ListCount - 1
      flag = False
      For j = 1 To bjcoulen
        If Trim(List2.List(i)) = Trim(bjcoulist(j).cname) Then
          cno = bjcoulist(j).cno
        End If
      Next
      For k = 1 To xkcoulen
        If Trim(cno) = Trim(stuxkcoulist(k).cno) And Trim(student) = Trim(stuxkcoulist(k).sno) Then
          flag = True
        End If
      Next
      If flag = False Then
        xkcoulen = xkcoulen + 1
        ReDim Preserve stuxkcoulist(xkcoulen)
        stuxkcoulist(xkcoulen).cno = cno
        stuxkcoulist(xkcoulen).sno = student
        stuxkcoulist(xkcoulen).cj = ""
      End If
    Next
  End If
  Open App.Path & "\xk.dat" For Output As #1
  For i = 1 To xkcoulen
    Write #1, stuxkcoulist(i).cno, stuxkcoulist(i).sno, stuxkcoulist(i).cj
  Next
  Close #1
  MsgBox ("已提交")
End Sub

Private Sub Command4_Click()
  Unload Me
  Form4.Show
End Sub

Private Sub Form_Load()
  Dim strfilename As String
  Dim n As Integer
  Dim bjno As String * 2
  Dim couyx As Boolean
  bjno = Mid(student, 3, 2)
  Open App.Path & "\xk.dat" For Input As #1
  Do While Not EOF(1)
    Input #1, xk1.cno, xk1.sno, xk1.cj
    If Trim(xk1.sno) = Trim(student) Then
      stucoulen = stucoulen + 1
    End If
  Loop
  Close #1
  ReDim stucoulist(stucoulen)
  Open App.Path & "\xk.dat" For Input As #1
  i = 0
  Do While Not EOF(1)
    Input #1, xk1.cno, xk1.sno, xk1.cj
    If Trim(xk1.sno) = Trim(student) Then
      i = i + 1
      stucoulist(i).cno = xk1.cno
      stucoulist(i).sno = xk1.sno
      stucoulist(i).cj = xk1.cj
    End If
  Loop
  Close #1
  Open App.Path & "\cou.dat" For Input As #1
  bjcoulen = 0
  Do While Not EOF(1)
    Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
    If Trim(c1.bj) = bjno Then
      bjcoulen = bjcoulen + 1
    End If
  Loop
  Close #1
  If bjcoulen = 0 Then
    Command1.Enabled = False
    Command2.Enabled = False
    a = MsgBox("课程信息为空", 49, "信息确认")
  Else
    ReDim bjcoulist(bjcoulen)
    Open App.Path & "\cou.dat" For Input As #1
    i = 0
    Do While Not EOF(1)
      Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
      If Trim(c1.bj) = bjno Then
        i = i + 1
        bjcoulist(i).cno = c1.cno
        bjcoulist(i).cname = c1.cname
        bjcoulist(i).js = c1.js
        bjcoulist(i).bj = c1.bj
        bjcoulist(i).lx = c1.lx
        bjcoulist(i).dt = c1.dt
      End If
    Loop
    Close #1
  End If
  For j = 1 To bjcoulen
    couyx = False
    For i = 1 To stucoulen
      If stucoulist(i).cno = bjcoulist(j).cno Then
        List2.AddItem bjcoulist(j).cname
        couyx = True
      End If
    Next
    If couyx = False Then
      List1.AddItem bjcoulist(j).cname
    End If
  Next
End Sub

Private Sub List1_Click()
  If Trim(List1.List(List1.ListIndex)) <> "" Then
    For i = 1 To bjcoulen
      If Trim(List1.List(List1.ListIndex)) = Trim(bjcoulist(i).cname) Then
        Text1.Text = bjcoulist(i).cno
        Text2.Text = bjcoulist(i).cname
        Text3.Text = bjcoulist(i).js
        Text4.Text = bjcoulist(i).bj
        Text5.Text = bjcoulist(i).lx
        Text6.Text = bjcoulist(i).dt
        Exit For
      End If
    Next
  End If
End Sub

Private Sub List2_Click()
  If Trim(List2.List(List2.ListIndex)) <> "" Then
    For i = 1 To bjcoulen
      If Trim(List2.List(List2.ListIndex)) = Trim(bjcoulist(i).cname) Then
        Text1.Text = bjcoulist(i).cno
        Text2.Text = bjcoulist(i).cname
        Text3.Text = bjcoulist(i).js
        Text4.Text = bjcoulist(i).bj
        Text5.Text = bjcoulist(i).lx
        Text6.Text = bjcoulist(i).dt
        Exit For
      End If
    Next
  End If
End Sub
