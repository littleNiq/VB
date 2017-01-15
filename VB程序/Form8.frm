VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00FFC0FF&
   Caption         =   "成绩录入"
   ClientHeight    =   6750
   ClientLeft      =   7455
   ClientTop       =   2055
   ClientWidth     =   9135
   LinkTopic       =   "Form8"
   ScaleHeight     =   6750
   ScaleWidth      =   9135
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   735
      Left            =   4560
      TabIndex        =   25
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   735
      Left            =   3960
      TabIndex        =   24
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   735
      Left            =   1200
      TabIndex        =   23
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   9
      Left            =   5880
      TabIndex        =   22
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   8
      Left            =   5880
      TabIndex        =   21
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   7
      Left            =   5880
      TabIndex        =   20
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   6
      Left            =   5880
      TabIndex        =   19
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   5
      Left            =   5880
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   17
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   3
      Left            =   2640
      TabIndex        =   16
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   15
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   9
      Left            =   3840
      TabIndex        =   12
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   8
      Left            =   3840
      TabIndex        =   11
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   6
      Left            =   3840
      TabIndex        =   9
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   5
      Left            =   3840
      TabIndex        =   8
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   4
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   3
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   5
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   1
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "提交"
      Height          =   735
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   0
      Picture         =   "Form8.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程名"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type teacoutype
  cno As String * 2
  cname As String * 6
  js As String * 6
  bj As String * 2
  lx As String * 1
  dt As String * 3
End Type
Private Type teaxktype
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
Dim c1 As teacoutype
Dim teacoulist() As teacoutype
Dim teacounlen As Integer
Dim teacno As String
Dim teasno As String
Dim all1 As teaxktype
Dim xk1 As teaxktype
Dim stu1 As stutype
Dim teaxklist() As xkstutype
Dim teaxklen As Integer
Dim allxklist() As teaxktype
Dim curpge As Integer
Dim pgecount As Integer
Dim pgemod As Integer
Dim allxknlen As Integer

Private Sub teaxkclea()
  For i = 0 To 9
    Text1(i).Text = ""
  Next
  For i = 0 To 9
    Text2(i).Text = ""
  Next
End Sub

Private Sub Combo1_Click()
  Call teaxkclea
  If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    For i = 1 To teacounlen
      If Trim(Combo1.List(Combo1.ListIndex)) = Trim(teacoulist(i).cname) Then
        teacno = teacoulist(i).cno
        Exit For
      End If
    Next
  End If
  Open App.Path & "\xk.dat" For Input As #1
    teaxklen = 0
    Do While Not EOF(1)
      Input #1, xk1.cno, xk1.sno, xk1.cj
      If Trim(xk1.cno) = Trim(teacno) Then
        teaxklen = teaxklen + 1
      End If
    Loop
    Close #1
    If teaxklen = 0 Then
      Command2.Enabled = False
      Command3.Enabled = False
      a = MsgBox("无人选修该课程", 49, "信息确认")
    Else
      ReDim teaxklist(teaxklen)
      i = 0
      Open App.Path & "\xk.dat" For Input As #1
      Do While Not EOF(1)
        Input #1, xk1.cno, xk1.sno, xk1.cj
        If Trim(xk1.cno) = Trim(teacno) Then
          teasno = xk1.sno
          Open App.Path & "\stu1.dat" For Input As #2
          Do While Not EOF(2)
            Input #2, stu1.sno, stu1.sname, stu1.bj
            If Trim(stu1.sno) = Trim(teasno) Then
              i = i + 1
              teaxklist(i).sno = stu1.sno
              teaxklist(i).sname = stu1.sname
              teaxklist(i).cj = xk1.cj
            End If
          Loop
          Close #2
       End If
     Loop
     Close #1
   End If
   pgecount = (teaxklen - 1) \ 10 + 1
   pgemod = teaxklen Mod 10
   curpge = 1
   If pgecount <= 1 Then
     For i = 1 To pgemod
       Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname
       Text2(i - 1).Text = teaxklist((curpge - 1) * 10 + i).cj
     Next
     Command2.Enabled = False
     Command3.Enabled = False
   Else
     For i = 1 To 10
       Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname
       Text2(i - 1).Text = teaxklist((curpge - 1) * 10 + i).cj
     Next
     Command3.Enabled = True
   End If
End Sub

Private Sub Command1_Click()
  For k = 1 To pgemod
     j = (curpge - 1) * 10 + k
     teaxklist(j).cj = Text2(k - 1)
  Next
  Open App.Path & "\xk.dat" For Input As #1
    allxknlen = 0
    Do While Not EOF(1)
      Input #1, all1.cno, all1.sno, all1.cj
        allxknlen = allxknlen + 1
    Loop
  Close #1
  ReDim allxklist(allxknlen)
  Open App.Path & "\xk.dat" For Input As #1
  i = 0
  Do While Not EOF(1)
    i = i + 1
    Input #1, allxklist(i).cno, allxklist(i).sno, allxklist(i).cj
  Loop
  Close #1
  For k = 1 To allxknlen
    For j = 1 To teaxklen
      If Trim(teaxklist(j).sno) = Trim(allxklist(k).sno) And Trim(teacno) = Trim(allxklist(k).cno) Then
        allxklist(k).cj = teaxklist(j).cj
      End If
    Next
  Next
  Open App.Path & "\xk.dat" For Output As #1
  For k = 1 To allxknlen
    Write #1, allxklist(k).cno, allxklist(k).sno, allxklist(k).cj
  Next
  Close #1
  MsgBox ("已提交")
End Sub

Private Sub Command4_Click()
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
    Do While Not EOF(1)
    Input #1, c1.cno, c1.cname, c1.js, c1.bj, c1.lx, c1.dt
    If Trim(c1.js) = Trim(teacher) Then
      teacounlen = teacounlen + 1
    End If
    Loop
    Close #1
    If teacounlen = 0 Then
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
      Command2.Enabled = False
      For i = 1 To teacounlen
        Combo1.AddItem teacoulist(i).cname
      Next
    End If
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
    Command2.Enabled = False
    Command3.Enabled = False
  End If
End Sub
Private Sub Command2_Click()
  If curpge > 1 Then
    Call teaxkclea
    curpge = curpge - 1
    For i = 1 To 10
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname
      Text2(i - 1).Text = teaxklist((curpge - 1) * 10 + i).cj
    Next
    Command3.Enabled = True
  Else
    Command2.Enabled = False
  End If
End Sub

Private Sub Command3_Click()
  If curpge < pgecount - 1 Then
    Call teaxkclea
    curpge = curpge + 1
    For i = 1 To 10
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname
      Text2(i - 1).Text = teaxklist((curpge - 1) * 10 + i).cj
    Next
    Command2.Enabled = True
  ElseIf curpge = pgecount - 1 Then
    Call teaxkclea
    curpge = curpge + 1
    For i = 1 To pgemod
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname
      Text2(i - 1).Text = teaxklist((curpge - 1) * 10 + i).cj
    Next
    Command2.Enabled = True
    Command3.Enabled = False
  End If
End Sub

