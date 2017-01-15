VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H00FFC0C0&
   Caption         =   "学生名单"
   ClientHeight    =   8070
   ClientLeft      =   7320
   ClientTop       =   3045
   ClientWidth     =   8385
   LinkTopic       =   "Form18"
   ScaleHeight     =   8070
   ScaleWidth      =   8385
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   735
      Left            =   4200
      TabIndex        =   14
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   735
      Left            =   1080
      TabIndex        =   13
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   975
      Left            =   5760
      TabIndex        =   12
      Top             =   6840
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   9
      Left            =   4080
      TabIndex        =   11
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   8
      Left            =   4080
      TabIndex        =   10
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   7
      Left            =   4080
      TabIndex        =   9
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   6
      Left            =   4080
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   5
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   4
      Left            =   840
      TabIndex        =   6
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   900
   End
End
Attribute VB_Name = "Form18"
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

Dim xk1 As teaxktype
Dim stu1 As stutype
Dim teaxklist() As xkstutype
Dim teaxklen As Integer

Dim curpge As Integer
Dim pgecount As Integer
Dim pgemod As Integer


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
       Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname + "--" + teaxklist((curpge - 1) * 10 + i).cj
     Next
     Command2.Enabled = False
     Command3.Enabled = False
   Else
     For i = 1 To 10
       Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname + "--" + teaxklist((curpge - 1) * 10 + i).cj
     Next
     Command3.Enabled = True
   End If
End Sub

Private Sub Command1_Click()
  Unload Me
  Form3.Show
End Sub

Private Sub Command2_Click()
  If curpge > 1 Then
    Call teaxkclea
    curpge = curpge - 1
    For i = 1 To 10
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname + "--" + teaxklist((curpge - 1) * 10 + i).cj
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
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname + "--" + teaxklist((curpge - 1) * 10 + i).cj
    Next
    Command2.Enabled = True
  ElseIf curpge = pgecount - 1 Then
    Call teaxkclea
    curpge = curpge + 1
    For i = 1 To pgemod
      Text1(i - 1).Text = teaxklist((curpge - 1) * 10 + i).sno + "--" + teaxklist((curpge - 1) * 10 + i).sname + "--" + teaxklist((curpge - 1) * 10 + i).cj
    Next
    Command2.Enabled = True
    Command3.Enabled = False
  End If
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

Private Sub teaxkclea()
  For i = 0 To 9
    Text1(i).Text = ""
  Next
End Sub
