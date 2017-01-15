VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00FF80FF&
   Caption         =   "课表查询"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form20"
   ScaleHeight     =   9150
   ScaleWidth      =   9555
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   0
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   1
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   2
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   3
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   4
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   5
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   6
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   7
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   8
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   9
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   10
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   11
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   12
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   13
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   14
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   15
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   16
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   17
      Left            =   4200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   18
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   5760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   1335
      Index           =   19
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全部显示"
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   855
      Left            =   7200
      TabIndex        =   0
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程名"
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "星期一"
      Height          =   180
      Left            =   1680
      TabIndex        =   31
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "星期二"
      Height          =   180
      Left            =   3120
      TabIndex        =   30
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "星期三"
      Height          =   180
      Left            =   4560
      TabIndex        =   29
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "星期四"
      Height          =   180
      Left            =   6000
      TabIndex        =   28
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "星期五"
      Height          =   180
      Left            =   7560
      TabIndex        =   27
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "一"
      Height          =   180
      Left            =   840
      TabIndex        =   26
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "二"
      Height          =   180
      Left            =   840
      TabIndex        =   25
      Top             =   3360
      Width           =   180
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "三"
      Height          =   180
      Left            =   840
      TabIndex        =   24
      Top             =   4800
      Width           =   180
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "四"
      Height          =   180
      Left            =   840
      TabIndex        =   23
      Top             =   6000
      Width           =   180
   End
End
Attribute VB_Name = "Form20"
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
Dim s1 As stuxktype
Dim stuxklen As Integer
Dim stuxklist() As stuxktype
Dim c1 As stucoutype
Dim stucoulist() As stucoutype
Private Sub Combo1_Click()
  Dim i As Integer
  Call kbcle
  If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    For i = 1 To stuxklen
      If Trim(Combo1.List(Combo1.ListIndex)) = Trim(stucoulist(i).cname) Then
        Call kbxs(i)
        Exit For
      End If
    Next
  End If
  Combo1.Text = ""
End Sub
Private Sub Command1_Click()
  Dim i As Integer
  Call kbcle
  For i = 1 To stuxklen
    Call kbxs(i)
  Next
End Sub

Private Sub Command2_Click()
  Form4.Show
  Unload Me
End Sub

Private Sub Form_Load()
  Dim i As Integer
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
  For i = 1 To stuxklen
    Call kbxs(i)
  Next
End Sub

Private Sub kbxs(a As Integer)
  Dim h As Integer
  Dim l As Integer
  Dim sj As String
  Dim kcxi As String
  sj = stucoulist(a).dt
  kcxi = "课程名：" + stucoulist(a).cname + "--教师" + stucoulist(a).js
  h = Val(Mid(sj, 3, 1)) / 2
  l = Val(Mid(sj, 1, 1))
  Text1(1 + ((h - 1) * 5) / 2 - 1) = kcxi
End Sub

Private Sub kbcle()
  For j = 0 To 19
    Text1(j).Text = ""
  Next
End Sub

