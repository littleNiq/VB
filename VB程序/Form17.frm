VERSION 5.00
Begin VB.Form Form17 
   BackColor       =   &H00FFFFC0&
   Caption         =   "课表"
   ClientHeight    =   8940
   ClientLeft      =   6030
   ClientTop       =   2295
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form17"
   ScaleHeight     =   8940
   ScaleWidth      =   10245
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      TabIndex        =   34
      Top             =   7680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全部显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1560
      TabIndex        =   21
      Top             =   480
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5280
      TabIndex        =   20
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   19
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   18
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   17
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   16
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   15
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   14
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   13
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   12
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   11
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   10
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   9
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   8
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   7
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   6
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   5
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   4
      Left            =   7200
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   3
      Left            =   5760
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   2
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   1
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "四"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   33
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   32
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   31
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   960
      TabIndex        =   30
      Top             =   2040
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "星期五"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7680
      TabIndex        =   29
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "星期四"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6120
      TabIndex        =   28
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "星期三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4680
      TabIndex        =   27
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "星期二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3240
      TabIndex        =   26
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "星期一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1800
      TabIndex        =   25
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   360
      TabIndex        =   23
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "请选择班级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4320
      TabIndex        =   22
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "Form17"
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
  Dim i As Integer
  Call kbcle
  If Trim(Combo1.List(Combo1.ListIndex)) <> "" Then
    For i = 1 To teacounlen
      If Trim(Combo1.List(Combo1.ListIndex)) = Trim(teacoulist(i).cname) Then
        Call kbxs(i)
        Exit For
      End If
    Next
  End If
  Combo1.Text = ""
  Combo2.Text = ""
End Sub

Private Sub Combo2_click()
  Dim i As Integer
  Call kbcle
  If Trim(Combo2.List(Combo2.ListIndex)) <> "" Then
    For i = 1 To teacounlen
      If Trim(Combo2.List(Combo2.ListIndex)) = Trim(teacoulist(i).bj) Then
        Call kbxs(i)
      End If
    Next
  End If
  Combo1.Text = ""
  Combo2.Text = ""
End Sub

Private Sub Command1_Click()
  Dim i As Integer
  Call kbcle
  For i = 1 To teacounlen
    Call kbxs(i)
  Next
End Sub

Private Sub Command2_Click()
  Form3.Show
  Unload Me
End Sub

Private Sub Form_Load()
   Dim strfilename As String
   Dim n As Integer
   Dim i As Integer
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
        Next j
        If n = 0 Then
          Combo2.AddItem teacoulist(i).bj
        End If
      Next i
      For i = 1 To teacounlen
        Call kbxs(i)
      Next i
    End If
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
  End If
End Sub

Private Sub kbxs(a As Integer)
  Dim h As Integer
  Dim l As Integer
  Dim sj As String
  Dim kcxi As String
  sj = teacoulist(a).dt
  kcxi = "课程名：" + teacoulist(a).cname + "--班级" + teacoulist(a).bj
  h = Val(Mid(sj, 3, 1)) / 2
  l = Val(Mid(sj, 1, 1))
  Text1((h - 1) * 5 + l - 1) = kcxi
End Sub

Private Sub kbcle()
  For j = 0 To 19
    Text1(j).Text = ""
  Next
End Sub
