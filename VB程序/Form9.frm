VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form9 
   Caption         =   "成绩统计"
   ClientHeight    =   6885
   ClientLeft      =   3600
   ClientTop       =   2190
   ClientWidth     =   15435
   LinkTopic       =   "Form9"
   ScaleHeight     =   6885
   ScaleWidth      =   15435
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   975
      Left            =   10800
      TabIndex        =   24
      Top             =   480
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   9960
      ScaleHeight     =   3795
      ScaleWidth      =   4755
      TabIndex        =   18
      Top             =   2160
      Width           =   4815
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3735
      Left            =   4560
      OleObjectBlob   =   "Form9.frx":0000
      TabIndex        =   17
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "成绩分析统计"
      Height          =   4455
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   3855
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   3600
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1080
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Label12"
         Height          =   180
         Left            =   3000
         TabIndex        =   23
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   180
         Left            =   3000
         TabIndex        =   22
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   180
         Left            =   3000
         TabIndex        =   21
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
         Height          =   180
         Left            =   3000
         TabIndex        =   20
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Label8"
         Height          =   180
         Left            =   3000
         TabIndex        =   19
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "总人数"
         Height          =   180
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   540
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2520
         Top             =   3120
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "不及格数"
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   720
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2520
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "及格人数"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Width           =   720
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2520
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "中等人数"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   720
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2520
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "良好人数"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   720
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2520
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "优秀人数"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "饼状图"
      Height          =   375
      Left            =   7680
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "柱状图"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择课程名"
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1080
   End
End
Attribute VB_Name = "Form9"
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

Dim y As Integer
Dim l As Integer
Dim z As Integer
Dim j As Integer
Dim b As Integer
Dim y1 As Currency
Dim l1 As Currency
Dim z1 As Currency
Dim j1 As Currency
Dim b1 As Currency
Private Sub Combo1_Click()
  y = 0: l = 0: z = 0: j = 0: b = 0
  y1 = 0: l1 = 0: z1 = 0: j1 = 0: b1 = 0
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
   For i = 1 To teaxklen
     If teaxklist(i).cj < 100 And teaxklist(i).cj >= 90 Then
       y = y + 1
     ElseIf teaxklist(i).cj < 90 And teaxklist(i).cj >= 80 Then
       l = l + 1
     ElseIf teaxklist(i).cj < 80 And teaxklist(i).cj >= 70 Then
       z = z + 1
     ElseIf teaxklist(i).cj < 70 And teaxklist(i).cj >= 60 Then
       j = j + 1
     Else
       b = b + 1
     End If
   Next
   y1 = (y / teaxklen)
   l1 = (l / teaxklen)
   z1 = (z / teaxklen)
   j1 = (j / teaxklen)
   b1 = (b / teaxklen)
   Label8.Caption = Str(y1 * 100) & "%"
   Label9.Caption = Str(l1 * 100) & "%"
   Label10.Caption = Str(z1 * 100) & "%"
   Label11.Caption = Str(j1 * 100) & "%"
   Label12.Caption = Str(b1 * 100) & "%"
   Text1.Text = y
   Text2.Text = l
   Text3.Text = z
   Text4.Text = j
   Text5.Text = b
   Text6.Text = teaxklen
End Sub

Private Sub Command1_Click()
  Dim arrvalue(5, 2)
  arrvalue(1, 1) = "优秀"
  arrvalue(1, 2) = y
  arrvalue(2, 1) = "良好"
  arrvalue(2, 2) = l
  arrvalue(3, 1) = "中等"
  arrvalue(3, 2) = z
  arrvalue(4, 1) = "及格"
  arrvalue(4, 2) = j
  arrvalue(5, 1) = "不及格"
  arrvalue(5, 2) = b
  MSChart1.ChartData = arrvalue
End Sub

Private Sub Command2_Click()
  Call ht(y1, l1, z1, j1, b1)
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
    End If
  Else
    a = MsgBox("课程信息为空", 49, "信息确认")
  End If
  Label8.Caption = ""
  Label9.Caption = ""
  Label10.Caption = ""
  Label11.Caption = ""
  Label12.Caption = ""
End Sub
Private Sub ht(a As Currency, b As Currency, c As Currency, d As Currency, e As Currency)
  Picture1.Cls
  Const pi = 3.14159
  Picture1.FillStyle = 0
  If a <> 0 Then
    Picture1.FillColor = vbRed
    Picture1.Circle (2000, 1500), 1000, vbRed, -2 * pi, -2 * pi * a
  End If
  If b <> 0 Then
    Picture1.FillColor = vbGreen
    If a = 0 Then
      Picture1.Circle (2000, 1500), 1000, vbGreen, -2 * pi, -2 * pi * (a + b)
    Else
      Picture1.Circle (2000, 1500), 1000, vbGreen, -2 * pi * a, -2 * pi * (a + b)
    End If
  End If
  If c <> 0 Then
    Picture1.FillColor = vbBlue
    If a = 0 And b = 0 Then
      Picture1.Circle (2000, 1500), 1000, vbBlue, -2 * pi, -2 * pi * (a + b + c)
    Else
      Picture1.Circle (2000, 1500), 1000, vbBlue, -2 * pi * (a + b), -2 * pi * (a + b + c)
    End If
  End If
  If d <> 0 Then
    Picture1.FillColor = vbYellow
    If a = 0 And b = 0 And c = 0 Then
      Picture1.Circle (2000, 1500), 1000, vbYellow, -2 * pi, -2 * pi * (a + b + c + d)
    Else
      Picture1.Circle (2000, 1500), 1000, vbYellow, -2 * pi * (a + b + c), -2 * pi * (a + b + c + d)
    End If
  End If
  If e <> 0 Then
    Picture1.FillColor = vbBlack
    If a = 0 And b = 0 And c = 0 And d = 0 Then
      Picture1.Circle (2000, 1500), 1000, vbBlack, -2 * pi, -2 * pi
    Else
      Picture1.Circle (2000, 1500), 1000, vbBlack, -2 * pi * (a + b + c + d), -2 * pi
    End If
  End If
End Sub
  
 
