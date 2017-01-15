VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "登录界面"
   ClientHeight    =   8700
   ClientLeft      =   3450
   ClientTop       =   1185
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   8700
   ScaleWidth      =   15795
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   12840
      Top             =   6120
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "换一个吧"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Tag             =   "&H8000000D&"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2520
      TabIndex        =   15
      Top             =   5880
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":1A90
      Left            =   5040
      List            =   "Form1.frx":1AA0
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   2880
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   1920
      TabIndex        =   9
      Text            =   " "
      Top             =   1440
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "点我退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "用户登陆"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "123"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   2325
      Left            =   12360
      Picture         =   "Form1.frx":1ABC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2925
   End
   Begin VB.Shape Shape7 
      Height          =   855
      Left            =   2160
      Top             =   7080
      Width           =   735
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   15240
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   15240
      Top             =   2760
      Width           =   375
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      BorderStyle     =   2  'Dash
      FillColor       =   &H00FFFF80&
      Height          =   615
      Left            =   360
      Top             =   6600
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      Height          =   375
      Left            =   14880
      Top             =   7800
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF80FF&
      FillColor       =   &H00FF80FF&
      Height          =   375
      Left            =   480
      Top             =   3240
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   375
      Left            =   14760
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label Label9 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "验证码"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF80&
      Caption         =   " 请选择"
      Height          =   495
      Left            =   4680
      TabIndex        =   13
      Top             =   4320
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0FF&
      Caption         =   "请输入密码"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "请输入ID"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "欢迎进入吉林师大教务网络管理系统"
      BeginProperty Font 
         Name            =   "方正姚体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   13935
   End
   Begin VB.Image Image1 
      Height          =   8835
      Left            =   0
      Picture         =   "Form1.frx":3B52
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15840
   End
   Begin VB.Label Label3 
      Caption         =   " "
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "密码"
      Height          =   180
      Left            =   2400
      TabIndex        =   4
      Top             =   3720
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   2280
      TabIndex        =   2
      Top             =   2400
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text4.Text = Label9.Caption Then
Dim yh As Boolean
    yh = True
    For i = 1 To usnlen
        If Trim(Text1.Text) = Trim(uslist(i).yhm) Then
          yh = False
          If Trim(Text2.Text) = Trim(uslist(i).mi) Then
            If uslist(i).qx = "1" Then
              administrator = uslist(i).yhm
              Form2.Show
              Unload Me
            ElseIf uslist(i).qx = "2" Then
              teacher = uslist(i).yhm
              Form3.Show
              Unload Me
            Else
              student = uslist(i).yhm
              Form4.Show
              Unload Me
            End If
            Exit For
          Else
            If MsgBox("密码不正确", 49, "信息确认") = vbOK Then '提示框，信息不正确
              Text2.Text = ""
            Else
              Unload Me
            End If
          End If
        End If
   Next i
   If yh Then
     a = MsgBox("用户不存在", 49, "信息确认")
   End If
   Else
   a = MsgBox("输入验证码错误", "19", "信息错误")
   Print a
End If
End Sub
Private Sub Command2_Click()
 Randomize
Label9.Caption = Int(Rnd * 9000 + 1000)
End Sub
Private Sub Command3_Click()
  Text2.Text = ""
  Text2.SetFocus
  Text1.Text = ""
  Text1.SetFocus
End Sub

Private Sub Command4_Click()
  End
End Sub
Private Sub Form_Load()
  Label9.Caption = Int(Rnd * 9000 + 1000)
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

Private Sub Text1_click()
Text1.Text = Combo1.Text
End Sub
Private Sub Timer1_Timer()
Text3.Text = "现在时间" & Time
End Sub
Private Sub Timer2_Timer()
If Shape5.Left < 600 Then
  Shape5.Left = 15000
End If
If Shape6.Left < 600 Then
  Shape6.Left = 15000
End If

If Shape1.Left < 600 Then
  Shape1.Left = 15000
End If
If Shape3.Left < 600 Then
  Shape3.Left = 15000
End If
If Shape2.Left > 16000 Then
  Shape2.Left = 500
End If
If Shape4.Left > 16000 Then
  Shape4.Left = 500
End If
Shape5.Left = Shape5.Left - 150
 Shape6.Left = Shape6.Left - 150
 Shape1.Left = Shape1.Left - 150
 Shape2.Left = Shape2.Left + 150
 Shape3.Left = Shape3.Left - 150
 Shape4.Left = Shape4.Left + 150
End Sub
