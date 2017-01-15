VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form10 
   BackColor       =   &H00FFC0FF&
   Caption         =   "添加学生信息"
   ClientHeight    =   7365
   ClientLeft      =   7140
   ClientTop       =   2490
   ClientWidth     =   8475
   LinkTopic       =   "Form10"
   ScaleHeight     =   7365
   ScaleWidth      =   8475
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4680
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1080
      TabIndex        =   6
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   4860
      Left            =   0
      Picture         =   "Form10.frx":0000
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   8475
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   1200
      TabIndex        =   5
      Top             =   2160
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   1200
      TabIndex        =   4
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "学号"
      Height          =   180
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   360
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type stutype
  sno As String * 6
  sname As String * 6
  bj As String * 2
End Type
Dim s1 As stutype
Dim stulist() As stutype
Dim stunlen As Integer
Dim oldname As String
Dim picname As String
Dim hh As Boolean
Private Sub Command1_Click()
  For i = 1 To stunlen
    If stulist(i).sno = Trim(Text1.Text) Then
      hh = False
    Else
    End If
  Next
  If hh Then
    stunlen = stunlen + 1
    ReDim Preserve stulist(stunlen)
    stulist(stunlen).sno = Trim(Text1.Text)
    stulist(stunlen).sname = Trim(Text2.Text)
    stulist(stunlen).bj = Trim(Text3.Text)
    picname = App.Path & "\" & "图片\" & stulist(stunlen).sno & ".jpg"
    Name oldname As picname
    Open App.Path & "\stu.dat" For Output As #1
    For i = 1 To stunlen
      Write #1, stulist(i).sno, stulist(i).sname, stulist(i).bj
    Next i
    Close #1
    MsgBox ("已添加")
  Else
    MsgBox ("学号重名")
  End If
  hh = True
End Sub

Private Sub Command2_Click()
  Form2.Show
  Unload Me
End Sub

Private Sub Form_Load()
  Open App.Path & "\stu.dat" For Input As #1
  stunlen = 0
  Do While Not EOF(1)
    Input #1, s1.sno, s1.sname, s1.bj
    stunlen = stunlen + 1
  Loop
  Close #1
  Print stunlen
  ReDim stulist(stunlen)
  Open App.Path & "\stu.dat" For Input As #1
  For i = 1 To stunlen
    Input #1, stulist(i).sno, stulist(i).sname, stulist(i).bj
  Next i
  Close #1
  hh = True
  
End Sub


Private Sub Image1_DblClick()
  CommonDialog1.Action = 1
  If CommonDialog1.FileName <> "" Then
    oldname = CommonDialog1.FileName
    Image1.Picture = LoadPicture(oldname)
    Command1.Enabled = True
  Else
    a = MsgBox("必须添加照片", 49, "信息确认")
  End If
End Sub

