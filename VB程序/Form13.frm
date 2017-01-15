VERSION 5.00
Begin VB.Form Form13 
   BackColor       =   &H00FF8080&
   Caption         =   "添加课程信息"
   ClientHeight    =   7020
   ClientLeft      =   7380
   ClientTop       =   2820
   ClientWidth     =   7770
   LinkTopic       =   "Form13"
   ScaleHeight     =   7020
   ScaleWidth      =   7770
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   13
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   12
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   4440
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   4440
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   3930
      Left            =   0
      Picture         =   "Form13.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   7770
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "时间"
      Height          =   180
      Left            =   3840
      TabIndex        =   11
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "类型"
      Height          =   180
      Left            =   3840
      TabIndex        =   10
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   3840
      TabIndex        =   9
      Top             =   600
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "任课老师"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "课名"
      Height          =   180
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "课号"
      Height          =   180
      Left            =   600
      TabIndex        =   6
      Top             =   600
      Width           =   360
   End
End
Attribute VB_Name = "Form13"
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
Dim hh As Boolean

Private Sub Command1_Click()
  For i = 1 To counlen
    If coulist(i).sno = Trim(Text1.Text) Then
      hh = False
    Else
    End If
  Next
  If hh Then
      counlen = counlen + 1
      ReDim Preserve coulist(counlen)
      coulist(counlen).sno = Trim(Text1.Text)
      coulist(counlen).sname = Trim(Text2.Text)
      coulist(counlen).tc = Trim(Text3.Text)
      coulist(counlen).bj = Trim(Text4.Text)
      coulist(counlen).lx = Trim(Text5.Text)
      coulist(counlen).sj = Trim(Text6.Text)
      Open App.Path & "\cou.dat" For Output As #1
      For i = 1 To counlen
        Write #1, coulist(i).sno, coulist(i).sname, coulist(i).tc, coulist(i).bj, coulist(i).lx, coulist(i).sj
      Next i
      Close #1
      MsgBox ("已添加")
  Else
     MsgBox ("课号重名")
  End If
  hh = True
  
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
  hh = True
  
End Sub
Private Sub Command2_Click()
  Unload Me
  Form2.Show
End Sub

