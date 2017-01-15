VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form11 
   Caption         =   "浏览修改学生信息"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form11"
   ScaleHeight     =   7665
   ScaleWidth      =   7875
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   840
      TabIndex        =   11
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "请输入班级"
      Height          =   180
      Left            =   3240
      TabIndex        =   10
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "班级"
      Height          =   180
      Left            =   600
      TabIndex        =   5
      Top             =   6960
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Left            =   600
      TabIndex        =   4
      Top             =   5640
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "学号"
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   360
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1935
   End
End
Attribute VB_Name = "Form11"
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
Dim gg As Integer
Dim oldname As String
Dim picname As String
Private Sub Command1_Click()
  List1.Clear
  For i = 1 To stunlen
    If Trim(Text5.Text) = Trim(stulist(i).bj) Then
      lxx = "学号：" + stulist(i).sno + "--姓名：" + stulist(i).sname
      List1.List(n) = lxx
      n = n + 1
    End If
  Next i
End Sub

Private Sub Command2_Click()
  stulist(gg).sno = Trim(Text1.Text)
  stulist(gg).sname = Trim(Text2.Text)
  stulist(gg).bj = Trim(Text3.Text)
  Open App.Path & "\stu.dat" For Output As #1
  For i = 1 To stunlen
    Write #1, stulist(i).sno, stulist(i).sname, stulist(i).bj
  Next i
  Close #1
  MsgBox ("已修改")
  Unload Me
  Form2.Show
End Sub

Private Sub Command3_Click()
  Unload Me
  Form2.Show
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
End Sub

Private Sub List1_Click()
  lxx = List1.List(List1.ListIndex)
  For i = 1 To stunlen
    If stulist(i).sno = Mid(lxx, 4, 6) Then
      Text1.Text = stulist(i).sno
      Text2.Text = stulist(i).sname
      Text3.Text = stulist(i).bj
      picname = App.Path & "\" & "\图片\" & stulist(stunlen).sno & ".jpg"
      Image1.Picture = LoadPicture(picname)
      gg = i
    End If
  Next i
End Sub
