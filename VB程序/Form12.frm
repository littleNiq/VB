VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form12 
   BackColor       =   &H00FFC0C0&
   Caption         =   "É¾³ýÑ§ÉúÐÅÏ¢"
   ClientHeight    =   9015
   ClientLeft      =   6630
   ClientTop       =   2460
   ClientWidth     =   8460
   LinkTopic       =   "Form12"
   ScaleHeight     =   9015
   ScaleWidth      =   8460
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6960
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÍË³ö"
      Height          =   975
      Left            =   4440
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "É¾³ý"
      Height          =   975
      Left            =   1200
      TabIndex        =   6
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "°à¼¶"
      Height          =   180
      Left            =   720
      TabIndex        =   5
      Top             =   3240
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "ÐÕÃû"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ñ§ºÅ"
      Height          =   180
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   360
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   4440
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "Form12"
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
Private Sub Command2_Click()
    Kill picname
    n = stunlen - 1
    For i = gg To n
      stulist(i).sno = stulist(i + 1).sno
      stulist(i).sname = stulist(i + 1).sname
      stulist(i).bj = stulist(i + 1).bj
    Next
    ReDim Preserve stulist(n)
    Open App.Path & "\stu.dat" For Output As #1
    For i = 1 To stunlen - 1
      Write #1, stulist(i).sno, stulist(i).sname, stulist(i).bj
    Next i
    Close #1
    MsgBox ("ÒÑÉ¾³ý")
    stunlen = n
    gg = 1
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
  gg = 1
End Sub

Private Sub Command4_Click()
  gg = gg - 1
  If gg = 0 Then gg = stunlen
  If gg = stunlen + 1 Then gg = 1
  Text1.Text = stulist(gg).sno
  Text2.Text = stulist(gg).sname
  Text3.Text = stulist(gg).bj
  picname = App.Path & "\" & "\Í¼Æ¬\" & stulist(gg).sno & ".jpg"
  Image1.Picture = LoadPicture(picname)
End Sub

Private Sub Command5_Click()
  gg = gg + 1
  If gg = 0 Then gg = stunlen
  If gg = stunlen + 1 Then gg = 1
  Text1.Text = stulist(gg).sno
  Text2.Text = stulist(gg).sname
  Text3.Text = stulist(gg).bj
  picname = App.Path & "\Í¼Æ¬\" & stulist(gg).sno & ".jpg"
  Image1.Picture = LoadPicture(picname)
End Sub

