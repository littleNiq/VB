VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form11 
   BackColor       =   &H00FFC0C0&
   Caption         =   "浏览修改学生信息"
   ClientHeight    =   7665
   ClientLeft      =   7500
   ClientTop       =   2115
   ClientWidth     =   7875
   LinkTopic       =   "Form11"
   ScaleHeight     =   7665
   ScaleWidth      =   7875
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   1560
      TabIndex        =   11
      Top             =   6600
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   840
      TabIndex        =   10
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   6360
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查询"
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "请输入班级"
      Height          =   180
      Left            =   3240
      TabIndex        =   9
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
Option Base 1
Private Type stutype
   sno As String * 6
   sname As String * 6
   bj As String * 2
End Type
Private Type zhongjie
   zj As String * 6
End Type
Dim si As stutype
Dim stulist() As stutype
Dim zhongj As zhongjie
Dim nlen As Integer
Dim picname As String
Dim oldname As String
Dim i As Integer
Dim weizhi As Integer
Dim pic As Boolean
Dim zhaopian As Boolean
Private Sub Command1_Click()
    List1.Clear
    For i = 1 To nlen
      If Trim(Text1.Text) = Trim(stulist(i).bj) Then
         lxx = "学号：" + stulist(i).sno + "--姓名：" + stulist(i).sname
         List1.List(n) = lxx
         n = n + 1
      End If
   Next
End Sub
Private Sub Command2_Click()
   If Len(Text2.Text) > 6 Then
     Print MsgBox("学号不超过六位数，请重新输入")
     Text2.Text = ""
     Exit Sub
   End If
   
   If zhaopian = True And pic = True Then
      zhongj.zj = Text2.Text
      Kill App.Path & "\图片\" & stulist(weizhi).sno & ".jpg"
      picname = App.Path & "\" & zhongj.zj & ".jpg"
      Name oldname As picname
   Else
      If zhaopian = True And pic = False Then
         Kill App.Path & "\图片\" & stulist(weizhi).sno & ".jpg"
         picname = App.Path & "\图片\" & stulist(weizhi).sno & ".jpg"
         Name oldname As picname
      Else
         If pic = True And zhaopian = False Then
            zhongj.zj = Text2.Text
            oldpicname = App.Path & "\图片\" & stulist(weizhi).sno & ".jpg"
            newpicname = App.Path & "\图片\" & zhongj.zj & ".jpg"
            Name oldpicname As newpicname
            Image1.Picture = LoadPicture(newpicname)
            pic = False
         End If
       End If
    End If
    stulist(weizhi).sno = Text2.Text
    stulist(weizhi).sname = Text3.Text
    stulist(weizhi).bj = Text4.Text
    Open App.Path & "\stu.dat" For Output As #1
    For k = 1 To nlen
         Write #1, stulist(k).sno, stulist(k).sname, stulist(k).bj
    Next
    Close #1
    Print MsgBox("信息已经更改")
    List1.Clear
    For i = 1 To nlen
         If Trim(Text1.Text) = Trim(stulist(i).bj) Then
            lxx = "学号：" + stulist(i).sno + "--姓名：" + stulist(i).sname
            List1.List(n) = lxx
            n = n + 1
         End If
    Next
End Sub

Private Sub Command3_Click()
  Unload Me
  Form2.Show
End Sub

Private Sub Form_Load()
   Open App.Path & "\stu.dat" For Input As #1
   nlen = 0
   Do While Not EOF(1)
      Input #1, si.sno, si.sname, si.bj
      nlen = nlen + 1
      usnlen = usnlen + 1
   Loop
   Close #1
   ReDim stulist(nlen)
   Open App.Path & "\stu.dat" For Input As #1
   For i = 1 To nlen
      Input #1, stulist(i).sno, stulist(i).sname, stulist(i).bj
   Next
   Close #1
   zhaopian = False
End Sub

Private Sub Image1_DblClick()
   CommonDialog2.Action = 1
   If CommonDialog2.FileName <> "" Then
      oldname = CommonDialog1.FileName
      Image1.Picture = LoadPicture(oldname)
      Command1.Enabled = True
      zhaopian = True
   Else
      a = MsgBox("必须添加照片", 49, "信息确认")
      Exit Sub
   End If
End Sub

Private Sub List1_Click()
   lxx = List1.List(List1.ListIndex)
   For i = 1 To nlen
      If stulist(i).sno = Mid(lxx, 4, 6) Then
         Text2.Text = stulist(i).sno
         Text3.Text = stulist(i).sname
         Text4.Text = stulist(i).bj
         Image1.Picture = LoadPicture(App.Path & "\图片\" & stulist(i).sno & ".jpg")
         weizhi = i
         Exit Sub
      End If
   Next
End Sub

Private Sub Text2_Change()
   pic = True
End Sub

