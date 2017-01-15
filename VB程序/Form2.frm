VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "管理员"
   ClientHeight    =   6735
   ClientLeft      =   7380
   ClientTop       =   2205
   ClientWidth     =   8760
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6735
   ScaleWidth      =   8760
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5640
      Picture         =   "Form2.frx":7433
      TabIndex        =   0
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   6735
      Left            =   0
      Picture         =   "Form2.frx":19110
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8745
   End
   Begin VB.Menu mun1 
      Caption         =   "用户管理"
      Begin VB.Menu mun11 
         Caption         =   "用户注册"
      End
      Begin VB.Menu mun12 
         Caption         =   "用户修改"
      End
      Begin VB.Menu mun13 
         Caption         =   "用户删除"
      End
   End
   Begin VB.Menu mun3 
      Caption         =   "学生信息管理"
      Begin VB.Menu mun31 
         Caption         =   "添加学生信息"
      End
      Begin VB.Menu mun32 
         Caption         =   "浏览修改学生信息"
      End
      Begin VB.Menu mun33 
         Caption         =   "删除学生信息"
      End
   End
   Begin VB.Menu mun4 
      Caption         =   "课程信息管理"
      Begin VB.Menu mun41 
         Caption         =   "添加课程信息"
      End
      Begin VB.Menu mun42 
         Caption         =   "修改课程信息"
      End
      Begin VB.Menu mun43 
         Caption         =   "删除课程信息"
      End
   End
   Begin VB.Menu m22 
      Caption         =   "文件管理"
      Begin VB.Menu s 
         Caption         =   "文件备份"
         Index           =   1
      End
      Begin VB.Menu s1 
         Caption         =   "文件还原"
         Index           =   11
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub mun11_Click()
  Form5.Show
  Unload Me
End Sub

Private Sub mun12_Click()
  Form6.Show
  Unload Me
End Sub

Private Sub mun13_Click()
  Form7.Show
  Unload Me
End Sub

Private Sub mun21_Click()
  Form8.Show
  Unload Me
End Sub

Private Sub mun22_Click()
  Form9.Show
  Unload Me
End Sub

Private Sub mun31_Click()
  Form10.Show
  Unload Me
End Sub

Private Sub mun32_Click()
  Form11.Show
  Unload Me
End Sub

Private Sub mun33_Click()
  Form12.Show
  Unload Me
End Sub


Private Sub mun41_Click()
  Form13.Show
  Unload Me
End Sub

Private Sub mun42_Click()
  Form14.Show
  Unload Me
End Sub

Private Sub mun43_Click()
  Form15.Show
  Unload Me
End Sub

Private Sub s_Click(Index As Integer)
Form23.Show

End Sub

Private Sub s1_Click(Index As Integer)
Form22.Show
End Sub
