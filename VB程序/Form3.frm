VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "教师管理"
   ClientHeight    =   7905
   ClientLeft      =   7125
   ClientTop       =   2835
   ClientWidth     =   7635
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   7635
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   975
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6675
      Left            =   0
      Picture         =   "Form3.frx":625D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7680
   End
   Begin VB.Menu mun1 
      Caption         =   "课程查询"
      Begin VB.Menu mun11 
         Caption         =   "开课信息查询"
      End
      Begin VB.Menu mun12 
         Caption         =   "课表查询"
      End
      Begin VB.Menu mun13 
         Caption         =   "学生名单"
      End
   End
   Begin VB.Menu mun2 
      Caption         =   "成绩管理"
      Begin VB.Menu mun21 
         Caption         =   "成绩录入"
      End
      Begin VB.Menu mun22 
         Caption         =   "成绩统计"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Form1.Show
  Unload Me
End Sub

Private Sub mun11_Click()
  Form16.Show
  Unload Me
End Sub

Private Sub mun12_Click()
  Form17.Show
  Unload Me
End Sub

Private Sub mun13_Click()
  Form18.Show
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
