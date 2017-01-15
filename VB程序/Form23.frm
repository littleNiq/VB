VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form23"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10185
   LinkTopic       =   "Form23"
   ScaleHeight     =   6180
   ScaleWidth      =   10185
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List2 
      Height          =   2220
      ItemData        =   "Form23.frx":0000
      Left            =   2880
      List            =   "Form23.frx":0002
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   2220
      ItemData        =   "Form23.frx":0004
      Left            =   480
      List            =   "Form23.frx":0014
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   " 取消"
      Height          =   735
      Left            =   3240
      TabIndex        =   3
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   " 备份"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   " 需备份文件"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "文件选取"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MkDir App.Path & "\" & Format(Now, "yyyy-mm-dd")
For i = 0 To List1.ListCount - 1
  mystr = Right(List1.List(i), Len(List1.List(i)) - InStrRev(List1.List(i), "\"))
  FileCopy List1.List(i), App.Path & "\" & Format(Now, "yyyy-mm-dd") & "\" & mystr
Next i
a = MsgBox("已备份成功", 49, "信息")
Unload Me
End Sub
End Sub
Private Sub Command2_Click()
  Unload Me
End Sub
Private Sub File1_Click()
   List1.AddItem File1.FileName
End Sub
 

