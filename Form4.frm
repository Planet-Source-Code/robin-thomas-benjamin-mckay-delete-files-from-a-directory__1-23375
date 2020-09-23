VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Folder File(s) contents"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Return To Selecting Folder"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Path = Form2.Dir1.Path
Label1.Caption = "Now viewing contents of -" + Form2.Dir1.Path
End Sub

