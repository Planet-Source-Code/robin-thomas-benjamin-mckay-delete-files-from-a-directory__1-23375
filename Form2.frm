VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Select Folder"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&View File Contents"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   3240
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Please double-click a folder to select it"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form1.Text1.Text = Dir1.Path
Unload Me
Form1.Command2.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form4.Show
End Sub
