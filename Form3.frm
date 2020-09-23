VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "About"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3960
      Picture         =   "Form3.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Feedback"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Thank you for using this application - Robin McKay"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Delete files from a directory with just ONE click!"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Delete Files 1.0 - Written By Robin McKay"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next
Shell ("start mailto:ian@imckay.fsnet.co.uk")
End Sub
