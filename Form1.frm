VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Files"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3900
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "delete"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "mail"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "exit"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&About Waste Disposer 2001"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete ALL Files"
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":086E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "/*.*"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a directory from which you wish to delete all files. This will permanently erase all files from that directory. "
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu SelectFolder 
         Caption         =   "&Select Folder..."
      End
      Begin VB.Menu border1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu HelpTopics 
         Caption         =   "&Help Topics"
      End
      Begin VB.Menu border2 
         Caption         =   "-"
      End
      Begin VB.Menu STFC 
         Caption         =   "&Support The Freeware Cause..."
      End
      Begin VB.Menu ReleaseNotes 
         Caption         =   "&Release Notes"
      End
      Begin VB.Menu border3 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
Form3.Show
End Sub

Private Sub Command1_Click()
Form2.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Integer
If Text1 = "" Then
    MsgBox "A directory has NOT been specified"
Else
    i = MsgBox("Are you sure?", vbYesNo + vbQuestion, "Confirm")
        Select Case i
            Case vbYes
                Kill Text1.Text + Label2.Caption
            Case vbNo
                MsgBox "File(s) in directory NOT deleted"
        End Select
End If
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Visible = False
End Sub

Private Sub HelpTopics_Click()
MsgBox "Select a folder you'd like to delete files from by going to the button that has the three dots." + vbCrLf + "Then, select a folder you wish to delete files from." + vbCrLf + "Click delete files and all files, NOT folders, will be deleted from that folder." + vbCrLf + vbCrLf + "Please note that files which are deleted do NOT go to the recycle bin but are permanently deleted from the directory.", vbInformation, "Help Topics"
End Sub

Private Sub ReleaseNotes_Click()
MsgBox "This is a BETA edition of Delete Files." + vbCrLf + vbCrLf + "If you can suggest any improvements to the program or to the source code, please let me know so that I can add this to a new edition." + vbCrLf + vbCrLf + "Thank you again for your interest and don't forget to support the freeware cause.", vbInformation, "Release Notes"
End Sub

Private Sub SelectFolder_Click()
Form2.Show
End Sub

Private Sub STFC_Click()
On Error Resume Next
Shell ("START HTTP://WWW.EXPAGE.COM/FREEWARECAUSE")
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then Command2.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "delete"
        Command2_Click
    Case "mail"
        Shell ("start mailto:ian@imckay.fsnet.co.uk")
    Case "exit"
        Unload Me
End Select
End Sub
