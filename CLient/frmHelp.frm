VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   Caption         =   "Help...."
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7110
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RText1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   14737632
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmHelp.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   6435
      TabIndex        =   1
      Top             =   0
      Width           =   6495
   End
   Begin VB.Menu mnuGenHelp 
      Caption         =   "&General Help"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuSetUP 
         Caption         =   "&Set-Up"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File Help"
      Begin VB.Menu mnuBrowse 
         Caption         =   "&Browsing"
         Begin VB.Menu mnuSavePaths 
            Caption         =   "&Saving File Paths"
         End
         Begin VB.Menu mnuLoadPaths 
            Caption         =   "&Loading Saved Paths"
         End
         Begin VB.Menu mnuSearch 
            Caption         =   "S&earching"
         End
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAtts 
         Caption         =   "Changing File At&tributes"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQDrive 
         Caption         =   "&Querying a drive for it's files"
      End
      Begin VB.Menu mnuQDirectory 
         Caption         =   "Q&uerying a drive for it's directorys"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMove 
         Caption         =   "&Moving a file"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "E&diting a text file"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copying a file"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Deleting a file"
      End
      Begin VB.Menu mnuRenameFile 
         Caption         =   "Renaming a &file"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateDir 
         Caption         =   "C&reating a directory"
      End
      Begin VB.Menu mnuRenameDir 
         Caption         =   "Re&naming a directory"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remo&ving a directory"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit Help"
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Top = frmClient.Top + frmClient.Top + 350
    Left = frmClient.Left + frmClient.Left + 800
    Form_Resize
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Resize()
  On Error Resume Next
   With RText1
     .Width = Me.ScaleWidth - 15
     Picture1.Width = .Width
     .Height = Me.ScaleHeight - 15
     Picture1.Height = .Height
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub mnuConnect_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Connect"
End Sub

Private Sub mnuCopy_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Copy"
End Sub

Private Sub mnuDelete_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Delete"
End Sub

Private Sub mnuEdit_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Edit"
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLoadPaths_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Load_Saved_Paths"
End Sub

Private Sub mnuMove_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Move"
End Sub

Private Sub mnuQDirectory_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "QueryD"
End Sub

Private Sub mnuQDrive_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Query"
End Sub

Private Sub mnuSavePaths_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Browse_Save"
End Sub

Private Sub mnuSearch_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Search"
End Sub

Private Sub mnuSetUP_Click()
    RText1.Text = ""
    RText1.Visible = False
    LoadHelpFile "Setup"
End Sub
