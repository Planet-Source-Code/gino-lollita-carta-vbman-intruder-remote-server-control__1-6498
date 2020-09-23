VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search..."
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3855
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   280
         Left            =   2640
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   280
         Left            =   2640
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " File name:"
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdCover 
      Caption         =   "C&over"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy to Clipboard"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1740
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Returned"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Results"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   525
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   Me.Height = 1590
   Unload Me
End Sub

Private Sub cmdCopy_Click()
       
      ' copy to clipboard
      Clipboard.Clear
      Clipboard.SetText List1.List(List1.ListIndex)
    
End Sub

Private Sub cmdCover_Click()
   Me.Height = 1590
End Sub

Private Sub cmdSearch_Click()
   Dim InValid As Integer
   On Error GoTo NotConnected
   List1.Clear
    
   If Text1 = "" Then
      InValid = 1
   End If
   
   If Text2 = "" Then
      InValid = 2
   End If
   
   If Text2 = "" And Text1 = "" Then
      InValid = 3
   End If
   
   If InValid = 1 Then
      MsgBox "Please enter a File name to search for.", , ""
      Exit Sub
   End If
   
   If InValid = 2 Then
      MsgBox "Please enter a Drive to search on.", , ""
      Exit Sub

   End If
   
   If InValid = 3 Then
      MsgBox "Please enter both a Drive and File name to search.", , ""
      Exit Sub
   End If
   
   
   ' uncover the list display
   frmSearch.Height = 4425
   
   Text2 = Left(Text2, 1) & ":\" 'Command  'Drive         'Filename
   SendData "Search," & Text2 & "|" & Text1
   
   Logit Search
   Label4 = "Searching Users Hard Drive...."
     
NotConnected:
   MsgBox "You are not connected!!", vbCritical, "Not Connected"
End Sub






Private Sub Form_Load()
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 1 Then
        List1.ToolTipText = List1.List(List1.ListIndex)
     End If
End Sub
