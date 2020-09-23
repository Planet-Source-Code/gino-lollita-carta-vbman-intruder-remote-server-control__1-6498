VERSION 5.00
Begin VB.Form frmDispCap 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2760
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "frmDispCap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Secs As Integer
Dim DnTime As Long


Private Sub Form_Load()
   ' set the specified Down Time
    'Timer1.Enabled = True
    DnTime = HowLong
    Secs = 0
End Sub

Private Sub Timer1_Timer()
     On Error Resume Next
     Secs = Secs + 1
     
     If Secs > DnTime Then
       Me.Visible = False
       ' clear the pic
       Set Picture1 = LoadPicture()
       ' clear the clipboard
       Clipboard.Clear
       Unload Me
       ' delete the screen capture
       Kill CAPTURE
       ' alert admin lockup done
       SendData "LockUp_Done,"
       'Timer1.Enabled = False
     End If
     
End Sub
