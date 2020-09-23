VERSION 5.00
Begin VB.Form MsgForm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3015
      Left            =   0
      ScaleHeight     =   2955
      ScaleWidth      =   6990
      TabIndex        =   0
      Top             =   -120
      Width           =   7050
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   4440
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   1560
         Top             =   4440
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   5
         Left            =   480
         Picture         =   "MsgForm.frx":0000
         Top             =   3360
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   4
         Left            =   480
         Picture         =   "MsgForm.frx":2AC7
         Top             =   3360
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   3
         Left            =   360
         Picture         =   "MsgForm.frx":52C7
         Top             =   3360
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   2
         Left            =   960
         Picture         =   "MsgForm.frx":7891
         Top             =   7680
         Width           =   4500
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "MsgForm.frx":9AFF
         Top             =   3360
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   1
         Left            =   -3360
         Picture         =   "MsgForm.frx":A845
         Top             =   7680
         Width           =   4500
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   0
         Left            =   -4200
         Picture         =   "MsgForm.frx":C646
         Top             =   2400
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   0
         Left            =   1200
         Picture         =   "MsgForm.frx":D29B
         Top             =   720
         Visible         =   0   'False
         Width           =   4500
      End
   End
End
Attribute VB_Name = "MsgForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub Form_Activate()
    If Me.Tag = "About" Then
       Caption = "About..."
       
    End If
End Sub

Private Sub Form_Load()
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Resize()
 On Error Resume Next
    Picture1.Move 10, 10, Me.ScaleWidth - 10, Me.ScaleHeight - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub Timer2_Timer()
    Static FirstRun As Integer
    
    FirstRun = FirstRun + 1
    
    If FirstRun = 1 Then
      Dim imgLeft
      imgLeft = Image1(0).Left
    End If
    
    Image1(0).Move Image1(0).Left + 120, Image1(0).Top
    
    If Image1(0).Left > Timer2.Tag And Timer3.Tag <> "stop" Then
       Timer2.Enabled = False
       Timer3.Enabled = True
       Pause 10000
       
       Image1(0).Picture = Image1(1).Picture
       Timer2.Interval = 20
       Timer2.Enabled = True
       Timer2.Tag = 6800
       
       If Image1(0).Left > 6800 Then
          Image1(0).Left = imgLeft
       End If
    End If
    
End Sub

Private Sub Timer3_Timer()
     Static i As Integer
            
     imgTXT(0).Visible = True
     
     If i > 5 Then
       i = 0
       Timer3.Tag = "stop"
       Timer3.Enabled = False
       Pause 9000
       imgTXT(0).Picture = imgTXT(0).Picture
       imgTXT(0).Visible = False
       Exit Sub
     End If
     imgTXT(0).Picture = imgTXT(i).Picture
     
     i = i + 1
End Sub

Private Sub txtSendMsg_Change()

End Sub
