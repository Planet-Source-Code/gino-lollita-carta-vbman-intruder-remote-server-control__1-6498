VERSION 5.00
Begin VB.Form frmMsgType 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Message Type"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5805
   Icon            =   "frmMsgType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Select Icon"
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   5535
      Begin VB.OptionButton optIcon 
         BackColor       =   &H00404040&
         Caption         =   "Confirm"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H00404040&
         Caption         =   "Critical"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton optIcon 
         BackColor       =   &H00404040&
         Caption         =   "Exclamation"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   4320
         Picture         =   "frmMsgType.frx":08CA
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2520
         Picture         =   "frmMsgType.frx":0D0C
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   600
         Picture         =   "frmMsgType.frx":114E
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdExChat 
      Caption         =   "Exit Chat"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Text            =   " "
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   6
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox txtSendMsg 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000C000&
      Height          =   825
      Left            =   2520
      TabIndex        =   5
      Text            =   " "
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Message Type"
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton Option4 
         BackColor       =   &H00404040&
         Caption         =   "Moving OK"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00404040&
         Caption         =   "Chat Dialog"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00404040&
         Caption         =   "Yes\No"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00404040&
         Caption         =   "Simple OK"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Message"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Caption"
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frmMsgType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IconChoice As String
Public called As Integer

Private Sub cmdExChat_Click()
    Option3.Value = False
    Option3_Click
    called = 0
End Sub

Private Sub cmdOK_Click()
   Option4.Value = False
   Option4_Click
End Sub

Private Sub cmdSend_Click()
    On Error GoTo NotConnected
   
    
    
    If Option1.Value = True Then
      SendData "Msg," & txtSendMsg
    ElseIf Option2.Value = True Then  ' Yes\No Dialog
      SendData "YN_Msg," & txtSendMsg & ";" & Text1
    ElseIf Option3.Value = True Then
      called = called + 1
      
      If called = 1 Then
        SendData "Chat," & txtSendMsg
      End If
      
      SendData "In_Msg," & txtSendMsg
      Text2.Text = Text2 & vbCrLf & UCase(txtSendMsg)
      ScrollChat txtSendMsg, frmMsgType.Text2.Text
      txtSendMsg = ""
    ElseIf Option4.Value = True Then
      SendData "Moving_Dialog," & txtSendMsg & ";" & Text1 & ":" & IconChoice
    End If
    
    Logit MsgSent
    
    Exit Sub
    
NotConnected:
   MsgBox "You are not connected!!", vbCritical, "Not Connected"
End Sub


Private Sub cmdDone_Click()
    txtSendMsg.Visible = False
    cmdSend.Visible = False
    cmdDone.Visible = False
    called = 0
    Unload Me
End Sub

Private Sub Form_Load()
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub optIcon_Click(Index As Integer)
    
    If Index = 0 Then
      IconChoice = "1"
    ElseIf Index = 1 Then
      IconChoice = "2"
    ElseIf Index = 2 Then
      IconChoice = "3"
    End If
    
End Sub

Private Sub Option1_Click()
   If Option1.Value = True Then
     Label1(0).Enabled = False
     Text1.Enabled = False
   Else
     Label1(0).Enabled = True
     Text1.Enabled = True
   End If
   
   
   If Option4.Value = False Then
     Me.Height = 2700
   End If
End Sub

Private Sub Option2_Click()
    
   If Option1.Value = False Then
     Label1(0).Enabled = True
     Text1.Enabled = True
   End If
   
   
   
   If Option4.Value = False Then
     Me.Height = 2700
   End If
End Sub

Private Sub Option3_Click()
   If Option3.Value = True Then
     Text2.Visible = True
     Frame1.Visible = False
     txtSendMsg.Left = Text2.Left
     txtSendMsg.Top = Text2.Height + 325
     txtSendMsg.Height = 290
     txtSendMsg.Width = Text2.Width
     cmdExChat.Visible = True
     
   Else
     Text2.Visible = False
     Frame1.Visible = True
     
     txtSendMsg.Left = 2520
     txtSendMsg.Top = 1080
     txtSendMsg.Width = 3135
     txtSendMsg.Height = 825
     cmdExChat.Visible = False
   End If
   
   
   If Option1.Value = False Then
     Label1(0).Enabled = True
     Text1.Enabled = True
   End If
   
   
   If Option4.Value = False Then
     Me.Height = 2700
   End If
   
End Sub

Private Sub Option4_Click()
   If Option4.Value = True Then
     Me.Height = 4520
   Else
     Me.Height = 2700
   End If
   
   If Option1.Value = False Then
     Label1(0).Enabled = True
     Text1.Enabled = True
   End If
End Sub
