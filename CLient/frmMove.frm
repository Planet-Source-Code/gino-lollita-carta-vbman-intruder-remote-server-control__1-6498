VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMove 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Move a file...."
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   4932
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2508
   ScaleWidth      =   4932
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   336
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5052
      _ExtentX        =   8911
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Load"
            Object.ToolTipText     =   "Load Path File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse"
            Object.ToolTipText     =   "Browse"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cancel"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "&Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   70
      TabIndex        =   0
      Top             =   360
      Width           =   4760
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To this destination:"
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Move this file........."
         ForeColor       =   &H0000C000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMove.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMove.frx":27B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMove.frx":4F64
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMove.frx":7716
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdMove_Click()
On Error GoTo NotConnected
   frmFBrowser.sb.Panels.Item(1) = "Moving " & frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
   SendData "Move," & Text1 & "|" & Text2
   frmClient.sb.Panels.Item(1).Text = "  Ready"
   
   
   Logit Move_
   Unload Me
   Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub



Private Sub Form_Activate()
    Me.Tag = "Move"
End Sub

Private Sub Form_Load()
    Dim Selection As String
    Selection = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
    
    If Selection = "" Then
       MsgBox "You must make a selection before this operation can be carried out.", , ""
       Unload Me
       Exit Sub
    Else
       Text1 = Selection
    End If
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Dim State
     frmDirs.List1.Visible = True
     'On Error GoTo NotConnected
     
     State = frmClient.tcpClient.State
   
     ' make sure there's a connection
     If State = 0 And Button.Key = "Cancel" Then
       ' let exit regardless
     ElseIf State <> 0 And Button.Key = "Cancel" Then
       ' let exit regardless
     ElseIf State = 0 And Button.Key <> "Cancel" Then
       GoTo NotConnected
     End If
    
    Select Case Button.Key
       Case "Load"
         frmDirs.Tag = "Move"
         LoadDirFile
       Case "Browse"
         Dim DriveL As String
         DriveL = InputBox("Enter the letter of the Drive you want to Query.")
    
         If DriveL = "" Then Exit Sub
    
         If DriveL <> "" And Len(DriveL) = 1 Then
           DriveL = DriveL + ":\"
         End If
    
         GatherDirs DriveL
       Case "Cancel"
         Unload Me
    End Select
    
    Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub
