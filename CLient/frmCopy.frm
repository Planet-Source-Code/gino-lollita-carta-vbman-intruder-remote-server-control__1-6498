VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCopy 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copy a file to..."
   ClientHeight    =   2508
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   5028
   Icon            =   "frmCopy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2508
   ScaleWidth      =   5028
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
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
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4780
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   1320
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
               Picture         =   "frmCopy.frx":08CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCopy.frx":307C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCopy.frx":582E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCopy.frx":7FE0
               Key             =   ""
            EndProperty
         EndProperty
      End
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
         TabIndex        =   1
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " ...into this location."
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Copy this file......"
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdCopy_Click()
    
    SendData "Copy," & Text1 & "|" & Text2
    frmClient.sb.Panels.Item(1).Text = "  Ready"
    Unload Me
    
End Sub

Private Sub Form_Activate()
    Me.Tag = "Copy"
End Sub

Private Sub Form_Load()
    Top = Screen.ActiveForm.Top + Screen.ActiveForm.ScaleTop + 600
    Left = Screen.ActiveForm.Left + Screen.ActiveForm.ScaleLeft + 600
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
     Dim State
     
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
            frmDirs.Tag = "Copy"
            frmDirs.List1.Visible = True
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
