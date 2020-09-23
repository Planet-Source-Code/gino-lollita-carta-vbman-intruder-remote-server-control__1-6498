VERSION 5.00
Begin VB.Form frmSetMail 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-Mail Notify"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4575
   Icon            =   "frmSetMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAdmin 
      BackColor       =   &H00000000&
      Caption         =   "Administrator Setup"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox chkEnable 
      BackColor       =   &H00000000&
      Caption         =   "Enable Email Notify"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Selections"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4335
      Begin VB.OptionButton optAlways 
         BackColor       =   &H00000000&
         Caption         =   "Notify everytime user goes online"
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton optOnce 
         BackColor       =   &H00000000&
         Caption         =   "Notify only once."
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Send to:"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSetMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Mail_Setup_Fname As String

Private Type Email
  sAddr As String
  bEnabled As Boolean
  bNotifyOnce As Boolean
  bNotifyAlways As Boolean
  bChkAdmn As Boolean
End Type

Dim MaiL As Email

Private Sub chkAdmin_Click()
   If chkAdmin.Value = vbChecked Then
      MaiL.bChkAdmn = True
   Else
      MaiL.bChkAdmn = False
   End If
   
End Sub

Private Sub chkEnable_Click()
   If chkEnable.Value = vbChecked Then
      Frame1.Enabled = True
      MaiL.bEnabled = True
   Else
      Frame1.Enabled = False
      MaiL.bEnabled = False
   End If
End Sub



Private Sub cmdSave_Click()
     
    MaiL.sAddr = Text1
    
    If MaiL.sAddr = "" Then
      MsgBox "E-Mail address required!", , "ERROR"
      Exit Sub
    End If
    
    ' ok to save data
    Mail_Setup_Fname = App.Path & "\SysConfig.sys"
    
    Open Mail_Setup_Fname For Binary As #1
      Put #1, , MaiL
    Close #1
    
    ' send the new data to the server to be saved
    ' and incorporated into the servers actions
    SendData "New_Email_Settings," & MaiL.bEnabled & ":" _
                                   & MaiL.bNotifyAlways & ":" _
                                   & MaiL.bNotifyOnce & ":" _
                                   & MaiL.sAddr & ":" _
                                   & MaiL.bChkAdmn & ":"
    
End Sub

Private Sub Form_Load()
    GetMailSettings
    
    With MaiL
      If .bEnabled Then
        chkEnable = vbChecked
      End If
      If .bChkAdmn Then
        chkAdmin = vbChecked
      End If
      optAlways = .bNotifyAlways
      optOnce = .bNotifyOnce
      Text1 = .sAddr
    End With
    
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub optAlways_Click()
    If optAlways Then
      MaiL.bNotifyAlways = True
      MaiL.bNotifyOnce = False
    End If
End Sub

Private Sub optOnce_Click()
    If optOnce Then
      MaiL.bNotifyOnce = True
      MaiL.bNotifyAlways = False
    End If
End Sub


Private Sub GetMailSettings()
   On Error GoTo OpenErr
   
   Open App.Path & "\SysConfig.sys" For Binary As #1
      Get #1, , MaiL
   Close #1
   
   Exit Sub
      
OpenErr:

   MsgBox "Error opening mail file"
End Sub
