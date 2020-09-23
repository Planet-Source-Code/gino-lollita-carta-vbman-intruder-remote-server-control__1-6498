VERSION 5.00
Begin VB.Form frmResolution 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select desired resolution"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "800X600"
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "&Change"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    SendData "Change_Resolution," & Combo1
    
    
    Dim Msg As String
    Msg = " -->Attempting to change resolution to " & Combo1 & " at: " & CurrentIP
    Logit Msg
End Sub

Private Sub Form_Load()
    With Combo1
      .AddItem "640X480"
      .AddItem "800X600"
      .AddItem "1024X768"
      .Text = "640X480"
    End With
End Sub



