VERSION 5.00
Begin VB.Form frmPing 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "(Ping)"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2505
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Ping Results"
      ForeColor       =   &H0000C000&
      Height          =   1455
      Left            =   50
      TabIndex        =   5
      Top             =   840
      Width           =   2415
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   915
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Frame3"
      Height          =   2295
      Left            =   50
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton Command3 
         Caption         =   "&Connect"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ping"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      Begin VB.ComboBox cboAddress 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "IP  or   DNS:"
         ForeColor       =   &H0000C000&
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
     PingHost frmPing.cboAddress
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    frmChooseConn.Combo1 = cboAddress
    frmChooseConn.Show
End Sub



Private Sub Form_Load()
    Dim ConList As String
    'Get the connection list from the registry
    ConList = GetSetting(App.Title, "Settings", "ConnectionList", "")
    'Update the combo box with recent IP addresses used
    UpdateCMB cboAddress, ConList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ConList As String
    
    ConList = GetCMB(cboAddress)
    'Write the recent IP address to the registry
    SaveSetting App.Title, "Settings", "ConnectionList", ConList
End Sub
