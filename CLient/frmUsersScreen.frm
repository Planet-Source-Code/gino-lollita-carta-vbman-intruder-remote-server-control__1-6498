VERSION 5.00
Begin VB.Form frmUsersScreen 
   Caption         =   "Screen Capture"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4215
   Icon            =   "frmUsersScreen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmUsersScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Form_Load()
    Picture1 = LoadPicture("C:\SCREEN.BMP")
    
    Me.Show
End Sub

Private Sub MNUExit_Click()
    Unload Me
End Sub

Private Sub MNURefresh_Click()
    'Call the capture button code on the main form
    'Call FRMAdmin.BTNScreen_Click
End Sub

