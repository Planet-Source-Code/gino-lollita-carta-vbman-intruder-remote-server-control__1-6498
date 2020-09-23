Attribute VB_Name = "modStartProc"
Option Explicit

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF
Private Const DEBUG_PROCESS = &H1
Private Const DEBUG_ONLY_THIS_PROCESS = &H2
Private Const CREATE_SUSPENDED = &H4
Private Const DETACHED_PROCESS = &H8
Private Const CREATE_NEW_CONSOLE = &H10
Private Const NORMAL_PRIORITY_CLASS = &H20
Private Const IDLE_PRIORITY_CLASS = &H40
Private Const HIGH_PRIORITY_CLASS = &H80
Private Const REALTIME_PRIORITY_CLASS = &H100
Private Const CREATE_NEW_PROCESS_GROUP = &H200
Private Const CREATE_NO_WINDOW = &H8000000
Private Const WAIT_FAILED = -1&
Private Const WAIT_OBJECT_0 = 0
Private Const WAIT_ABANDONED = &H80&
Private Const WAIT_ABANDONED_0 = &H80&
Private Const WAIT_TIMEOUT = &H102&
Private Const SW_SHOW = 5

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessBynum Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public ThisPort As Integer

'Sub Main()
     
'  ThisPort = Command()
'
'  Load Form1
'End Sub



Public Function CreateProc(sFile As String, Optional sCommands = "") As Boolean
' starts a process and waits until process is idle.
' returns true if process is successfully loaded.
' use the optional second parameter if you are running
' something with command line parameters
    Dim res&
    Dim sinfo As STARTUPINFO
    Dim pinfo As PROCESS_INFORMATION
    
    With sinfo
      .cb = Len(sinfo)
      .lpReserved = vbNullString
      .lpDesktop = vbNullString
      .lpTitle = vbNullString
      .dwFlags = 0
    End With
    res = CreateProcessBynum(sFile, IIf(Len(sCommands), sFile & " " & sCommands, vbNullString), 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, sinfo, pinfo)
    If res Then        'Launched
       WaitForTerm2 pinfo
       CreateProc = True
    Else
       'Terminated
       CreateProc = False
    End If
End Function
    
Public Sub WaitForTerm2(pinfo As PROCESS_INFORMATION)
    ' Let the process initialize
    WaitForInputIdle pinfo.hProcess, INFINITE
    ' We don't need the thread handle
    CloseHandle pinfo.hThread
End Sub
