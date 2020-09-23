Attribute VB_Name = "MTaskList"
' *********************************************************************
'  Copyright ©1998 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
' modified to work with the Listview control... gh0ul '2000
Option Explicit
'
' Required Win32 API Declarations
'
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'
' Constant used to determine window owner.
'
Private Const GWL_HWNDPARENT = (-8)
'
' Listbox messages
'
Private Const LB_ADDSTRING = &H180
Private Const LB_SETITEMDATA = &H19A
'
' Private variables needed to support enumeration
'
Private m_hWnd As Long

Public Filling As Integer

Public Function FillTaskListBox(lst As ListBox, How As Integer) As Long
   '
   ' Fill an array with the current running Tasks
   '
   Filling = How
   lst.Clear
   Call EnumWindows(AddressOf EnumWindowsProc, lst.hWnd)
   FillTaskListBox = lst.ListCount
End Function

Private Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
   Static WindowText As String
   Static nRet As Long
   Static ICount As Integer
   
   '
   ' Make sure we meet visibility requirements.
   '
   'If Filling = WIN_ALL Then GoTo DoAll
   If IsWindowVisible(hWnd) Then
'DoAll:
      '
      ' It shouldn't have any parent window, either.
      '
      If GetParent(hWnd) = 0 Then
         '
         ' And, finally, it shouldn't have an owner.
         '
         If GetWindowLong(hWnd, GWL_HWNDPARENT) = 0 Then
            '
            ' Retrieve windowtext (caption)
            '
            WindowText = Space$(256)
            nRet = GetWindowText(hWnd, WindowText, Len(WindowText))
            If nRet Then
               '
               ' Clean up window text and add to list.
               '
               ICount = ICount + 1
               WindowText = Left$(WindowText, nRet) & " " & hWnd
               nRet = SendMessage(lParam, LB_ADDSTRING, 0, ByVal WindowText)
               Call SendMessage(lParam, LB_SETITEMDATA, nRet, ByVal hWnd)
               
               
            End If
         End If
      End If
   End If
   '
   ' Return True to continue enumeration.
   '
   EnumWindowsProc = True
End Function


