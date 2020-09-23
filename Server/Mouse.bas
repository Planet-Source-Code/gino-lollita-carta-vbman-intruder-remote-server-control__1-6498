Attribute VB_Name = "Mouse"
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_MOVE = &H1
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const TWIPS_PER_INCH = 1440
Private Const POINTS_PER_INCH = 72
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const MOUSE_MICKEYS = 65535
Public Enum enReportStyle
    rsPixels
    rsTwips
    rsInches
    rsPoints
End Enum
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub FunnyMouse()
    Dim i As Long, X As Long, Y As Long
    GetScreenRes X, Y
    For i = 1 To 50
        Call MouseMove(Rnd * X, Rnd * Y)
        Sleep (100)
    Next i
End Sub

Public Sub GetScreenRes(ByRef X As Long, ByRef Y As Long, Optional ByVal ReportStyle As enReportStyle)
    X = GetSystemMetrics(SM_CXSCREEN)
    Y = GetSystemMetrics(SM_CYSCREEN)
    If Not IsMissing(ReportStyle) Then
        If ReportStyle <> rsPixels Then
            X = X * Screen.TwipsPerPixelX
            Y = Y * Screen.TwipsPerPixelY
            If ReportStyle = rsInches Or ReportStyle = rsPoints Then
                X = X \ TWIPS_PER_INCH
                Y = Y \ TWIPS_PER_INCH
                If ReportStyle = rsPoints Then
                    X = X * POINTS_PER_INCH
                    Y = Y * POINTS_PER_INCH
                End If
            End If
        End If
    End If
End Sub

' Converts pixel X coordinates to mickeys
Public Function PixelXToMickey(ByVal pixX As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tX As Single
    Dim tpixX As Single
    Dim tMickeys As Single
    GetScreenRes X, Y
    tMickeys = MOUSE_MICKEYS
    tX = X
    tpixX = pixX
    PixelXToMickey = CLng((tMickeys / tX) * tpixX)
End Function

' Converts pixel Y coordinates to mickeys
Public Function PixelYToMickey(ByVal pixY As Long) As Long
    Dim X As Long
    Dim Y As Long
    Dim tY As Single
    Dim tpixY As Single
    Dim tMickeys As Single
    GetScreenRes X, Y
    tMickeys = MOUSE_MICKEYS
    tY = Y
    tpixY = pixY
    PixelYToMickey = CLng((tMickeys / tY) * tpixY)
End Function

Public Sub MouseMove(ByRef xPixel As Long, ByRef yPixel As Long)
    Dim cbuttons As Long
    Dim dwExtraInfo As Long
    mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, PixelXToMickey(xPixel), PixelYToMickey(yPixel), cbuttons, dwExtraInfo
End Sub

