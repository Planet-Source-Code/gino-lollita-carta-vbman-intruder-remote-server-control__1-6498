Attribute VB_Name = "modWriteDT"
Option Explicit



Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Public Sub WriteOnDT(strMsg As String)
Dim hdc As Long
Dim tR As RECT
Dim lCol As Long

    ' First get the Desktop DC:
    hdc = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
            
    ' Draw text on it:
    tR.Left = Int((700 * Rnd) + 1)
    tR.Top = Int((20 * Rnd) + 1)
    tR.Right = 640
    tR.Bottom = 32
    lCol = GetTextColor(hdc)
    SetTextColor hdc, &HFF&
    DrawText hdc, strMsg, Len(strMsg), tR, 0
    SetTextColor hdc, lCol
    
    ' Make sure you do this to release the GDI
    ' resource:
    DeleteDC hdc
    
End Sub
