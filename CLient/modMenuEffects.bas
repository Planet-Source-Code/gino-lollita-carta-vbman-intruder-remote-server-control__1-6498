Attribute VB_Name = "modMenuEffects"
Option Explicit

Declare Function GetMenu Lib "user32" _
(ByVal hwnd As Long) As Long

Declare Function GetSubMenu Lib "user32" _
(ByVal hmenu As Long, ByVal nPos As Long) As Long

Public Declare Function GetMenuItemID Lib "user32" _
(ByVal hmenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hmenu As Long, ByVal nPosition As Long, _
ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, _
ByVal hBitmapChecked As Long) As Long

Public Const MF_BITMAP = &H4&

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Declare Function GetMenuItemCount Lib "user32" _
(ByVal hmenu As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" _
Alias "GetMenuItemInfoA" (ByVal hmenu As Long, _
ByVal un As Long, ByVal b As Boolean, _
lpMenuItemInfo As MENUITEMINFO) As Boolean

Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_STRING = &H0&



