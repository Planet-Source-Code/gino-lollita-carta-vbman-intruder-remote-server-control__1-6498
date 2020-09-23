VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmFBrowser 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "File Browser"
   ClientHeight    =   2904
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   6516
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFBrowser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2904
   ScaleWidth      =   6516
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   396
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6516
      _ExtentX        =   11494
      _ExtentY        =   699
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   16
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Load"
            Object.ToolTipText     =   "Open"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BrowseFile"
            Object.ToolTipText     =   "Browse User Files"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "BrowseDir"
            Object.ToolTipText     =   "Browse users directorys"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.Tag             =   "Search Files"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Move"
            Object.ToolTipText     =   "Move"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Edit"
            Object.Tag             =   "Edit"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Copy"
            Object.Tag             =   "Copy"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Rename"
            Object.Tag             =   "Rename"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Clear"
            Object.ToolTipText     =   "CLear"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2100
      Left            =   0
      TabIndex        =   5
      Top             =   420
      Visible         =   0   'False
      Width           =   6495
   End
   Begin MSComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2655
      Width           =   6510
      _ExtentX        =   11494
      _ExtentY        =   466
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Text            =   "Ready..."
            TextSave        =   "Ready..."
            Object.ToolTipText     =   "Status Bar"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtDrive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Text            =   "C:\"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblValid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a valid drive to query:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2385
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2640
      Top             =   3000
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   1920
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":04E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":07FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":0CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":100C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":1326
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":1500
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":181A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":19F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":1BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":1DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmFBrowser.frx":1F82
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub cmdFileActions_Click(Index As Integer)
   
        
   Select Case Index
       Case 0
       Case 1
       Case 2
       Case 3
       Case 4
       Case 5
    End Select
    
    
End Sub

Private Sub cmdOK_Click()
   On Error GoTo NotConnected
    Frame2.Visible = False
    sb.Panels.Item(1).Text = " Gathering file information from the SERVER, please wait...."
    Timer1.Enabled = False
    
    ' send the command to the server
    SendData "Get_Users_Files," & txtDrive
    With List1
       .Visible = True
       .Clear
    End With
    Data.Drive_Name = txtDrive
    ClearArray Data, "Files"
    
    Logit GetFiles
    Exit Sub
    
    
NotConnected:
     MsgBox "You are not connected!", vbCritical, "Client not connected!"
     sb.Panels.Item(1).Text = " Last command attempt failed.  Waiting... time is of the essence."
End Sub


Private Sub Command1_Click()
    Frame2.Visible = False
    ' if there was something there before..... show it again
    If List1.ListCount > 0 Then List1.Visible = True
End Sub

Private Sub Form_Load()
     'Me.Top = frmClient.Top + 250
     'Me.Left = frmClient.Left + 500
     ' ask the server for the current File paths
     FrmCnt = FrmCnt + 1
End Sub



Private Sub Form_Resize()

    On Error Resume Next
 
    With List1
      .Height = Me.Height - (sb.Height + Toolbar1.Height + 218)
      .Width = Me.ScaleWidth - 10
    End With
    
    sb.Panels.Item(1).MinWidth = Me.ScaleWidth - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    List1.Clear
    FrmCnt = FrmCnt - 1
End Sub



Private Sub List1_Click()
    
    If UCase(Right(List1.List(List1.ListIndex), 4)) = ".EXE" Or _
       UCase(Right(List1.List(List1.ListIndex), 4)) = ".TXT" Or _
       UCase(Right(List1.List(List1.ListIndex), 4)) = "HTML" Then
       frmClient.mnuRun.Enabled = True
    Else
       frmClient.mnuRun.Enabled = False
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
       PopupMenu frmClient.mnuPopUp2
    End If
End Sub



Sub Down()
    Dim DrPath As String
    ' download a file.
    ' make sure a file has been selected.
    
    '--- basic algorythm:
    '
    ' If a file has been selected, send that filepath to the server
    ' with a command to download it
    DLoadFName = List1.List(List1.ListIndex)
    If DLoadFName = "" Then MsgBox "You must select a file to Download.", vbExclamation, "Select a file": Exit Sub
    DrPath = ShowBrowser() ' get a folder to save to
    DLoadSaveName = DrPath & GetFileNameFromPath(DLoadFName)
    
    SendData "DownLoad_This_,start;" & DLoadFName
End Sub

Sub Run()
    SendData "Run," & List1.List(List1.ListIndex)
End Sub

Sub FileExit()
    Unload Me
End Sub

Sub FileLoad()
    List1.Clear
    ClearArray Data, "Files"
    sb.Panels.Item(1).Text = "  Loading file..."
    LoadPathProc
End Sub

Sub MoveFile()
    ' only call delete if a file has been selected.
    
   Dim FName As String
   
   FName = List1.List(List1.ListIndex)
   
   If FName = "" Then Exit Sub
           
   sb.Panels.Item(1).Text = "Preparing to move: " & FName
   frmMove.Show
   
End Sub

Sub Copy()
    ' only call delete if a file has been selected.
    
   Dim FName As String
   
   FName = List1.List(List1.ListIndex)
   
   If FName = "" Then Exit Sub

   sb.Panels.Item(1).Text = "Preparing to Copy: " & FName
   Dim Selection As String
   Selection = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
    
   If Selection = "" Then
      MsgBox "You must make a selection before this operation can be carried out.", , ""
      Exit Sub
   Else
      frmCopy.Text1 = Selection
      frmCopy.Show
      frmDirs.Tag = "Copy"
   End If
End Sub

Sub Del()
   ' only call delete if a file has been selected.
    
   Dim FName As String
   
   FName = List1.List(List1.ListIndex)
   
   If FName = "" Then Exit Sub
   
   sb.Panels.Item(1).Text = "Preparing to Delete the file: " & FName
   DoDelete
   
End Sub

Sub Rename()
   ' only call delete if a file has been selected.
    
   Dim FName As String
   
   FName = List1.List(List1.ListIndex)
   
   If FName = "" Then Exit Sub
   
   Dim NewName, OldName
           
   sb.Panels.Item(1).Text = "Preparing to rename: " & FName
   OldName = FName
   NewName = InputBox("Enter a new name for this file, including extension.", "New Name")
           
   If NewName = "" Then
      MsgBox "Name was not changed.", vbInformation
      Exit Sub
   End If
           
   NewName = GetPathOnly(FName) & NewName
           
   ' rename the file
   Name OldName As NewName
   MsgBox "The file: " & FName & vbCrLf & _
          "has been renamed to: " & NewName, vbInformation
End Sub



Private Sub Timer1_Timer()
   sb.Panels.Item(1).Text = " Retrieving file path #" & List1.ListCount - 1 & _
             " of " & Data.Num_Files & " from the Data Server."
             
   If List1.ListCount >= Data.Num_Files Then
     Timer1.Enabled = False
     sb.Panels.Item(1).Text = "  Ready"
   End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
       Static i As Integer
   Dim State As Integer
   Dim FName As String
   Dim DrivePath As String
   
   FName = List1.List(List1.ListIndex)
   
   On Error Resume Next
   State = frmClient.tcpClient.State
   
   ' make sure there's a connection
   If State = 0 And Button.Key = "Exit" Or Button.Key = "Load" Then
     ' let exit regardless
   ElseIf State <> 0 And Button.Key = "Exit" Or Button.Key = "Load" Then
     ' let exit regardless
   ElseIf State = 0 And Button.Key <> "Exit" Then
     GoTo NotConnected
   End If
   
   
    Select Case Button.Key
       Case "Load"
           FileLoad
       Case "Move"
           sb.Panels.Item(1).Text = "Preparing to move: " & FName
           frmMove.Show
       Case "Delete"
           sb.Panels.Item(1).Text = "Preparing to Delete the file: " & FName
           DoDelete
       Case "Edit"
           sb.Panels.Item(1).Text = "Loading: (" & FName & ")"
           frmEdit.Show
       Case "Copy"
           sb.Panels.Item(1).Text = "Preparing to Copy: " & FName
           Dim Selection As String
           Selection = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
    
           If Selection = "" Then
             MsgBox "You must make a selection before this operation can be carried out.", , ""
             Exit Sub
           Else
             frmCopy.Text1 = Selection
             frmCopy.Show
             frmDirs.Tag = "Copy"
           End If
       Case "Rename" 'Rename
           Dim NewName, OldName
           
           sb.Panels.Item(1).Text = "Preparing to rename: " & FName
           OldName = FName
           NewName = InputBox("Enter a new name for this file, including extension.", "New Name")
           
           If NewName = "" Then
             MsgBox "Name was not changed.", vbInformation
             Exit Sub
           End If
           
           DrivePath = GetPathOnly(FName)
           
                     
           If Right(DrivePath, 1) <> "\" Then
             NewName = DrivePath & "\" & NewName
           Else
             NewName = DrivePath & NewName
           End If
           
           
           SendData "Rename," & OldName & ";" & NewName
                
       Case "Save"
           SavePaths Data, frmFBrowser, True
       Case "BrowseFile"
           Frame2.Visible = True
           List1.Visible = False
           txtDrive.SetFocus
       Case "BrowseDir"
           frmDirs.Tag = "Browse"
           frmDirs.Form_Load
           frmDirs.Show
       Case "Search"
           frmSearch.Show
       Case "Refresh"
           ' Redfresh the current view
           Data.Drive_Name = Left(List1.List(2), 3)
           SendData "Get_Users_Files," & Data.Drive_Name
           sb.Panels.Item(1).Text = "Refreshing the file list for drive " & Data.Drive_Name
    
           With List1
            .Visible = True
            .Clear
           End With
    
           Logit Refresh_
     
       Case "Clear"
           List1.Clear
           ClearArray Data, "Files"
       
       Case "Exit"
           Unload Me
    End Select
    
    Exit Sub
    
    
NotConnected:
     MsgBox "You are not connected!", vbCritical, "Error"
End Sub
