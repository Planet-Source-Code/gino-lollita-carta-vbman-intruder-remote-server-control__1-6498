VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDirs 
   BackColor       =   &H00000000&
   Caption         =   "Directorys on Drive:"
   ClientHeight    =   2688
   ClientLeft      =   168
   ClientTop       =   228
   ClientWidth     =   6264
   Icon            =   "frmDirs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2688
   ScaleWidth      =   6264
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
      Height          =   1872
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   6255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":1350
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":166A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":426E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":4B48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":72FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDirs.frx":7454
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2400
      Top             =   1080
   End
   Begin MSComctlLib.StatusBar Sb 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   6
      Top             =   2436
      Width           =   6264
      _ExtentX        =   11049
      _ExtentY        =   445
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
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1560
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtDrive 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblValid 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a valid drive to query:"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2385
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   336
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6228
      _ExtentX        =   10986
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   166
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Load"
            Object.ToolTipText     =   "Load"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Path Files"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse"
            Object.ToolTipText     =   "Browes Users Directorys"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select"
            Object.ToolTipText     =   "Select"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Create"
            Object.ToolTipText     =   "New Directory"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear"
            Object.ToolTipText     =   "Clear Screen"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmDirs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Dim DriveL As String
    DriveL = txtDrive
    If DriveL = "" Then Exit Sub
    
    DriveL = Left(DriveL, 1) & ":\"
    sb.Panels.Item(1) = " Gathering file information from the SERVER, please wait...."
    GatherDirs DriveL
    With List1
       .Visible = True
       .Clear
    End With
    Frame2.Visible = False
    ClearArray Data, "Dirs"
End Sub

Private Sub Command1_Click()
    Frame2.Visible = False
    ' if there was something there before..... show it again
    If List1.ListCount > 0 Then List1.Visible = True
End Sub

Private Sub Form_Activate()
    If Me.Tag = "" Then
       Toolbar1.Buttons.Item(3).Enabled = False
    Else
       Toolbar1.Buttons.Item(3).Enabled = True
    End If
       
End Sub

Public Sub Form_Load()
    FrmCnt = FrmCnt + 1
    
    If Me.Tag = "Browse" Then
        Frame2.Visible = True
        With List1
          .Visible = False
          .Clear
        End With
    
        Caption = "Viewing folders on " & Dir_.Drive_Name
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
     
    With List1
      .Height = Me.Height - (sb.Height + Toolbar1.Height + 218)
      .Width = Me.ScaleWidth - 10
    End With
    Toolbar1.Width = Me.ScaleWidth
    sb.Panels.Item(1).MinWidth = Me.ScaleWidth - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
     List1.Clear
     FrmCnt = FrmCnt - 1
End Sub




Private Sub Timer1_Timer()
    Static i As Integer
    
    i = i + 1
    sb.Panels.Item(1) = " Retrieving Directory #" & List1.ListCount & " of " & Dir_.Num_Dirs & " total Folders."

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim State
     
     On Error GoTo NotConnected
     
     State = frmClient.tcpClient.State
   
     ' make sure there's a connection
     If State = 0 And Button.Key = "Exit" Then
       ' let exit regardless
     ElseIf State <> 0 And Button.Key = "Exit" Or Button.Key = "Load" Then
       ' let exit regardless
     ElseIf State = 0 And Button.Key <> "Exit" Or Button.Key = "Load" Then
       GoTo NotConnected
     End If
     
    Select Case Button.Key
       Case "Load"
          List1.Visible = True
          List1.Clear
          ClearArray Dir_, "Dirs"
          LoadDirFile
        
       Case "Select"
          If Me.Tag = "Copy" Then
            frmCopy.Text2 = List1.List(List1.ListIndex) & _
            GetName(frmCopy.Text1)
            frmCopy.ZOrder 0
            Unload Me
          ElseIf Me.Tag = "Move" Then
            frmMove.Text2 = List1.List(List1.ListIndex) & _
            GetName(frmMove.Text1)
            frmMove.ZOrder 0
            Unload Me
          ElseIf Me.Tag = "Upload" Then
            '
            'upload LoacalFname to the selected directory
            UpLoadFile LocalFName, List1.List(List1.ListIndex)
            Unload Me
            
          End If
     
          Timer1.Enabled = False
          'Unload Me
       
       Case "Browse"
           List1.Visible = False
           Frame2.Visible = True
       Case "Refresh"
           ' Refresh the current view
           Dir_.Drive_Name = Left(List1.List(2), 3)
           sb.Panels.Item(1).Text = "Refreshing the file list for drive " & Data.Drive_Name
           GatherDirs Dir_.Drive_Name
    
           With List1
            .Visible = True
            .Clear
           End With
    
           Logit Refresh_
           
       Case "Clear"
           ClearArray Dir_, "Dirs"
           List1.Clear
           sb.Panels.Item(1).Text = "  Ready..."
       
       Case "Save"
           SavePaths Dir_, frmDirs, False
           sb.Panels.Item(1) = "  Ready"
       
       Case "Create"
           Dim newDirPath As String
           Dim NewDir As String
           newDirPath = List1.List(List1.ListIndex)
           
           NewDir = InputBox("Enter a name for the new directory.", "New Directory")
           
           If NewDir = "" Then
             MsgBox "Invalid Directory name!", vbCritical, "Bad Name"
             Exit Sub
           End If
           
           SendData "Create_Dir," & newDirPath & NewDir
           
           Dim Msg As String
           Msg = " -->Attempting to create " & newDirPath & NewDir & " on " & CurrentIP
           
           Logit Msg
       Case "Exit"
           Unload Me
    End Select
    
    
     Exit Sub
NotConnected:
   MsgBox "You are not connected!", vbCritical, "Client not connected!"
End Sub
