VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPicIMg 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select image to display"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6705
   Icon            =   "frmPicWave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   3165
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "  Ready..."
            TextSave        =   "  Ready..."
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdRetrieve 
      Caption         =   "&Retrieve"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Display"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   1680
      Width           =   855
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000C000&
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmPicIMg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLoad_Click()
    FileName = App.Path & "\ImgList.dat"
    Dim i As Integer
    
    If FileExists(FileName) = False Then
      MsgBox "No file to load!", vbCritical, "File Not Found"
      Exit Sub
    End If
    List1.Clear
    LoadFile FileName, "Pics"
End Sub

Private Sub cmdRetrieve_Click()
    SendData "Get_Pic_Paths,"
    Screen.MousePointer = vbHourglass
    frmPicIMg.StatusBar1.Panels(1).Text = "  Retrieving Image file paths from host...."
    
    Dim Msg As String
    Msg = " -->Attempting to retrieve all .jpg image files from: " & CurrentIP
    Logit Msg
End Sub

Private Sub cmdSave_Click()
   Dim i As Integer
   On Error GoTo Trap:
   
   FileName = App.Path & "\ImgList.dat"
   
   Data.Num_Files = List1.ListCount
      
   For i = 1 To Data.Num_Files
      Data.FileName(i) = List1.List(i)
   Next
      
   Open FileName For Binary As #1
      Put #1, , Data
   Close #1
   
   MsgBox "Image Path List Saved", vbInformation, "Image List Saved"
   
   Exit Sub
Trap:
   MsgBox err.Description, vbCritical, "ERROR #" & err.Number

End Sub

Private Sub cmdSelect_Click()
    SendData "Show_Picture," & List1.List(List1.ListIndex)
    StatusBar1.Panels(1).Text = "  Now Displaying [" & List1.List(List1.ListIndex) & "]"
    
    
    Dim Msg As String
    Msg = " -->Display Image: (" & List1.List(List1.ListIndex) & ") at: " & CurrentIP
    Logit Msg
End Sub

Private Sub Form_Load()
    FrmCnt = FrmCnt + 1
    Form_Resize
    
End Sub

Private Sub Form_Resize()
    StatusBar1.Panels(1).MinWidth = Me.ScaleWidth - 25
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCnt = FrmCnt - 1
End Sub
