VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmTskManager 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Host Monitor"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Height          =   222
      Left            =   2580
      Picture         =   "frmTskManager.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   222
      Left            =   2860
      Picture         =   "frmTskManager.frx":00EE
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   240
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6165
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Current Proccesses"
      TabPicture(0)   =   "frmTskManager.frx":01DC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdEnd"
      Tab(0).Control(1)=   "List1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Drives"
      TabPicture(1)   =   "frmTskManager.frx":01F8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TreeView1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "System"
      TabPicture(2)   =   "frmTskManager.frx":0214
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picContainer"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox picContainer 
         Height          =   3255
         Left            =   -74880
         ScaleHeight     =   3195
         ScaleWidth      =   2235
         TabIndex        =   6
         Top             =   120
         Width           =   2295
         Begin VB.VScrollBar VScroll1 
            Height          =   3200
            Left            =   2000
            TabIndex        =   7
            Top             =   0
            Width           =   240
         End
         Begin VB.PictureBox picInner 
            BorderStyle     =   0  'None
            Height          =   7335
            Left            =   0
            ScaleHeight     =   7335
            ScaleWidth      =   2055
            TabIndex        =   8
            Top             =   0
            Width           =   2055
            Begin VB.CommandButton Command1 
               Caption         =   "Retrieve Sys Info"
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1695
            End
            Begin VB.Frame Frame1 
               Height          =   1455
               Left            =   120
               TabIndex        =   10
               Top             =   2520
               Width           =   1695
               Begin VB.CommandButton Command2 
                  Caption         =   "Get Additional Space"
                  Height          =   495
                  Left            =   120
                  TabIndex        =   12
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.ComboBox Combo1 
                  BackColor       =   &H00C0C0C0&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   11
                  Text            =   "Combo1"
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Default C:\"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   13
                  Top             =   240
                  Width           =   780
               End
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   2160
               Width           =   1680
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Number of Display Colors:"
               Height          =   195
               Index           =   3
               Left            =   75
               TabIndex        =   28
               Top             =   4080
               Width           =   1815
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   27
               Top             =   6120
               Width           =   1695
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   26
               Top             =   4920
               Width           =   1695
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   25
               Top             =   6720
               Width           =   1695
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   24
               Top             =   5520
               Width           =   1695
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   23
               Top             =   4320
               Width           =   1695
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Page File size:"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   22
               Top             =   5880
               Width           =   1425
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Memory Load:"
               Height          =   195
               Index           =   6
               Left            =   120
               TabIndex        =   21
               Top             =   4680
               Width           =   1005
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total RAM:"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   20
               Top             =   6480
               Width           =   810
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Current Resolution:"
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   19
               Top             =   5280
               Width           =   1350
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "HDD Space:"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   18
               Top             =   1920
               Width           =   915
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   17
               Top             =   1560
               Width           =   1695
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Computer Name:"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   16
               Top             =   1320
               Width           =   1185
            End
            Begin VB.Label SysInfo 
               BorderStyle     =   1  'Fixed Single
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   15
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Windows Version:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   14
               Top             =   720
               Width           =   1275
            End
         End
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "&End Task"
         Height          =   375
         Left            =   -74880
         TabIndex        =   5
         Top             =   3000
         Width           =   975
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   2805
         Left            =   70
         TabIndex        =   4
         Top             =   75
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   4948
         _Version        =   327682
         Style           =   7
         ImageList       =   "ImageList2"
         Appearance      =   1
      End
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   -74930
         TabIndex        =   3
         Top             =   75
         Width           =   2265
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   4515
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   480
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":054A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":0B7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":0E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTskManager.frx":11B2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public KeyName As String
Public LabelName As String
Public Relative As String
Dim LastValue As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdEnd_Click()
   '
   Dim app_ As String
   app_ = List1.List(List1.ListIndex)
   
   If app_ = "" Then Exit Sub
   
   SendData "Close_App," & app_
   
   List1.RemoveItem List1.ListIndex
End Sub

Private Sub cmdRefresh_Click()
      List1.Clear
      TreeView1.Nodes.Clear
      SendData "Get_Tasks,"
      Pause 1000
      SendData "Get_Drives,"
      
      SSTab1.TabEnabled(0) = True
      SSTab1.TabEnabled(1) = True
      SSTab1.TabEnabled(2) = True
      
      Dim Msg3 As String
      Msg3 = " -->Refreshing Drive and current task information..."
      Logit Msg3
End Sub

Private Sub Command1_Click()
    '
    ' send for the sys Info
    SendData "Get_SysInfo,"
    
End Sub

Private Sub Command3_Click()
   Me.WindowState = vbMinimized
End Sub

Private Sub Command4_Click()
    If Me.Height = 610 Then
      Me.Height = 5010
   ElseIf Me.Height > 610 Then
      Me.Height = 610
   End If
End Sub

Private Sub Form_Load()
   
   SSTab1.TabEnabled(0) = False
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   
   sb.Panels(1).MinWidth = Me.ScaleWidth - 25
   
VScroll1.Max = picInner.Height - picContainer.Height
VScroll1.SmallChange = 80
VScroll1.LargeChange = picInner.Height / 3 'VScroll1.Max / 20
'Form_Paint
End Sub



Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
      '
      Dim i As Integer
      Dim ImgNum As Integer
      ImgNum = 6
      Relative = "new"
      If Node.Child.Text = "" Then
        If Drives.totalNumDrives <> 0 Then
          TreeView1.Nodes.Remove Node.Child.Index
          'TreeView1.Nodes.Add "Main Branch", tvwChild, Relative, Drives.IndvDrives(1), 7
                             
          ' create a branch for each drive
          For i = 1 To Drives.totalNumDrives
             If (StrComp(Drives.IndvDrives(i), "A:\ [FLOPPY]", vbTextCompare) = 0) Then
                 ImgNum = 2
             ElseIf (StrComp(Drives.IndvDrives(i), "B:\ [FLOPPY]", vbTextCompare) = 0) Then
                 ImgNum = 3
             ElseIf (Right(Drives.IndvDrives(i), 8) = "[CD-ROM]") Then
                 ImgNum = 4
             End If
             
             TreeView1.Nodes.Add Node.Key, tvwChild, , Drives.IndvDrives(i), ImgNum
          Next
          
          
        End If
      End If
End Sub
''







' scrolloing pic box
Private Sub VScroll1_Change()
    picInner.Top = -VScroll1.Value
End Sub

