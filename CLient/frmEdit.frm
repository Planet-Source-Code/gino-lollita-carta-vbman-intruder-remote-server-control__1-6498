VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00000000&
   Caption         =   "Edit any text file....."
   ClientHeight    =   2850
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   7440
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2850
   ScaleWidth      =   7440
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmEdit.frx":08CA
      Top             =   0
      Width           =   7455
   End
   Begin VB.Menu muFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnuDeskTop 
         Caption         =   "Save &on users Desktop"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "&Cut"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSel 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "C&lear All"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu muEditCard 
         Caption         =   "C&alling Card"
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Selection As String

 
Private Sub Form_Load()
    FrmCnt = FrmCnt + 1
    Selection = frmFBrowser.List1.List(frmFBrowser.List1.ListIndex)
    
    If Selection = "" Then
       MsgBox "You must make a selection before this operation can be carried out.", , ""
       Unload Me
       Exit Sub
    Else
       If LCase(Right(Selection, 3)) <> "txt" And _
          LCase(Right(Selection, 3)) <> "ini" And _
          LCase(Right(Selection, 3)) <> "bat" And _
          LCase(Right(Selection, 3)) <> "---" And _
          LCase(Right(Selection, 3)) <> "bak" And _
          LCase(Right(Selection, 3)) <> "inf" And _
          LCase(Right(Selection, 3)) <> "dos" And _
          LCase(Right(Selection, 3)) <> "old" And _
          LCase(Right(Selection, 3)) <> "log" And _
          LCase(Right(Selection, 3)) <> "chk" And _
          LCase(Right(Selection, 3)) <> "b~k" And _
          LCase(Right(Selection, 3)) <> "htm" And _
          LCase(Right(Selection, 4)) <> "html" And _
          LCase(Right(Selection, 3)) <> "bas" And _
          LCase(Right(Selection, 3)) <> "cpp" And _
          LCase(Right(Selection, 2)) <> "js" And _
          LCase(Right(Selection, 1)) <> "c" And _
          LCase(Right(Selection, 3)) <> "sys" Then
          Text1.Text = "This file canont be edited, it is not a text based file, or it is not in text format."
          TextChanged = False
       Else
          Caption = Selection
          LoadTextFile Selection
       End If
    End If
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    Text1.Move 10, 0, Me.ScaleWidth - 10, Me.ScaleHeight - 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Choice As Integer
    
    If TextChanged Then
      Choice = MsgBox("The file [" & Selection & "] has changed." & vbCrLf & _
                       "would you like to save it before exiting?", vbYesNo, "Save")
      If Choice = 6 Then
         ' yes
         SaveTextFile Selection
      Else
         ' No
      End If
    End If
    
    FrmCnt = FrmCnt - 1
End Sub

Private Sub mnuEditClear_Click()
    Text1 = ""
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End Sub

Private Sub mnuEditPaste_Click()
    Text1.SelText = Clipboard.GetText()
End Sub


Private Sub mnuEditSel_Click()
   Text1.SelStart = 1
   Text1.SelLength = Len(Text1)
End Sub





Private Sub mnuFileSave_Click()
    SaveTextFile Selection
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim FName As String
    
    FName = InputBox("Enter a name to save this file under." & vbCrLf & vbCrLf & _
                      "Full path will be used if given.", "Save As...")
                      
    If FName = "" Then ' Cancel
      Exit Sub
    Else
      SaveTextFile FName
    End If
    
   
Errh:
   MsgBox err.Description, vbCritical, "ERROR"
                     
End Sub

Private Sub muEditCard_Click()
    Dim Name As String
    Dim CCard As String
    ' insert a calling card at the end of the file
    Name = InputBox("Enter your name.... if you dare. Or simply choose cancel.", "Name")
    
    If Name = "" Then
       CCard = "REM  This system has been hacked.... you are infected by the Intruder v.1!! CopyRight 1999"
    Else
       CCard = "REM  This system has been hacked by " & Name & ". You are infected by the Intruder v.1!! CopyRight 1999"
    End If
    ' append to the bottom
    Text1 = Text1 & vbCrLf & vbCrLf & "     " & CCard
End Sub

Private Sub Text1_Change()
   TextChanged = True
End Sub
