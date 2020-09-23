Attribute VB_Name = "modHelp"
Option Explicit

Sub LoadHelpFile(Loc As String)
  
    '
    Dim FData As String
    Dim RTData As String
    
    Open App.Path & "\Help1.txt" For Binary As #1
       FData = Input(LOF(1), 1)
    Close
    
    ' load in only the specific data requested
    RTData = Retrieve(Loc, FData)
    frmHelp.RText1 = RTData
    FormatText RTData
    frmHelp.RText1.Visible = True
End Sub



Function Retrieve(Loc As String, FData As String) As String
    Const Start As String = "---Start "
    Const End_ As String = "1999"
   
    Dim StartPos As Integer
    Dim StartBuf As Integer
    Dim EndPos As Integer
    Dim EndBuf As Integer
    
    Dim i As Integer
    Dim TempStr As String, TempStr2 As String
    
    ' setup the start buf
    StartBuf = Len(Start) + Len(Loc)
    EndBuf = Len(End_)
    
    
      ' locate the start position
      For i = 0 To Len(FData)
         TempStr = Left(FData, i)
              
         If Right(TempStr, StartBuf) = Start & Loc Then
            StartPos = Len(TempStr) + 5
            Exit For
         End If
      Next
      
      ' get the end position
      For i = StartPos To Len(FData)
         TempStr2 = Left(FData, i)
             
         If Right(TempStr2, EndBuf) = End_ Then
            EndPos = Len(TempStr2) - (StartPos - 1)
            Exit For
         End If
      Next
           
      Retrieve = Mid(FData, StartPos, EndPos)
    
End Function


Sub FormatText(StrData As String)
   Dim StopPos As Integer, StopPos2 As Integer, _
       StopPos3 As Integer
   
   Const StartPos As Integer = 1
   
   ' format header
   StopPos = InStr(StartPos, StrData, ">", vbTextCompare)
   
   With frmHelp.RText1
     .SelLength = Len(.Text)  ' select all
     .SelColor = vbBlack      ' turn all black
     .SelLength = StopPos     ' get stop position for header
     .SelFontSize = 14        ' set headers size
     .SelBold = True          ' make it bold
     .SelColor = vbBlue       ' make only the header blue
   End With
   
   ' format subHeader
   StopPos2 = InStr(StopPos, StrData, ":", vbTextCompare)
   
   With frmHelp.RText1
     .SelStart = StopPos
     .SelLength = StopPos2 - StopPos
     .SelFontSize = 11
     .SelItalic = True
     .SelColor = vbActiveTitleBar
     .SelLength = 0
   End With
      
   
End Sub
