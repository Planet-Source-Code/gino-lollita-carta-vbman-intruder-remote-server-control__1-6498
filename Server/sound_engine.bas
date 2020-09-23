Attribute VB_Name = "Sound_Engine"

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlaySound(sndFilename As String)
   On Error GoTo err
   
   sndPlaySound sndFilename, 3  'The 3 prevents the system from freezing during playback
   Exit Sub

err:
    SendData "Wave_Error," & err.Description
End Sub

