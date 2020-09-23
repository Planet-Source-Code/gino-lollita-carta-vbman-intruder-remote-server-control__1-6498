VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   780
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Declare Dynamic Link
'First Reference It By Clicking Project in the Toolbar
'Click Reference And Search For Devtools Dll
Dim DevTool As New RegSvr 'Linking The Dll To The Project

Private Sub Form_Load()
DevTool.AboutMe ("c:\temp") 'Store About File In C:\Temp
End Sub

Public Function SampleApps()
'Do Not Execute This Function Unless You Know What Youre Doing

With DevTool
.DoChangeUser
.DoChangeWallPaper ("C:\temp\Sample.bmp") 'Changes Wallpaper With Given BMP FIle
.DoOpenCd_RomDrive 'Opens Cd rom Drive
.DoCloseCd_RomDrive 'Closes Cd Rom Drive
.DoCreateUrl ("c:\Windows\Desktop\Test.url"), ("www.developit.demon.nl") 'Creates Shorcut To InternetPage
.DoDialupNow 'Dials Users Deafault DUN Connection
.DoDownloadRegsvr32 ("C:\temp") 'Retrieves Regsvr32.exe from The Dll And Stores It In C:\Temp
.DoFileCryption ("C:\temp\First.txt"), ("C:\temp\Second.txt"), ("devtools"), True
'Encrypts First.txt To Second.Txt with Password ,to Decrypt Place Second.txt AS first
'File and as second file place name.txt
.DoFindCD_RomDrive ' Finds Users CD Rom Drive Letter
.DoHideTaskBar 'Hides The Windows TaskBar
.DoPolicyChangeMenuSpeed ("199") ' changes the menu speed in windows
'Fore The Policy Features Do
DevTool.AboutMe ("c:\temp") 'Store About File In C:\Temp

.DoPolicyNoRun True ' The Run Command Has been Disabled From The StartMenu
' true is Disabled False is enabeled
End With

' For Questions And Answers Mailto : Devtools@developit.demon.nl

End Function
