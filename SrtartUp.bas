Attribute VB_Name = "SrtartUp"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'NEED THIS TO STAY ON TOP
'Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public MapFileUniqueID As String 'for hlp file
Public Const HELP_CONTENTS = &H3&
Public Declare Function WinHelp Lib "user32.dll" Alias "WinHelpA" (ByVal hWndMain As Long, ByVal lpHelpFile As String, ByVal uCommand As Long, dwData As Any) As Long



Sub Main()

frmSplash.Show
frmSplash.Refresh
Sleep 3000
Load Form1
Form1.Show
Unload frmSplash

End Sub




