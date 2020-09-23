Attribute VB_Name = "SysTray"
Private Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Const WM_MOUSEMOVE = &H200
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2

Private SysTray As NOTIFYICONDATA

Private Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Sub SystemTrayDeleteIcon(control As control)
Shell_NotifyIcon NIM_DELETE, SysTray
End Sub

Public Sub SystemTrayAddIcon(control As control, Form As Form)
        SysTray.cbSize = Len(SysTray)
        SysTray.hwnd = control.hwnd 'control to receive messages from
        SysTray.uID = vbNull
        SysTray.uFlags = NIF_ICON Or NIF_MESSAGE 'flags needed
        SysTray.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
        SysTray.hIcon = control.Picture 'the icon to display
        Shell_NotifyIcon NIM_ADD, SysTray 'add the icon to system tray
End Sub

