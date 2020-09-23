Attribute VB_Name = "Form1SysMenuMod"
'Note all the following code must be placed into new module
'WARNING: If u place exit button in your app, do not place End under it, but u must place Unload Me instead

'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo
'Code type: Sys menu twister
'Menu identification number: sys1066310
'Structure saved in: C:\ProjektiVB6\Api Menu(Ver31)\AppSysMenuReplacement.Map
'---------------------------------------------------------------------------------------------------------------------------
Public pOldProc As Long  ' pointer to Form1's previous window procedure
Public SourceForm As Form
Public appdestructed As Boolean
Public Sub SysMenuModify(hwnd As Long)
SysMenuRestoreDefault (hwnd) 'first reverse sys menu to original state'
'NOW MODIFY IT.......
' Handles to the popup menus to display
Dim hPopupMenu1 As Long
Dim hPopupMenu2 As Long
Dim hPopupMenu3 As Long
Dim hPopupMenu4 As Long
Dim mii1 As MENUITEMINFO ' Structure that will describe menu items
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value

'Create the popup menus which are initialy empty.

hSysMenu = GetSystemMenu(hwnd, 0)
OrigSysMenuCount = GetMenuItemCount(hSysMenu)
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()
hPopupMenu4 = CreatePopupMenu()

'Create the structure which is the base for all added menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'get items into created menus & describe their properties

With mii1 '(Voodoo to Sys Tray)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1009
.dwTypeData = "Voodoo to Sys Tray"
.cch = Len("Voodoo to Sys Tray")
.hSubMenu = 0
End With
retval = InsertMenuItem(hSysMenu, 0, 1, mii1)

With mii1 '(Voodoo minimize)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1008
.dwTypeData = "Voodoo minimize"
.cch = Len("Voodoo minimize")
.hSubMenu = 0
End With
retval = InsertMenuItem(hSysMenu, 1, 1, mii1)

With mii1 '(Voodoo close)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1007
.dwTypeData = "Voodoo close"
.cch = Len("Voodoo close")
.hSubMenu = 0
End With
retval = InsertMenuItem(hSysMenu, 2, 1, mii1)

With mii1 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1006
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hSysMenu, 3, 1, mii1)

With mii1 '(Thanks and links:)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1005
.dwTypeData = "Thanks and links:"
.cch = Len("Thanks and links:")
.hSubMenu = hPopupMenu3
End With
retval = InsertMenuItem(hSysMenu, 4, 1, mii1)

With mii1 '(Thanks and links:/Paul Kuliniewicz)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1004
.dwTypeData = "Paul Kuliniewicz"
.cch = Len("Paul Kuliniewicz")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu3, 0, 1, mii1)

With mii1 '(Thanks and links:/Paul Kuliniewicz/Beam me up Scoty)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1001
.dwTypeData = "Beam me up Scoty"
.cch = Len("Beam me up Scoty")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii1)

With mii1 '(Thanks and links:/Paul Kuliniewicz/KPD - Team)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1003
.dwTypeData = "KPD - Team"
.cch = Len("KPD - Team")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu3, 1, 1, mii1)

With mii1 '(Thanks and links:/Paul Kuliniewicz/KPD - Team/Beam me up Scoty)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1000
.dwTypeData = "Beam me up Scoty"
.cch = Len("Beam me up Scoty")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1 '(Thanks and links:/Paul Kuliniewicz/KPD - Team/M.C. ---> e-mail: kozlicki@yahoo.com)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT
.wID = 1002
.dwTypeData = "M.C. ---> e-mail: kozlicki@yahoo.com"
.cch = Len("M.C. ---> e-mail: kozlicki@yahoo.com")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 2, 1, mii1)

NewSysMenuCount = GetMenuItemCount(hSysMenu)
For i = 0 To OrigSysMenuCount
RemoveMenu hSysMenu, NewSysMenuCount - i, MF_BYPOSITION Or MF_REMOVE
Next i
'The following code is for adding pictures into menus, if there are any!
'------------------------------------------------------------
retval = SetMenuItemBitmaps(hSysMenu, 1009, 1, SourceForm.MenuPicBoxSys0831066(0), SourceForm.MenuPicBoxSys0831066(0))
retval = SetMenuItemBitmaps(hSysMenu, 1008, 1, SourceForm.MenuPicBoxSys0831066(1), SourceForm.MenuPicBoxSys0831066(1))
retval = SetMenuItemBitmaps(hSysMenu, 1007, 1, SourceForm.MenuPicBoxSys0831066(2), SourceForm.MenuPicBoxSys0831066(2))
retval = SetMenuItemBitmaps(hSysMenu, 1005, 1, SourceForm.MenuPicBoxSys0831066(3), SourceForm.MenuPicBoxSys0831066(3))
'------------------------------------------------------------
' Set the custom window procedure to process Form1's messages.
pOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub SysMenuRestoreDefault(hwnd As Long)
' Before unloading, restore the default system menu and remove the
' custom window procedure.
Dim retval As Long  ' return value
' Replace the previous window procedure to prevent crashing.
retval = SetWindowLong(hwnd, GWL_WNDPROC, pOldProc)
' Remove the modifications made to the system menu.
retval = GetSystemMenu(hwnd, 1)
End Sub
'The following function acts as Form1's window procedure to process messages.
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim hSysMenu As Long     ' handle to Form1's system menu
Dim mii As MENUITEMINFO  ' menu item information for Always On Top
Dim retval As Long       ' return value
Select Case uMsg
Case 160 'mouse moving ower caption bar area
      If wParam = 8 Or wParam = 9 Or wParam = 20 Then 'min, restore and close button
     SysMenuRestoreDefault (SourceForm.hwnd)
     SetTimer SourceForm.hwnd, 0, 1, AddressOf Changer
     DrawMenuBar SourceForm.hwnd 'force refresh of caption bar
     End If
     WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
Case WM_INITMENU 'before sys menu is displayed

SysMenuModify (hwnd) 'rebuild it before displaying it

Case WM_SYSCOMMAND 'when you click any of added items !!!!
'Here is the place where things happens
'wParam = magicNumber identifying our selected item
Select Case wParam
Case 1009 '(Voodoo to Sys Tray)

Form1.Visible = False
SystemTrayAddIcon Form1.SysTrayPic, Form1

Case 1008 '(Voodoo minimize)

MCCloseForm Form1, "Min"

Case 1007 '(Voodoo close)
appdestructed = True
'Closing Procedure

MCCloseForm Form1, 2

Case 1001 '(Thanks and links:/Paul Kuliniewicz/Beam me up Scoty)

Shell "start http://www.vbapi.com"

Case 1000 '(Thanks and links:/Paul Kuliniewicz/KPD - Team/Beam me up Scoty)

Shell "start http://www.allapi.net/"

Case 1002 '(Thanks and links:/Paul Kuliniewicz/KPD - Team/M.C. ---> e-mail: kozlicki@yahoo.com)

Case Else
' Some other item was selected.  Let the previous window procedure process it.
WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
End Select
Case Else
' Some other item was selected.  Let the previous window procedure process it.
WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
End Select
End Function
Public Sub Changer()
Dim Mes As Msg
WaitMessage
PeekMessage Mes, SourceForm.hwnd, 160, 512, PM_NOREMOVE 'And Message.wParam = 3 Then
If Mes.Message = 160 And wParam <> 8 And wParam <> 9 And wParam <> 20 Then  'min restore and close button
'change sys menu at this point
         KillTimer SourceForm.hwnd, 0 'kill the timer
         SysMenuModify (SourceForm.hwnd)
End If



If Mes.Message = 512 Then 'user moved mouse over form area
      SysMenuRestoreDefault (SourceForm.hwnd)
      SetTimer SourceForm.hwnd, 0, 1, AddressOf Changer
      DrawMenuBar SourceForm.hwnd 'force refresh of caption bar
End If
End Sub

