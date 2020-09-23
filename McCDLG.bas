Attribute VB_Name = "McCDLG"
'Module name: MC CDLG
'This module created by M.C
'November, 2000
'mainly I made it easy to understand and access some CDLG-s



'main source from:
    'KPD-Team 1998
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public CurrentOpenedFile As String
Public LastPathOpened As String
Public CurentMapFilePathAndName As String 'stores last opened file


Public Function CdlgFileToOpenOrSave(Action As String, MyForm As Form, Filters() As String, StartDir As String, DlgBoxTitle As String)
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    'Set the parent window
    OFName.hwndOwner = MyForm.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    
    'Read from array & compose apropriate filter string
    For i = 0 To UBound(Filters, 2)
    ComposedFilter = ComposedFilter & Filters(1, i) & " (" & Filters(2, i) & ")" + Chr$(0) + Filters(2, i) + Chr$(0)
    Next i
    
    OFName.lpstrFilter = ComposedFilter
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = StartDir
    'Set the title
    OFName.lpstrTitle = DlgBoxTitle
    'No flags
    OFName.flags = 0
    
    Select Case Action
    
        Case "Open"
        'Show the 'Open File'-dialog
        If GetOpenFileName(OFName) Then
            CdlgFileToOpenOrSave = Trim$(OFName.lpstrFile)
            CdlgFileToOpenOrSave = Left(CdlgFileToOpenOrSave, Len(CdlgFileToOpenOrSave) - 1)
            CurentMapFilePathAndName = CdlgFileToOpenOrSave
        Else
            'do nothing as cancel was pressed
        End If
        
        Case "Save"
        If CurentMapFilePathAndName <> "" Then
        OFName.lpstrFile = CurentMapFilePathAndName
        End If
        
        'Show the 'Save File'-dialog
        If GetSaveFileName(OFName) Then
            CdlgFileToOpenOrSave = OFName.lpstrFile
          
            'CdlgFileToOpenOrSave = Trim$(OFName.lpstrFile)
            'CdlgFileToOpenOrSave = Left(CdlgFileToOpenOrSave, Len(CdlgFileToOpenOrSave) - 1)
        Else
             'do nothing as cancel was pressed
        End If
   End Select
End Function
Public Function ShowColor(MyForm As Form)
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hwndOwner = MyForm.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function

Public Function SuckOutFileName(PathAndFilename As String)
j = Len(PathAndFilename)
Do
    If Mid(PathAndFilename, j, 1) = "\" Then
    SuckOutFileName = Right(PathAndFilename, Len(PathAndFilename) - j)
    Exit Do
    End If
j = j - 1
Loop
End Function

Public Function SuckOutFilePath(PathAndFilename As String)
j = 0
j = Len(PathAndFilename)
Do
    If Mid(PathAndFilename, j, 1) = "\" Then
    SuckOutFilePath = Left(PathAndFilename, Len(PathAndFilename) - (Len(PathAndFilename) - j))
    Exit Do
    End If
j = j - 1
Loop
End Function
