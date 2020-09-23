VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "M.C. Menu Voodoo"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   7320
      ScaleHeight     =   4035
      ScaleWidth      =   3795
      TabIndex        =   51
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ListBox List7 
         Height          =   2400
         Left            =   3600
         TabIndex        =   56
         Top             =   720
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.ListBox List6 
         Height          =   2400
         Left            =   360
         TabIndex        =   53
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   1200
         TabIndex        =   52
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "The following VB menus has been found. Select any of them or all."
         Height          =   375
         Left            =   600
         TabIndex        =   54
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1515
      ScaleWidth      =   9435
      TabIndex        =   31
      Top             =   5040
      Width           =   9495
      Begin VB.CommandButton SelectionButton 
         Caption         =   "Comments !"
         Enabled         =   0   'False
         Height          =   255
         Index           =   3
         Left            =   4800
         TabIndex        =   36
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton SelectionButton 
         Caption         =   "Form_Load"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   35
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton SelectionButton 
         Caption         =   "General section of form"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   34
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton SelectionButton 
         Caption         =   "Main code"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1095
         Left            =   0
         TabIndex        =   32
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   1931
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         TextRTF         =   $"Form1.frx":0CCA
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   6360
         Picture         =   "Form1.frx":0E1E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click here to vote for !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7080
         TabIndex        =   37
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00000000&
      Caption         =   "General menu behaviour settings"
      ForeColor       =   &H000000FF&
      Height          =   5175
      Left            =   1680
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00C0C0C0&
         Height          =   1335
         Index           =   2
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   6915
         TabIndex        =   38
         Top             =   2640
         Width           =   6975
         Begin VB.CommandButton Command9 
            Caption         =   "ReadMe1st"
            Height          =   255
            Left            =   4080
            TabIndex        =   48
            Top             =   120
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Delete color combination"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   47
            Top             =   960
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Save color combination"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   4080
            TabIndex        =   46
            Top             =   680
            Width           =   2175
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Get bkg color"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   960
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Get text color"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   680
            Width           =   1455
         End
         Begin VB.ListBox List5 
            Enabled         =   0   'False
            Height          =   840
            Left            =   2040
            TabIndex        =   41
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Colorized by me"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Default colors"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Item text"
            Height          =   255
            Index           =   2
            Left            =   4080
            TabIndex        =   43
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Favourites:"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   42
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command7 
         Caption         =   "OK"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   4440
         Width           =   2535
      End
      Begin VB.PictureBox Picture3 
         Height          =   855
         Index           =   1
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   6915
         TabIndex        =   16
         Top             =   1560
         Width           =   6975
         Begin VB.OptionButton Option7 
            Caption         =   "PopUpMenu will be triggered with left click."
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Value           =   -1  'True
            Width           =   5775
         End
         Begin VB.OptionButton Option7 
            Caption         =   "PopUpMenu will be triggered with right click."
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   120
            Width           =   5775
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   855
            Left            =   4080
            TabIndex        =   30
            Top             =   -240
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1095
         Index           =   0
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   6915
         TabIndex        =   12
         Top             =   360
         Width           =   6975
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the center/down side where mouse will be clicked. "
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   15
            Top             =   720
            Width           =   5895
         End
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the right/down side where mouse will be clicked. "
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   5895
         End
         Begin VB.OptionButton Option6 
            Caption         =   "PopUpMenu will appear to the left/down side where mouse will be clicked. "
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   5895
         End
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing here works for sys menu !"
         ForeColor       =   &H0080FF80&
         Height          =   255
         Left            =   2760
         TabIndex        =   66
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture6 
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4515
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   4575
      Begin VB.PictureBox SysTrayPic 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   3240
         Picture         =   "Form1.frx":1AE8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   65
         Top             =   120
         Width           =   540
      End
      Begin VB.PictureBox MenuPicBoxSys0831066 
         Height          =   255
         Index           =   3
         Left            =   1200
         Picture         =   "Form1.frx":27B2
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox MenuPicBoxSys0831066 
         Height          =   255
         Index           =   2
         Left            =   840
         Picture         =   "Form1.frx":29FC
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox MenuPicBoxSys0831066 
         Height          =   255
         Index           =   1
         Left            =   480
         Picture         =   "Form1.frx":2C46
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox MenuPicBoxSys0831066 
         Height          =   255
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":2E90
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox MenuPicBox0101284205 
         Height          =   255
         Index           =   2
         Left            =   2640
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   59
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox MenuPicBox0101284205 
         Height          =   255
         Index           =   1
         Left            =   2160
         Picture         =   "Form1.frx":32C2
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   58
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox MenuPicBox0101284205 
         AutoSize        =   -1  'True
         Height          =   255
         Index           =   0
         Left            =   1680
         Picture         =   "Form1.frx":350C
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   57
         Top             =   120
         Width           =   255
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   3
         Left            =   1200
         Picture         =   "Form1.frx":3756
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   29
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   2
         Left            =   840
         Picture         =   "Form1.frx":3948
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   28
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "Form1.frx":3B3A
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   27
         Top             =   120
         Width           =   240
      End
      Begin VB.PictureBox MenuPicContainer 
         AutoSize        =   -1  'True
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "Form1.frx":3D2C
         ScaleHeight     =   180
         ScaleWidth      =   180
         TabIndex        =   26
         Top             =   120
         Width           =   240
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   24
      Top             =   3960
      Width           =   1575
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00808080&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
      Begin VB.CommandButton Command19 
         Caption         =   "Sys Twister"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Vb Sucker"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Make replacement"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Generate Code"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Paste code direct under event where you want POPUPmenu to appear, i.e. Command1_Click. Must have module with pub declares."
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Public Declares .."
         Height          =   255
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Declarations ready to be pasted into module.Once u have mod. u yust need 'Generate code' button."
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Private declares"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Paste this into declaration section of new form. "
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ListBox List4 
      Height          =   4740
      Left            =   8520
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   4740
      ItemData        =   "Form1.frx":3F1E
      Left            =   7560
      List            =   "Form1.frx":3F20
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load structure"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save structure"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Exit"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   4740
      ItemData        =   "Form1.frx":3F22
      Left            =   6600
      List            =   "Form1.frx":3F24
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   4740
      ItemData        =   "Form1.frx":3F26
      Left            =   1680
      List            =   "Form1.frx":3F28
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Suck in Vb menus"
      Height          =   375
      Left            =   0
      TabIndex        =   49
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Pic"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   8520
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Item state"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   7560
      TabIndex        =   8
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Item type"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Structure building box"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Under generall section of your form place:
Dim APICheck1 As Boolean 'Bold
Dim APICheck2 As Boolean 'Normal
Dim APICheck3 As Boolean 'Enabled
Dim APICheck4 As Boolean 'Disabled
Dim APICheck5 As Boolean 'Grayed
Dim APICheck6 As Boolean 'Make them ALL/Grayed
Dim APICheck7 As Boolean 'Checked ?
Dim APICheck8 As Boolean 'Checked ?/Not Checked
Dim APICheck9 As Boolean 'Checked ?/On Start Checked/With Checkmark
Dim APICheck10 As Boolean 'Checked ?/On Start Checked/With RadioButton
Dim APICheck11 As Boolean 'Checked ?/On Start Checked/With Pictures
Dim APICheck12 As Boolean 'Checked ?/On Start Checked/On Start Unchecked
Dim APICheck13 As Boolean 'Checked ?/On Start Checked/On Start Unchecked/With Checkmark
Dim APICheck14 As Boolean 'Checked ?/On Start Checked/On Start Unchecked/With RadioButton
Dim APICheck15 As Boolean 'Checked ?/On Start Checked/On Start Unchecked/With Pictures
Dim APICheck24 As Boolean 'Picture ?
Dim APICheck25 As Boolean 'Picture ?/Yes
Dim APICheck26 As Boolean 'Picture ?/No
Dim APICheck27 As Boolean 'Insert Column Break
Dim APICheck28 As Boolean 'Insert Column Break/With dividing line
Dim APICheck29 As Boolean 'Insert Column Break/Without dividing line



Dim controler
Dim ClickNoCanDo As Boolean 'some click trouble repair
Dim VariableIndex As Integer 'for checked items needed variables
Dim OCLC As Integer 'OutputCodeLineCounter

'for sizing outputcode box
Dim defaultpicture1top As Integer
Dim defaultpicture1height As Integer
Dim defaultRichTextBox1top As Integer
Dim defaultRichTextBox1height As Integer
Dim MaximizeYes As Boolean
'----------------------------------------------------------
'For Vb Sucking menus
Dim FileVbOpened As String
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Function GetWindowsDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDirectory = AddBackslash(Left$(s, i - 1))
End Function
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function
'above is not esential for this project, allso not mycode


Private Function GetItemLevel(ItemCaption As String)
If Left(ItemCaption, 1) <> "-" Then
GetItemLevel = 0
Else:
    For i = 0 To Len(ItemCaption)
           If Mid(ItemCaption, i + 1, 1) <> "-" Then GetItemLevel = Counter / 4: Exit For
           Counter = Counter + 1
    Next i
End If
End Function
Private Function GetItemCaption(ItemCaption As String)
If Left(ItemCaption, 1) <> "-" Then
GetItemCaption = ItemCaption
Else:
    For i = 0 To Len(ItemCaption)

           If Mid(ItemCaption, i + 1, 1) <> "-" Then GetItemCaption = Right(ItemCaption, Len(ItemCaption) - Counter): Exit For
           Counter = Counter + 1
    Next i
End If

End Function
Private Sub Command1_Click()
'PrivateDeclaresForCode
Open App.Path & "\MenuPrivateDeclares.txt" For Output As #1    ' Open file for output

'1.declaration section
Print #1, "'Declaration section"

Print #1, "Private Declare Function CreatePopupMenu Lib" & Chr(34) & "user32.dll" & Chr(34) & " ()  As Long"

Print #1, "Private Declare Function DestroyMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " (ByVal hMenu As Long) As Long"

Print #1, "Private Type MENUITEMINFO"

Print #1, "        cbSize As Long"

Print #1, "        fMask As Long"

Print #1, "        fType As Long"

Print #1, "        fState As Long"

Print #1, "        wID As Long"

Print #1, "        hSubMenu As Long"

Print #1, "        hbmpChecked As Long"

Print #1, "        hbmpUnchecked As Long"

Print #1, "        dwItemData As Long"

Print #1, "        dwTypeData As String"

Print #1, "        cch As Long"

Print #1, "End Type"


'Constant Definitions
 
Print #1, "Private Const MIIM_STATE = &H1"

Print #1, "Private Const MIIM_ID = &H2"

Print #1, "Private Const MIIM_SUBMENU = &H4"

Print #1, "Private Const MIIM_CHECKMARKS = &H8"

Print #1, "Private Const MIIM_DATA = &H20"

Print #1, "Private Const MIIM_TYPE = &H10"

Print #1, "Private Const MFT_BITMAP = &H4"

Print #1, "Private Const MFT_MENUBARBREAK = &H20"

Print #1, "Private Const MFT_MENUBREAK = &H40"

Print #1, "Private Const MFT_OWNERDRAW = &H100"

Print #1, "Private Const MFT_RADIOCHECK = &H200"

Print #1, "Private Const MFT_RIGHTJUSTIFY = &H4000"

Print #1, "Private Const MFT_RIGHTORDER = &H2000"

Print #1, "Private Const MFT_SEPARATOR = &H800"

Print #1, "Private Const MFT_STRING = &H0"

Print #1, "Private Const MFS_CHECKED = &H8"

Print #1, "Private Const MFS_DEFAULT = &H1000"

Print #1, "Private Const MFS_DISABLED = &H2"

Print #1, "Private Const MFS_ENABLED = &H0"

Print #1, "Private Const MFS_GRAYED = &H1"

Print #1, "Private Const MFS_HILITE = &H80"

Print #1, "Private Const MFS_UNCHECKED = &H0"

Print #1, "Private Const MFS_UNHILITE = &H0"

'functions = API-s

Print #1, "Private Declare Function InsertMenuItem Lib " & Chr(34) & "user32.dll" & Chr(34) & " Alias " & Chr(34) & "InsertMenuItemA" & Chr(34) & " _"

Print #1, "(ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long"

Print #1, "Private Declare Function TrackPopupMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " _"

Print #1, "(ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long," & " _"

Print #1, "ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long) As Long"

Print #1, "Private Const TPM_RIGHTALIGN = &H8&"

Print #1, "Private Const TPM_CENTERALIGN = &H4&"

Print #1, "Private Const TPM_LEFTALIGN = &H0"

Print #1, "Private Const TPM_TOPALIGN = &H0"

Print #1, "Private Const TPM_NONOTIFY = &H80"

Print #1, "Private Const TPM_RETURNCMD = &H100"

Print #1, "Private Const TPM_LEFTBUTTON = &H0"

Print #1, "Private Const  TPM_RIGHTBUTTON = &H2&"

Print #1, "Private Type POINT_TYPE"

Print #1, "x As Long"

Print #1, "y As Long"

Print #1, "End Type"

Print #1, "Private Declare Function GetCursorPos Lib " & Chr(34) & "user32.dll" & Chr(34) & " (lpPoint As POINT_TYPE) As Long"

Print #1, "Private Declare Function AppendMenu Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "AppendMenuA" & Chr(34) & " (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long"

Print #1, "Private Declare Function SetMenuItemBitmaps Lib " & Chr(34) & "user32" & Chr(34) & " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long"

Print #1, "'Menu colors"
Print #1, "Private Declare Function SetSysColors Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long"
Print #1, "Private Declare Function GetSysColor Lib " & Chr(34) & "user32" & Chr(34) & "(ByVal nIndex As Long) As Long"
Print #1, "Private Const COLOR_MENUTEXT = 7"
Print #1, "Private Const COLOR_MENU = 4"

Print #1, "'Sys menu stuff"
Print #1, "Private Declare Function GetSystemMenu Lib "; user32; " (ByVal hWnd As Long, ByVal bRevert As Long) As Long"
Print #1, "Private Declare Function GetMenuItemCount Lib "; user32; " (ByVal hMenu As Long) As Long"
Print #1, "Private Declare Function GetMenuItemInfo Lib "; user32; " Alias "; GetMenuItemInfoA; " (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long"
'''''''''''''''''
Print #1, "Private Declare Function SetWindowLong Lib " & Chr(34) & " user32" & Chr(34) & "Alias" & Chr(34) & " SetWindowLongA" & Chr(34) & " (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long"
Print #1, "Private Const GWL_WNDPROC = -4"
Print #1, "Private Declare Function CallWindowProc Lib " & Chr(34) & " user32.dll" & Chr(34) & "; Alias" & Chr(34) & " CallWindowProcA" & Chr(34) & " (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long"
Print #1, "Private Const WM_SYSCOMMAND = &H112"
Print #1, "Private Const WM_INITMENU = &H116"

Print #1, "Private Declare Function RemoveMenu Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long"
Print #1, "Private Const MF_REMOVE = &H1000&"

Print #1, "Private Type RECT"
Print #1, "    Left As Long"
Print #1, "    Top As Long"
Print #1, "    Right As Long"
Print #1, "    Bottom As Long"
Print #1, "End Type"

Print #1, "Private Declare Function DrawMenuBar Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long) As Long"
Print #1, "Private Declare Function CreateRectRgn Lib " & Chr(34) & " gdi32" & Chr(34) & " (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
Print #1, "Private Declare Function GetWindowRect Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, lpRect As RECT) As Long"
Print #1, "Private Declare Function PtInRegion Lib " & Chr(34) & " gdi32" & Chr(34) & " (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long"

'for complete replacement of sys menu
Print #1, "Private Declare Function PeekMessage Lib " & Chr(34) & " user32" & Chr(34) & "; Alias" & Chr(34) & "; PeekMessageA" & Chr(34) & " (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long"
Print #1, "Private Declare Function WaitMessage Lib " & Chr(34) & " user32" & Chr(34) & " () As Long"
Print #1, "Private Declare Function GetMessage Lib " & Chr(34) & " user32" & Chr(34) & "; Alias" & Chr(34) & "; GetMessageA" & Chr(34) & " (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long"
Print #1, "Private Declare Function SetTimer Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long"
Print #1, "Private Declare Function KillTimer Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long"

Print #1, "Private Type POINTAPI"
Print #1, "    x As Long"
Print #1, "    y As Long"
Print #1, "End Type"
Print #1, "Private Type Msg"
Print #1, "    hwnd As Long"
Print #1, "    Message As Long"
Print #1, "    wParam As Long"
Print #1, "    lParam As Long"
Print #1, "    time As Long"
Print #1, "    pt As POINTAPI"
Print #1, "End Type"

Print #1, "Private Const PM_REMOVE = &H1"
Print #1, "Private Const PM_NOREMOVE = &H0"

'Needed to add buttons to caption bar
Print #1, "Private Const DFC_CAPTION = 1"

Print #1, "Private Const DFCS_CAPTIONCLOSE = &H0"
Print #1, "Private Const DFCS_CAPTIONMAX = &H2"
Print #1, "Private Const DFCS_CAPTIONMIN = &H1"
Print #1, "Private Const DFCS_CAPTIONRESTORE = &H3"

Print #1, "Private Const DFCS_PUSHED = &H200"

Print #1, "Private Const SM_CYCAPTION = 4"


Print #1, "Private Declare Function DrawFrameControl Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long"
Print #1, "Private Declare Function GetSystemMetrics Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal nIndex As Long) As Long"
Print #1, "Private Declare Function GetWindowDC Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long) As Long"
Print #1, "Private Declare Function ReleaseDC Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal hdc As Long) As Long"
Print #1, "Private Declare Function SetRect Lib " & Chr(34) & " user32" & Chr(34) & " (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"


Close #1

RichTextBox1.LoadFile App.Path & "\MenuPrivateDeclares.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText

MsgBox "Private declares are sitting on clipboard and waiting for you, anyway you can look at them in below text box.", , "M.C. Menu Voodoo "

End Sub

Private Sub ApiPopUpMenuCodeGenerator() 'code maker

End Sub

Private Sub MCMenuVoodoo(process As String)

'ApiPopUpMenuCodeGenerator

SelectionButton(0).Enabled = False
SelectionButton(1).Enabled = False
SelectionButton(2).Enabled = False
SelectionButton(3).Enabled = False

SelectionButton(2).Caption = "General section of form"

'avoid disaster
If List1.ListCount = 0 Then
MsgBox " No items in list box, ending....", , "M.C. Menu Voodoo "
Exit Sub
End If

'avoid another disaster
If MapFileUniqueID = "" Then
MsgBox "Structure is not saved jet, save it and then countinue. If the structure was sucked from vb menus, repeat the whole process again, starting from sucking. ", , "M.C. Menu Voodoo "
Exit Sub
End If


Dim magicnumber As Integer

magicnumber = 1000 ' start value of it, could be any other that comes to your mind"
MenuAppendNum = 1


'SAVE STUFF INTO TABLE --------------------------------------------------------------------

'dbs stuff---------------------
Dim dbs As Database
Dim rst As Recordset
Set dbs = OpenDatabase(App.Path & "\my.mdb")
Set rst = dbs.OpenRecordset("SaveConstruct")
'------------------------------
' Clear table.
dbs.Execute "DELETE * FROM SaveConstruct ;"

For i = 0 To List1.ListCount - 1
With rst


.AddNew
'ControlIndex
![ControlIndex] = i
'items string
![Container] = List1.List(i)
'level
![Level] = GetItemLevel(List1.List(i))
'caption = cut off ---- signs from string if any
![Caption] = GetItemCaption(List1.List(i))

'-----inserting this in version 2.00----TRANSLATION-----------------------------
Dim TypeOutputString As String
Dim StateOutputString As String

'1.first translate state
'Type translated......
        TypeOutputString = "MFT_STRING"
        
        Select Case Mid(List2.List(i), 2, 1)
        Case "R"
        TypeOutputString = TypeOutputString & " Or MFT_RADIOCHECK"
        Case "C"
        'do nothing
        Case Else
        End Select

        Select Case Mid(List2.List(i), 3, 1)
        Case "L"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBARBREAK"
        Case "N"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBREAK"
        Case Else
        End Select

'2.second translate state
'State translated......
        Select Case Mid(List3.List(i), 1, 1)
        Case "E"
        StateOutputString = "MFS_ENABLED"
        Case "D"
        StateOutputString = "MFS_DISABLED"
        Case "G"
        StateOutputString = "MFS_GRAYED"
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 2, 1)
        Case "B"
        StateOutputString = StateOutputString & " Or MFS_DEFAULT"
        Case "N"
        'do nothing
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 3, 1)
        Case "C"
        'this disabled in ver 3.0 - reason: substantial shrinking of output code with help of IIf function
        'StateOutputString = StateOutputString & " Or MFS_CHECKED"
        ![Checked] = "MFS_CHECKED"
        Case "U"
        'this disabled in ver 3.0 - reason: substantial shrinking of output code with help of IIf function
        'StateOutputString = StateOutputString & " Or MFS_UNCHECKED"
        ![Checked] = "MFS_UNCHECKED"
        Case Else
        End Select
        
'-------------------END OF TRANSLATION---------------------------------------------------


'-----inserting this in version 3.00----variable naming for checked items-----------------

If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
VariableIndex = VariableIndex + 1
stringvar = CStr(VariableIndex)
![CheckedVarIndex] = MapFileUniqueID & stringvar
End If

'-----end of variable naming for checked items----------------------------------------------


'type of
        If GetItemCaption(List1.List(i)) = "/separator/" Then
        ![ItemType] = "MFT_SEPARATOR"
        Else
        ![ItemType] = TypeOutputString
        End If
'State
![ItemState] = StateOutputString

'Picture Yes/no
![Picture] = List4.List(i)

'Place = order in which items in menu should appear
        b = 0
        c = 0
        a = GetItemLevel(List1.List(i))
        Do
        If i = 0 Then ![Place] = 0: Exit Do
        b = b + 1
        If i - b < 0 Then ![Place] = c: Exit Do

                Select Case GetItemLevel(List1.List(i - b))
                Case Is = a
                c = c + 1
                Case Is < a
                ![Place] = c
                c = 0
                Exit Do
                Case Is > a
            
                Case Else
                Beep '? - if this happen then throw your computer away, lol
                End Select
        Loop Until i - b < 0
.Update
End With
Next i
'END SAVING  STUFF INTO TABLE


DoEvents
'dbs stuff---------------------
'Dim rst1 As Recordset
'Set rst1 = dbs.OpenRecordset("SaveConstruct")

    
'find max level in list1
MaxItemLevel = 0
For i = 0 To List1.ListCount
        If GetItemLevel(List1.List(i)) > MaxItemLevel Then
        MaxItemLevel = GetItemLevel(List1.List(i))
        End If
Next i


s = MaxItemLevel
'now read from bottom to top & get data into table
For j = 0 To s
       rst.MoveLast
       d = rst![Caption]
       Do
                     If rst![Level] = MaxItemLevel Then
                        Do
                           If rst![Level] = MaxItemLevel Then
                                With rst
                                .Edit
                                 ![MenuNum] = MenuAppendNum 'which menu item fits in
                                 ![MenuName] = "hPopupMenu" & MenuAppendNum 'which menu item fits in
                                 ![ItemMagicNum] = magicnumber 'allso for separator
                                .Update
                                End With
                          magicnumber = magicnumber + 1
                           End If
                           rst.MovePrevious
                        If rst.BOF Then Exit For
                        Loop Until rst![Level] < MaxItemLevel
                   rst.MoveNext
                   MenuAppendNum = MenuAppendNum + 1
                   End If
        rst.MovePrevious
        Loop Until rst.BOF
MaxItemLevel = MaxItemLevel - 1
'now next level
Next j

'Enter sub menus data - which menu is to be opened when mouse moves over specific item
For i = 0 To List1.ListCount

        Select Case GetItemCaption(List1.List(i))
        Case Is = "/separator/" 'separators can't have submenus !
            
                       If GetItemLevel(List1.List(i + 1)) > GetItemLevel(List1.List(i)) Then 'problem at sight
                       t = i
                       Do
                       t = t - 1
                       Loop Until GetItemCaption(List1.List(t)) <> "/separator/"
                       'now we have menu (t) and it's submenu (i+1)... get it into database
                        
                        rst.MoveFirst
                        Do
                                If rst![Caption] = GetItemCaption(List1.List(i + 1)) And rst![ControlIndex] = i + 1 Then
                                SubMen = rst![MenuNum]
                                Exit Do
                                End If
                                rst.MoveNext
                        Loop
                        
                        rst.MoveFirst
                        Do
                                If rst![Caption] = GetItemCaption(List1.List(t)) And rst![ControlIndex] = t Then
                                     With rst
                                     .Edit
                                     ![ItemSubMenu] = SubMen
                                     .Update
                                     End With
                                     Exit Do
                                End If
                        rst.MoveNext
                        Loop
                
                
                Else: GoTo 11 'do nothing as 'separators can't have submenus !
                End If
        
        Case Else
        If GetItemLevel(List1.List(i + 1)) = (GetItemLevel(List1.List(i)) + 1) Then
        'if next one level higher as curent item then .....
                rst.MoveFirst
                Do
                        If rst![Caption] = GetItemCaption(List1.List(i + 1)) And rst![ControlIndex] = i + 1 Then
                        SubMen = rst![MenuNum]
                        Exit Do
                        End If
                        rst.MoveNext
                Loop
                
                rst.MoveFirst
                Do
                        If rst![Caption] = GetItemCaption(List1.List(i)) And rst![ControlIndex] = i Then
                             With rst
                             .Edit
                             ![ItemSubMenu] = SubMen
                             .Update
                             End With
                             Exit Do
                        End If
                rst.MoveNext
                Loop
       End If
       End Select
11
Next i

'--------------------------------------------------------------------------------------------
'added into ver 3.0
'Coplete Path Finder Procedure
'Before writing code - get some data that will place output code coments on steroids
'rst1 is steel opened ....., so use it
rst.MoveFirst
For i = 0 To List1.ListCount - 1
    depth = rst![Level]
    Select Case depth
    Case 0
            With rst
            .Edit
            ![CompletePath] = ![Caption]
            .Update
            End With
    Case Else
                  Dim FCP As String 'final complete path
                  FCP = GetItemCaption(List1.List(i))
                          f = i
                          CurrentLevel = GetItemLevel(List1.List(i))
                          Do
                                 f = f - 1
                                 If GetItemLevel(List1.List(f)) < CurrentLevel And _
                                 GetItemCaption(List1.List(f)) <> "/Separator/" Then
                                 FCP = GetItemCaption(List1.List(f)) & "/" & FCP
                                 End If
                                 CurrentLevel = GetItemLevel(List1.List(f))
                                 
                          Loop Until GetItemLevel(List1.List(f)) = 0 'at 0 exit loop
                                'path is complete - Print it down
                                With rst
                                .Edit
                                ![CompletePath] = FCP
                                .Update
                                End With
                                                    
    End Select
    
rst.MoveNext
Next i

'--------------------------------------------------------------------------------------------
'START WRITING CODE
'--------------------------------------------------------------------------------------------

'only god knows why these lines are here, but they must be !
'comment made one month after ver ... coding - learn from it, lol
If controler = 0 Then
'Text2.SetFocus
'Text2.Text = ""
controler = 1
End If

    'lets see how many menus is to be created
    
     NumOfMenusToBeCreated = 1
     rst.MoveFirst
     For i = 0 To List1.ListCount - 1
     If rst![MenuNum] > NumOfMenusToBeCreated Then NumOfMenusToBeCreated = rst![MenuNum]
     rst.MoveNext
     Next i
     
'here is where output code will go....
Open App.Path & "\MenuMainCode.txt" For Output As #1    ' Open file for output

 

OCLC = 0 'code lines counter
'2. The core of the thing

      Print #1, "'--------------------------------------------------------------------------------------------------------------------------"
      Print #1, "'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo "
      Print #1, "'Menu identification number: " & MapFileUniqueID
      Print #1, "'Structure saved in: " & CurentMapFilePathAndName
      Print #1, "'---------------------------------------------------------------------------------------------------------------------------"
      OCLC = OCLC + 3
            Print #1, "' Handles to the popup menus to display"
            
            For i = 1 To NumOfMenusToBeCreated
            Print #1, "Dim hPopupMenu" & i & " As Long"
            OCLC = OCLC + 1
            Next i
                
         
      Print #1, "Dim mii1 As MENUITEMINFO ' Structure that will describe menu items"
      Print #1, "Dim curpos As POINT_TYPE  ' holds the current mouse coordinates"
      Print #1, "Dim menusel As Long       ' ID of what the user selected in the popup menu"
      Print #1, "Dim retval As Long        ' generic return value"
      OCLC = OCLC + 4
      
    'Create the popup menus which are initialy empty.
    Print #1, ""
    Print #1, "'Create the popup menus which are initialy empty."
    Print #1, ""
    
    For i = 1 To NumOfMenusToBeCreated
    Print #1, "hPopupMenu" & i & " = CreatePopupMenu()"
    OCLC = OCLC + 1
    Next i
    
    
      'Structure of menus to be displayed
      Print #1, ""
      Print #1, "'Create the structure which is the base for all menus:"
      
      Print #1, "With mii1"
      
      Print #1, ".cbSize = Len(mii1)' The size of this structure."
      
      Print #1, ".fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU' Which elements of the structure to use."
      
      Print #1, "End With"
      Print #1, ""
      OCLC = OCLC + 4
       
DoEvents

rst.MoveFirst



'get items into created menus ! & describe their properties"

Print #1, "'get items into created menus & describe their properties"
Print #1, ""
'-------------------------------------------
'-------------------------------------------
'-------------------------------------------
For i = 0 To List1.ListCount - 1
            Print #1, "With mii1" & "'(" & rst![CompletePath] & ")"
            
            'type thing
            Print #1, ".fType =" & rst![ItemType]
            
            'state thing
            Select Case rst![Checked]
            Case "MFS_UNCHECKED", "MFS_CHECKED"
            Print #1, ".fState =" & rst![ItemState] & " Or IIf( APICheck" & rst![CheckedVarIndex] & ", " & " MFS_CHECKED" & ", " & " MFS_UNCHECKED" & ")"
            Case Else 'normal item not checked at all
            Print #1, ".fState =" & rst![ItemState]
            End Select
            
            Print #1, ".wID =" & rst![ItemMagicNum]
            Print #1, ".dwTypeData =" & Chr(34) & rst![Caption] & Chr(34) ' Display the following text for the item."
            Print #1, ".cch = Len(" & Chr(34) & rst![Caption] & Chr(34) & ")"
            
            'if there is a submenu for this item then......
            If rst![ItemSubMenu] > 0 Then
            Print #1, ".hSubMenu = hPopupMenu" & rst![ItemSubMenu]
            Else
            Print #1, ".hSubMenu = 0"
            End If
            Print #1, "End With"

      'well where to send the item ?
      Print #1, "retval = InsertMenuItem(" & rst![MenuName] & "," & rst![Place] & ",1, mii1" & ")"
      Print #1, ""
      OCLC = OCLC + 9
      rst.MoveNext
Next i


Print #1, "'The following code is for adding pictures into menus, if there are any!"
Print #1, "'------------------------------------------------------------"


'HERE ADD PICTURES IF ANY
rst.MoveFirst 'back to beginning
Dim picindex As Integer
Dim PicMsgBoxWarning As Boolean
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst![Picture] = "Yes" Then
                If Mid(List2.List(i), 2, 1) = "P" Then
                
                Print #1, "retval = SetMenuItemBitmaps(" & rst![MenuName] & "," & rst![ItemMagicNum] & ",1, MenuPicBox" & MapFileUniqueID & "(" & picindex & ")," & "MenuPicBox" & MapFileUniqueID & "(" & picindex + 1 & "))"
                OCLC = OCLC + 1
                picindex = picindex + 2
                Else
                
                Print #1, "retval = SetMenuItemBitmaps(" & rst![MenuName] & "," & rst![ItemMagicNum] & ",1, MenuPicBox" & MapFileUniqueID & "(" & picindex & ")," & "MenuPicBox" & MapFileUniqueID & "(" & picindex & "))"
                OCLC = OCLC + 1
                picindex = picindex + 1
                
                End If
                PicMsgBoxWarning = True
        End If
rst.MoveNext
Next i


Print #1, "'------------------------------------------------------------"




   
 
         'see what are the general menu behaviour settings
                       Dim BehaviourString As String
                       BehaviourString = "TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD"
                          
                       If Option6(0).Value = True Then BehaviourString = BehaviourString & " Or TPM_RIGHTALIGN"
                       If Option6(1).Value = True Then BehaviourString = BehaviourString & " Or TPM_LEFTALIGN"
                       If Option6(2).Value = True Then BehaviourString = BehaviourString & " Or TPM_CENTERALIGN"

                       If Option7(0).Value = True Then BehaviourString = BehaviourString & " Or TPM_LEFTBUTTON"
                       If Option7(1).Value = True Then BehaviourString = BehaviourString & " Or TPM_RIGHTBUTTON"
        
        'colorized menu stuff
        If Option1(1).Value = True Then 'case there are colors
        Print #1, "DefaultMenuColor = GetSysColor(COLOR_MENU)"
        Print #1, "DefaultMenuTextColor = GetSysColor(COLOR_MENUTEXT)"
        'set to new
        Print #1, "SetSysColors 1, COLOR_MENU," & Label5(2).BackColor
        Print #1, "SetSysColors 1, COLOR_MENUTEXT," & Label5(2).ForeColor
        OCLC = OCLC + 4
        End If
        
        ' Determine where the mouse cursor currently is, in order to have
        ' the popup menu appear at that point.
         Print #1, ""
         Print #1, "'LET'S FINALY SHOW THE MENU"
         
         
         Print #1, "retval = GetCursorPos(curpos)"
         
        ' Display the popup menu at the mouse cursor.  Instead of sending messages
        ' to window Form1, have the function merely return the ID of the user's selection.
        Print #1, "menusel = TrackPopupMenu(hPopupMenu" & NumOfMenusToBeCreated & "," & BehaviourString & ",curpos.x, curpos.y, 0, Form1.hWnd, 0 )"
        
      ' Before acting upon the user's selection, destroy the popup menu now.
        Print #1, "retval = DestroyMenu(hPopupMenu" & NumOfMenusToBeCreated & ")"
        OCLC = OCLC + 3
        
        'colorized menu stuff - return to default
        If Option1(1).Value = True Then 'case there are colors
        Print #1, "SetSysColors 1, COLOR_MENU, DefaultMenuColor"
        Print #1, "SetSysColors 1, COLOR_MENUTEXT, DefaultMenuTextColor"
        OCLC = OCLC + 2
        End If
        
        
        'create event handling
        Print #1, ""
        Print #1, "'------------------------------------------------------------------------------------------------"
        
        Print #1, "'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!"
        
        Print #1, "'------------------------------------------------------------------------------------------------"
        
        Print #1, "Select Case menusel"
                      OCLC = OCLC + 1
                      
                         rst.MoveFirst
                         For i = 0 To List1.ListCount - 1
                         Select Case rst![Caption]
                         Case "/separator/" 'separator = do nothing as you can't select it
                         Case Else
                            If rst![ItemSubMenu] = 0 Then
                                    
                                    Print #1, "Case " & rst![ItemMagicNum] & "'(" & rst![CompletePath] & ")"
                                    OCLC = OCLC + 1
                                    'now what to do ?
                                    
                                    
                                        
                                        Select Case process
                                        '***********************************
                                        Case "NormalCode"
                                        '***********************************
                                            'if it is marked checked
                                            If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                            markmessage = True
                                            
                                            Print #1, "'Do change appropriate variable value now"
                                            Print #1, "APICheck" & rst![CheckedVarIndex] & " = IIf( APICheck" & rst![CheckedVarIndex] & ", " & " False" & ", " & " True" & ")"
                                            Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                            Print #1, "Stop"
                                            Print #1, "'And for your convinience....."
                                            Print #1, "If APICheck" & rst![CheckedVarIndex] & " = True Then'in case item become unchecked ....."
                                            Print #1, "Else 'in case item become checked ....."
                                            Print #1, "End If"
                                            OCLC = OCLC + 6
                                            
                                            Else:
                                            Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                            Print #1, "Stop"
                                            OCLC = OCLC + 1
                                            End If
                                            
                                        '***********************************
                                        Case "CodeWithCodeFromFrm"
                                        '***********************************
                                            If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                            markmessage = True
                                            End If
                                            
                                            Open App.Path & "\MenuAPIMenuSuckedCode.txt" For Input As #2
                                            Do While Not EOF(2)
                                            Line Input #2, textline
                                                    If Left(textline, 4) = "Case" And Right(textline, Len("(" & rst![CompletePath] & ")")) = "(" & rst![CompletePath] & ")" Then
                                                         Do
                                                         Line Input #2, textline
                                                                If Left(textline, 4) = "Case" Then GoTo avoid
                                                         Print #1, textline
                                                         OCLC = OCLC + 1
                                                         Loop
                                                    End If
                                            Loop
                                            Close #2
                                            
                                            'We come to end of file so this must be a new item ...
                                                        
                                                        'if it is marked checked
                                                        If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                                        markmessage = True
                                                        Print #1, "'Do change appropriate variable value now"
                                                        Print #1, "APICheck" & rst![CheckedVarIndex] & " = IIf( APICheck" & rst![CheckedVarIndex] & ", " & " False" & ", " & " True" & ")"
                                                        Print #1, "Stop"
                                                        Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                                        Print #1, "'And for your convinience....."
                                                        Print #1, "If APICheck" & rst![CheckedVarIndex] & " = True Then'in case item become unchecked ....."
                                                        Print #1, "Else 'in case item become checked ....."
                                                        Print #1, "End If"
                                                        OCLC = OCLC + 6
                                                        Else
                                                        Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                                        Print #1, "Stop"
                                                        OCLC = OCLC + 1
                                                        End If
                                                        
avoid:
                                            Close #2
                                            
                                        '***********************************
                                        Case "VBMenuSuckedCode"
                                        '***********************************
                                        If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                            markmessage = True
                                            End If
                                            
                                            Open App.Path & "\MenuVBMenuSuckedCode.txt" For Input As #2
                                            Do While Not EOF(2)
                                            Line Input #2, textline
                                                    If textline = List7.List(i) Then
                                                         Do
                                                         Line Input #2, textline
                                                                If textline = "End Sub" Then GoTo avoid1
                                                         Print #1, textline
                                                         OCLC = OCLC + 1
                                                         Loop
                                                    End If
                                            Loop
                                            Close #2
                                            
                                            'We come to end of file so this must be a new item ...
                                                        
                                                        'if it is marked checked
                                                        If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                                        markmessage = True
                                                        Print #1, "'Do change appropriate variable value now"
                                                        Print #1, "APICheck" & rst![CheckedVarIndex] & " = IIf( APICheck" & rst![CheckedVarIndex] & ", " & " False" & ", " & " True" & ")"
                                                        Print #1, "Stop"
                                                        Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                                        Print #1, "'And for your convinience....."
                                                        Print #1, "If APICheck" & rst![CheckedVarIndex] & " = True Then'in case item become unchecked ....."
                                                        Print #1, "Else 'in case item become checked ....."
                                                        Print #1, "End If"
                                                        OCLC = OCLC + 6
                                                        Else
                                                        Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                                        Print #1, "Stop"
                                                        OCLC = OCLC + 1
                                                        End If
                                                        
avoid1:
                                            Close #2
                                        End Select
                             End If
                         End Select
                         rst.MoveNext
                         Next i
       
       Print #1, "Case Else"
       
       Print #1, "End Select"
       OCLC = OCLC + 2
       
   Close #1

SelectionButton(0).Enabled = True

'MESSAGE PROCEDURE


If PicMsgBoxWarning = True Then 'comments for pictures
PicMsgBoxWarning = False
SelectionButton(3).Enabled = True
'here is where code will go....
Open App.Path & "\MenuPicturesComment.txt" For Output As #2    ' Open file for output

Print #2, "'You have choosen to have some pics on your menu"
Print #2, "'so you have to put following array of picture controls into your projects."
Print #2, "'Comments which picture fits into specific array element are allso here."
OCLC = OCLC + 3


rst.MoveFirst
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst![Picture] = "Yes" Then
                If Mid(List2.List(i), 2, 1) = "P" Then
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")" & " ,When Checked"
                        
                        picindex = picindex + 1
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")" & " ,When UnChecked"
                        
                        picindex = picindex + 1
                        
                        Else
                        
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")"
                        
                        picindex = picindex + 1
                        
                End If
    End If
rst.MoveNext
Next i
Close #2
End If

'if you have some items checked
If markmessage = True Then
markmessage = False
SelectionButton(1).Enabled = True
SelectionButton(2).Enabled = True
Open App.Path & "\MenuGeneralSectionCode.txt" For Output As #2    ' Open file for output

Print #2, "'--------------------------------------------------------------------------------------------------------------------------"
Print #2, "'Generall section declarations for M.C. Menu Voodoo:"
Print #2, "'Menu identification number: " & MapFileUniqueID
Print #2, "'Structure saved in: " & CurentMapFilePathAndName
Print #2, "'---------------------------------------------------------------------------------------------------------------------------"



'general section of form
rst.MoveFirst

For i = 0 To List1.ListCount - 1
      If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
     
      Print #2, "Dim APICheck" & rst![CheckedVarIndex] & " as Boolean" & "'" & rst![CompletePath]
      OCLC = OCLC + 1
      End If
rst.MoveNext
Next i
Print #2, "'--------------------------------------------------------------------------------------------------------------------------"
Print #2, "'End of Generall section declarations for M.C. Menu Voodoo"
Print #2, "'--------------------------------------------------------------------------------------------------------------------------"

Close #2


Open App.Path & "\MenuFormLoadCode.txt" For Output As #2    ' Open file for output
Print #2, "'--------------------------------------------------------------------------------------------------------------------------"
Print #2, "'Form_Load initial variables values for M.C. Menu Voodoo:"
Print #2, "'Menu identification number: " & MapFileUniqueID
Print #2, "'Structure saved in: " & CurentMapFilePathAndName
Print #2, "'This is not neccesary, only if you like it"
Print #2, "'---------------------------------------------------------------------------------------------------------------------------"

'form_load event
rst.MoveFirst

For i = 0 To List1.ListCount - 1
      If rst![Checked] = "MFS_CHECKED" Then
 
      Print #2, "APICheck" & rst![CheckedVarIndex] & " = True" & "'" & rst![CompletePath]
      OCLC = OCLC + 1
      ElseIf rst![Checked] = "MFS_UNCHECKED" Then
  
      Print #2, "APICheck" & rst![CheckedVarIndex] & " = False" & "'" & rst![CompletePath]
      OCLC = OCLC + 1
      End If
rst.MoveNext
Next i
Print #2, "'--------------------------------------------------------------------------------------------------------------------------"
Print #2, "'End of Form_Load initial variables values for M.C. Menu Voodoo:"
Print #2, "'--------------------------------------------------------------------------------------------------------------------------"

Close #2
End If

'end of message procedure




'and on the end, dbs closing stuff
rst.Close
Set rst = Nothing
dbs.Close

'next one solves hung up of this sub in case that you run it
'again while stuff is already in text2, forgot how  = did no comment, but it works
controler = 0
VariableIndex = 0 'this not beeing here caused big bug, that now disappeared


'show the MainCode to user
RichTextBox1.LoadFile App.Path & "\MenuMainCode.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText

Select Case process
Case "NormalCode"
MsgBox "Main code has been automaticaly copyed to clipboard. Clicking any other buttons above outputtext box does the same." & Chr(13) & OCLC & " lines of code generated.", , "M.C. Menu Voodoo "
Case "CodeWithCodeFromFrm"
MsgBox "This is a replacement code for previously existed menu code i.e. Main one allso include code that you have manualy written in previous version of this menu structure." & Chr(10) & "If you meanwhile deleted some items, the code that existed there is no longer here." & Chr(10) & "If you added some new items, clicking them at run time will bring you to Stop command." & Chr(10) & "Replace all parts of code in your project i.e. main code, general section, ... !" & Chr(13) & OCLC & " lines of code generated.", , "M.C. Menu Voodoo "
End Select

End Sub

Private Sub Command10_Click()
MCMenuVoodoo ("NormalCode")
End Sub

Private Sub Command11_Click() 'exit
Unload Form1

End Sub









Private Sub Command12_Click()
WinHelp Me.hwnd, App.Path & "\Voodoo.hlp", HELP_CONTENTS, ByVal 0
End Sub

Private Sub Command14_Click()


'erase if anything already there
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List6.Clear

'OPEN FILE WITH CDLG NO OCX INCLUDED

'first dimension is not to be changed, second is to be set to num of filters that
'you will add to dialog - 1 !
Dim ArrayWithFileFilters(2, 1) As String 'array holfding file filters
'first filter & its description
ArrayWithFileFilters(1, 0) = "Frm Files"
ArrayWithFileFilters(2, 0) = "*.frm"
'second filter & its description
ArrayWithFileFilters(1, 1) = "All Files"
ArrayWithFileFilters(2, 1) = "*.*"
'x filter & its description............

If LastPathOpened = "" Then LastPathOpened = App.Path
'now call our dialog public function
Dim ReturnedFile As String
ReturnedFile = CdlgFileToOpenOrSave("Open", Me, ArrayWithFileFilters, LastPathOpened, "M.C. Menu Voodoo")
If ReturnedFile = "" Then Exit Sub 'if cancel was pressed
'END OF OPEN FILE WITH CDLG NO OCX INCLUDED
FileVbOpened = ReturnedFile

CurrentOpenedFile = SuckOutFileName(ReturnedFile)
Me.Caption = "(" & CurrentOpenedFile & ")" & " M.C. Menu Voodoo"
LastPathOpened = SuckOutFilePath(ReturnedFile)



Open ReturnedFile For Input As #1     ' Open file for output.


Line Input #1, textline
If textline <> "VERSION 5.00" Then MsgBox "Probably wont work fine as it is not correct (vb 6.0) frm,  but hey why not try it anyway", , "M.C. Menu Voodoo "


List6.AddItem "'Take Them All'"
Do While Not EOF(1)
Line Input #1, textline 'get line into variable
TrimmedTextLine = LTrim(textline)

        If Left(TrimmedTextLine, 13) = "Begin VB.Menu" Then 'menu detected
        'calculate the difference ... which give as depth of item ..
        depth = Len(textline) - Len(TrimmedTextLine)
        realdepth = (depth / 3) - 1 'the depth of item
        
        'get caption which is in next line
        Line Input #1, textline 'this one should hold item caption
        TrimmedTextLine = LTrim(textline)
        ICaption = Right(TrimmedTextLine, Len(TrimmedTextLine) - 21)
        ICaption = Mid(ICaption, 1, Len(ICaption) - 1) 'some corrections
                     
                     If realdepth = 0 Then
                     List6.AddItem ICaption
                     End If
        End If
Loop

Close #1

If List6.ListCount = 1 Then
List6.Clear
MsgBox "No VB menus found in this frm !", , "M.C. Menu Voodoo "
Exit Sub
End If

Picture2.Visible = True
End Sub

Private Sub Command15_Click()

'Check if there is a menu in list1
If List1.ListCount = 0 Then
MsgBox " No menu structure opened, ending procedure....", , "M.C. Menu Voodoo "
Exit Sub
End If

'missing MapFileUniqueID
If MapFileUniqueID = "" Then
MsgBox "Missing MapFileUniqueID, ending...." & Chr(10) & "Causes and solutions:" & Chr(10) & "1. You have structure that is new and not saved ! " & Chr(10) & "2. You have structure that is sucked from vb menus and not saved ! ", , "M.C. Menu Voodoo "
Exit Sub
End If

'SAVE FILE WITH CDLG NO OCX INCLUDED

'first dimension of  ArrayWithFileFilters is not to be changed, second is to
'be set to num of filters that you will add to dialog - 1 !
Dim ArrayWithFileFilters(2, 1) As String 'array holfding file filters
'first filter & its description
ArrayWithFileFilters(1, 0) = "Frm Files"
ArrayWithFileFilters(2, 0) = "*.Frm"
'second filter & its description
ArrayWithFileFilters(1, 1) = "All Files"
ArrayWithFileFilters(2, 1) = "*.*"
'x filter & its description............

MsgBox "The purpose of this is to make replacement menu code for menu code that already exist in your project (in case u decided to change menu structure after you added tons of hand writen code). The code generated will contain allso your hand made code. The point is to replace code without mixing all up (at the same time enabling you to change items properties, adding items, deleting them...) and saving a lot of time. DO NOT CHANGE CAPTIONS OF PREVIOUSLY EXISTING ITEMS BEFORE RUNING THIS, because they are the orientation for this coding. Now select .frm file in which this menu already exist. New code will be generated."

'now call our dialog public function
Dim ReturnedFile As String
ReturnedFile = CdlgFileToOpenOrSave("Open", Me, ArrayWithFileFilters, App.Path, "M.C. Menu Voodoo")
'END OF Save FILE WITH CDLG NO OCX INCLUDED

If ReturnedFile = "" Then Exit Sub ' cancel was pressed

Open ReturnedFile For Input As #1     ' Open file for output.
'copy out existing API menu code from your frm
Do While Not EOF(1)
            Line Input #1, textline
            If Left(textline, 28) = "'Menu identification number:" And Right(textline, 10) = MapFileUniqueID Then
                    Do
                    Line Input #1, textline
                            If textline = "Select Case menusel" Then
                            Open App.Path & "\MenuAPIMenuSuckedCode.txt" For Output As #2    ' Open file for output.
                                    Do
                                    Line Input #1, textline
                                    Print #2, textline
                                    Loop Until textline = "End Sub"
                                    Close #2
                                    Close #1
                                    GoTo stepout '''''''''''
                            End If
                    Loop
            End If
Loop

'if code comes to this line, this is the sign that there is no menu with corect MENU ID !
MsgBox "Appropriate menu code with right MapFileUniqueID not found in this .frm. Ending procedure.....", , "M.C. Menu Voodoo "
Close #1
Exit Sub

stepout:
Close #1

MCMenuVoodoo ("CodeWithCodeFromFrm")

End Sub

Private Sub Command16_Click()
List6.Clear
Picture2.Visible = False
End Sub

Private Sub Command17_Click()
Command17.Enabled = False
MCMenuVoodoo ("VBMenuSuckedCode")
End Sub






Private Sub Command19_Click()
'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo
'Menu identification number: 0096294581
'Structure saved in: C:\ProjektiVB6\Api Menu(Ver31)\SysTwisterMenu.map
'---------------------------------------------------------------------------------------------------------------------------
' Handles to the popup menus to display
Dim hPopupMenu1 As Long
Dim mii1 As MENUITEMINFO ' Structure that will describe menu items
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value

'Create the popup menus which are initialy empty.

hPopupMenu1 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'get items into created menus & describe their properties

With mii1 '(Erase and replace)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002
.dwTypeData = "Erase and replace"
.cch = Len("Erase and replace")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1 '(Insert at the start)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001
.dwTypeData = "Insert at the start"
.cch = Len("Insert at the start")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii1 '(Add at the end)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000
.dwTypeData = "Add at the end"
.cch = Len("Add at the end")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, mii1)

'The following code is for adding pictures into menus, if there are any!
'------------------------------------------------------------
'------------------------------------------------------------

'LET'S FINALY SHOW THE MENU
retval = GetCursorPos(curpos)
curpos.y = Command19.Top + Picture5.Top + 21
curpos.x = Command19.Left + Picture5.Left
menusel = TrackPopupMenu(hPopupMenu1, TMP_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, curpos.x, curpos.y, 0, Form1.hwnd, 0)
retval = DestroyMenu(hPopupMenu1)

'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Select Case menusel
Case 1002 '(Erase and replace)
MCMenuVoodooSystem ("EraseAndReplace")
Case 1001 '(Insert at the start)
MCMenuVoodooSystem ("Insert")
Case 1000 '(Add at the end)
MCMenuVoodooSystem ("Add")
Case Else
End Select
End Sub

Private Sub MCMenuVoodooSystem(proccess As String)

'This code here is basicaly a copy of Private Sub MCMenuVoodoo(Process As String)
'modified for use to to twist system menu

'ApiPopUpMenuCodeGenerator
SelectionButton(0).Enabled = False
SelectionButton(1).Enabled = False
SelectionButton(2).Enabled = False
SelectionButton(3).Enabled = False

SelectionButton(2).Caption = "Form_QueryUnload"


'avoid disaster
If List1.ListCount = 0 Then
MsgBox " No items in list box, ending....", , "M.C. Menu Voodoo "
Exit Sub
End If


'avoid another disaster
If MapFileUniqueID = "" Then
MsgBox "Structure is not saved jet, save it and then countinue. If the structure was sucked from vb menus, repeat the whole process again, starting from sucking. ", , "M.C. Menu Voodoo "
Exit Sub
End If


Dim magicnumber As Integer

magicnumber = 1000 ' start value of it, could be any other that comes to your mind"
MenuAppendNum = 1


'SAVE STUFF INTO TABLE -the stuf from ur structure-------------------------------------------

'dbs stuff---------------------
Dim dbs As Database
Dim rst As Recordset
Set dbs = OpenDatabase(App.Path & "\my.mdb")
Set rst = dbs.OpenRecordset("SaveConstruct")
'------------------------------
' Clear table.
dbs.Execute "DELETE * FROM SaveConstruct ;"

For i = 0 To List1.ListCount - 1
With rst


.AddNew
'ControlIndex
![ControlIndex] = i
'items string
![Container] = List1.List(i)
'level
![Level] = GetItemLevel(List1.List(i))
'caption = cut off ---- signs from string if any
![Caption] = GetItemCaption(List1.List(i))

'-----inserting this in version 2.00----TRANSLATION-----------------------------
Dim TypeOutputString As String
Dim StateOutputString As String

'1.first translate state
'Type translated......
        TypeOutputString = "MFT_STRING"
        
        Select Case Mid(List2.List(i), 2, 1)
        Case "R"
        TypeOutputString = TypeOutputString & " Or MFT_RADIOCHECK"
        Case "C"
        'do nothing
        Case Else
        End Select

        Select Case Mid(List2.List(i), 3, 1)
        Case "L"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBARBREAK"
        Case "N"
        TypeOutputString = TypeOutputString & " Or MFT_MENUBREAK"
        Case Else
        End Select

'2.second translate state
'State translated......
        Select Case Mid(List3.List(i), 1, 1)
        Case "E"
        StateOutputString = "MFS_ENABLED"
        Case "D"
        StateOutputString = "MFS_DISABLED"
        Case "G"
        StateOutputString = "MFS_GRAYED"
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 2, 1)
        Case "B"
        StateOutputString = StateOutputString & " Or MFS_DEFAULT"
        Case "N"
        'do nothing
        Case Else
        End Select
        
        Select Case Mid(List3.List(i), 3, 1)
        Case "C"
        'this disabled in ver 3.0 - reason: substantial shrinking of output code with help of IIf function
        'StateOutputString = StateOutputString & " Or MFS_CHECKED"
        ![Checked] = "MFS_CHECKED"
        Case "U"
        'this disabled in ver 3.0 - reason: substantial shrinking of output code with help of IIf function
        'StateOutputString = StateOutputString & " Or MFS_UNCHECKED"
        ![Checked] = "MFS_UNCHECKED"
        Case Else
        End Select
        
'-------------------END OF TRANSLATION---------------------------------------------------


'-----inserting this in version 3.00----variable naming for checked items-----------------

If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
VariableIndex = VariableIndex + 1
stringvar = CStr(VariableIndex)
![CheckedVarIndex] = MapFileUniqueID & stringvar
End If

'-----end of variable naming for checked items----------------------------------------------


'type of
        If GetItemCaption(List1.List(i)) = "/separator/" Then
        ![ItemType] = "MFT_SEPARATOR"
        Else
        ![ItemType] = TypeOutputString
        End If
'State
![ItemState] = StateOutputString

'Picture Yes/no
![Picture] = List4.List(i)

'Place = order in which items in menu should appear
        b = 0
        c = 0
        a = GetItemLevel(List1.List(i))
        Do
        If i = 0 Then ![Place] = 0: Exit Do
        b = b + 1
        If i - b < 0 Then ![Place] = c: Exit Do

                Select Case GetItemLevel(List1.List(i - b))
                Case Is = a
                c = c + 1
                Case Is < a
                ![Place] = c
                c = 0
                Exit Do
                Case Is > a
            
                Case Else
                Beep '? - if this happen then throw your computer away, lol
                End Select
        Loop Until i - b < 0
.Update
End With
Next i
'END SAVING  STUFF INTO TABLE


DoEvents
'dbs stuff---------------------
'Dim rst1 As Recordset
'Set rst1 = dbs.OpenRecordset("SaveConstruct")

    
'find max level in list1
MaxItemLevel = 0
For i = 0 To List1.ListCount
        If GetItemLevel(List1.List(i)) > MaxItemLevel Then
        MaxItemLevel = GetItemLevel(List1.List(i))
        End If
Next i

 



s = MaxItemLevel
'now read from bottom to top & get data into table
For j = 0 To s
       rst.MoveLast
       d = rst![Caption]
       Do
                     If rst![Level] = MaxItemLevel Then
                        Do
                           If rst![Level] = MaxItemLevel Then
                                With rst
                                .Edit
                                 ![MenuNum] = MenuAppendNum 'which menu item fits in
                                 ![MenuName] = "hPopupMenu" & MenuAppendNum 'which menu item fits in
                                 ![ItemMagicNum] = magicnumber 'allso for separator
                                .Update
                                End With
                          magicnumber = magicnumber + 1
                           End If
                           rst.MovePrevious
                        If rst.BOF Then Exit For
                        Loop Until rst![Level] < MaxItemLevel
                   rst.MoveNext
                   MenuAppendNum = MenuAppendNum + 1
                   End If
        rst.MovePrevious
        Loop Until rst.BOF
MaxItemLevel = MaxItemLevel - 1
'now next level
Next j

'Enter sub menus data - which menu is to be opened when mouse moves over specific item
For i = 0 To List1.ListCount

        Select Case GetItemCaption(List1.List(i))
        Case Is = "/separator/" 'separators can't have submenus !
            
                       If GetItemLevel(List1.List(i + 1)) > GetItemLevel(List1.List(i)) Then 'problem at sight
                       t = i
                       Do
                       t = t - 1
                       Loop Until GetItemCaption(List1.List(t)) <> "/separator/"
                       'now we have menu (t) and it's submenu (i+1)... get it into database
                        
                        rst.MoveFirst
                        Do
                                If rst![Caption] = GetItemCaption(List1.List(i + 1)) And rst![ControlIndex] = i + 1 Then
                                SubMen = rst![MenuNum]
                                Exit Do
                                End If
                                rst.MoveNext
                        Loop
                        
                        rst.MoveFirst
                        Do
                                If rst![Caption] = GetItemCaption(List1.List(t)) And rst![ControlIndex] = t Then
                                     With rst
                                     .Edit
                                     ![ItemSubMenu] = SubMen
                                     .Update
                                     End With
                                     Exit Do
                                End If
                        rst.MoveNext
                        Loop
                
                
                Else: GoTo 11 'do nothing as 'separators can't have submenus !
                End If
        
        Case Else
        If GetItemLevel(List1.List(i + 1)) = (GetItemLevel(List1.List(i)) + 1) Then
        'if next one level higher as curent item then .....
                rst.MoveFirst
                Do
                        If rst![Caption] = GetItemCaption(List1.List(i + 1)) And rst![ControlIndex] = i + 1 Then
                        SubMen = rst![MenuNum]
                        Exit Do
                        End If
                        rst.MoveNext
                Loop
                
                rst.MoveFirst
                Do
                        If rst![Caption] = GetItemCaption(List1.List(i)) And rst![ControlIndex] = i Then
                             With rst
                             .Edit
                             ![ItemSubMenu] = SubMen
                             .Update
                             End With
                             Exit Do
                        End If
                rst.MoveNext
                Loop
       End If
       End Select
11
Next i

'--------------------------------------------------------------------------------------------
'added into ver 3.0
'Coplete Path Finder Procedure
'Before writing code - get some data that will place output code coments on steroids
'rst1 is steel opened ....., so use it
rst.MoveFirst
For i = 0 To List1.ListCount - 1
    depth = rst![Level]
    Select Case depth
    Case 0
            With rst
            .Edit
            ![CompletePath] = ![Caption]
            .Update
            End With
    Case Else
                  Dim FCP As String 'final complete path
                  FCP = GetItemCaption(List1.List(i))
                          f = i
                          CurrentLevel = GetItemLevel(List1.List(i))
                          Do
                                 f = f - 1
                                 If GetItemLevel(List1.List(f)) < CurrentLevel And _
                                 GetItemCaption(List1.List(f)) <> "/Separator/" Then
                                 FCP = GetItemCaption(List1.List(f)) & "/" & FCP
                                 End If
                                 CurrentLevel = GetItemLevel(List1.List(f))
                                 
                          Loop Until GetItemLevel(List1.List(f)) = 0 'at 0 exit loop
                                'path is complete - Print it down
                                With rst
                                .Edit
                                ![CompletePath] = FCP
                                .Update
                                End With
                                                    
    End Select
    
rst.MoveNext
Next i

'--------------------------------------------------------------------------------------------
'START WRITING CODE
'--------------------------------------------------------------------------------------------

'only god knows why these lines are here, but they must be !
'comment made one month after ver ... coding - learn from it, lol
If controler = 0 Then
controler = 1
End If

    'lets see how many menus is to be created
    
     NumOfMenusToBeCreated = 1
     rst.MoveFirst
     For i = 0 To List1.ListCount - 1
     If rst![MenuNum] > NumOfMenusToBeCreated Then NumOfMenusToBeCreated = rst![MenuNum]
     rst.MoveNext
     Next i
     
'some sys needed modifications

rst.MoveFirst
For i = 0 To List1.ListCount - 1
    
        With rst
        If ![MenuName] = "hPopupMenu" & NumOfMenusToBeCreated Then
        .Edit
        ![MenuName] = "hsysmenu"
        .Update
        End If
        End With
    
rst.MoveNext
Next i


OCLC = 0
'2. The core of the thing
      'Form Load
      Open App.Path & "\MenuFormLoadCode.txt" For Output As #1    ' Open file for output
      Print #1, "Set SourceForm = Me"
      Print #1, "SysMenuModify (Me.hwnd)"
      Close #1
      
      'Form_QueryUnload
      Open App.Path & "\MenuFormQueryUnloadCode.txt" For Output As #1    ' Open file for output
      Print #1, "SysMenuRestoreDefault(Me.hwnd)"
      Close #1
      
      'here is where main output code will go....
      Open App.Path & "\MenuMainCode.txt" For Output As #1    ' Open file for output
      Print #1, "'Note all the following code must be placed into new module"
      Print #1, "'WARNING: If u place exit button in your app, do not place End under it, but u must place Unload Me instead"
      Print #1,
      
      Print #1, "'--------------------------------------------------------------------------------------------------------------------------"
      Print #1, "'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo "
      Print #1, "'Code type: Sys menu twister"
      Print #1, "'Menu identification number: sys" & Right(MapFileUniqueID, 7)
      Print #1, "'Structure saved in: " & CurentMapFilePathAndName
      Print #1, "'---------------------------------------------------------------------------------------------------------------------------"

'first make some variables public
rst.MoveFirst
For i = 0 To List1.ListCount - 1
      If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
     
      Print #1, "Public APICheck" & "Sys" & Right(rst![CheckedVarIndex], 7) & " as Boolean" & "'" & rst![CompletePath]
      
      End If
rst.MoveNext
Next i

        Print #1, "Public pOldProc As Long  ' pointer to Form1's previous window procedure"
        Print #1, "Public SourceForm as Form  ' pointer to Form1's previous window procedure"
       ' Print #1, "Public DefaultMenuColor as long'colors if any"
       ' Print #1, "Public DefaultMenuTextColor as long'colors if any"
'**********************************************************************************************************
'MAIN SUB THAT MODIFY THINGS IN SYS MENU
'**********************************************************************************************************
      
            Print #1, "Public Sub SysMenuModify(hwnd as long)"
            Print #1, "SysMenuRestoreDefault (hwnd) 'first reverse sys menu to original state'"
            Print #1, "'NOW MODIFY IT......."
            Print #1, "' Handles to the popup menus to display"
            
            For i = 1 To NumOfMenusToBeCreated
            Print #1, "Dim hPopupMenu" & i & " As Long"
            Next i
        
      Print #1, "Dim mii1 As MENUITEMINFO ' Structure that will describe menu items"
      Print #1, "Dim curpos As POINT_TYPE  ' holds the current mouse coordinates"
      Print #1, "Dim menusel As Long       ' ID of what the user selected in the popup menu"
      Print #1, "Dim retval As Long        ' generic return value"
   
    'Create the popup menus which are initialy empty.
    Print #1, ""
    Print #1, "'Create the popup menus which are initialy empty."
    Print #1, ""
    
    ' Get a handle to the system menu.
    Print #1, "hSysMenu = GetSystemMenu(hwnd, 0)"
    'How many items are currently in it.
    Print #1, "OrigSysMenuCount = GetMenuItemCount(hSysMenu)"
    
    
   
    
    'create the menus
    For i = 1 To NumOfMenusToBeCreated
    Print #1, "hPopupMenu" & i & " = CreatePopupMenu()"
    Next i
    
    
      'Structure of menus to be displayed
      Print #1, ""
      Print #1, "'Create the structure which is the base for all added menus:"
      Print #1, "With mii1"
      Print #1, ".cbSize = Len(mii1)' The size of this structure."
      Print #1, ".fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU' Which elements of the structure to use."
      Print #1, "End With"
      Print #1, ""
      
       
DoEvents

rst.MoveFirst

'get items into created menus ! & describe their properties"

Print #1, "'get items into created menus & describe their properties"
Print #1, ""
'-------------------------------------------
'-------------------------------------------
'-------------------------------------------
For i = 0 To List1.ListCount - 1
      
            Print #1, "With mii1" & "'(" & rst![CompletePath] & ")"
            
            'type thing
            Print #1, ".fType =" & rst![ItemType]
            
            'state thing
            Select Case rst![Checked]
            Case "MFS_UNCHECKED", "MFS_CHECKED"
            Print #1, ".fState =" & rst![ItemState] & " Or IIf( APICheck" & "Sys" & Right(rst![CheckedVarIndex], 7) & ", " & " MFS_CHECKED" & ", " & " MFS_UNCHECKED" & ")"
            
            Case Else 'normal item not checked at all
            Print #1, ".fState =" & rst![ItemState]
            
            End Select
            
                     
            Print #1, ".wID =" & rst![ItemMagicNum]
            
            Print #1, ".dwTypeData =" & Chr(34) & rst![Caption] & Chr(34) ' Display the following text for the item."
            
            Print #1, ".cch = Len(" & Chr(34) & rst![Caption] & Chr(34) & ")"
            
            'if there is a submenu for this item then......
            If rst![ItemSubMenu] > 0 Then
            Print #1, ".hSubMenu = hPopupMenu" & rst![ItemSubMenu]
            
            Else
            Print #1, ".hSubMenu = 0"
            
            End If
            Print #1, "End With"
            
      
      
      
      'well where to send the item ? To the sys menu !
      
      
      Select Case proccess
      Case "EraseAndReplace", "Insert"
      Print #1, "retval = InsertMenuItem(" & rst![MenuName] & "," & rst![Place] & ",1, mii1" & ")"
      Case "Add"
      If rst![MenuName] = "hsysmenu" Then
      Print #1, "retval = InsertMenuItem(" & rst![MenuName] & "," & rst![Place] & " + OrigSysMenuCount" & ",1, mii1" & ")"
      Else
      Print #1, "retval = InsertMenuItem(" & rst![MenuName] & "," & rst![Place] & ",1, mii1" & ")"
      End If
      End Select
      
      
      
      'insert at the start
      
      
      'add at the end
      'If rst![MenuName] = "hsysmenu" Then
      'Print #1, "retval = InsertMenuItem(" & rst![MenuName] & "," & rst![Place] & OrigSysMenuCount & ",1, mii1" & ")"
      'End If
      
      
   
      Print #1, ""
      rst.MoveNext

      
Next i

    Print #1, "NewSysMenuCount = GetMenuItemCount(hSysMenu)"

    'if total replacement then erase all original items
    If proccess = "EraseAndReplace" Then
    Print #1, "For i = 0 to OrigSysMenuCount"
    Print #1, "RemoveMenu hSysMenu, NewSysMenuCount-i,MF_BYPOSITION or MF_REMOVE"
    Print #1, "Next i"
    End If




Print #1, "'The following code is for adding pictures into menus, if there are any!"
Print #1, "'------------------------------------------------------------"


'HERE ADD PICTURES IF ANY
rst.MoveFirst 'back to beginning
Dim picindex As Integer
Dim PicMsgBoxWarning As Boolean
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst![Picture] = "Yes" Then
                If Mid(List2.List(i), 2, 1) = "P" Then
                
                Print #1, "retval = SetMenuItemBitmaps(" & rst![MenuName] & "," & rst![ItemMagicNum] & ",1, SourceForm.MenuPicBox" & "Sys" & Left(MapFileUniqueID, 7) & "(" & picindex & ")," & "SourceForm.MenuPicBox" & "Sys" & Left(MapFileUniqueID, 7) & "(" & picindex + 1 & "))"
                
                picindex = picindex + 2
                Else
                
                Print #1, "retval = SetMenuItemBitmaps(" & rst![MenuName] & "," & rst![ItemMagicNum] & ",1, SourceForm.MenuPicBox" & "Sys" & Left(MapFileUniqueID, 7) & "(" & picindex & ")," & "SourceForm.MenuPicBox" & "Sys" & Left(MapFileUniqueID, 7) & "(" & picindex & "))"
                
                picindex = picindex + 1
                
                End If
        PicMsgBoxWarning = True
        End If
rst.MoveNext
Next i


Print #1, "'------------------------------------------------------------"
    
                        Print #1, "' Set the custom window procedure to process Form1's messages."
                        Print #1, "pOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)"
                        Print #1, "End Sub"
                        'end of  Public Sub SysMenuModify


'**********************************************************************************************************
'Restorer of default system menu
'**********************************************************************************************************
Print #1,
Print #1, "Public Sub SysMenuRestoreDefault(hwnd As Long)"
Print #1, "' Before unloading, restore the default system menu and remove the"
Print #1, "' custom window procedure."
Print #1, "Dim retval As Long  ' return value"
Print #1, "' Replace the previous window procedure to prevent crashing."
Print #1, "retval = SetWindowLong(hwnd, GWL_WNDPROC, pOldProc)"
Print #1, "' Remove the modifications made to the system menu."
Print #1, "retval = GetSystemMenu(hwnd, 1)"
Print #1, "End Sub"
        
        
'**********************************************************************************************************
'WINDOW PROC function - enables ur app to read commands from sys menu
'**********************************************************************************************************

        
        'The function itself
        Print #1, "'The following function acts as Form1's window procedure to process messages."
        Print #1, "Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long"
        Print #1, "Dim hSysMenu As Long     ' handle to Form1's system menu"
        Print #1, "Dim mii As MENUITEMINFO  ' menu item information for Always On Top"
        Print #1, "Dim retval As Long       ' return value"
        
        Print #1, "Select Case uMsg"
        
        
        
        If proccess = "EraseAndReplace" Then
        '*************************************************************
        Print #1, "Case 160 'mouse moving ower caption bar area"

        Print #1, "      If wParam = 8 Or wParam = 9 Or wParam = 20 Then 'min, restore and close button"
        Print #1, "     SysMenuRestoreDefault (SourceForm.hwnd)"
        Print #1, "     SetTimer SourceForm.hwnd, 0, 1, AddressOf Changer"
        Print #1, "     DrawMenuBar SourceForm.hwnd 'force refresh of caption bar"
        Print #1, "     End If"
        Print #1, "     WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)"
        '*************************************************************
        End If
        
        
        
        
        Print #1, "Case WM_INITMENU 'before sys menu is displayed"
        Print #1,
        Print #1, "SysMenuModify (hwnd) 'rebuild it before displaying it"
        Print #1,
        Print #1, "Case WM_SYSCOMMAND 'when you click any of added items !!!!"
        Print #1, "'Here is the place where things happens"
        
        Print #1, "'wParam = magicNumber identifying our selected item"
        
        'here is the way around to restore sys menu looks each time u select any item in it
        
        Print #1, "Select case wParam"
                         rst.MoveFirst
                         For i = 0 To List1.ListCount - 1
                         Select Case rst![Caption]
                         Case "/separator/" 'separator = do nothing as you can't select it
                         Case Else
                            If rst![ItemSubMenu] = 0 Then
                                    
                                    Print #1, "Case " & rst![ItemMagicNum] & "'(" & rst![CompletePath] & ")"
                                    Print #1,
                                            'now what to do ?
                                            'if it is marked checked
                                            If rst![Checked] = "MFS_CHECKED" Or rst![Checked] = "MFS_UNCHECKED" Then
                                            markmessage = True
                                            
                                            
                                            Print #1, "'******************************************************************"
                                            Print #1, "'The lines underlined with stars enables Stop to stop at all otherwise Vb would crash here"
                                            Print #1, "'all stars underlined lines MUST ! be erased at the point that you replace Stop with something else"
                                            Print #1, "'Then be certain that there are no errors in that code, otherwise Vb crash again - save project offen!"
                                            Print #1, "retval = SetWindowLong(hwnd, GWL_WNDPROC, pOldProc)"
                                            Print #1, "'******************************************************************"
                                            Print #1,
                                            Print #1, "'Do change appropriate variable value now"
                                            Print #1, "APICheck" & "Sys" & Right(rst![CheckedVarIndex], 7) & " = IIf( APICheck" & "Sys" & Right(rst![CheckedVarIndex], 7) & ", " & " False" & ", " & " True" & ")"
                                            Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                            Print #1, "'And for your convinience....."
                                            Print #1, "If APICheck" & "Sys" & Right(rst![CheckedVarIndex], 7) & " = True Then'in case item become unchecked ....."
                                            Print #1, "Else 'in case item become checked ....."
                                            Print #1, "End If"
                                            Print #1, "Stop"
                                            Print #1,
                                            Print #1, "SysMenuModify (hwnd)"
                                            Print #1, "'******************************************************************"
                                            'if Option1(1).Value = True Then 'case there are colors
                                            'Print #1, "SetSysColors 1, COLOR_MENU, DefaultMenuColor"
                                            'Print #1, "SetSysColors 1, COLOR_MENUTEXT, DefaultMenuTextColor"
                                            'End If
                                        
                                            OCLC = OCLC + 6
                                            
                                            Else:
                                            Print #1, "'The lines underlined with stars enables Stop to stop at all otherwise Vb would crash here"
                                            Print #1, "'all stars underlined lines  MUST ! be erased at the point that you replace Stop with something else"
                                            Print #1, "'Then be certain that there are no errors in that code, otherwise Vb crash again - save project offen!"
                                            Print #1, "retval = SetWindowLong(hwnd, GWL_WNDPROC, pOldProc)"
                                            Print #1, "'******************************************************************"
                                            Print #1,
                                            Print #1, "'Here is the place to put the code to be executed clicking:  " & rst![CompletePath] & "."
                                            Print #1, "Stop"
                                            Print #1,
                                            Print #1, "SysMenuModify (hwnd)"
                                            Print #1, "'******************************************************************"
                                            End If
                                            
                                            
                          End If
                          End Select
                          rst.MoveNext
                          Next i
                         
       
    
        Print #1, "Case Else"
        Print #1, "' Some other item was selected.  Let the previous window procedure process it."
        Print #1, "WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)"
        Print #1, "End Select"
        Print #1, "Case Else"
        Print #1, "' Some other item was selected.  Let the previous window procedure process it."
        Print #1, "WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)"
        Print #1, "End Select"
        Print #1, "End Function"
        
        
        
       If proccess = "EraseAndReplace" Then
       '***********************************************
       Print #1, "Public Sub Changer()"
       Print #1, "Dim Mes As Msg"
       Print #1, "WaitMessage"
       Print #1, "PeekMessage Mes, SourceForm.hwnd, 160, 512, PM_NOREMOVE 'And Message.wParam = 3 Then"
       Print #1, "If Mes.Message = 160 And wParam <> 8 And wParam <> 9 And wParam <> 20 Then  'min restore and close button"
       Print #1, "'change sys menu at this point"
       Print #1, "         KillTimer SourceForm.hwnd, 0 'kill the timer"
       Print #1, "         SysMenuModify (SourceForm.hwnd)"
       Print #1, "End If"
        
       Print #1, "If Mes.Message = 512 Then 'user moved mouse over form area"
       Print #1, "      SysMenuRestoreDefault (SourceForm.hwnd)"
       Print #1, "      SetTimer SourceForm.hwnd, 0, 1, AddressOf Changer"
       Print #1, "      DrawMenuBar SourceForm.hwnd 'force refresh of caption bar"
       Print #1, "End If"

       Print #1, "End Sub"
       '***********************************************
       End If
        
        
 Close #1
      
            
        
      
        
    
        
       
'blablablabla....................
       
   

SelectionButton(0).Enabled = True

'MESSAGE PROCEDURE


If PicMsgBoxWarning = True Then 'comments for pictures
PicMsgBoxWarning = False
SelectionButton(3).Enabled = True
'here is where code will go....
Open App.Path & "\MenuPicturesComment.txt" For Output As #2    ' Open file for output

Print #2, "'You have choosen to have some pics on your menu"
Print #2, "'so you have to put following array of picture controls into your projects."
Print #2, "'Comments which picture fits into specific array element are allso here."
OCLC = OCLC + 3


rst.MoveFirst
picindex = 0
For i = 0 To List1.ListCount - 1
        If rst![Picture] = "Yes" Then
                If Mid(List2.List(i), 2, 1) = "P" Then
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")" & " ,When Checked"
                        
                        picindex = picindex + 1
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")" & " ,When UnChecked"
                        
                        picindex = picindex + 1
                        
                        Else
                        
                        Print #2, "'MenuPicBox" & MapFileUniqueID & "(" & picindex & ")" & " = (" & rst![CompletePath] & ")"
                        
                        picindex = picindex + 1
                        
                End If
    End If
rst.MoveNext
Next i
Close #2
End If






'and on the end, dbs closing stuff
rst.Close
Set rst = Nothing
dbs.Close

'next one solves hung up of this sub in case that you run it
'again while stuff is already in text2, forgot how  = did no comment, but it works
controler = 0
VariableIndex = 0 'this not beeing here caused big bug, that now disappeared


'show the MainCode to user
RichTextBox1.LoadFile App.Path & "\MenuMainCode.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText

SelectionButton(1).Enabled = True
SelectionButton(2).Enabled = True

MsgBox "The code has been automaticaly copyed to clipboard. It is presumed that you already have MenuVoodoo.bas in your project.The best is that you add new module and paste all this to it, name it Form1SysMenuMod.bas or something like that.Then in your 'Form1' load event place: Set SourceForm = Me and SysMenuModify(hwnd). That is all i.e clicking ur form sysmenu should bring up modified one, that allso does things." & Chr(13) & OCLC & " lines of code generated.", , "M.C. Menu Voodoo "

End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
Label5(2).ForeColor = ShowColor(Me)
Case 1
Label5(2).BackColor = ShowColor(Me)
Case 2 'save color
        getstring = InputBox("Enter the name of new color combination", "SaveMenuColors", List5.List(List5.ListIndex))
        If getstring = "" Then Exit Sub 'if cancel
        
        'dbs stuff---------------------
        Dim dbs As Database
        Dim rst As Recordset
        Set dbs = OpenDatabase(App.Path & "\my.mdb")
        Set rst = dbs.OpenRecordset("MenuColors")
        
        'check first if replacing existing savings
        
        rst.MoveFirst
        Do
            If rst![Name] = getstring Then
                 With rst
                .Edit
                ![TextColor] = Label5(2).ForeColor
                ![BackGroundColor] = Label5(2).BackColor
                .Update
                 End With
             MsgBox getstring & " updated.", , "M.C. Menu Voodoo "
             Exit Sub
            End If
        rst.MoveNext
        Loop Until rst.EOF
        
        
        'if then adding new
        
        rst.MoveLast
        
        With rst
        .AddNew
        ![Name] = getstring
        ![TextColor] = Label5(2).ForeColor
        ![BackGroundColor] = Label5(2).BackColor
        .Update
        End With
        
        'refresh listbox
        List5.Clear
        rst.MoveFirst
        Do
        List5.AddItem rst![Name]
        rst.MoveNext
        Loop Until rst.EOF
        
        rst.Close
        dbs.Close
Case Else
End Select
End Sub

Private Sub Command3_Click()
'SAVE FILE WITH CDLG NO OCX INCLUDED

'first dimension is not to be changed, second is to be set to num of filters that
'you will add to dialog - 1 !
Dim ArrayWithFileFilters(2, 1) As String 'array holfding file filters
'first filter & its description
ArrayWithFileFilters(1, 0) = "Map Files"
ArrayWithFileFilters(2, 0) = "*.Map"
'second filter & its description
ArrayWithFileFilters(1, 1) = "All Files"
ArrayWithFileFilters(2, 1) = "*.*"
'x filter & its description............

'now call our dialog public function
Dim ReturnedFile As String
ReturnedFile = CdlgFileToOpenOrSave("Save", Me, ArrayWithFileFilters, App.Path, "M.C. Menu Voodoo")
'END OF Save FILE WITH CDLG NO OCX INCLUDED

If ReturnedFile = "" Then Exit Sub ' cancel was pressed

CurentMapFilePathAndName = ReturnedFile

'see if file is already there
Dim FileAlreadyExist As String
FileAlreadyExist = Dir(CurentMapFilePathAndName)

If FileAlreadyExist = "" Then 'doe's not exist jet
        'generate randome MapFileUniqueID
        Dim Id As Long
        Randomize
        Id = Int(Rnd * 1000000000)
        MapFileUniqueID = CStr(Id)
        If Len(MapFileUniqueID) < 10 Then
        MapFileUniqueID = String(10 - Len(MapFileUniqueID), "0") & MapFileUniqueID
        End If
Else 'it is already there - so we don't want to change unique ID
        'open file and suck out unique ID
        Open ReturnedFile For Input As #1     ' Open file for output.
        Input #1, a, b, c, d
        MapFileUniqueID = b
        Close #1
End If

CurrentOpenedFile = ReturnedFile
CurrentOpenedFile1 = SuckOutFileName(ReturnedFile)
Me.Caption = "(" & CurrentOpenedFile1 & ")" & ", MapFileUniqueID= " & MapFileUniqueID & ", M.C. Menu Voodoo"


'now savestuff
Open ReturnedFile For Output As #1    ' Open file for output.
'MapFileUniqueID
Write #1, "Map file Unique ID", MapFileUniqueID, "", ""

'lists contence
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i), List2.List(i), List3.List(i), List4.List(i)
Next i

'And general settings
Write #1, "GS", "", "", ""
If Option6(0).Value = True Then Write #1, "TPM_RIGHTALIGN = &H8&", "", "", ""
If Option6(1).Value = True Then Write #1, "TPM_LEFTALIGN = &H0", "", "", ""
If Option6(2).Value = True Then Write #1, "TPM_CENTERALIGN = &H4&", "", "", ""

If Option7(0).Value = True Then Write #1, "TPM_RIGHTBUTTON = &H2&", "", "", ""
If Option7(1).Value = True Then Write #1, "TPM_LEFTBUTTON = &H0", "", "", ""

If Option1(0).Value = True Then Write #1, "DEFAULT MENU COLORS", "", "", ""
If Option1(1).Value = True Then Write #1, "USER DEFINED MENU COLORS", Label5(2).BackColor, Label5(2).ForeColor, ""

Close #1   ' Close file.

End Sub

Private Sub Command4_Click()
'erase if anything already there
List1.Clear
List2.Clear
List3.Clear
List4.Clear

On Error GoTo errorhandler

'OPEN FILE WITH CDLG NO OCX INCLUDED

'first dimension is not to be changed, second is to be set to num of filters that
'you will add to dialog - 1 !
Dim ArrayWithFileFilters(2, 1) As String 'array holfding file filters
'first filter & its description
ArrayWithFileFilters(1, 0) = "Map Files"
ArrayWithFileFilters(2, 0) = "*.Map"
'second filter & its description
ArrayWithFileFilters(1, 1) = "All Files"
ArrayWithFileFilters(2, 1) = "*.*"
'x filter & its description............

'now call our dialog public function
Dim ReturnedFile As String
ReturnedFile = CdlgFileToOpenOrSave("Open", Me, ArrayWithFileFilters, App.Path, "M.C. Menu Voodoo")

'END OF OPEN FILE WITH CDLG NO OCX INCLUDED
CurrentOpenedFile = SuckOutFileName(ReturnedFile)
Open ReturnedFile For Input As #1     ' Open file for output.
    
    i = 0
    
    Input #1, a, b, c, d
    MapFileUniqueID = b
    
    'getstuf into form caption bar
    Me.Caption = "(" & CurrentOpenedFile & ")" & ", MapFileUniqueID= " & MapFileUniqueID & ", M.C. Menu Voodoo"
    
    CurentMapFilePathAndName = ReturnedFile
    Do
    Input #1, a, b, c, d
    If a = "GS" Then Exit Do
    List1.AddItem a
    List2.AddItem b
    List3.AddItem c
    List4.AddItem d
    i = i + 1
    Loop
    
    Input #1, a, b, c, d
    
    If a = "TPM_RIGHTALIGN = &H8&" Then Option6(0).Value = True
    If a = "TPM_LEFTALIGN = &H0" Then Option6(1).Value = True
    If a = "TPM_CENTERALIGN = &H4&" Then Option6(2).Value = True
    
    Input #1, a, b, c, d

    If a = "TPM_RIGHTBUTTON = &H2&" Then Option7(0).Value = True
    If a = "TPM_LEFTBUTTON = &H0" Then Option7(1).Value = True
    
    Input #1, a, b, c, d
    
    'menu colors
    If a = "DEFAULT MENU COLORS" Then
    Option1(0).Value = True
    Label5(2).BackColor = 12632256
    Label5(2).ForeColor = 0
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    List5.Enabled = False
    End If
    
    If a = "USER DEFINED MENU COLORS" Then
    Option1(1).Value = True
    Command2(0).Enabled = True
    Command2(1).Enabled = True
    Command2(2).Enabled = True
    Command2(3).Enabled = True
    List5.Enabled = True
    Label5(2).BackColor = b
    Label5(2).ForeColor = c
    End If
    
Close #1   ' Close file.
Exit Sub
errorhandler:
Select Case Err
Case 62
'appears always
Case Else
End Select
Close #1   ' Close file.

End Sub




Private Sub Command7_Click() 'close general settings
Frame7.Visible = False
End Sub

Public Sub Command8_Click()

'code to fit into module
Open App.Path & "\MenuPublicDeclares.txt" For Output As #1    ' Open file for output


'1.declaration section
Print #1, "'Declaration section"

Print #1, "Public Declare Function CreatePopupMenu Lib" & Chr(34) & "user32.dll" & Chr(34) & " ()  As Long"

Print #1, "Public Declare Function DestroyMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " (ByVal hMenu As Long) As Long"

Print #1, "Public Type MENUITEMINFO"

Print #1, "        cbSize As Long"

Print #1, "        fMask As Long"

Print #1, "        fType As Long"

Print #1, "        fState As Long"

Print #1, "        wID As Long"

Print #1, "        hSubMenu As Long"

Print #1, "        hbmpChecked As Long"

Print #1, "        hbmpUnchecked As Long"

Print #1, "        dwItemData As Long"

Print #1, "        dwTypeData As String"

Print #1, "        cch As Long"

Print #1, "End Type"


'Constant Definitions
 
Print #1, "Public Const MIIM_STATE = &H1"

Print #1, "Public Const MIIM_ID = &H2"

Print #1, "Public Const MIIM_SUBMENU = &H4"

Print #1, "Public Const MIIM_CHECKMARKS = &H8"

Print #1, "Public Const MIIM_DATA = &H20"

Print #1, "Public Const MIIM_TYPE = &H10"

Print #1, "Public Const MFT_BITMAP = &H4"

Print #1, "Public Const MFT_MENUBARBREAK = &H20"

Print #1, "Public Const MFT_MENUBREAK = &H40"

Print #1, "Public Const MFT_OWNERDRAW = &H100"

Print #1, "Public Const MFT_RADIOCHECK = &H200"

Print #1, "Public Const MFT_RIGHTJUSTIFY = &H4000"

Print #1, "Public Const MFT_RIGHTORDER = &H2000"

Print #1, "Public Const MFT_SEPARATOR = &H800"

Print #1, "Public Const MFT_STRING = &H0"

Print #1, "Public Const MFS_CHECKED = &H8"

Print #1, "Public Const MFS_DEFAULT = &H1000"

Print #1, "Public Const MFS_DISABLED = &H2"

Print #1, "Public Const MFS_ENABLED = &H0"

Print #1, "Public Const MFS_GRAYED = &H1"

Print #1, "Public Const MFS_HILITE = &H80"

Print #1, "Public Const MFS_UNCHECKED = &H0"

Print #1, "Public Const MFS_UNHILITE = &H0"

'functions = API-s

Print #1, "Public Declare Function InsertMenuItem Lib " & Chr(34) & "user32.dll" & Chr(34) & " Alias " & Chr(34) & "InsertMenuItemA" & Chr(34) & " _"

Print #1, "(ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long"

Print #1, "Public Declare Function TrackPopupMenu Lib " & Chr(34) & "user32.dll" & Chr(34) & " _"

Print #1, "(ByVal hMenu As Long, ByVal uFlags As Long, ByVal x As Long, ByVal y As Long," & " _"

Print #1, "ByVal nReserved As Long, ByVal hWnd As Long, ByVal prcRect As Long) As Long"

Print #1, "Public Const TPM_RIGHTALIGN = &H8&"

Print #1, "Public Const TPM_CENTERALIGN = &H4&"

Print #1, "Public Const TPM_LEFTALIGN = &H0"

Print #1, "Public Const TPM_TOPALIGN = &H0"

Print #1, "Public Const TPM_NONOTIFY = &H80"

Print #1, "Public Const TPM_RETURNCMD = &H100"

Print #1, "Public Const TPM_LEFTBUTTON = &H0"

Print #1, "Public Const  TPM_RIGHTBUTTON = &H2&"

Print #1, "Public Type POINT_TYPE"

Print #1, "x As Long"

Print #1, "y As Long"

Print #1, "End Type"

Print #1, "Public Declare Function GetCursorPos Lib " & Chr(34) & "user32.dll" & Chr(34) & " (lpPoint As POINT_TYPE) As Long"

Print #1, "Public Declare Function AppendMenu Lib " & Chr(34) & "user32" & Chr(34) & " Alias " & Chr(34) & "AppendMenuA" & Chr(34) & " (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long"
 'for adding pictures
Print #1, "Public Declare Function SetMenuItemBitmaps Lib " & Chr(34) & "user32" & Chr(34) & " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long"

Print #1, "Menu colors"
Print #1, "Public Declare Function SetSysColors Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long"
Print #1, "Public Declare Function GetSysColor Lib " & Chr(34) & "user32" & Chr(34) & "(ByVal nIndex As Long) As Long"
Print #1, "Public Const COLOR_MENUTEXT = 7"
Print #1, "Public Const COLOR_MENU = 4"

Print #1, "'Sys menu stuff"
Print #1, "Public Declare Function GetSystemMenu Lib " & Chr(34) & " user32  (ByVal hWnd As Long, ByVal bRevert As Long) As Long"
Print #1, "Public Declare Function GetMenuItemCount Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hMenu As Long) As Long"
Print #1, "Public Declare Function GetMenuItemInfo Lib " & Chr(34) & " user32" & Chr(34) & " Alias" & Chr(34) & " GetMenuItemInfoA" & Chr(34) & " (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long"
'''''''''''''''''
Print #1, "Public Declare Function SetWindowLong Lib " & Chr(34) & " user32" & Chr(34) & " Alias" & Chr(34) & " & Chr(34) & " & Chr(34) & " SetWindowLongA" & Chr(34) & " (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long"
Print #1, "Public Const GWL_WNDPROC = -4"
Print #1, "Public Declare Function CallWindowProc Lib " & Chr(34) & " user32.dll" & Chr(34) & " Alias" & Chr(34) & " & Chr(34) & " & Chr(34) & " CallWindowProcA" & Chr(34) & " (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long"
Print #1, "Public Const WM_SYSCOMMAND = &H112"
Print #1, "Public Const WM_INITMENU = &H116"

Print #1, "Public Declare Function RemoveMenu Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long"
Print #1, "Public Const MF_REMOVE = &H1000&"

Print #1, "Public Type RECT"
Print #1, "    Left As Long"
Print #1, "    Top As Long"
Print #1, "    Right As Long"
Print #1, "    Bottom As Long"
Print #1, "End Type"

Print #1, "Public Declare Function DrawMenuBar Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long) As Long"
Print #1, "Public Declare Function CreateRectRgn Lib " & Chr(34) & " gdi32" & Chr(34) & " (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
Print #1, "Public Declare Function GetWindowRect Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, lpRect As RECT) As Long"
Print #1, "Public Declare Function PtInRegion Lib " & Chr(34) & " gdi32" & Chr(34) & " (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long"

'for complete replacement of sys menu
Print #1, "Public Declare Function PeekMessage Lib " & Chr(34) & " user32" & Chr(34) & " Alias" & Chr(34) & " & Chr(34) & " & Chr(34) & " PeekMessageA" & Chr(34) & " (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long"
Print #1, "Public Declare Function WaitMessage Lib " & Chr(34) & " user32" & Chr(34) & " () As Long"
Print #1, "Public Declare Function GetMessage Lib " & Chr(34) & " user32" & Chr(34) & " Alias" & Chr(34) & " & Chr(34) & " & Chr(34) & " GetMessageA" & Chr(34) & " (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long"
Print #1, "Public Declare Function SetTimer Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long"
Print #1, "Public Declare Function KillTimer Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long"

Print #1, "Public Type POINTAPI"
Print #1, "    x As Long"
Print #1, "    y As Long"
Print #1, "End Type"
Print #1, "Public Type Msg"
Print #1, "    hwnd As Long"
Print #1, "    Message As Long"
Print #1, "    wParam As Long"
Print #1, "    lParam As Long"
Print #1, "    time As Long"
Print #1, "    pt As POINTAPI"
Print #1, "End Type"

Print #1, "Public Const PM_REMOVE = &H1"
Print #1, "Public Const PM_NOREMOVE = &H0"

'Needed to add buttons to caption bar
Print #1, "Public Const DFC_CAPTION = 1"

Print #1, "Public Const DFCS_CAPTIONCLOSE = &H0"
Print #1, "Public Const DFCS_CAPTIONMAX = &H2"
Print #1, "Public Const DFCS_CAPTIONMIN = &H1"
Print #1, "Public Const DFCS_CAPTIONRESTORE = &H3"

Print #1, "Public Const DFCS_PUSHED = &H200"

Print #1, "Public Const SM_CYCAPTION = 4"


Print #1, "Public Declare Function DrawFrameControl Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long"
Print #1, "Public Declare Function GetSystemMetrics Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal nIndex As Long) As Long"
Print #1, "Public Declare Function GetWindowDC Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long) As Long"
Print #1, "Public Declare Function ReleaseDC Lib " & Chr(34) & " user32" & Chr(34) & " (ByVal hwnd As Long, ByVal hdc As Long) As Long"
Print #1, "Public Declare Function SetRect Lib " & Chr(34) & " user32" & Chr(34) & " (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"


Close #1

RichTextBox1.LoadFile App.Path & "\MenuPublicDeclares.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText

MsgBox "Public declares are sitting on clipboard and waiting for you, anyway you can look at them in below text box.", , "M.C. Menu Voodoo "


End Sub


Private Sub Command9_Click()
MsgBox "This stuff works, but changes color of menus on all, opened windows, so as result when ur colorised menu poping up all windows slightly flicker which is ugly, but when ur menu closed, default settings are restored. So don't worry, it just looks a kind of non proffessional.", , "M.C. Menu Voodoo "
End Sub



Private Sub Form_Activate()
Set SourceForm = Me
SysMenuModify (Me.hwnd)

End Sub


Private Sub Form_Load()


'cover whole screen with our form
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Me.Height = Screen.Height - (Screen.TwipsPerPixelY * 25) 'let some place for sys tray etc.


RichTextBox1.Text = ""

'declarations of menu variables:
APICheck1 = False 'Bold
APICheck2 = True 'Normal
APICheck3 = True 'Enabled
APICheck4 = False 'Disabled
APICheck5 = False 'Grayed
APICheck7 = False 'Checked ?
APICheck8 = False 'Checked ?/Not Checked
APICheck9 = False 'Checked ?/On Start Checked/With Checkmark
APICheck10 = False 'Checked ?/On Start Checked/With RadioButton
APICheck11 = False 'Checked ?/On Start Checked/With Pictures
APICheck12 = False 'Checked ?/On Start Checked/On Start Unchecked
APICheck13 = False 'Checked ?/On Start Checked/On Start Unchecked/With Checkmark
APICheck14 = False 'Checked ?/On Start Checked/On Start Unchecked/With RadioButton
APICheck15 = False 'Checked ?/On Start Checked/On Start Unchecked/With Pictures
APICheck24 = False 'Picture ?
APICheck25 = False 'Picture ?/Yes
APICheck26 = True 'Picture ?/No
APICheck27 = False 'Insert Column Break
APICheck28 = False 'Insert Column Break/With dividing line
APICheck29 = False 'Insert Column Break/Without dividing line
APICheck30 = False 'Insert Column Break/Delete col. break



'for sizing outputcode box
defaultpicture1top = Picture1.Top
defaultpicture1height = Picture1.Height
defaultRichTextBox1top = RichTextBox1.Top
defaultRichTextBox1height = RichTextBox1.Height
MaximizeYes = False
'----------------------------------------------------------
'center Picture2
Picture2.Left = (List1.Left + List1.Width / 2) - (Picture2.Width / 2)
Picture2.Top = (List1.Top + List1.Height / 2) - (Picture2.Height / 2)
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


SysMenuRestoreDefault (Me.hwnd)





End Sub


Private Sub Form_Terminate()
If appdestructed = True Then Exit Sub
MCCloseForm Form1, 2

End Sub

Private Sub Form_Unload(Cancel As Integer)

SysMenuRestoreDefault (Me.hwnd)

End Sub

Private Sub Image1_Click()
Select Case MaximizeYes
Case False
Picture1.Top = 0
Picture1.Height = Me.Height
RichTextBox1.Height = Picture1.Height - 800
MaximizeYes = True
Case True
Picture1.Top = defaultpicture1top
Picture1.Height = defaultpicture1height
RichTextBox1.Height = defaultRichTextBox1height
MaximizeYes = False
End Select



End Sub

Private Sub Label4_Click()
Open App.Path & "\InetAdress.txt" For Input As #1    ' Open file for output.
Input #1, a
Close #1
Shell "start " & a
End Sub



Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If ClickNoCanDo = True Then Exit Sub
'teh following lines trigger some click events so inhibit that...
ClickNoCanDo = True
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
ClickNoCanDo = False
'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo
'Menu identification number: 0495708703
'Structure saved in: C:\ProjektiVB6\Api Menu(Ver31)\M1.map
'---------------------------------------------------------------------------------------------------------------------------

Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim hPopupMenu2 As Long ' handle to the popup menu to display
Dim hPopupMenu3 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim mii2 As MENUITEMINFO   ' describes menu items to add
Dim mii3 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value


'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()

'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With

'Make all structures equal
mii2 = mii1
mii3 = mii1

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "Item"
.cch = Len("Item")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)

With mii1
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "Separator"
.cch = Len("Separator")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "Item"
.cch = Len("Item")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii2)

With mii2
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "Separator"
.cch = Len("Separator")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 1, 1, mii2)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1015 ' Assign this item an item identifier.
.dwTypeData = "Move Right"
.cch = Len("Move Right")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 0, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1014 ' Assign this item an item identifier.
.dwTypeData = "Move Left"
.cch = Len("Move Left")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 1, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1013 ' Assign this item an item identifier.
.dwTypeData = "Move Up"
.cch = Len("Move Up")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 2, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1012 ' Assign this item an item identifier.
.dwTypeData = "Move Down"
.cch = Len("Move Down")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 3, 1, mii3)

With mii3
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1011 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 4, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "Menu Behaviour"
.cch = Len("Menu Behaviour")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 5, 1, mii3)

With mii3
.fType = MFT_STRING Or MFT_MENUBARBREAK
.fState = MFS_ENABLED
.wID = 1009 ' Assign this item an item identifier.
.dwTypeData = "Add"
.cch = Len("Add")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu3, 6, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1008 ' Assign this item an item identifier.
.dwTypeData = "Insert"
.cch = Len("Insert")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu3, 7, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1007 ' Assign this item an item identifier.
.dwTypeData = "Delete"
.cch = Len("Delete")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 8, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1006 ' Assign this item an item identifier.
.dwTypeData = "Change String"
.cch = Len("Change String")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 9, 1, mii3)

With mii3
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1005 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 10, 1, mii3)

With mii3
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "Close Menu"
.cch = Len("Item properties")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 11, 1, mii3)

'The following code is for adding pictures into menus!
'------------------------------------------------------------
'------------------------------------------------------------

retval = SetMenuItemBitmaps(hPopupMenu3, 1015, 1, MenuPicContainer(0), MenuPicContainer(0))
retval = SetMenuItemBitmaps(hPopupMenu3, 1014, 1, MenuPicContainer(1), MenuPicContainer(1))
retval = SetMenuItemBitmaps(hPopupMenu3, 1013, 1, MenuPicContainer(2), MenuPicContainer(2))
retval = SetMenuItemBitmaps(hPopupMenu3, 1012, 1, MenuPicContainer(3), MenuPicContainer(3))
'------------------------------------------------------------
'------------------------------------------------------------

retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu3, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_CENTERALIGN Or TPM_RIGHTBUTTON, curpos.x, curpos.y, 0, Form1.hwnd, 0)
retval = DestroyMenu(hPopupMenu3)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Select Case menusel

Case 1001 '(Insert/Item)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Can't insert as there is nothing in list !", , "M.C. Menu Voodoo "
                Exit Sub
                End If
                
                Dim default
                default = ""
                Message = "Type in new item caption"
                Title = "Input box"
                myvalue = InputBox(Message, Title, default)
                If myvalue <> "" Then
                List1.AddItem myvalue, List1.ListIndex
                List2.AddItem "S  ", List1.ListIndex - 1 'MFT_STRING or MFT_CHECKED
                List3.AddItem "EN ", List1.ListIndex - 1 'MFS_ENABLED
                List4.AddItem "No", List1.ListIndex - 1
                End If

Case 1000 '(Insert/Separator)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Can't insert as there is nothing in list !", , "M.C. Menu Voodoo "
                Exit Sub
                End If
List1.AddItem "/separator/", List1.ListIndex
List2.AddItem "SC ", List1.ListIndex - 1
List3.AddItem "EN ", List1.ListIndex - 1
List4.AddItem "No", List1.ListIndex - 1

Case 1003 '(Add/Item)

 'Dim default
            default = ""
            Message = "Type in new item caption"
            Title = "Input box"
            myvalue = InputBox(Message, Title, default)
            
            If myvalue <> "" Then
            List1.AddItem myvalue
            List2.AddItem "S  " 'MFT_STRING or  MFT_CHECKED
            List3.AddItem "EN "
            List4.AddItem "No"
            End If

Case 1002 '(Add/Separator)

'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "Having separator on top of menu is a dumb thing. Can't permit that !", , "M.C. Menu Voodoo "
                Exit Sub
                End If
List1.AddItem "/separator/"
List2.AddItem "S  "
List3.AddItem "EN "
List4.AddItem "No"

Case 1015 '(Move Right)
If List1.ListCount = 0 Then Exit Sub
List1.List(List1.ListIndex) = "----" & List1.List(List1.ListIndex)
Case 1014 '(Move Left)
If List1.ListCount = 0 Then Exit Sub
If Left(List1.List(List1.ListIndex), 1) = "-" Then
List1.List(List1.ListIndex) = Right(List1.List(List1.ListIndex), Len(List1.List(List1.ListIndex)) - 4)
End If
Case 1013 '(Move Up)
        If List1.ListIndex = 0 Then Exit Sub 'avoid mistakes
        ClickNoCanDo = True
        Dim ItemDataArray(3) As String
        'remember it
        ItemDataArray(0) = List1.List(List1.ListIndex)
        ItemDataArray(1) = List2.List(List1.ListIndex)
        ItemDataArray(2) = List3.List(List1.ListIndex)
        ItemDataArray(3) = List4.List(List1.ListIndex)
        'remember current index
        CurIndex = List1.ListIndex
        'Delete it
        List4.RemoveItem List1.ListIndex 'Action
        List2.RemoveItem List1.ListIndex 'Action
        List3.RemoveItem List1.ListIndex 'Action
        List1.RemoveItem List1.ListIndex 'Action
        'set index
        List1.ListIndex = CurIndex - 1
        List2.ListIndex = CurIndex - 1
        List3.ListIndex = CurIndex - 1
        List4.ListIndex = CurIndex - 1
        'insert it back one item up
        List1.AddItem ItemDataArray(0), List1.ListIndex
        List2.AddItem ItemDataArray(1), List2.ListIndex
        List3.AddItem ItemDataArray(2), List3.ListIndex
        List4.AddItem ItemDataArray(3), List4.ListIndex
        'set index
        List1.ListIndex = CurIndex - 1
        List2.ListIndex = CurIndex - 1
        List3.ListIndex = CurIndex - 1
        List4.ListIndex = CurIndex - 1
        'List1.ListIndex

'If List1.ListCount = 0 Then Exit Sub
'If List1.ListIndex > 0 Then
'List1.Selected(List1.ListIndex - 1) = True 'highlight one item up
'End If
ClickNoCanDo = False
Case 1012 '(Move Down)
If List1.ListIndex = List1.ListCount - 1 Then Exit Sub 'avoid mistakes
ClickNoCanDo = True
        Dim ItemDataArray1(3) As String
        'remember it
        ItemDataArray1(0) = List1.List(List1.ListIndex)
        ItemDataArray1(1) = List2.List(List1.ListIndex)
        ItemDataArray1(2) = List3.List(List1.ListIndex)
        ItemDataArray1(3) = List4.List(List1.ListIndex)
        'remember current index
        CurIndex = List1.ListIndex
        'Delete it
        List4.RemoveItem List1.ListIndex 'Action
        List2.RemoveItem List1.ListIndex 'Action
        List3.RemoveItem List1.ListIndex 'Action
        List1.RemoveItem List1.ListIndex 'Action
        'set index
        If CurIndex = List1.ListCount - 1 Then
         List1.ListIndex = CurIndex
        List2.ListIndex = CurIndex
        List3.ListIndex = CurIndex
        List4.ListIndex = CurIndex
        Else
        List1.ListIndex = CurIndex + 1
        List2.ListIndex = CurIndex + 1
        List3.ListIndex = CurIndex + 1
        List4.ListIndex = CurIndex + 1
        End If
        'insert it back one item up
        If CurIndex = List1.ListCount - 1 Then
        List1.AddItem ItemDataArray1(0)
        List2.AddItem ItemDataArray1(1)
        List3.AddItem ItemDataArray1(2)
        List4.AddItem ItemDataArray1(3)
        Else
        List1.AddItem ItemDataArray1(0), List1.ListIndex
        List2.AddItem ItemDataArray1(1), List2.ListIndex
        List3.AddItem ItemDataArray1(2), List3.ListIndex
        List4.AddItem ItemDataArray1(3), List4.ListIndex
        End If
        'set index
        List1.ListIndex = CurIndex + 1
        List2.ListIndex = CurIndex + 1
        List3.ListIndex = CurIndex + 1
        List4.ListIndex = CurIndex + 1
ClickNoCanDo = False

Case 1010 '(Menu Behaviour)

Frame7.Visible = True
'open colors table

'dbs stuff---------------------
Dim dbs As Database
Dim rst As Recordset
Set dbs = OpenDatabase(App.Path & "\my.mdb")
Set rst = dbs.OpenRecordset("MenuColors")

rst.MoveFirst
Do
List5.AddItem rst![Name]
rst.MoveNext
Loop Until rst.EOF

rst.Close
dbs.Close



Case 1007 '(Delete)

 'user errors .......
                If List1.ListCount = 0 Then
                MsgBox "You can't delete as there is nothing to delete !", , "M.C. Menu Voodoo "
                Exit Sub
                End If
                If List1.List(List1.ListIndex) = "" Then
                MsgBox "No item selected !", , "M.C. Menu Voodoo "
                Exit Sub
                End If
        List4.RemoveItem List1.ListIndex 'Action
        List2.RemoveItem List1.ListIndex 'Action
        List3.RemoveItem List1.ListIndex 'Action
        List1.RemoveItem List1.ListIndex 'Action

Case 1006 '(ChangeString)

If List1.ListIndex = -1 Then MsgBox "Can't change as there no selection", , "M.C. Menu Voodoo ": Exit Sub 'on empty list, or nozhing selected
StrA = InputBox("Enter new item string to replace current one", , List1.List(List1.ListIndex))
If StrA = "" Then Exit Sub
List1.List(List1.ListIndex) = StrA

Case 1004 '(Close Menu)
'nothing as just closing menu

Case Else 'from select case menusel , far up
End Select

End Sub

Private Sub List1_Scroll()
'get list1.topindex
a = SendMessage(List1.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&
End Sub












Sub PropertyMenu()
'First read property values of current item selected to be showed in menu that will pop up


'Checked......
Select Case Mid(List3.List(List3.ListIndex), 3, 1)

Case "C" 'on start checked
APICheck7 = True 'Checked ?
APICheck8 = False 'Checked ?/Not Checked

        Select Case Mid(List2.List(List2.ListIndex), 2, 1)
        Case "R"
        APICheck10 = True 'Checked ?/On Start Checked/With RadioButton
        APICheck14 = False 'Checked ?/On Start Checked/On Start Unchecked/With RadioButton
        Case "P"
        APICheck11 = True 'Checked ?/On Start Checked/With Pictures
        APICheck15 = False 'Checked ?/On Start Checked/On Start Unchecked/With Picturess
        Case "C"
        APICheck9 = True 'Checked ?/On Start Checked/With Checkmark
        APICheck13 = False 'Checked ?/On Start Checked/On Start Unchecked/With Checkmark
        End Select

Case "U" 'On start unchecked
APICheck7 = True 'Checked ?
APICheck8 = False 'Checked ?/Not Checked

Select Case Mid(List2.List(List2.ListIndex), 2, 1)
        Case "R"
        APICheck14 = True 'Checked ?/On Start Checked/On Start Unchecked/With RadioButton
        APICheck10 = False 'Checked ?/On Start Checked/With RadioButton
        Case "P"
        APICheck15 = True 'Checked ?/On Start Checked/On Start Unchecked/With Picturess
        APICheck11 = False 'Checked ?/On Start Checked/With Pictures
        Case "C"
        APICheck13 = True 'Checked ?/On Start Checked/On Start Unchecked/With Checkmark
        APICheck9 = False 'Checked ?/On Start Checked/With Checkmark
        End Select
Case " " 'not checked at all
APICheck7 = False 'Checked ?
APICheck8 = True 'Checked ?/Not Checked
APICheck14 = False 'Checked ?/On Start Checked/On Start Unchecked/With RadioButton
APICheck10 = False 'Checked ?/On Start Checked/With RadioButton
APICheck15 = False 'Checked ?/On Start Checked/On Start Unchecked/With Picturess
APICheck11 = False 'Checked ?/On Start Checked/With Pictures
APICheck13 = False 'Checked ?/On Start Checked/On Start Unchecked/With Checkmark
APICheck9 = False 'Checked ?/On Start Checked/With Checkmark
End Select

'COLUMN BREAK

Select Case Mid(List2.List(List2.ListIndex), 3, 1)
Case "L"
APICheck27 = True 'Insert Column Break
APICheck28 = True 'Insert Column Break/With dividing line
APICheck29 = False 'Insert Column Break/Without dividing line
Case "N"
APICheck27 = True 'Insert Column Break
APICheck28 = False 'Insert Column Break/With dividing line
APICheck29 = True 'Insert Column Break/Without dividing line
Case " "
APICheck27 = False 'Insert Column Break
APICheck28 = False 'Insert Column Break/Without dividing line
APICheck29 = False 'Insert Column Break/With dividing line
End Select

' Enabled / DISABLED / GRAYED

Select Case Mid(List3.List(List3.ListIndex), 1, 1)
Case "E"
APICheck3 = True 'Enabled
APICheck4 = False 'Disabled
APICheck5 = False 'Grayed
Case "D"
APICheck3 = False 'Enabled
APICheck4 = True 'Disabled
APICheck5 = False 'Grayed
Case "G"
APICheck3 = False 'Enabled
APICheck4 = False 'Disabled
APICheck5 = True 'Grayed
End Select

'Bold/Normal
Select Case Mid(List3.List(List3.ListIndex), 2, 1)
Case "B"
APICheck1 = True 'Bold
APICheck2 = False 'Normal
Case "N"
APICheck1 = False 'Bold
APICheck2 = True 'Normal
End Select

'pictures
Select Case List4.List(List4.ListIndex)
Case "No"
APICheck24 = False 'Picture ?
APICheck25 = False 'Picture ?/Yes
APICheck26 = True 'Picture ?/No
Case "Yes"
APICheck24 = True 'Picture ?
APICheck25 = True 'Picture ?/Yes
APICheck26 = False 'Picture ?/No
End Select



'--------------------------------------------------------------------------------------------------------------------------
'CODE AUTOGENERATED WITH:  M.C. Menu Voodoo
'Menu identification number: 1234567890
'Structure saved in: C:\ProjektiVB6\Api Menu(Ver31)\StateVer3.MAP
'---------------------------------------------------------------------------------------------------------------------------
Dim hPopupMenu1 As Long ' handle to the popup menu to display
Dim hPopupMenu2 As Long ' handle to the popup menu to display
Dim hPopupMenu3 As Long ' handle to the popup menu to display
Dim hPopupMenu4 As Long ' handle to the popup menu to display
Dim hPopupMenu5 As Long ' handle to the popup menu to display
Dim hPopupMenu6 As Long ' handle to the popup menu to display
Dim hPopupMenu7 As Long ' handle to the popup menu to display
Dim hPopupMenu8 As Long ' handle to the popup menu to display
Dim hPopupMenu9 As Long ' handle to the popup menu to display
Dim hPopupMenu10 As Long ' handle to the popup menu to display
Dim hPopupMenu11 As Long ' handle to the popup menu to display
Dim mii1 As MENUITEMINFO   ' describes menu items to add
Dim mii2 As MENUITEMINFO   ' describes menu items to add
Dim mii3 As MENUITEMINFO   ' describes menu items to add
Dim mii4 As MENUITEMINFO   ' describes menu items to add
Dim mii5 As MENUITEMINFO   ' describes menu items to add
Dim mii6 As MENUITEMINFO   ' describes menu items to add
Dim mii7 As MENUITEMINFO   ' describes menu items to add
Dim mii8 As MENUITEMINFO   ' describes menu items to add
Dim mii9 As MENUITEMINFO   ' describes menu items to add
Dim mii10 As MENUITEMINFO   ' describes menu items to add
Dim mii11 As MENUITEMINFO   ' describes menu items to add
Dim curpos As POINT_TYPE  ' holds the current mouse coordinates
Dim menusel As Long       ' ID of what the user selected in the popup menu
Dim retval As Long        ' generic return value
'Create the popup menus which are initialy empty.
hPopupMenu1 = CreatePopupMenu()
hPopupMenu2 = CreatePopupMenu()
hPopupMenu3 = CreatePopupMenu()
hPopupMenu4 = CreatePopupMenu()
hPopupMenu5 = CreatePopupMenu()
hPopupMenu6 = CreatePopupMenu()
hPopupMenu7 = CreatePopupMenu()
hPopupMenu8 = CreatePopupMenu()
hPopupMenu9 = CreatePopupMenu()
hPopupMenu10 = CreatePopupMenu()
hPopupMenu11 = CreatePopupMenu()
'Create the structure which is the base for all menus:
With mii1
.cbSize = Len(mii1) ' The size of this structure.
.fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE Or MIIM_SUBMENU ' Which elements of the structure to use.
End With
'Make all structures equal
mii2 = mii1
mii3 = mii1
mii4 = mii1
mii5 = mii1
mii6 = mii1
mii7 = mii1
mii8 = mii1
mii9 = mii1
mii10 = mii1
mii11 = mii1

With mii11 '(Bold)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck1, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1047 ' Assign this item an item identifier.
.dwTypeData = "Bold"
.cch = Len("Bold")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu11, 0, 1, mii11)

With mii11 '(Normal)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck2, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1046 ' Assign this item an item identifier.
.dwTypeData = "Normal"
.cch = Len("Normal")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu11, 1, 1, mii11)

With mii11 '(Make them ALL)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1045 ' Assign this item an item identifier.
.dwTypeData = "Make them ALL"
.cch = Len("Make them ALL")
.hSubMenu = hPopupMenu10
End With
retval = InsertMenuItem(hPopupMenu11, 2, 1, mii11)
With mii10 '(Make them ALL/Bold)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1029 ' Assign this item an item identifier.
.dwTypeData = "Bold"
.cch = Len("Bold")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu10, 0, 1, mii10)
With mii10 '(Make them ALL/Normal)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1028 ' Assign this item an item identifier.
.dwTypeData = "Normal"
.cch = Len("Normal")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu10, 1, 1, mii10)
With mii11 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1044 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 3, 1, mii11)

With mii11 '(Enabled)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck3, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1043 ' Assign this item an item identifier.
.dwTypeData = "Enabled"
.cch = Len("Enabled")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 4, 1, mii11)

With mii11 '(Disabled)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck4, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1042 ' Assign this item an item identifier.
.dwTypeData = "Disabled"
.cch = Len("Disabled")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu11, 5, 1, mii11)

With mii11 '(Grayed)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck5, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1041 ' Assign this item an item identifier.
.dwTypeData = "Grayed"
.cch = Len("Grayed")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu11, 6, 1, mii11)

With mii11 '(Make them ALL)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1040 ' Assign this item an item identifier.
.dwTypeData = "Make them ALL"
.cch = Len("Make them ALL")
.hSubMenu = hPopupMenu9
End With
retval = InsertMenuItem(hPopupMenu11, 7, 1, mii11)
With mii9 '(Make them ALL/Enabled)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1027 ' Assign this item an item identifier.
.dwTypeData = "Enabled"
.cch = Len("Enabled")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu9, 0, 1, mii9)
With mii9 '(Make them ALL/Disabled)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1026 ' Assign this item an item identifier.
.dwTypeData = "Disabled"
.cch = Len("Disabled")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu9, 1, 1, mii9)


With mii9 '(Make them ALL/Grayed)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1025 ' Assign this item an item identifier.
.dwTypeData = "Grayed"
.cch = Len("Grayed")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu9, 2, 1, mii9)

With mii11 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1039 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 8, 1, mii11)

With mii11 '(Checked ?)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck7, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1038 ' Assign this item an item identifier.
.dwTypeData = "Checked ?"
.cch = Len("Checked ?")
.hSubMenu = hPopupMenu8
End With
retval = InsertMenuItem(hPopupMenu11, 9, 1, mii11)


With mii8 '(Checked ?/Not Checked)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck8, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1024 ' Assign this item an item identifier.
.dwTypeData = "Not Checked"
.cch = Len("Not Checked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu8, 0, 1, mii8)

With mii8 '(Checked ?/On Start Checked)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1023 ' Assign this item an item identifier.
.dwTypeData = "On Start Checked"
.cch = Len("On Start Checked")
.hSubMenu = hPopupMenu4
End With
retval = InsertMenuItem(hPopupMenu8, 1, 1, mii8)

With mii4 '(Checked ?/On Start Checked/With Checkmark)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck9, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1011 ' Assign this item an item identifier.
.dwTypeData = "With Checkmark"
.cch = Len("With Checkmark")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu4, 0, 1, mii4)


With mii4 '(Checked ?/On Start Checked/With RadioButton)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck10, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1010 ' Assign this item an item identifier.
.dwTypeData = "With RadioButton"
.cch = Len("With RadioButton")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu4, 1, 1, mii4)


With mii4 '(Checked ?/On Start Checked/With Pictures)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck11, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1009 ' Assign this item an item identifier.
.dwTypeData = "With Pictures"
.cch = Len("With Pictures")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu4, 2, 1, mii4)

With mii8 '(Checked ?/On Start Checked/On Start Unchecked)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck12, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1022 ' Assign this item an item identifier.
.dwTypeData = "On Start Unchecked"
.cch = Len("On Start Unchecked")
.hSubMenu = hPopupMenu3
End With
retval = InsertMenuItem(hPopupMenu8, 2, 1, mii8)

With mii3 '(Checked ?/On Start Checked/On Start Unchecked/With Checkmark)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck13, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1008 ' Assign this item an item identifier.
.dwTypeData = "With Checkmark"
.cch = Len("With Checkmark")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu3, 0, 1, mii3)

With mii3 '(Checked ?/On Start Checked/On Start Unchecked/With RadioButton)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck14, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1007 ' Assign this item an item identifier.
.dwTypeData = "With RadioButton"
.cch = Len("With RadioButton")
.hSubMenu = 0
End With

retval = InsertMenuItem(hPopupMenu3, 1, 1, mii3)


With mii3 '(Checked ?/On Start Checked/On Start Unchecked/With Pictures)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck15, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1006 ' Assign this item an item identifier.
.dwTypeData = "With Pictures"
.cch = Len("With Pictures")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu3, 2, 1, mii3)

With mii11 '(Make Them ALL)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1037 ' Assign this item an item identifier.
.dwTypeData = "Make Them ALL"
.cch = Len("Make Them ALL")
.hSubMenu = hPopupMenu7
End With
retval = InsertMenuItem(hPopupMenu11, 10, 1, mii11)

With mii7 '(Make Them ALL/Not Checked)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck16, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1021 ' Assign this item an item identifier.
.dwTypeData = "Not Checked"
.cch = Len("Not Checked")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu7, 0, 1, mii7)

With mii7 '(Make Them ALL/On Start Checked)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1020 ' Assign this item an item identifier.
.dwTypeData = "On Start Checked"
.cch = Len("On Start Checked")
.hSubMenu = hPopupMenu2
End With
retval = InsertMenuItem(hPopupMenu7, 1, 1, mii7)

With mii2 '(Make Them ALL/On Start Checked/With Checkmark)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck17, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1005 ' Assign this item an item identifier.
.dwTypeData = "With Checkmark"
.cch = Len("With Checkmark")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 0, 1, mii2)


With mii2 '(Make Them ALL/On Start Checked/With RadioButton)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck18, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1004 ' Assign this item an item identifier.
.dwTypeData = "With RadioButton"
.cch = Len("With RadioButton")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 1, 1, mii2)

With mii2 '(Make Them ALL/On Start Checked/With Pictures)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck19, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1003 ' Assign this item an item identifier.
.dwTypeData = "With Pictures"
.cch = Len("With Pictures")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu2, 2, 1, mii2)

With mii7 '(Make Them ALL/On Start Checked/On Start Unchecked)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck20, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1019 ' Assign this item an item identifier.
.dwTypeData = "On Start Unchecked"
.cch = Len("On Start Unchecked")
.hSubMenu = hPopupMenu1
End With
retval = InsertMenuItem(hPopupMenu7, 2, 1, mii7)


With mii1 '(Make Them ALL/On Start Checked/On Start Unchecked/With Checkmark)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck21, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1002 ' Assign this item an item identifier.
.dwTypeData = "With Checkmark"
.cch = Len("With Checkmark")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 0, 1, mii1)


With mii1 '(Make Them ALL/On Start Checked/On Start Unchecked/With RadioButton)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck22, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1001 ' Assign this item an item identifier.
.dwTypeData = "With RadioButton"
.cch = Len("With RadioButton")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 1, 1, mii1)


With mii1 '(Make Them ALL/On Start Checked/On Start Unchecked/With Pictures)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck23, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1000 ' Assign this item an item identifier.
.dwTypeData = "With Pictures"
.cch = Len("With Pictures")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu1, 2, 1, mii1)

With mii11 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1036 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 11, 1, mii11)

With mii11 '(Picture ?)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck24, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1035 ' Assign this item an item identifier.
.dwTypeData = "Picture ?"
.cch = Len("Picture ?")
.hSubMenu = hPopupMenu6
End With
retval = InsertMenuItem(hPopupMenu11, 12, 1, mii11)

With mii6 '(Picture ?/Yes)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck25, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1018 ' Assign this item an item identifier.
.dwTypeData = "Yes"
.cch = Len("Yes")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu6, 0, 1, mii6)

With mii6 '(Picture ?/No)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck26, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1017 ' Assign this item an item identifier.
.dwTypeData = "No"
.cch = Len("No")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu6, 1, 1, mii6)

With mii6 '(Picture ?/Yes for all)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1016 ' Assign this item an item identifier.
.dwTypeData = "Yes for all"
.cch = Len("Yes for all")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu6, 2, 1, mii6)
With mii6 '(Picture ?/No for all)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1015 ' Assign this item an item identifier.
.dwTypeData = "No for all"
.cch = Len("No for all")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu6, 3, 1, mii6)
With mii11 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1034 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 13, 1, mii11)

With mii11 '(Insert Column Break)
.fType = MFT_STRING
.fState = MFS_ENABLED Or MFS_DEFAULT Or IIf(APICheck27, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1033 ' Assign this item an item identifier.
.dwTypeData = "Insert Column Break"
.cch = Len("Insert Column Break")
.hSubMenu = hPopupMenu5
End With
retval = InsertMenuItem(hPopupMenu11, 14, 1, mii11)

With mii5 '(Insert Column Break/With dividing line)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck28, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1014 ' Assign this item an item identifier.
.dwTypeData = "With dividing line"
.cch = Len("With dividing line")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 0, 1, mii5)

With mii5 '(Insert Column Break/Without dividing line)
.fType = MFT_STRING
.fState = MFS_ENABLED Or IIf(APICheck29, MFS_CHECKED, MFS_UNCHECKED)
.wID = 1013 ' Assign this item an item identifier.
.dwTypeData = "Without dividing line"
.cch = Len("Without dividing line")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 1, 1, mii5)

With mii5 '(Insert Column Break/Delete col. break)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1012 ' Assign this item an item identifier.
.dwTypeData = "Delete col. break"
.cch = Len("Delete col. break")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu5, 2, 1, mii5)

With mii11 '(/separator/)
.fType = MFT_SEPARATOR
.fState = MFS_ENABLED
.wID = 1032 ' Assign this item an item identifier.
.dwTypeData = "/separator/"
.cch = Len("/separator/")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 15, 1, mii11)
With mii11 '(Close menu)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1031 ' Assign this item an item identifier.
.dwTypeData = "Close menu"
.cch = Len("Close menu")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 16, 1, mii11)
With mii11 '(Help)
.fType = MFT_STRING
.fState = MFS_ENABLED
.wID = 1030 ' Assign this item an item identifier.
.dwTypeData = "Help"
.cch = Len("Help")
.hSubMenu = 0
End With
retval = InsertMenuItem(hPopupMenu11, 17, 1, mii11)
'The following code is for adding pictures into menus, if there are any!
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
'------------------------------------------------------------
retval = GetCursorPos(curpos)
menusel = TrackPopupMenu(hPopupMenu11, TPM_TOPALIGN Or TPM_NONOTIFY Or TPM_RETURNCMD Or TPM_RIGHTALIGN Or TPM_RIGHTBUTTON, curpos.x, curpos.y, 0, Form1.hwnd, 0)
retval = DestroyMenu(hPopupMenu11)
'------------------------------------------------------------------------------------------------
'DOWN BELOW  PUT IN YOUR CODE MANUALY !!!!
'------------------------------------------------------------------------------------------------
Dim IncomingString As String
Dim outputstring As String

Select Case menusel
Case 1047 '(Bold)
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "B" & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
        If APICheck1 = True Then   'in case item become unchecked .....
        APICheck1 = False
        Else 'in case item become checked .....
        APICheck1 = True
        End If
Case 1046 '(Normal)
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "N" & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
        If APICheck2 = True Then 'in case item become unchecked .....
        APICheck2 = False
        Else 'in case item become checked .....
        APICheck2 = True
        End If
Case 1029 '(Make them ALL/Bold)
        For i = 0 To List3.ListCount - 1
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "B" & Mid(IncomingString, 3, 1)
        List3.List(i) = outputstring
        Next i
Case 1028 '(Make them ALL/Normal)
        For i = 0 To List3.ListCount - 1
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "N" & Mid(IncomingString, 3, 1)
        List3.List(i) = outputstring
        Next i
Case 1043 '(Enabled)
        IncomingString = List3.List(List3.ListIndex)
        outputstring = "E" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
        If APICheck3 = True Then 'in case item become unchecked .....
        APICheck3 = False
        Else 'in case item become checked .....
        APICheck3 = True
        End If
Case 1042 '(Disabled)
        IncomingString = List3.List(List3.ListIndex)
        outputstring = "D" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
        If APICheck4 = True Then 'in case item become unchecked .....
        APICheck4 = False
        Else 'in case item become checked .....
        APICheck4 = True
        End If
Case 1041 '(Grayed)
        IncomingString = List3.List(List3.ListIndex)
        outputstring = "G" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(List3.ListIndex) = outputstring
        If APICheck5 = True Then 'in case item become unchecked .....
        APICheck5 = False
        Else 'in case item become checked .....
        APICheck5 = True
        End If
Case 1027 '(Make them ALL/Enabled)
        For i = 0 To List3.ListCount - 1
        IncomingString = List3.List(i)
        outputstring = "E" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(i) = outputstring
        Next i
Case 1026 '(Make them ALL/Disabled)
        For i = 0 To List3.ListCount - 1
        IncomingString = List3.List(i)
        outputstring = "D" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(i) = outputstring
        Next i
Case 1025 '(Make them ALL/Grayed)
        For i = 0 To List3.ListCount - 1
        IncomingString = List3.List(i)
        outputstring = "G" & Mid(IncomingString, 2, 1) & Mid(IncomingString, 3, 1)
        List3.List(i) = outputstring
        Next i
Case 1024 '(Checked ?/Not Checked)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & " "
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & " " & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        If APICheck8 = True Then 'in case item become unchecked .....
        APICheck8 = False
        Else 'in case item become checked .....
        APICheck8 = True
        End If
Case 1011 '(Checked ?/On Start Checked/With Checkmark)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "C" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'list4
        List4.List(List4.ListIndex) = "No"
        
        If APICheck9 = True Then 'in case item become unchecked .....
        APICheck9 = False
        Else 'in case item become checked .....
        APICheck9 = True
        End If
Case 1010 '(Checked ?/On Start Checked/With RadioButton)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "R" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'list4
        List4.List(List4.ListIndex) = "No"
        If APICheck10 = True Then 'in case item become unchecked .....
        APICheck10 = False
        Else 'in case item become checked .....
        APICheck10 = True
        End If
Case 1009 '(Checked ?/On Start Checked/With Pictures)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "P" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'List4
        List4.List(List4.ListIndex) = "Yes"
        
        If APICheck11 = True Then 'in case item become unchecked .....
        APICheck11 = False
        Else 'in case item become checked .....
        APICheck11 = True
        End If
Case 1008 '(Checked ?/On Start Checked/On Start Unchecked/With Checkmark)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "C" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'list4
        List4.List(List4.ListIndex) = "No"
        If APICheck13 = True Then 'in case item become unchecked .....
        APICheck13 = False
        Else 'in case item become checked .....
        APICheck13 = True
        End If

Case 1007 '(Checked ?/On Start Checked/On Start Unchecked/With RadioButton)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "R" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'list4
        List4.List(List4.ListIndex) = "No"
        If APICheck14 = True Then 'in case item become unchecked .....
        APICheck14 = False
        Else 'in case item become checked .....
        APICheck14 = True
        End If
Case 1006 '(Checked ?/On Start Checked/On Start Unchecked/With Pictures)
        'list3
        IncomingString = List3.List(List3.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(List3.ListIndex) = outputstring
        'List2
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & "P" & Mid(IncomingString, 3, 1)
        List2.List(List2.ListIndex) = outputstring
        'List4
        List4.List(List4.ListIndex) = "Yes"
        If APICheck15 = True Then 'in case item become unchecked .....
        APICheck15 = False
        Else 'in case item become checked .....
        APICheck15 = True
        End If
Case 1021 '(Make Them ALL/Not Checked)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & " "
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & " " & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        Next i

Case 1005 '(Make Them ALL/On Start Checked/With Checkmark)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "C" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "No"
        Next i

Case 1004 '(Make Them ALL/On Start Checked/With RadioButton)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "R" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "No"
        Next i
If APICheck18 = True Then
'in case item become unchecked .....
APICheck18 = False
Else
'in case item become checked .....
APICheck18 = True
End If
Case 1003 '(Make Them ALL/On Start Checked/With Pictures)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "C"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "P" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "Yes"
        Next i
If APICheck19 = True Then
'in case item become unchecked .....
APICheck19 = False
Else
'in case item become checked .....
APICheck19 = True
End If
Case 1002 '(Make Them ALL/On Start Checked/On Start Unchecked/With Checkmark)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "C" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "No"
        Next i
If APICheck21 = True Then
'in case item become unchecked .....
APICheck21 = False
Else
'in case item become checked .....
APICheck21 = True
End If
Case 1001 '(Make Them ALL/On Start Checked/On Start Unchecked/With RadioButton)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "R" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "No"
        Next i
If APICheck22 = True Then
'in case item become unchecked .....
APICheck22 = False
Else
'in case item become checked .....
APICheck22 = True
End If
Case 1000 '(Make Them ALL/On Start Checked/On Start Unchecked/With Pictures)
        For i = 0 To List1.ListCount - 1
        'list3
        IncomingString = List3.List(i)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "U"
        List3.List(i) = outputstring
        'List2
        IncomingString = List2.List(i)
        outputstring = Mid(IncomingString, 1, 1) & "P" & Mid(IncomingString, 3, 1)
        List2.List(i) = outputstring
        'list4
        List4.List(i) = "Yes"
        Next i
If APICheck23 = True Then
'in case item become unchecked .....
APICheck23 = False
Else
'in case item become checked .....
APICheck23 = True
End If
Case 1018 '(Picture ?/Yes)
        If List1.List(List4.ListIndex) = "/separator/" Then MsgBox " Separators can't have pictures !", , "M.C. Menu Voodoo ": Exit Sub
        List4.List(List4.ListIndex) = "Yes"
        If APICheck25 = True Then 'in case item become unchecked .....
        APICheck25 = False
        Else 'in case item become checked .....
        APICheck25 = True
        End If
Case 1017 '(Picture ?/No)
        List4.List(List4.ListIndex) = "No"
        If APICheck26 = True Then 'in case item become unchecked .....
        APICheck26 = False
        Else 'in case item become checked .....
        APICheck26 = True
        End If
Case 1016 '(Picture ?/Yes for all)
        For i = 0 To List4.ListCount - 1
                If List1.List(i) <> "/separator/" And Mid(List2.List(i), 2, 1) <> "C" And Mid(List2.List(i), 2, 1) <> "R" Then
                List4.List(i) = "Yes"
                End If
        Next i
        MsgBox "Checked items which check style is not Picture cant have pictures, because they would owerride checkmarks, allso separators can't have pictures. To all the rest pictures have been added.", , "M.C. Menu Voodoo "
Case 1015 '(Picture ?/No for all)
        For i = 0 To List4.ListCount - 1
        List4.List(i) = "No"
        Next i
Case 1014 '(Insert Column Break/With dividing line)
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "L"
        List2.List(List2.ListIndex) = outputstring
        If APICheck28 = True Then 'in case item become unchecked .....
        APICheck28 = False
        Else 'in case item become checked .....
        APICheck28 = True
        End If
Case 1013 '(Insert Column Break/Without dividing line)
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & "N"
        List2.List(List2.ListIndex) = outputstring
        If APICheck29 = True Then 'in case item become unchecked .....
        APICheck29 = False
        Else 'in case item become checked .....
        APICheck29 = True
        End If
Case 1012 '(Insert Column Break/Delete col. break)
        IncomingString = List2.List(List2.ListIndex)
        outputstring = Mid(IncomingString, 1, 1) & Mid(IncomingString, 2, 1) & " "
        List2.List(List2.ListIndex) = outputstring
        
Case 1031 '(Close menu)
        'do nothing as it closes itself
Case 1030 '(Help)
MsgBox "In ItemType list box:" & Chr(10) & _
        "1) S = String(MFT_STRING), that is default, must be here in any case" & Chr(10) & _
        "2) R = RadioButton(MFT_RADIOCHECK), type of checkmark" & Chr(10) & _
        "    C = Normal checkmark(nothinga as it is default type of checkmark)" & Chr(10) & _
        "    P = Pictures(nothing - you will use 2 different .bmp)" & Chr(10) & _
        "3) L = Column break with dividing Line(MFT_MENUBARBREAK)" & Chr(10) & _
        "    N = Column break without dividing Line(MFT_MENUBREAK)" & Chr(10) & _
        "" & Chr(10) & _
        "In ItemState list box:" & Chr(10) & _
        "1) E = Enabled(MFS_ENABLED) - this is added here as default" & Chr(10) & _
        "    D = Disabled(MFS_DISABLED)" & Chr(10) & _
        "    G = Grayed(MFS_GRAYED)" & Chr(10) & _
        "2) B = Bold(MFS_DEFAULT)" & Chr(10) & _
        "    N = not bold(nothing as this is default)" & Chr(10) & _
        "    Upper two are optional on second place" & Chr(10) & _
        "3) C = Checked(MFS_CHECKED)" & Chr(10) & _
        "    U = Unchecked(MFS_UNCHECHED)" & Chr(10) & _
        "    Upper two are optional on third place", , "M.C. Menu Voodoo "
Case Else
End Select
End Sub



Private Sub List2_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List2.ListIndex
List3.ListIndex = List2.ListIndex
List4.ListIndex = List2.ListIndex
ClickNoCanDo = False
PropertyMenu
End Sub

Private Sub List2_Scroll()
'get list2.topindex
a = SendMessage(List2.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list2 topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub

Private Sub List3_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List3.ListIndex
List2.ListIndex = List3.ListIndex
List4.ListIndex = List3.ListIndex
ClickNoCanDo = False
PropertyMenu
End Sub

Private Sub List3_Scroll()
'get list3.topindex
a = SendMessage(List3.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List4.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub

Private Sub List4_Click()
If ClickNoCanDo = True Then Exit Sub
'teh following lines triggere some click events so inhibit that...
ClickNoCanDo = True
List1.ListIndex = List4.ListIndex
List2.ListIndex = List4.ListIndex
List3.ListIndex = List4.ListIndex
ClickNoCanDo = False
PropertyMenu

End Sub

Private Sub List4_Scroll()
'get list4.topindex
a = SendMessage(List4.hwnd, LB_GETTOPINDEX, ByVal 0&, ByVal 0&)
'set list topindex = list1 topindex
SendMessage List1.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List2.hwnd, LB_SETTOPINDEX, a, ByVal 0&
SendMessage List3.hwnd, LB_SETTOPINDEX, a, ByVal 0&

End Sub

Private Sub List5_Click()
'dbs stuff---------------------
Dim dbs As Database
Dim rst As Recordset
Set dbs = OpenDatabase(App.Path & "\my.mdb")
Set rst = dbs.OpenRecordset("MenuColors")
rst.MoveFirst
Do
        If rst![Name] = List5.List(List5.ListIndex) Then
        Label5(2).ForeColor = rst![TextColor]
        Label5(2).BackColor = rst![BackGroundColor]
        Exit Do
        End If
rst.MoveNext
Loop Until rst.EOF
rst.Close
dbs.Close

End Sub

Private Sub List6_Click()
Command17.Enabled = True
List7.Clear

Open FileVbOpened For Input As #1

 


Select Case List6.List(List6.ListIndex)
Case "'Take Them All'"
           Do While Not EOF(1)
        Line Input #1, textline 'get line into variable
        TrimmedTextLine = LTrim(textline)
        
                If Left(TrimmedTextLine, 13) = "Begin VB.Menu" Then 'menu detected
                'get its name
                MenuItemName = Right(TrimmedTextLine, Len(TrimmedTextLine) - 14)
                'calculate the difference ... which give as depth of item ..
                depth = Len(textline) - Len(TrimmedTextLine)
                realdepth = (depth / 3) - 1 'the depth of item
                
                'get caption which is in next line
                Line Input #1, textline 'this one should hold item caption
                TrimmedTextLine = LTrim(textline)
                ICaption = Right(TrimmedTextLine, Len(TrimmedTextLine) - 21)
                ICaption = Mid(ICaption, 1, Len(ICaption) - 1) 'some corrections
                 List1.AddItem String(realdepth * 4, "-") & ICaption
                 List2.AddItem "SC " 'MFT_STRING or MFT_CHECKED
                 List3.AddItem "EN " 'MFS_ENABLED
                 List4.AddItem "No" 'Pictures
                 
                 'menu item name - use it to suck out the code
                 'MenuItemName = Right(TrimmedTextLine, Len(TrimmedTextLine) - 14)
                 List7.AddItem "Private Sub " & Left(MenuItemName, Len(MenuItemName) - 1) & "_Click()"
            
                End If
        Loop
        Close #1
           '*****'store existing menu code into file******
           
           Open App.Path & "\MenuVBMenuSuckedCode.txt" For Output As #2 'file to write to
           For i = 0 To List7.ListCount - 1
           Open FileVbOpened For Input As #1 'file to read from
           
           Do While Not EOF(1)
           Line Input #1, textline 'get line into variable
         
               If textline = List7.List(i) Then
               
                     Print #2, textline
                     Do
                     Line Input #1, textline
                     If textline = "End Sub" Then Print #2, textline:  Print #2, "*******************": GoTo 10
                     Print #2, textline
                     Loop
               End If
           Loop
10
           Close #1
           
           Next i
           Close #2
           '*****'end of store existing menu code into file******
      
Case Else

        Do While Not EOF(1)
        Line Input #1, textline 'get line into variable
        TrimmedTextLine = LTrim(textline)
                
            
                If Left(TrimmedTextLine, 13) = "Begin VB.Menu" Then 'menu detected
                
                        
                        'get name
                            MenuItemName = Right(TrimmedTextLine, Len(TrimmedTextLine) - 14)
                            If MenuItemName = List6.List(List6.ListIndex) & " " Then 'found the start of structure
                                    
                                    Do Until textline = "   End"
                                        
                                            Line Input #1, textline 'get line into variable
                                            TrimmedTextLine = LTrim(textline)
                                        If Left(TrimmedTextLine, 13) = "Begin VB.Menu" Then 'menu detected
                                            MenuItemName = Right(TrimmedTextLine, Len(TrimmedTextLine) - 14)
                                            'the depth of item
                                            depth = Len(textline) - Len(TrimmedTextLine)
                                            realdepth = (depth / 3) - 2
                                            
                                            'get caption which is in next line
                                            Line Input #1, textline 'this one should hold item caption
                                            TrimmedTextLine = LTrim(textline)
                                            ICaption = Right(TrimmedTextLine, Len(TrimmedTextLine) - 21)
                                            ICaption = Mid(ICaption, 1, Len(ICaption) - 1) 'some corrections
                        
                                            List1.AddItem String(realdepth * 4, "-") & ICaption
                                            List2.AddItem "SC " 'MFT_STRING or MFT_CHECKED
                                            List3.AddItem "EN " 'MFS_ENABLED
                                            List4.AddItem "No" 'Pictures
                                            List7.AddItem "Private Sub " & Left(MenuItemName, Len(MenuItemName) - 1) & "_Click()"
                                        End If
                                    Loop
                            End If
              
              End If
              Loop
              Close #1
     '*****'store existing menu code into file******
           
           Open App.Path & "\MenuVBMenuSuckedCode.txt" For Output As #2 'file to write to
           For i = 0 To List7.ListCount - 1
           Open FileVbOpened For Input As #1 'file to read from
           
           Do While Not EOF(1)
           Line Input #1, textline 'get line into variable
         
               If textline = List7.List(i) Then
               
                     Print #2, textline
                     Do
                     Line Input #1, textline
                     If textline = "End Sub" Then Print #2, textline: Print #2, "*******************": GoTo 11
                     Print #2, textline
                     Loop
               End If
           Loop
11
           Close #1
           
           Next i
           Close #2
           '*****'end of store existing menu code into file******

End Select









Picture2.Visible = False
Close #1
Me.Caption = "M.C. Menu Voodoo"

MsgBox "To insure everything will work right,BEFORE GENERATING any CODE, save this structure to map file ! As you might noticed Vb Sucker button becomed enabled. Clicking it will generate structure code and at the same time suck in appropriate hand written code from frm. where vb menu exist.", , "M.C. Menu Voodoo"





End Sub




Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
Label5(2).BackColor = 12632256
Label5(2).ForeColor = 0
Command2(0).Enabled = False
Command2(1).Enabled = False
Command2(2).Enabled = False
Command2(3).Enabled = False
List5.Enabled = False
Case 1
Command2(0).Enabled = True
Command2(1).Enabled = True
Command2(2).Enabled = True
Command2(3).Enabled = True
List5.Enabled = True
Case Else
End Select
End Sub

Private Sub SelectionButton_Click(Index As Integer)
Select Case Index
Case 0 'main code
RichTextBox1.LoadFile App.Path & "\MenuMainCode.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText
Case 1 'form_load
RichTextBox1.LoadFile App.Path & "\MenuFormLoadCode.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText
Case 2 'general section of form
      Select Case SelectionButton(2).Caption
      Case "General section of form"
      RichTextBox1.LoadFile App.Path & "\MenuGeneralSectionCode.txt"
      Clipboard.Clear
      Clipboard.SetText RichTextBox1.Text, vbCFText
      Case Else 'FormQueryUnload
      RichTextBox1.LoadFile App.Path & "\MenuFormQueryUnloadCode.txt"
      Clipboard.Clear
      Clipboard.SetText RichTextBox1.Text, vbCFText
      End Select
Case 3 ' pictures
RichTextBox1.LoadFile App.Path & "\MenuPicturesComment.txt"
Clipboard.Clear
Clipboard.SetText RichTextBox1.Text, vbCFText

End Select
End Sub

Private Sub Timer1_Timer()
Dim Mes As Msg
'DoEvents
WaitMessage
PeekMessage Mes, Me.hwnd, 161, 161, PM_NOREMOVE 'And Message.wParam = 3 Then
If Mes.Message = 161 And Mes.wParam = 3 Then
Beep
        'change sys menu at this point
        Set SourceForm = Me
        SysMenuModify (Me.hwnd)
Timer1.Enabled = False
End If
'DoEvents
'Text1.Text = Mes.Message
End Sub



Private Sub SysTrayPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Dim Msg As Long
        Msg = (x And &HFF) * &H100
        Select Case Msg
        Case 3840 'left mouse button down
        Form1.Visible = True
        SystemTrayDeleteIcon SysTrayPic
        Case Else
        End Select
End Sub
