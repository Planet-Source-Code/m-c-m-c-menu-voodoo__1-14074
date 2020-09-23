VERSION 5.00
Begin VB.Form StartUpForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "StartUpForm"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   18
      Top             =   2280
      Width           =   1200
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   11
      Left            =   5280
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   10
      Left            =   4800
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   9
      Left            =   4320
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   8
      Left            =   3840
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   7
      Left            =   3360
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   6
      Left            =   2880
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   5
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   4
      Left            =   1920
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   3
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   2
      Left            =   960
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox SysTrayPic 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   480
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   2
         Text            =   "8"
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Delay in seconds"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox FinalPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   960
      ScaleHeight     =   4095
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.PictureBox PictContainer 
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   480
         ScaleHeight     =   1335
         ScaleWidth      =   1575
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   360
            Top             =   600
         End
      End
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1335
         Left            =   1680
         ScaleHeight     =   1335
         ScaleWidth      =   2775
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   135
         Left            =   2400
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image Image1 
         Height          =   1575
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   6840
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu PicMnu 
      Caption         =   "PicMnu"
      Begin VB.Menu LetMouseGo 
         Caption         =   "Let mouse go"
      End
      Begin VB.Menu DeleteImage 
         Caption         =   "Delete image"
      End
      Begin VB.Menu SaveImage 
         Caption         =   "Save image"
      End
      Begin VB.Menu ResizeImage 
         Caption         =   "Resize image(shrink)"
      End
      Begin VB.Menu MoveImage 
         Caption         =   "Move image"
      End
      Begin VB.Menu GoGetAnotherPicture 
         Caption         =   "Go get another picture"
      End
   End
   Begin VB.Menu FormMnu 
      Caption         =   "FormMnu"
      Begin VB.Menu GoGetNewPictureNoDelay 
         Caption         =   "Go get new picture"
      End
      Begin VB.Menu GoGetPic 
         Caption         =   "Go get new picture (10 sec. delay)"
      End
      Begin VB.Menu SaveEntireScreenShot 
         Caption         =   "Save entire screen shot"
      End
      Begin VB.Menu CutOutAndSave 
         Caption         =   "Cut out and save"
      End
      Begin VB.Menu ExitMe 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "StartUpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim a ' to show count down in forms caption bar
Dim mouseexit As Boolean 'avoid some problems at interaction of menu and my bas
Dim CutCounting As Boolean 'vil break counting if user click in systray
Dim ImageIdentifier As Byte ' images index
Dim ImageDeleted As Boolean ' prevents capturing mouse into disappeard image
Dim BeenThere As Boolean





Private Sub DeleteImage_Click()
Unload Image1(ImageIdentifier)
ImageDeleted = True
End Sub

Private Sub ExitMe_Click()
End
End Sub

Private Sub FinalPicture_Click()

Select Case ActionTaken

Case "GLUER"

NumberOfLoadedImages = NumberOfLoadedImages + 1
Load Image1(NumberOfLoadedImages)

'Picture3.Visible = False
DoEvents
'Set FinalPicture.Picture = hDCToPicture(GetDC(Picture3.hwnd), (startX / Screen.TwipsPerPixelX), (startY / Screen.TwipsPerPixelY), (Picture3.Width / Screen.TwipsPerPixelX), (Picture3.Height / Screen.TwipsPerPixelY))
Image1(NumberOfLoadedImages).Width = Picture3.Width
Image1(NumberOfLoadedImages).Height = Picture3.Height
Image1(NumberOfLoadedImages).Top = Picture3.Top
Image1(NumberOfLoadedImages).Left = Picture3.Left
Image1(NumberOfLoadedImages).Picture = Picture3.Picture
Picture3.Visible = False
Image1(NumberOfLoadedImages).Visible = True
ActionTaken = "NONE"
Case "NONE"
ClipCursor ByVal 0&
PopupMenu FormMnu
ExitMe.Visible = True
Case Else

End Select
'Me.MousePointer = 0
'ActionTaken = "NONE"
End Sub

Private Sub Form_Activate()
'Dim c As control
'For Each c In Me
'If TypeOf c Is Menu Then
'          If c.Visible = False Then
'          List1.AddItem c.Caption
'          End If
'End If
'Next

FinalPicture.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If ActionTaken = "RESIZER" Then
                 Select Case KeyCode
                 Case vbKeyDown
                 Image1(ImageIdentifier).Height = Image1(ImageIdentifier).Height - Screen.TwipsPerPixelY
                 Case vbKeyUp
                 Image1(ImageIdentifier).Height = Image1(ImageIdentifier).Height + Screen.TwipsPerPixelY
                 Case vbKeyRight
                 Image1(ImageIdentifier).Width = Image1(ImageIdentifier).Width + Screen.TwipsPerPixelX
                 Case vbKeyLeft
                 Image1(ImageIdentifier).Width = Image1(ImageIdentifier).Width - Screen.TwipsPerPixelX
                 Case Else
                 End Select


End If



End Sub

Private Sub Form_Load()
'-----------------------------------
ActionTaken = "NONE"
'-----------------------------------
Me.Left = 0
Me.Top = 0
Me.Width = Screen.Width
Me.Height = Screen.Height
FinalPicture.Top = 0
FinalPicture.Left = 0
FinalPicture.Width = Me.Width
FinalPicture.Height = Me.Height
SystemTrayIcon SysTrayPic(11), Me, "ADD"
CommonDialog1.InitDir = App.Path

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SystemTrayIcon SysTrayPic(11), Me, "DELETE"

End Sub

Private Sub GoGetAnotherPicture_Click()
GoGetPic_Click
End Sub

Private Sub GoGetNewPictureNoDelay_Click()
mouseexit = True
ClipCursor ByVal 0
'DO STUFF TO GLUE THINGS INTO OUR PICTURE
ActionTaken = "GLUER"

a = 0 'count down
Me.Hide
'do stuff
Me.Visible = False
Picture1.Visible = False
Form1.Show
Load Form1
DoEvents

End Sub

Private Sub GoGetPic_Click()

mouseexit = True
ClipCursor ByVal 0
'DO STUFF TO GLUE THINGS INTO OUR PICTURE
ActionTaken = "GLUER"

a = 0 'count down
Me.Hide
'Timer1.Interval = Text1.Text * 1000
Timer2.Interval = 1000

Timer2.Enabled = True 'will count down
Me.Visible = False
End Sub


Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
  Case 2 'left click exit image
  'ClipCursor ByVal 0
  Case 1
  ImageIdentifier = Index
  ActionTaken = "NONE"
  Me.MousePointer = 0

  PopupMenu PicMnu
  
  If mouseexit = True Or ImageDeleted = True Then GoTo Skip
  MCCaptureMouseCursorIntoNestedArea StartUpForm, FinalPicture, Image1(Index)
Skip:
mouseexit = False
ImageDeleted = False
  Case Else
  End Select
End Sub

Private Sub LetMouseGo_Click()
mouseexit = True
ClipCursor ByVal 0
End Sub

Private Sub FinalPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case ActionTaken
Case "GLUER"
Picture3.Left = X - 1
Picture3.Top = Y - 1
startX = X
startY = Y

Case Else
End Select
End Sub


Private Sub MoveImage_Click()
'take it up into picturebox
Picture3.Width = Image1(ImageIdentifier).Width
Picture3.Height = Image1(ImageIdentifier).Height
Set Picture3.Picture = Image1(ImageIdentifier).Picture

Picture3.ZOrder 0
'delete it from current position
DeleteImage_Click
'enable pic3 to be moved together with mouse
Picture3.Visible = True
Picture3.Refresh


ActionTaken = "GLUER"
End Sub

Private Sub Picture3_Click()
Select Case ActionTaken

Case "GLUER"

NumberOfLoadedImages = NumberOfLoadedImages + 1
Load Image1(NumberOfLoadedImages)

'Picture3.Visible = False
DoEvents
'Set FinalPicture.Picture = hDCToPicture(GetDC(Picture3.hwnd), (startX / Screen.TwipsPerPixelX), (startY / Screen.TwipsPerPixelY), (Picture3.Width / Screen.TwipsPerPixelX), (Picture3.Height / Screen.TwipsPerPixelY))
Image1(NumberOfLoadedImages).Width = Picture3.Width
Image1(NumberOfLoadedImages).Height = Picture3.Height
Image1(NumberOfLoadedImages).Top = Picture3.Top
Image1(NumberOfLoadedImages).Left = Picture3.Left
Image1(NumberOfLoadedImages).Picture = Picture3.Picture
Picture3.Visible = False
Image1(NumberOfLoadedImages).Visible = True
ActionTaken = "NONE"
Case Else
End Select
Me.MousePointer = 0
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case ActionTaken
Case "GLUER"
Picture3.Left = X + Picture3.Left
Picture3.Top = Y + Picture3.Top
startX = X
startY = Y
Case Else
End Select
End Sub

Private Sub ResizeImage_Click()
  'MCCaptureMouseCursorIntoNestedArea Me, FinalPicture, Image1(ImageIdentifier)

'Me.MousePointer = 15
'Me.MouseIcon = Image2.Picture
ActionTaken = "RESIZER" 'effect in mouse move - finalpicture
End Sub

Private Sub SaveEntireScreenShot_Click()
On Error GoTo ErrHandler
CommonDialog1.ShowSave
'without following line nothing right is saved !
DoEvents
Set Me.FinalPicture.Picture = hDCToPicture(GetDC(FinalPicture.hWnd), (0 / Screen.TwipsPerPixelX), (0 / Screen.TwipsPerPixelY), (Me.Width / Screen.TwipsPerPixelX), (Me.Height / Screen.TwipsPerPixelY))

SavePicture FinalPicture.Image, CommonDialog1.FileName
Exit Sub
ErrHandler:  'User pressed the Cancel button
End Sub

Private Sub SaveImage_Click()
On Error GoTo ErrHandler
CommonDialog1.ShowSave
SavePicture Image1(ImageIdentifier).Picture, CommonDialog1.FileName
Exit Sub
ErrHandler:  'User pressed the Cancel button
End Sub


Private Sub SysTrayPic_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Msg As Long
Msg = (X And &HFF) * &H100

Select Case Msg
       Case 3840 'click
       CutCounting = True
       Case Else
End Select
End Sub

Private Sub Timer2_Timer()
If CutCounting = True Then GoTo ExitThisSub

'count

SysTrayPic(11).Picture = SysTrayPic(Text1.Text - a).Picture
DoEvents
SystemTrayIcon SysTrayPic(11), Me, "MODIFY"
a = a + 1

If a > (Text1.Text) Then
ExitThisSub:
CutCounting = False
Timer2.Enabled = False
'do stuff
Me.Visible = False
Picture1.Visible = False
Form1.Show
Load Form1
Exit Sub
End If
End Sub

