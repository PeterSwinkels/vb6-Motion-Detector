VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   Icon            =   "Interface.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CurrentViewBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox MotionViewBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox PreviousViewBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramMainMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu ViewMainMenu 
      Caption         =   "&View"
      Begin VB.Menu ViewMenu 
         Caption         =   "&Current View"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu ViewMenu 
         Caption         =   "&Motion Viewer"
         Index           =   1
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu OptionsMainMenu 
      Caption         =   "&Options"
      Begin VB.Menu DisableWarningMenu 
         Caption         =   "Disable &Warning"
         Shortcut        =   ^W
      End
      Begin VB.Menu MotionDetectionMenu 
         Caption         =   "Motion &Detection"
         Shortcut        =   ^D
      End
      Begin VB.Menu OptionsMainMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu VideoCompressionMenu 
         Caption         =   "Video &Compression"
         Shortcut        =   ^V
      End
      Begin VB.Menu VideoFormatMenu 
         Caption         =   "Video &Format"
         Shortcut        =   ^F
      End
      Begin VB.Menu VideoSourceMenu 
         Caption         =   "Video &Source"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main window.
Option Explicit

'This enumeration lists the available views.
Private Enum ViewsE
   CurrentView      'Current view.
   MotionView       'Motion viewer.
End Enum



'This procedure selects the specified view.
Private Sub SelectView(NewView As ViewsE)
On Error GoTo ErrorTrap
Static SelectedView As Long
   
   ViewMenu(SelectedView).Checked = False
   ViewMenu(NewView).Checked = True
   
   If NewView = CurrentView Then
      CurrentViewBox.Visible = True
      MotionViewBox.Visible = False
   ElseIf NewView = MotionView Then
      CurrentViewBox.Visible = False
      MotionViewBox.Visible = True
   End If
   
   SelectedView = NewView
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure adjusts this window to the new size of the current view picture box.
Private Sub CurrentViewBox_Resize()
On Error GoTo ErrorTrap
   With CurrentViewBox
      Me.Width = .Width * Screen.TwipsPerPixelX
      Me.Height = .Height * Screen.TwipsPerPixelY
      Me.Width = Me.Width + ((.ScaleWidth - Me.ScaleWidth) * Screen.TwipsPerPixelX)
      Me.Height = Me.Height + ((.ScaleHeight - Me.ScaleHeight) * Screen.TwipsPerPixelY)
      MotionViewBox.Width = .Width
      MotionViewBox.Height = .Height
      PreviousViewBox.Width = .Width
      PreviousViewBox.Height = .Height
   End With
   
   Me.Left = (Screen.Width / 2) - (Me.Width / 2)
   Me.Top = (Screen.Height / 2) - (Me.Height / 2)
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure disabled/enables motion warnings.
Private Sub DisableWarningMenu_Click()
On Error GoTo ErrorTrap
   DisableWarning = Not DisableWarning
   DisableWarningMenu.Checked = DisableWarning
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   GrabFrame CurrentViewBox
   
   SelectView MotionView
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure gives the command to stop the video capture device when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   CaptureWindow , StopCapture:=True
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   With App
      MsgBox .Comments, vbInformation, .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision)
   End With
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure requests the user to specify the motion detection related settings.
Private Sub MotionDetectionMenu_Click()
On Error GoTo ErrorTrap
Dim NewColorThreshold As Long
Dim NewMotionThreshold As Long

   NewColorThreshold = CLng(Val(InputBox$("Color difference threshold (1-255)", , CStr(ColorThreshold))))
   NewMotionThreshold = CLng(Val(InputBox$("Warning motion threshold (1-100)", , CStr(MotionThreshold))))
   EMailAddress = InputBox$("Send warning e-mails to (if none is specified, a message is displayed):", , EMailAddress)
   
   If NewColorThreshold >= 1 And NewColorThreshold <= 255 Then ColorThreshold = NewColorThreshold
   If NewMotionThreshold >= 1 And NewMotionThreshold <= 100 Then MotionThreshold = NewMotionThreshold
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure closes this window.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure opens the video compression dialog window.
Private Sub VideoCompressionMenu_Click()
On Error GoTo ErrorTrap
   CheckForError SendMessageA(CaptureWindow(), WM_CAP_DLG_VIDEOCOMPRESSION, CLng(0), CLng(0))
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure opens the video format dialog window.
Private Sub VideoFormatMenu_Click()
On Error GoTo ErrorTrap
   CheckForError SendMessageA(CaptureWindow(), WM_CAP_DLG_VIDEOFORMAT, CLng(0), CLng(0))
   CaptureWindow , StopCapture:=True
   AdjustSize CurrentViewBox
   Restart = True
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure opens the video source dialog window.
Private Sub VideoSourceMenu_Click()
On Error GoTo ErrorTrap
   CheckForError SendMessageA(CaptureWindow(), WM_CAP_DLG_VIDEOSOURCE, CLng(0), CLng(0))
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure gives the command change the view to the user's selection.
Private Sub ViewMenu_Click(Index As Integer)
On Error GoTo ErrorTrap
   SelectView CLng(Index)
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


