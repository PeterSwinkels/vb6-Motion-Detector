Attribute VB_Name = "CoreModule"
'This module contains this program's main code.
Option Explicit

'Defines the Microsoft Windows API constants, functions and structures used by this program.
Private Type BITMAPINFOHEADER
   bmSize As Long
   bmWidth As Long
   bmHeight As Long
   bmPlanes As Integer
   bmBitCount As Integer
   bmCompression As Long
   bmSizeImage As Long
   bmXPelsPerMeter As Long
   bmYPelsPerMeter As Long
   bmClrUsed As Long
   bmClrImportant As Long
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

Public Type CAPSTATUS
   uiImageWidth As Long
   uiImageHeight As Long
   fLiveWindow As Long
   fOverlayWindow As Long
   fScale As Long
   ptScroll As POINTAPI
   fUsingDefaultPalette As Long
   fAudioHardware As Long
   fCapFileExists As Long
   dwCurrentVideoFrame As Long
   dwCurrentVideoFramesDropped As Long
   dwCurrentWaveSamples As Long
   dwCurrentTimeElapsedMS As Long
   hPalCurrent As Long
   fCapturingNow As Long
   dwReturn As Long
   wNumVideoAllocated As Long
   wNumAudioAllocated As Long
End Type
  
Private Type RGBTRIPLE
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
End Type
  
Private Type BITMAPINFO
   bmHeader As BITMAPINFOHEADER
   bmColors(0 To 255) As RGBTRIPLE
End Type

Public Const WM_CAP_DLG_VIDEOCOMPRESSION As Long = 1070&
Public Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065&
Public Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066&
Private Const BI_RGB As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0&
Private Const ERROR_FILE_NOT_FOUND As Long = 2&
Private Const ERROR_IO_PENDING As Long = 997&
Private Const ERROR_SUCCESS As Long = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const MAX_STRING As Long = 65535
Private Const WM_CAP_DRIVER_CONNECT As Long = 1034&
Private Const WM_CAP_DRIVER_DISCONNECT As Long = 1035&
Private Const WM_CAP_EDIT_COPY As Long = 1054&
Private Const WM_CAP_GET_STATUS As Long = 1078&
Private Const WM_CAP_GRAB_FRAME As Long = 1084&
Private Const WM_CLOSE As Long = 16&
Private Const WS_CHILD As Long = &H40000000

Public Declare Function SendMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function capCreateCaptureWindowA Lib "Avicap32.dll" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetDIBits Lib "Gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function IsWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetDIBits Lib "Gdi32.dll" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long


'Defines the constants, structures, and, variables used by this program.
Private Const NO_HANDLE As Long = 0   'Defines "no handle".

'This structure defines a RGB color.
Private Type RGBStr
   Red As Long     'Defines the red color component.
   Green As Long   'Defines the green color component.
   Blue As Long    'Defines the blue color component.
End Type

'This structure defines the settings.
Private Type SettingsStr
   ColorThreshold As Long      'Defines the difference threshold between a pixel's current and previous color.
   DisableWarning As Boolean   'Indicates whether motion warnings should be disabled.
   EMailAddress As String      'Defines the address to which warning e-mails are sent.
   MotionThreshold As Long     'Defines the motion threshold above which a warning is triggered.
End Type

Public Restart As Boolean        'Indicates whether the motion detection needs to be restarted.
Public Settings As SettingsStr   'Contains the settings.
'This procedure adjusts the specified picture box to the size of frames returned by the image capture device.
Public Sub AdjustSize(PictureBoxV As PictureBox)
On Error GoTo ErrorTrap
Dim Status As CAPSTATUS

   Status = GetCaptureStatus()
   PictureBoxV.Width = Status.uiImageWidth
   PictureBoxV.Height = Status.uiImageHeight
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure manages the capture window.
Public Function CaptureWindow(Optional StartCapture As Boolean = False, Optional ParentH As Long = NO_HANDLE, Optional StopCapture As Boolean = False) As Long
On Error GoTo ErrorTrap
Static CaptureWindowH As Long

   If StartCapture Then
      CaptureWindowH = CheckForError(capCreateCaptureWindowA(vbNullString, WS_CHILD, CLng(0), CLng(0), CLng(0), CLng(0), ParentH, CLng(0)), ERROR_FILE_NOT_FOUND)
      If Not CaptureWindowH = NO_HANDLE Then CheckForError SendMessageA(CaptureWindowH, WM_CAP_DRIVER_CONNECT, CLng(0), CLng(0)), ERROR_IO_PENDING
   ElseIf StopCapture Then
      CheckForError SendMessageA(CaptureWindowH, WM_CAP_DRIVER_DISCONNECT, CLng(0), CLng(0))
      CheckForError SendMessageA(CaptureWindowH, WM_CLOSE, CLng(0), CLng(0))
      CaptureWindowH = NO_HANDLE
   End If

   CaptureWindow = CaptureWindowH
   Exit Function
   
ErrorTrap:
   HandleError
End Function

'This procedure checks whether an error has occurred during the most recent Windows API call.
Public Function CheckForError(ReturnValue As Long, Optional Ignored As Long = ERROR_SUCCESS) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = Ignored) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description & vbCrLf
      Message = Message & "Return value: " & CStr(ReturnValue)
      MsgBox Message, vbExclamation
   End If
   
   CheckForError = ReturnValue
Exit Function

ErrorTrap:
   HandleError
End Function



'This procedure starts the motion detector.
Private Sub DetectMotion(CurrentViewBox As PictureBox, MotionViewBox As PictureBox, PreviousViewBox As PictureBox)
On Error GoTo ErrorTrap
Dim BitmapInformation As BITMAPINFO
Dim CurrentPixels() As RGBTRIPLE
Dim DifferencePixels() As RGBTRIPLE
Dim MotionLevel As Long
Dim PixelCount As Long
Dim PreviousPixels() As RGBTRIPLE

   AdjustSize CurrentViewBox

   With BitmapInformation.bmHeader
      .bmSize = Len(BitmapInformation.bmHeader)
      .bmWidth = CurrentViewBox.ScaleWidth
      .bmHeight = CurrentViewBox.ScaleHeight
      .bmBitCount = 24
      .bmClrImportant = 0
      .bmClrUsed = 0
      .bmCompression = BI_RGB
      .bmPlanes = 1
      .bmSizeImage = 0
      .bmXPelsPerMeter = 0
      .bmYPelsPerMeter = 0
   End With

   With CurrentViewBox
      PixelCount = .ScaleWidth * .ScaleHeight

      ReDim CurrentPixels(0 To PixelCount - 1) As RGBTRIPLE
      ReDim DifferencePixels(0 To PixelCount - 1) As RGBTRIPLE
      ReDim PreviousPixels(0 To PixelCount - 1) As RGBTRIPLE

      Do While DoEvents() > 0
         PreviousViewBox.Picture = .Image
         GrabFrame CurrentViewBox
         
         If Restart Then Exit Do
         CheckForError GetDIBits(PreviousViewBox.hDC, PreviousViewBox.Image, CLng(0), PreviousViewBox.ScaleHeight, PreviousPixels(0), BitmapInformation, DIB_RGB_COLORS)
         CheckForError GetDIBits(CurrentViewBox.hDC, CurrentViewBox.Image, CLng(0), CurrentViewBox.ScaleHeight, CurrentPixels(0), BitmapInformation, DIB_RGB_COLORS)

         MotionLevel = CLng((100 / PixelCount) * GetDifference(CurrentPixels(), PreviousPixels(), DifferencePixels()))

         If Restart Then Exit Do
         CheckForError SetDIBits(MotionViewBox.hDC, MotionViewBox.Image, CLng(0), MotionViewBox.ScaleHeight, DifferencePixels(0), BitmapInformation, DIB_RGB_COLORS)
         CurrentViewBox.Parent.Caption = App.Title & " - Motion level: " & CStr(MotionLevel) & " - Threshold: " & CStr(Settings.MotionThreshold) & " - Color Difference Threshold: " & CStr(Settings.ColorThreshold)
           
         If MotionLevel >= Settings.MotionThreshold Then
            If Not Settings.DisableWarning Then Warning MotionLevel, SaveSnapShot(CurrentViewBox), CurrentViewBox.Parent.DisableWarningMenu
            GrabFrame PreviousViewBox
            CurrentViewBox.Picture = PreviousViewBox.Image
         End If
      Loop
   End With
   
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure returns the image capture device's status.
Private Function GetCaptureStatus() As CAPSTATUS
On Error GoTo ErrorTrap
Dim Status As CAPSTATUS

   If IsWindow(CaptureWindow()) Then CheckForError SendMessageA(CaptureWindow(), WM_CAP_GET_STATUS, Len(Status), Status), ERROR_IO_PENDING
  
   GetCaptureStatus = Status
   Exit Function
   
ErrorTrap:
   HandleError
End Function

'This procedure determines the difference between two pixel sets and returns the result.
Private Function GetDifference(Pixels() As RGBTRIPLE, OtherPixels() As RGBTRIPLE, ByRef DifferencePixels() As RGBTRIPLE) As Long
On Error GoTo ErrorTrap
Dim Difference As RGBStr
Dim DifferenceCount As Long
Dim Index As Long

   DifferenceCount = 0
   For Index = LBound(DifferencePixels()) To UBound(DifferencePixels())
      Difference.Red = Abs(CLng(Pixels(Index).rgbRed) - CLng(OtherPixels(Index).rgbRed))
      Difference.Green = Abs(CLng(Pixels(Index).rgbGreen) - CLng(OtherPixels(Index).rgbGreen))
      Difference.Blue = Abs(CLng(Pixels(Index).rgbBlue) - CLng(OtherPixels(Index).rgbBlue))
      
      If (Difference.Red + Difference.Green + Difference.Blue) / 3 > Settings.ColorThreshold Then
         DifferencePixels(Index).rgbRed = &HFF
         DifferencePixels(Index).rgbGreen = &HFF
         DifferencePixels(Index).rgbBlue = &HFF
         
         DifferenceCount = DifferenceCount + 1
      Else
         DifferencePixels(Index).rgbRed = &H0
         DifferencePixels(Index).rgbGreen = &H0
         DifferencePixels(Index).rgbBlue = &H0
      End If
   Next Index
         
   GetDifference = DifferenceCount
   Exit Function
   
ErrorTrap:
   HandleError
End Function


'This procedure grabs a single frame from the image capture device.
Public Sub GrabFrame(Target As PictureBox)
On Error GoTo ErrorTrap
   
   SendMessageA CaptureWindow(), WM_CAP_GRAB_FRAME, CLng(0), CLng(0)
   SendMessageA CaptureWindow(), WM_CAP_EDIT_COPY, CLng(0), CLng(0)
   
   Target.Picture = Clipboard.GetData(vbCFBitmap)
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Message As String
   
   Message = Err.Description & vbCr & "Error code: " & Err.Number
   
   On Error Resume Next
   MsgBox Message, vbExclamation
   End
End Sub

'This procedure initializes this program.
Private Sub Main()
On Error GoTo ErrorTrap
   
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   With Settings
      .ColorThreshold = 12
      .DisableWarning = False
      .EMailAddress = vbNullString
      .MotionThreshold = 10
   End With
   
   Do
      Restart = False
      If Not CaptureWindow(StartCapture:=True, ParentH:=InterfaceWindow.hwnd) = NO_HANDLE Then
         InterfaceWindow.Show
         With InterfaceWindow
            DetectMotion .CurrentViewBox, .MotionViewBox, .PreviousViewBox
         End With
      End If
   Loop While Restart

   Exit Sub

ErrorTrap:
   HandleError
End Sub


'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With

   ProgramInformation = Information
   Exit Function

ErrorTrap:
   HandleError
End Function

'This procedure saves a snapshot.
Private Function SaveSnapShot(Source As PictureBox) As String
On Error GoTo ErrorTrap
Dim SnapShotFile As String

   SnapShotFile = CurDir$
   If Not Right$(SnapShotFile, 1) = "\" Then SnapShotFile = SnapShotFile & "\"
   SnapShotFile = SnapShotFile & "Snapshot.bmp"
   SavePicture Source.Picture, SnapShotFile

   SaveSnapShot = SnapShotFile
   Exit Function

ErrorTrap:
   HandleError
End Function

'This procedure sends a warning e-mail.
Private Sub SendEMail(Message As String, SnapShotFile As String)
Dim Messages As Object
Dim Session As Object

   Set Messages = CreateObject("msmapi.mapimessages")
   Set Session = CreateObject("msmapi.mapisession")
   
   Session.DownLoadMail = False
   Session.LogonUI = False
   Session.NewSession = True
   Session.SignOn
   
   With Messages
      .SessionID = Session.SessionID
      .Compose

      .AddressResolveUI = False
      .AttachmentPathName = SnapShotFile
      .MsgIndex = -1
      .MsgNoteText = Message
      .MsgSubject = "Motion Detector"
      .RecipAddress = Settings.EMailAddress
      .RecipDisplayName = Settings.EMailAddress
      .ResolveName
      .Send vDialog:=False
   End With
   
   Session.SignOff
   
   Set Messages = Nothing
   Set Session = Nothing

   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure either displays a warning or gives the command to send an e-mail.
Private Sub Warning(MotionLevel As Long, SnapShotFile As String, WarningMenu As Menu)
On Error GoTo ErrorTrap
Dim Message As String

   Message = "Motion (level " & CStr(MotionLevel) & ") detected." & vbCrLf & "Time: " & CStr(Now)
   
   If Settings.EMailAddress = vbNullString Then
      Beep
      Message = Message & vbCrLf & "Continue displaying warnings?"
      Settings.DisableWarning = (MsgBox(Message, vbExclamation Or vbYesNo) = vbNo)
      WarningMenu.Checked = Settings.DisableWarning
   Else
      SendEMail Message, SnapShotFile
   End If
   
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

