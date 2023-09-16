Attribute VB_Name = "MotionDetectorModule"
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
Private Const NOHANDLE As Long = 0 'Defines "no handle".

'This structure defines a RGB color.
Private Type RGBStr
   Red As Long     'Defines the red color component.
   Green As Long   'Defines the green color component.
   Blue As Long    'Defines the blue color component.
End Type

Public ColorThreshold As Long          'Contains the difference threshold between a pixel's current and previous color.
Public DisableWarning As Boolean       'Indicates whether motion warnings should be disabled.
Public EMailAddress As String          'Contains the address to which warning e-mails are sent.
Public MotionThreshold As Long         'Contains the motion threshold above which a warning is triggered.
Public Restart As Boolean              'Indicates whether the motion detection needs to be restarted.

'This procedure adjusts the specified picture box to the size of frames returned by the image capture device.
Public Sub AdjustSize(PictureBoxV As PictureBox)
Dim Status As CAPSTATUS

   Status = GetCaptureStatus()
   PictureBoxV.Width = Status.uiImageWidth
   PictureBoxV.Height = Status.uiImageHeight
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub


'This procedure manages the capture window.
Public Function CaptureWindow(Optional StartCapture As Boolean = False, Optional StopCapture As Boolean = False) As Long
On Error GoTo ErrorTrap
Static CaptureWindowH As Long

   If StartCapture Then
      CaptureWindowH = CheckForError(capCreateCaptureWindowA(vbNullString, WS_CHILD, CLng(0), CLng(0), CLng(0), CLng(0), InterfaceWindow.hwnd, CLng(0)), ERROR_FILE_NOT_FOUND)
      If Not CaptureWindowH = NOHANDLE Then CheckForError SendMessageA(CaptureWindowH, WM_CAP_DRIVER_CONNECT, CLng(0), CLng(0)), ERROR_IO_PENDING
   ElseIf StopCapture Then
      CheckForError SendMessageA(CaptureWindowH, WM_CAP_DRIVER_DISCONNECT, CLng(0), CLng(0))
      CheckForError SendMessageA(CaptureWindowH, WM_CLOSE, CLng(0), CLng(0))
      CaptureWindowH = NOHANDLE
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
Private Sub DetectMotion()
On Error GoTo ErrorTrap
Dim BitmapInformation As BITMAPINFO
Dim CurrentColors() As RGBTRIPLE
Dim Difference As RGBStr
Dim DifferenceColors() As RGBTRIPLE
Dim DifferenceCount As Long
Dim Index As Long
Dim MotionLevel As Long
Dim PixelCount As Long
Dim PreviousColors() As RGBTRIPLE

   AdjustSize InterfaceWindow.CurrentViewBox

   With BitmapInformation.bmHeader
      .bmSize = Len(BitmapInformation.bmHeader)
      .bmWidth = InterfaceWindow.CurrentViewBox.ScaleWidth
      .bmHeight = InterfaceWindow.CurrentViewBox.ScaleHeight
      .bmBitCount = 24
      .bmClrImportant = 0
      .bmClrUsed = 0
      .bmCompression = BI_RGB
      .bmPlanes = 1
      .bmSizeImage = 0
      .bmXPelsPerMeter = 0
      .bmYPelsPerMeter = 0
   End With

   With InterfaceWindow
      PixelCount = .CurrentViewBox.ScaleWidth * .CurrentViewBox.ScaleHeight

      ReDim CurrentColors(0 To PixelCount) As RGBTRIPLE
      ReDim PreviousColors(0 To PixelCount) As RGBTRIPLE

      Do While DoEvents() > 0
         .PreviousViewBox.Picture = .CurrentViewBox.Image
         GrabFrame .CurrentViewBox
         
         If Restart Then Exit Do
         CheckForError GetDIBits(.PreviousViewBox.hDC, .PreviousViewBox.Image, CLng(0), .PreviousViewBox.ScaleHeight, PreviousColors(0), BitmapInformation, DIB_RGB_COLORS)
         CheckForError GetDIBits(.CurrentViewBox.hDC, .CurrentViewBox.Image, CLng(0), .CurrentViewBox.ScaleHeight, CurrentColors(0), BitmapInformation, DIB_RGB_COLORS)

         ReDim DifferenceColors(0 To PixelCount) As RGBTRIPLE
         DifferenceCount = 0
         For Index = LBound(DifferenceColors()) To UBound(DifferenceColors())
            Difference.Red = Abs(CLng(CurrentColors(Index).rgbRed) - CLng(PreviousColors(Index).rgbRed))
            Difference.Green = Abs(CLng(CurrentColors(Index).rgbGreen) - CLng(PreviousColors(Index).rgbGreen))
            Difference.Blue = Abs(CLng(CurrentColors(Index).rgbBlue) - CLng(PreviousColors(Index).rgbBlue))
            
            If (Difference.Red + Difference.Green + Difference.Blue) / 3 > ColorThreshold Then
               DifferenceColors(Index).rgbRed = 255
               DifferenceColors(Index).rgbGreen = 255
               DifferenceColors(Index).rgbBlue = 255
               DifferenceCount = DifferenceCount + 1
            End If
         Next Index
         MotionLevel = CLng((100 / PixelCount) * DifferenceCount)
   
         If Restart Then Exit Do
         CheckForError SetDIBits(.MotionViewBox.hDC, .MotionViewBox.Image, CLng(0), .MotionViewBox.ScaleHeight, DifferenceColors(0), BitmapInformation, DIB_RGB_COLORS)
         .Caption = App.Title & " - Motion level: " & CStr(MotionLevel) & " - Threshold: " & CStr(MotionThreshold) & " - Color Difference Threshold: " & CStr(ColorThreshold)
           
         If MotionLevel >= MotionThreshold Then
            If Not DisableWarning Then SendWarning MotionLevel, SaveSnapShot(InterfaceWindow.CurrentViewBox), InterfaceWindow.DisableWarningMenu
            GrabFrame .PreviousViewBox
            .CurrentViewBox.Picture = .PreviousViewBox.Image
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
   
   ColorThreshold = 12
   DisableWarning = False
   MotionThreshold = 10
   EMailAddress = vbNullString

   Do
      Restart = False
      If Not CaptureWindow(StartCapture:=True) = NOHANDLE Then
         InterfaceWindow.Show
         DetectMotion
      End If
   Loop While Restart

   Exit Sub
ErrorTrap:
   HandleError
End Sub

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
Private Sub SendWarning(MotionLevel As Long, SnapShotFile As String, WarningMenu As Menu)
On Error GoTo ErrorTrap
Dim Message As String
Dim Messages As Object
Dim Session As Object

   Message = "Motion (level " & CStr(MotionLevel) & ") detected." & vbCrLf & "Time: " & CStr(Now)
   
   If EMailAddress = vbNullString Then
      Beep
      Message = Message & vbCrLf & "Continue displaying warnings?"
      DisableWarning = (MsgBox(Message, vbExclamation Or vbYesNo) = vbNo)
      WarningMenu.Checked = DisableWarning
   Else
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
         .RecipAddress = EMailAddress
         .RecipDisplayName = EMailAddress
         .ResolveName
         .Send vDialog:=False
      End With
      
      Session.SignOff
      
      Set Messages = Nothing
      Set Session = Nothing
   End If
   Exit Sub
   
ErrorTrap:
   HandleError
End Sub

