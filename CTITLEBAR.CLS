VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Filename: cTitlebar.cls
' Author:   Steve McMahon
' Date:     25 November 1998
'
' Mostly based on Ben Baird's code, see below.
' Changes:
'   -Change Subclass method to SSUBTMR
'   -Added WM_STYLECHANGED and WM_SETTEXT to allow titlebar to be
'   repainted when caption changes.
'   -Icon is now drawn transparent rather than with black background
'   -Added facility to change start/end and text colours (although if
'   the end colour isn't the title bar colour, it doesn't really work)
'   -Removed WM_SIZE subclass this prevented the form
'    from being resized and/or lost the Form_Resize message
'    and was also unnecessary
'   -Now determines what buttons are present when drawing gradient
'   -Check for tool window and appropriately smaller Tbar
'   -Added owner draw background support
'
' ==========================================================================

'//================================================
'+ Gradient titlebar example, version 1.1
'+ Author: Ben Baird
'+ Comments: Uploaded to Visual Basic Thunder
'            with Excalibur CodeLib on Monday,
'            September 1, 1997.
'            Thanks goes to Eric Dimayuga for
'            letting me bounce some ideas off
'            him. He also assisted me with the
'            theory of the code.
'
'+ If you release this code or any modified version
'  of it to the private, I would appreciate some
'  credit for the original code.
'+ FIXED IN THIS VERSION:
'   - Some trouble with the GradientGetCapsFont routine,
'   mainly a problem with the LOGFONT declaration.
'   - Added some more cleanup code.
'//================================================

Dim GradhWnd As Long, GradIcon As Long
Dim DrawDC As Long, tmpDC As Long
Dim hRgn As Long
Dim tmpGradFont As Long
Private m_bInSubClass As Boolean
Private m_bMDI As Boolean
Private m_pic As StdPicture
Private m_lHdc As Long
Private m_lHBmpOld As Long
Private m_lhPalOld As Long
Private m_lBitmapW As Long
Private m_lBitmapH As Long
Private m_sFileName As String
Private m_bCustomDraw As Boolean
Private m_bTextTransparent As Boolean
Implements ISubclass
Private m_emr As EMsgResponse
Private m_oColor(1 To 6) As OLE_COLOR
Public Enum EGradTitleBarColors
    eActiveStartColor = 1
    eActiveEndColor = 2
    eActiveText = 3
    einActivestartcolor = 4
    eInactiveEndColor = 5
    eInActiveText = 6
End Enum
Public Enum EGradTitleBarDrawStage
    eDrawBackground
    eDrawIcon
    eDrawText
End Enum
Public Event CustomDraw(ByVal eDrawStage As EGradTitleBarDrawStage, ByRef bDoDefault As Boolean, ByVal lHDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal bActive As Boolean)

Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type RECT
        left As Long
        tOp As Long
        Right As Long
        Bottom As Long
End Type
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 32
End Type

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
End Type
Dim CaptionFont As LOGFONT
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GCL_WNDPROC = (-24)
Private Const GCL_HICON = (-14)
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000 'WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_POPUP = &H80000000
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_EX_TOOLWINDOW = &H80&

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

Private Declare Function RectInRegion Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As DRAWTEXTPARAMS) As Long
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CMETRICS = 44
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CXBORDER = 5
Private Const SM_CXCURSOR = 13
Private Const SM_CXDLGFRAME = 7
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CXFRAME = 32
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CXHSCROLL = 21
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CXICONSPACING = 38
Private Const SM_CXMIN = 28
Private Const SM_CXMINTRACK = 34
'Private Const SM_CXSCREEN = 0
Private Const SM_CXSMSIZE = 30
Private Const SM_CXSIZEFRAME = SM_CXFRAME
Private Const SM_CXVSCROLL = 2
Private Const SM_CYBORDER = 6
'Private Const SM_CYCAPTION = 4
Private Const SM_CYCURSOR = 14
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Private Const SM_CYFRAME = 33
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYHSCROLL = 3
Private Const SM_CYICON = 12
Private Const SM_CYICONSPACING = 39
Private Const SM_CYKANJIWINDOW = 18
'Private Const SM_CYMENU = 15
Private Const SM_CYMIN = 29
Private Const SM_CYMINTRACK = 35
'Private Const SM_CYSCREEN = 1
Private Const SM_CYSMSIZE = 31
Private Const SM_CYSIZEFRAME = SM_CYFRAME
Private Const SM_CYVSCROLL = 20
Private Const SM_CYVTHUMB = 9
Private Const SM_DBCSENABLED = 42
Private Const SM_DEBUG = 22
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_MOUSEPRESENT = 19
Private Const SM_PENWINDOWS = 41
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_SWAPBUTTON = 23
Private Const SM_CYSMCAPTION = 51


Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const DFC_CAPTION = 1
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_INACTIVE = &H100

Private Const WM_ACTIVATE = &H6
Private Const WM_SIZE = &H5
Private Const WM_GETFONT = &H31
Private Const WM_STYLECHANGED = &H7D
Private Const WM_GETTEXT = &HD
Private Const WM_SETTEXT = &HC

'  WM_ACTIVATE state values
Private Const WA_INACTIVE = 0
Private Const WA_ACTIVE = 1
Private Const WA_CLICKACTIVE = 2

Private Const WM_NCHITTEST = &H84
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86

Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIGETACTIVE = &H229
Public Property Let TextTransparent(ByVal bState As Boolean)
    m_bTextTransparent = bState
End Property
Public Property Get TextTransparent() As Boolean
    TextTransparent = m_bTextTransparent
End Property
Public Property Let CustomDraw(ByVal bState As Boolean)
    m_bCustomDraw = bState
End Property
Public Property Get CustomDraw() As Boolean
    CustomDraw = m_bCustomDraw
End Property

Public Property Get Color(ByVal eType As EGradTitleBarColors) As OLE_COLOR
    Color = m_oColor(eType)
End Property
Public Property Let Color(ByVal eType As EGradTitleBarColors, oColor As OLE_COLOR)
    If (m_oColor(eType) <> oColor) Then
        m_oColor(eType) = oColor
        If (GradhWnd <> 0) Then
            SendMessage GradhWnd, WM_NCPAINT, 0, 0
        End If
    End If
End Property

Public Property Get Picture() As StdPicture
   Set Picture = m_pic
End Property
Public Property Let Picture(oPic As StdPicture)
    ' Load a picture from a StdPicture object:
    pClearUp
    If (pbEnsurePicture()) Then
        Set m_pic = oPic
        If (Err.Number = 0) Then
            pbGetBitmapIntoDC
        End If
    End If
End Property
Private Function pbEnsurePicture() As Boolean
On Error Resume Next
    pbEnsurePicture = True
    If (m_pic Is Nothing) Then
        Set m_pic = New StdPicture
        If (Err.Number <> 0) Then
            pErr 3, "Unable to allocate memory for picture object."
            pbEnsurePicture = False
        Else
        End If
    End If
On Error GoTo 0
    Exit Function
End Function
Private Function pbGetBitmapIntoDC() As Boolean
Dim tB As BITMAP
Dim lHDC As Long, lHwnd As Long

    ' Make a DC to hold the picture bitmap which we can blt from:
    lHwnd = GetDesktopWindow()
    lHDC = GetDC(lHwnd)
    m_lHdc = CreateCompatibleDC(lHDC)
    ReleaseDC lHwnd, lHDC
    If (m_lHdc <> 0) Then
        ' Get size of bitmap:
        GetObjectAPI m_pic.handle, LenB(tB), tB
        m_lBitmapW = tB.bmWidth
        m_lBitmapH = tB.bmHeight
        
        ' Select bitmap into DC:
        m_lHBmpOld = SelectObject(m_lHdc, m_pic.handle)
        If (m_lHBmpOld <> 0) Then
            ' Select the palette into the DC:
            m_lhPalOld = SelectObject(m_lHdc, m_pic.hPal)
            pbGetBitmapIntoDC = True
        Else
            pClearUp
            pErr 2, "Unable to select bitmap into DC"
        End If
    Else
        pErr 1, "Unable to create compatible DC"
    End If
End Function
Private Sub pClearUp()
    ' Clear reference to the filename:
    m_sFileName = ""
    ' If we have a DC, then clear up:
    If (m_lHdc <> 0) Then
        ' Select the bitmap out of DC:
        If (m_lHBmpOld <> 0) Then
            SelectObject m_lHdc, m_lHBmpOld
            ' The original bitmap does not have to deleted because it is owned by m_pic
        End If
        ' Select the palette out of the DC:
        If (m_lhPalOld <> 0) Then
            SelectObject m_lHdc, m_lhPalOld
            ' The original palette does not have to deleted because it is owned by m_pic
        End If
        ' Remove the DC:
        DeleteObject m_lHdc
    End If
End Sub
Private Sub pErr(lNumber As Long, sMsg As String)
    Err.Raise vbObjectError + 1048 + lNumber, App.EXEName & ".cTitleBar", sMsg & ", [" & lNumber & "]"
End Sub

Private Function GradientCallback(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim OldBMP As Long, NewBMP As Long
Dim rcWnd As RECT

Select Case wMsg
Case WM_NCACTIVATE, WM_MDIACTIVATE

    'Debug.Print "ACTIVATE"
   
  GetWindowRect GradhWnd, rcWnd

  'Create memory DC to draw the titlebar in.
  tmpDC = GetWindowDC(GradhWnd)
  DrawDC = CreateCompatibleDC(tmpDC)
  NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.left, 50)
  OldBMP = SelectObject(DrawDC, NewBMP)
  With rcWnd
   hRgn = CreateRectRgn(.left, .tOp, .Right, .Bottom)
   SelectClipRgn tmpDC, hRgn
   OffsetClipRgn tmpDC, -.left, -.tOp
  End With
  'Find out what color the titlebar needs
  'to be...
  If wParam And GetParent(GradhWnd) = 0 Then
   DrawGradient TranslateColor(Color(eActiveStartColor)), TranslateColor(Color(eActiveEndColor)), True
  ElseIf wParam = GradhWnd And GetParent(GradhWnd) <> 0 Then
   DrawGradient TranslateColor(Color(einActivestartcolor)), TranslateColor(Color(eInactiveEndColor)), False
  ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, 0, 0) = GradhWnd Then
   DrawGradient TranslateColor(Color(eActiveStartColor)), TranslateColor(Color(eActiveEndColor)), True
  Else
   DrawGradient TranslateColor(Color(einActivestartcolor)), TranslateColor(Color(eInactiveEndColor)), False
  End If
  'Cleanup
  SelectObject DrawDC, OldBMP
  DeleteObject NewBMP
  DeleteDC DrawDC
  OffsetClipRgn tmpDC, rcWnd.left, rcWnd.tOp
  GetClipRgn tmpDC, hRgn
  ReleaseDC GradhWnd, tmpDC
  DeleteObject hRgn
  tmpDC = 0
  GradientCallback = 1
  Exit Function

Case WM_NCPAINT
  
  'Debug.Print "NCPAINT"
  
  'Basically same as above.
  GetWindowRect GradhWnd, rcWnd
  tmpDC = GetWindowDC(GradhWnd)
  DrawDC = CreateCompatibleDC(tmpDC)
  NewBMP = CreateCompatibleBitmap(tmpDC, rcWnd.Right - rcWnd.left, 50)
  OldBMP = SelectObject(DrawDC, NewBMP)
  With rcWnd
   hRgn = CreateRectRgn(.left, .tOp, .Right, .Bottom)
   SelectClipRgn tmpDC, hRgn
   OffsetClipRgn tmpDC, -.left, -.tOp
  End With
  'Get the color to paint the caption with.
  If GetActiveWindow() = GradhWnd Then
   DrawGradient TranslateColor(Color(eActiveStartColor)), TranslateColor(Color(eActiveEndColor)), True
  ElseIf SendMessage(GetParent(GradhWnd), WM_MDIGETACTIVE, 0, 0) = GradhWnd Then
   DrawGradient TranslateColor(Color(eActiveStartColor)), TranslateColor(Color(eActiveEndColor)), True
  Else
   DrawGradient TranslateColor(Color(einActivestartcolor)), TranslateColor(Color(eInactiveEndColor)), False
  End If
  'Cleanup
  SelectObject DrawDC, OldBMP
  DeleteObject NewBMP
  DeleteDC DrawDC
  OffsetClipRgn tmpDC, rcWnd.left, rcWnd.tOp
  GetClipRgn tmpDC, hRgn
  
  'Call the old proc. This will only
  'draw the titlebar's min/max/close buttons
  'because we told it not to do the rest (this
  'eliminates flicker.)
  GradientCallback = CallOldWindowProc(hwnd, WM_NCPAINT, hRgn, lParam)
  ReleaseDC GradhWnd, tmpDC
  DeleteObject hRgn
  tmpDC = 0
  Exit Function
Case WM_SIZE, WM_STYLECHANGED, WM_SETTEXT
  'Whoa, we need to paint the caption.
  If hwnd = GradhWnd Then SendMessage GradhWnd, WM_NCPAINT, 0, 0
End Select

End Function

Public Sub RefreshTitleBar()
If GradhWnd <> 0 Then SendMessage GradhWnd, WM_NCPAINT, 0, 0
End Sub

Public Sub GradientForm(frm As Object)

   If (m_bInSubClass) Then
      GradientReleaseForm
   End If
   
   GradhWnd = frm.hwnd
   GradIcon = frm.Icon
   
   If (TypeOf frm Is MDIForm) Then
      m_bMDI = True
   End If
   AttachMessage Me, GradhWnd, WM_NCPAINT
   AttachMessage Me, GradhWnd, WM_NCACTIVATE
   AttachMessage Me, GradhWnd, WM_MDIACTIVATE
   ' Check for change in form caption:
   AttachMessage Me, GradhWnd, WM_STYLECHANGED
   AttachMessage Me, GradhWnd, WM_SETTEXT
   
   m_bInSubClass = True
   
   GradientGetCapsFont

End Sub

Public Sub GradientReleaseForm()

   If (m_bInSubClass) Then
      DetachMessage Me, GradhWnd, WM_NCPAINT
      DetachMessage Me, GradhWnd, WM_NCACTIVATE
      DetachMessage Me, GradhWnd, WM_MDIACTIVATE
      DetachMessage Me, GradhWnd, WM_STYLECHANGED
      DetachMessage Me, GradhWnd, WM_SETTEXT
      GradhWnd = 0
   End If

End Sub

Private Function DrawGradient( _
      ByVal Color1 As Long, _
      ByVal Color2 As Long, _
      ByVal bActive As Boolean _
   ) As Long

Dim i As Integer
Dim DestWidth As Long, DestHeight As Long
Dim StartPnt As Integer, EndPnt As Integer
Dim PixelStep As Long, XBorder As Long
Dim WndRect As RECT
Dim OldFont As Long
Dim fText As String
Dim iBmTop As Long
Dim lS As Long
Dim lBtnWidth As Long
Dim bDoDefault As Boolean
Dim bShowBitmap As Boolean

On Error Resume Next

   GetWindowRect GradhWnd, WndRect
   With WndRect
      DestWidth = .Right - .left
   End With
   
   bShowBitmap = ((m_lHdc <> 0) And (DestWidth > (m_lBitmapW * 4)))
      
   'Get height of caption bar
   lS = GetWindowLong(GradhWnd, GWL_EXSTYLE)
   If ((lS And WS_EX_TOOLWINDOW) = WS_EX_TOOLWINDOW) Then
       DestHeight = GetSystemMetrics(SM_CYSMCAPTION)
   Else
       DestHeight = GetSystemMetrics(SM_CYCAPTION)
   End If
   'Get the text of the form's caption
   fText = Space$(255)
   GetWindowText GradhWnd, fText, 255
   fText = Trim$(fText)
   'Get the width of the border
   XBorder = GetSystemMetrics(SM_CXFRAME)
   'The width of the area we need to paint:
   DestWidth = DestWidth - (XBorder * 2) + 6
   lS = GetWindowLong(GradhWnd, GWL_STYLE)
   If ((lS And WS_MINIMIZEBOX) = WS_MINIMIZEBOX) Or ((lS And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX) Then
       lBtnWidth = lBtnWidth + (GetSystemMetrics(SM_CXSMSIZE)) * 2
   End If
   If ((lS And WS_SYSMENU) = WS_SYSMENU) Then
       lBtnWidth = lBtnWidth + (GetSystemMetrics(SM_CXSMSIZE))
   End If
   If (lBtnWidth <> 0) Then lBtnWidth = lBtnWidth + 1
   DestWidth = DestWidth - lBtnWidth
   
   
   'Where the painting begins:
   StartPnt = XBorder
   'Where the painting ends:
   EndPnt = XBorder + DestWidth - 4
   
   'How many steps do we need to
   'paint the titlebar?
   If (bShowBitmap) Then
       PixelStep = DestWidth - 16 - m_lBitmapW
   Else
       PixelStep = DestWidth \ 8
   End If
   ReDim Colors(PixelStep) As Long
   'Create gradient colors in the array
   GradateColors Colors(), Color1, Color2
   
   Dim rct As RECT
   Dim hBr As Long
   
   With rct
      .tOp = XBorder
      .left = XBorder
      .Right = XBorder + (DestWidth \ PixelStep)
      .Bottom = XBorder + DestHeight - 1
      bDoDefault = True
      If (m_bCustomDraw) Then
          RaiseEvent CustomDraw(eDrawBackground, bDoDefault, DrawDC, XBorder, XBorder, EndPnt - XBorder, DestHeight - 1, bActive)
      End If
      If (bDoDefault) Then
          If (bShowBitmap) Then
              ' the first box is going to extend all the way from
              ' the icon to the end of the bitmap, and the gradient
              ' is only doing to start afterwards:
              rct.Right = rct.left + 16 + m_lBitmapW
          End If
          hBr = CreateSolidBrush(Colors(0))
          FillRect DrawDC, rct, hBr
          DeleteObject hBr
          If (bShowBitmap) Then
              rct.left = rct.Right
              rct.Right = rct.left + (DestWidth \ PixelStep)
          End If
          For i = 1 To PixelStep - 1
           'Paint the titlebar in increments, increasing
           'the color index with each iteration.
           hBr = CreateSolidBrush(Colors(i))
           FillRect DrawDC, rct, hBr
           'Cleanup
           DeleteObject hBr
           'Prepare for the next iteration
           OffsetRect rct, (DestWidth \ PixelStep), 0
           If i = PixelStep - 2 Then .Right = EndPnt
          Next
      End If
      
      bDoDefault = True
      If (m_bCustomDraw) Then
          RaiseEvent CustomDraw(eDrawIcon, bDoDefault, DrawDC, XBorder, XBorder, EndPnt - XBorder, DestHeight - 1, bActive)
      End If
      If (bDoDefault) Then
          If GradIcon <> 0 Then 'Paint the icon
           'Move the caption text's start point over
           'to make room for the icon
           .left = XBorder + GetSystemMetrics(SM_CXSMSIZE) + 2
           DrawIconEx DrawDC, XBorder + 1, XBorder + 1, GradIcon, GetSystemMetrics(SM_CXSMSIZE) - 2, GetSystemMetrics(SM_CYSMSIZE) - 2, ByVal 0&, ByVal 0&, DI_NORMAL
          Else
           'No icon
           .left = XBorder
          End If
      End If
   
      'If we have a picture:
      If (bShowBitmap) Then
         ' Draw it and shift left & right
         iBmTop = 2 + .tOp + ((.Bottom - .tOp - m_lBitmapH) \ 2)
         BitBlt DrawDC, .left, iBmTop, m_lBitmapW, m_lBitmapH, m_lHdc, 0, 0, SRCCOPY
         .left = .left + m_lBitmapW + 2
      Else
         .left = .left + 2
      End If
   
      bDoDefault = True
      If (m_bCustomDraw) Then
         RaiseEvent CustomDraw(eDrawText, bDoDefault, DrawDC, XBorder, XBorder, EndPnt - XBorder - 10, DestHeight - 1, bActive)
      End If
      If (bDoDefault) Then
         'If getting the caption font failed, use the font
         'from the gradient caption form.
         If CaptionFont.lfHeight = 0 And tmpGradFont = 0 Then
            tmpGradFont = SendMessage(GradhWnd, WM_GETFONT, 0, 0)
         ElseIf tmpGradFont = 0 Then
            tmpGradFont = CreateFontIndirect(CaptionFont)
         End If
         OldFont = SelectObject(DrawDC, tmpGradFont)
         'This is like setting FontTransparent on a Form to True:
         SetBkMode DrawDC, 1
         'Use a white caption, since the background is black
         'on the left side
         If (bActive) Then
            SetTextColor DrawDC, TranslateColor(m_oColor(eActiveText)) ' &HFFFFFF 'RGB(255, 255, 255)
         Else
            SetTextColor DrawDC, TranslateColor(m_oColor(eInActiveText)) '&HC0C0C0 'RGB(128, 128, 128)
         End If
         .Right = .Right - 10
         'Draw the caption text
         If (m_bTextTransparent) Then
            SetBkMode DrawDC, TRANSPARENT
         Else
            SetBkMode DrawDC, OPAQUE
            If (bActive) Then
                SetBkColor DrawDC, TranslateColor(m_oColor(eActiveStartColor))
            Else
                SetBkColor DrawDC, TranslateColor(m_oColor(einActivestartcolor))
            End If
         End If
         DrawText DrawDC, fText, Len(fText) - 1, rct, DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_VCENTER
         'Cleanup
         SelectObject DrawDC, OldFont
         DeleteObject tmpGradFont
         tmpGradFont = 0
      End If
   
      .left = XBorder
      .Right = .Right + 12
      If tmpDC <> 0 Then
         'Blit our work from the memory DC to the form's
         'window DC to finish the job.
         BitBlt tmpDC, .left, .tOp, .Right - .left - 10, .Bottom - .tOp, DrawDC, .left, .tOp, vbSrcCopy
         'Tell windows that we already painted most of
         'the titlebar.
         ExcludeClipRect tmpDC, XBorder, XBorder, .Right - .left - 7, .Bottom - .tOp + 4
   
      End If
   
   End With

End Function

Private Sub GradateColors(Colors() As Long, ByVal Color1 As Long, ByVal Color2 As Long)

'Alright, I admit -- this routine was
'taken from a VBPJ issue a few months back.

Dim i As Integer
Dim dblR As Double, dblG As Double, dblB As Double
Dim addR As Double, addG As Double, addB As Double
Dim bckR As Double, bckG As Double, bckB As Double

   dblR = CDbl(Color1 And &HFF)
   dblG = CDbl(Color1 And &HFF00&) / 255
   dblB = CDbl(Color1 And &HFF0000) / &HFF00&
   bckR = CDbl(Color2 And &HFF&)
   bckG = CDbl(Color2 And &HFF00&) / 255
   bckB = CDbl(Color2 And &HFF0000) / &HFF00&
   
   addR = (bckR - dblR) / UBound(Colors)
   addG = (bckG - dblG) / UBound(Colors)
   addB = (bckB - dblB) / UBound(Colors)
   
   For i = 0 To UBound(Colors)
      dblR = dblR + addR
      dblG = dblG + addG
      dblB = dblB + addB
      If dblR > 255 Then dblR = 255
      If dblG > 255 Then dblG = 255
      If dblB > 255 Then dblB = 255
      If dblR < 0 Then dblR = 0
      If dblG < 0 Then dblG = 0
      If dblG < 0 Then dblB = 0
      Colors(i) = RGB(dblR, dblG, dblB)
   Next

End Sub

Private Sub GradientGetCapsFont()

'Tries to retrieve the Windows caption font
'in the current Appearance scheme. Doesn't
'seem to work all the time, so if anyone knows
'why I'd appreciate being told.

Dim NCM As NONCLIENTMETRICS
Dim lfNew As LOGFONT

   NCM.cbSize = Len(NCM)
   Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
   
   If NCM.iCaptionHeight = 0 Then
      CaptionFont.lfHeight = 0
   Else
      CaptionFont = NCM.lfCaptionFont
   End If

End Sub
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1 'CLR_INVALID
    End If
End Function

Private Sub Class_Initialize()
    m_bCustomDraw = True
    m_bTextTransparent = True
    m_oColor(eActiveStartColor) = 0
    m_oColor(einActivestartcolor) = 0
    m_oColor(eActiveEndColor) = vbActiveTitleBar
    m_oColor(eInactiveEndColor) = vbInactiveTitleBar
    m_oColor(eActiveText) = &HFFFFFF
    m_oColor(eInActiveText) = &HC0C0C0
End Sub

Private Sub Class_Terminate()
   GradientReleaseForm
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   m_emr = RHS
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   'Debug.Print "Get MsgResponse", CurrentMessage
Select Case CurrentMessage
Case WM_NCACTIVATE, WM_MDIACTIVATE
  'This will cause a slight flicker because we
  'let Windows paint the caption before we do.
  'We don't let this happen in the WM_NCPAINT message,
  'which is called more often than NCACTIVATE.
   m_emr = emrPreprocess
   
Case WM_NCPAINT
  'Call the old proc. This will only
  'draw the titlebar's min/max/close buttons
  'because we told it not to do the rest (this
  'eliminates flicker.
  m_emr = emrConsume

Case Else
   m_emr = emrPostProcess
End Select

ISubclass_MsgResponse = m_emr
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   'Debug.Print "Get Msg" & vbCrLf
   ISubclass_WindowProc = GradientCallback(hwnd, iMsg, wParam, lParam)
End Function
