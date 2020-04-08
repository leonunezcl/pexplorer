Attribute VB_Name = "mDeclares"
Option Explicit

' ======================================================================
' cPOPMENU TYPE:
' Menu information:
Public Type tMenuItem
   sHelptext As String
   sCaption As String
   sAccelerator As String
   sShortCutDisplay As String
   sKey As String
   sTag As String ' New
   iShortCutShiftMask As Integer
   iShortCutShiftKey As Integer
   lID As Long
   lActualID As Long       ' The ID gets modified if we add a sub-menu to the hMenu of the popup
   lItemData As Long
   lIndex As Long
   lParentId As Long
   lIconIndex As Long
   bChecked As Boolean
   bEnabled As Boolean
   hMenu As Long
   lWidth As Long
   bCreated As Boolean
   bIsAVBMenu As Boolean
   lShortCutStartPos As Long
   bMarkToDestroy As Boolean
   bMenuBarBreak As Boolean
   bMenuBreak As Boolean
   bDefault As Boolean
   bIsPresent As Boolean
End Type

' ======================================================================
' API DECLARES, TYPES AND CONSTANTS:
Type POINTAPI
   x As Long
   y As Long
End Type
'Type RECT
'   left As Long
'   tOp As Long
'   Right As Long
'   Bottom As Long
'End Type
' ======================================================================

' =======================================================================
' MENU Declares:
' =======================================================================

' Menu flag constants:
Public Const MF_APPEND = &H100&
Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CALLBACKS = &H8000000
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_CONV = &H40000000
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_END = &H80
Public Const MF_ERRORS = &H10000000
Public Const MF_GRAYED = &H1&
Public Const MF_HELP = &H4000&
Public Const MF_HILITE = &H80&
Public Const MF_HSZ_INFO = &H1000000
Public Const MF_INSERT = &H0&
Public Const MF_LINKS = &H20000000
Public Const MF_MASK = &HFF000000
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MOUSESELECT = &H8000&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_POSTMSGS = &H4000000
Public Const MF_REMOVE = &H1000&
Public Const MF_SENDMSGS = &H2000000
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MF_UNHILITE = &H0&
Public Const MF_USECHECKBITMAPS = &H200&
Public Const MF_DEFAULT = &H1000&
' New versions of the names...
Public Const MFS_GRAYED = &H3&
Public Const MFS_DISABLED = MFS_GRAYED
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MFS_ENABLED = MF_ENABLED
Public Const MFS_UNCHECKED = MF_UNCHECKED
Public Const MFS_UNHILITE = MF_UNHILITE
Public Const MFS_DEFAULT = MF_DEFAULT

Public Const MIIM_STATE = &H1&
Public Const MIIM_ID = &H2&
Public Const MIIM_SUBMENU = &H4&
Public Const MIIM_CHECKMARKS = &H8&
Public Const MIIM_TYPE = &H10&
Public Const MIIM_DATA = &H20&

Public Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100

' Owner draw information:
Public Const ODS_CHECKED = &H8
Public Const ODS_DISABLED = &H4
Public Const ODS_FOCUS = &H10
Public Const ODS_GRAYED = &H2
Public Const ODS_SELECTED = &H1
Public Const ODT_BUTTON = 4
Public Const ODT_COMBOBOX = 3
Public Const ODT_LISTBOX = 2
Public Const ODT_MENU = 1

Type MEASUREITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemWidth As Long
   itemHeight As Long
   ItemData As Long
End Type

Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Type MENUITEMTEMPLATE
   mtOption As Integer
   mtID As Integer
   mtString As Byte
End Type

Type MENUITEMTEMPLATEHEADER
   versionNumber As Integer
   offset As Integer
End Type

Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Declare Function CreateMenu Lib "user32" () As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long

Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long

Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Declare Function TrackPopupMenuByLong Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type
Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS) As Long

' =======================================================================
' General API Declares:
' =======================================================================
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetLastError Lib "kernel32" () As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Const GW_CHILD = 5
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA As Long = 48&
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


' =======================================================================
' GDI Declares:
' =======================================================================
' GDI object functions:
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Const BITSPIXEL = 12
    Public Const LOGPIXELSX = 88    '  Logical pixels/inch in X
    Public Const LOGPIXELSY = 90    '  Logical pixels/inch in Y
' System metrics:
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    Public Const SM_CXICON = 11
    Public Const SM_CYICON = 12
    Public Const SM_CXFRAME = 32
    'Public Const SM_CYCAPTION = 4
    Public Const SM_CYFRAME = 33
    Public Const SM_CYBORDER = 6
    Public Const SM_CXBORDER = 5

' Region paint and fill functions:
Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
    Public Const FLOODFILLBORDER = 0
    Public Const FLOODFILLSURFACE = 1

' Pen functions:
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    Public Const PS_DASH = 1
    Public Const PS_DASHDOT = 3
    Public Const PS_DASHDOTDOT = 4
    Public Const PS_DOT = 2
    Public Const PS_SOLID = 0
    Public Const PS_NULL = 5

' Brush functions:
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long

' Line functions:
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
    Public Const BF_LEFT = &H1
    Public Const BF_BOTTOM = &H8
    Public Const BF_RIGHT = &H4
    Public Const BF_TOP = &H2
    Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
    Public Const BDR_INNER = &HC
    Public Const BDR_OUTER = &H3
    Public Const BDR_RAISED = &H5
    Public Const BDR_RAISEDINNER = &H4
    Public Const BDR_RAISEDOUTER = &H1
    Public Const BDR_SUNKEN = &HA
    Public Const BDR_SUNKENINNER = &H8
    Public Const BDR_SUNKENOUTER = &H2
    Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
    Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
    Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
    Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

' Colour functions:
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
    Public Const OPAQUE = 2
    Public Const TRANSPARENT = 1
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
    Public Const COLOR_ACTIVEBORDER = 10
    Public Const COLOR_ACTIVECAPTION = 2
    Public Const COLOR_ADJ_MAX = 100
    Public Const COLOR_ADJ_MIN = -100
    Public Const COLOR_APPWORKSPACE = 12
    Public Const COLOR_BACKGROUND = 1
    Public Const COLOR_BTNFACE = 15
    Public Const COLOR_BTNHIGHLIGHT = 20
    Public Const COLOR_BTNSHADOW = 16
    Public Const COLOR_BTNTEXT = 18
    Public Const COLOR_CAPTIONTEXT = 9
    Public Const COLOR_GRAYTEXT = 17
    Public Const COLOR_HIGHLIGHT = 13
    Public Const COLOR_HIGHLIGHTTEXT = 14
    Public Const COLOR_INACTIVEBORDER = 11
    Public Const COLOR_INACTIVECAPTION = 3
    Public Const COLOR_INACTIVECAPTIONTEXT = 19
    Public Const COLOR_MENU = 4
    Public Const COLOR_MENUTEXT = 7
    Public Const COLOR_SCROLLBAR = 0
    Public Const COLOR_WINDOW = 5
    Public Const COLOR_WINDOWFRAME = 6
    Public Const COLOR_WINDOWTEXT = 8
    Public Const COLORONCOLOR = 3

' Shell Extract icon functions:
Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

' GDI icon functions:
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

' Blitting functions
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Const SRCCOPY = &HCC0020
    Public Const SRCINVERT = &H660046
    Public Const BLACKNESS = &H42
    Public Const WHITENESS = &HFF0062
    Public Const SRCAND = &H8800C6
    Public Const SRCERASE = &H440328
    Public Const SRCPAINT = &HEE0086
    
Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Declare Function LoadBitmapBynum Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function LoadImageByNum Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
    Public Const LR_LOADMAP3DCOLORS = &H1000
    Public Const LR_LOADFROMFILE = &H10
    Public Const LR_LOADTRANSPARENT = &H20
    Public Const IMAGE_BITMAP = 0

' Text functions:
Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Public Const DT_BOTTOM = &H8
    Public Const DT_CENTER = &H1
    Public Const DT_LEFT = &H0
    Public Const DT_CALCRECT = &H400
    Public Const DT_WORDBREAK = &H10
    Public Const DT_VCENTER = &H4
    Public Const DT_TOP = &H0
    Public Const DT_TABSTOP = &H80
    Public Const DT_SINGLELINE = &H20
    Public Const DT_RIGHT = &H2
    Public Const DT_NOCLIP = &H100
    Public Const DT_INTERNAL = &H1000
    Public Const DT_EXTERNALLEADING = &H200
    Public Const DT_EXPANDTABS = &H40
    Public Const DT_CHARSTREAM = 4
Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As Long, ByVal nCount As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Boolean
    Public Const DI_MASK = 1
    Public Const DI_IMAGE = 2
    Public Const DI_NORMAL = 3
    Public Const DI_COMPAT = 4
    Public Const DI_DEFAULTSIZE = 8

Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
    Public Const SW_SHOWNOACTIVATE = 4

' Scrolling and region functions:
Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal hSavedDC As Long) As Long

Public Const LF_FACESIZE = 32
Type LOGFONT
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
        lfFaceName(LF_FACESIZE) As Byte
End Type
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700
Public Const FF_DONTCARE = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const DEFAULT_CHARSET = 1
Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Declare Function DrawState Lib "user32" Alias "DrawStateA" _
   (ByVal hdc As Long, _
   ByVal hBrush As Long, _
   ByVal lpDrawStateProc As Long, _
   ByVal lParam As Long, _
   ByVal wParam As Long, _
   ByVal x As Long, _
   ByVal y As Long, _
   ByVal cX As Long, _
   ByVal cY As Long, _
   ByVal fuFlags As Long) As Long
'/* Image type */
Public Const DST_COMPLEX = &H0
Public Const DST_TEXT = &H1
Public Const DST_PREFIXTEXT = &H2
Public Const DST_ICON = &H3
Public Const DST_BITMAP = &H4

' /* State type */
Public Const DSS_NORMAL = &H0
Public Const DSS_UNION = &H10         ' /* Gray string appearance */
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const DSS_RIGHT = &H8000

' =======================================================================
' Image list Declares:
' =======================================================================
' Create/Destroy functions:
Declare Function ImageList_Create Lib "COMCTL32.DLL" ( _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal fMask As Long, _
        ByVal cInitial As Long, _
        ByVal cGrow As Long _
    ) As Long
Public Const ILC_MASK = 1&
Public Const ILC_COLOR = 0&
Public Const ILC_COLORDDB = &HFE&
Public Const ILC_COLOR4 = &H4&
Public Const ILC_COLOR8 = &H8&
Public Const ILC_COLOR16 = &H10&
Public Const ILC_COLOR24 = &H18&
Public Const ILC_COLOR32 = &H20&
Public Const ILC_PALETTE = &H800&

Declare Function ImageList_Destroy Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
    ) As Long
    
' Add functions:
Declare Function ImageList_Add Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal hBmpMask As Long _
    ) As Long
Declare Function ImageList_AddMasked Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hBmp As Long, _
        ByVal crMask As Long _
    ) As Long
Declare Function ImageList_AddIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal hIcon As Long _
    ) As Long
Declare Function ImageList_LoadImage Lib "COMCTL32.DLL" ( _
        ByVal hInst As Long, _
        ByVal lpBmp As String, _
        ByVal cX As Long, _
        ByVal cGrow As Long, _
        ByVal crMask As Long, _
        ByVal uType As Long, _
        ByVal uFlags As Long _
    ) As Long
    
' Modification/deletion functions:
Declare Function ImageList_Remove Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long _
    ) As Long
Declare Function ImageList_Replace Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hBmpImage As Long, _
        ByVal hBmpMask As Long _
    ) As Long
Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hIcon As Long _
    ) As Long
    
' Image information functions:
Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
    ) As Long
Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        prcImage As RECT _
    ) As Long
Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        cX As Long, _
        cY As Long _
    ) As Long
Type IMAGEINFO
    hBitmapImage As Long
    hBitmapMask As Long
    cPlanes As Long
    cBitsPerPixel As Long
    rcImage As RECT
End Type
Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        pImageInfo As IMAGEINFO _
    )
    
' Create a new icon based on an image list icon:
Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal diIgnore As Long _
    ) As Long
    
' Merge and move functions:
Declare Function ImageList_Merge Lib "COMCTL32.DLL" ( _
        ByVal hIml1 As Long, _
        ByVal i As Long, _
        ByVal hIml2 As Long, _
        ByVal i2 As Long, _
        ByVal dx As Long, _
        ByVal dy As Long _
    ) As Long
Declare Sub ImageList_CopyDitherImage Lib "COMCTL32.DLL" ( _
        ByVal hImlDst As Long, _
        ByVal iDst As Integer, _
        ByVal xDst As Long, _
        ByVal yDst As Long, _
        ByVal hImlSrc As Long, _
        ByVal iSrc As Long _
    )
Declare Function ImageList_AddFromImageList Lib "COMCTL32.DLL" ( _
        ByVal hImlDest As Long, _
        ByVal hImlSrc As Long, _
        ByVal iSrc As Long _
    ) As Long
    
' Get/Set Background Colour:
Declare Function ImageList_SetBkColor Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal clrBk As Long _
    ) As Long
Public Const CLR_NONE = -1
Public Const CLR_DEFAULT = -16777216
Public Const CLR_HILIGHT = -16777216
Declare Function ImageList_GetBkColor Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long _
    ) As Long

' Draw:
Declare Function ImageList_Draw Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal hdcDst As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal fStyle As Long _
    ) As Long
Type IMAGELISTDRAWPARAMS
    cbSize As Long
    hIml As Long
    i As Long
    hdcDst As Long
    x As Long
    y As Long
    cX As Long
    cY As Long
    xBitmap As Long '        // x offest from the upperleft of bitmap
    yBitmap As Long '        // y offset from the upperleft of bitmap
    rgbBk As Long
    rgbFg As Long
    fStyle As Long
    dwRop As Long
End Type
Declare Function ImageList_DrawIndirect Lib "COMCTL32.DLL" (pimldp As IMAGELISTDRAWPARAMS) As Long

Declare Function ImageList_SetOverlayImage Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal iImage As Long, _
        ByVal iOverlay As Long _
    ) As Long
Public Const ILD_NORMAL = 0
Public Const ILD_TRANSPARENT = 1
Public Const ILD_BLEND25 = 2
Public Const ILD_SELECTED = 4
Public Const ILD_FOCUS = 4
Public Const ILD_MASK = &H10&
Public Const ILD_IMAGE = &H20&
Public Const ILD_ROP = &H40&
Public Const ILD_OVERLAYMASK = 3840

Declare Function ImageList_BeginDrag Lib "COMCTL32.DLL" ( _
        ByVal hIml As Long, _
        ByVal i As Long, _
        ByVal dxHotSpot As Long, _
        ByVal dyHotSpot As Long _
    ) As Long
Declare Function ImageList_DragMove Lib "COMCTL32.DLL" ( _
        ByVal x As Long, _
        ByVal y As Long _
    ) As Long
Declare Function ImageList_DragShow Lib "COMCTL32.DLL" ( _
        ByVal fShow As Long _
    ) As Long
Declare Function ImageList_EndDrag Lib "COMCTL32.DLL" () As Long



Public Sub ImageListDrawIcon( _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        Optional ByVal bSelected As Boolean = False, _
        Optional ByVal bBlend25 As Boolean = False _
    )
Dim lFlags As Long
Dim lR As Long

    lFlags = ILD_TRANSPARENT
    If (bSelected) Then
        lFlags = lFlags Or ILD_SELECTED
    End If
    If (bBlend25) Then
        lFlags = lFlags Or ILD_BLEND25
    End If
    lR = ImageList_Draw( _
            hIml, _
            iIconIndex, _
            hdc, _
            lX, _
            lY, _
            lFlags)
    If (lR = 0) Then
        Debug.Print "Failed to draw Image: " & iIconIndex & " onto hDC " & hdc, "ImageListDrawIcon"
    End If
End Sub
Public Sub ImageListDrawIconDisabled( _
        ByVal hdc As Long, _
        ByVal hIml As Long, _
        ByVal iIconIndex As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lSize As Long _
    )
Dim lR As Long
Dim hIcon As Long

   hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
   lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
   DestroyIcon hIcon
   
End Sub

Public Sub TileArea( _
        ByVal hdcTo As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal Width As Long, _
        ByVal Height As Long, _
        ByVal hDcSrc As Long, _
        ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long _
    )
Dim lSrcX As Long
Dim lSrcY As Long
Dim lSrcStartX As Long
Dim lSrcStartY As Long
Dim lSrcStartWidth As Long
Dim lSrcStartHeight As Long
Dim lDstX As Long
Dim lDstY As Long
Dim lDstWidth As Long
Dim lDstHeight As Long

    lSrcStartX = (x Mod SrcWidth)
    lSrcStartY = (y Mod SrcHeight)
    lSrcStartWidth = (SrcWidth - lSrcStartX)
    lSrcStartHeight = (SrcHeight - lSrcStartY)
    lSrcX = lSrcStartX
    lSrcY = lSrcStartY
    
    lDstY = y
    lDstHeight = lSrcStartHeight
    
    Do While lDstY < (y + Height)
        If (lDstY + lDstHeight) > (y + Height) Then
            lDstHeight = y + Height - lDstY
        End If
        lDstWidth = lSrcStartWidth
        lDstX = x
        lSrcX = lSrcStartX
        Do While lDstX < (x + Width)
            If (lDstX + lDstWidth) > (x + Width) Then
                lDstWidth = x + Width - lDstX
                If (lDstWidth = 0) Then
                    lDstWidth = 4
                End If
            End If
            'If (lDstWidth > Width) Then lDstWidth = Width
            'If (lDstHeight > Height) Then lDstHeight = Height
            BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, lSrcY, SRCCOPY
            lDstX = lDstX + lDstWidth
            lSrcX = 0
            lDstWidth = SrcWidth
        Loop
        lDstY = lDstY + lDstHeight
        lSrcY = 0
        lDstHeight = SrcHeight
    Loop
End Sub



