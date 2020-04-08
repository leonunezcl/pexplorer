Attribute VB_Name = "Apis"
Option Explicit

Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long, ByVal bErase As Long)
Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Long, ByVal t As Long)
Private Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
    ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Public Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal hRgn As Long) As Long
Public Declare Function OffsetClipRgn Lib "gdi32" (ByVal hdc As Long, _
    ByVal x As Long, ByVal y As Long) As Long
Public Const RGN_OR = 2
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional parameter
    lpClass As String 'Optional parameter
    hkeyClass As Long 'Optional parameter
    dwHotKey As Long 'Optional parameter
    hIcon As Long 'Optional parameter
    hProcess As Long 'Optional parameter
End Type

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" _
        (SEI As SHELLEXECUTEINFO) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Sub ShowProgress(Mode As Boolean)

    Dim rc As RECT

    Main.staBar.Panels(3).Visible = Mode
    
    If Mode Then
        With Main.pgbStatus
            .Left = Main.staBar.Panels(3).Left
            .Top = Main.staBar.Top + 2
            .Width = Main.staBar.Panels(3).Width
            .Visible = True
            .Max = 100
            .Value = 1
            .ZOrder 0
        End With
    Else
        Main.pgbStatus.Visible = False
    End If
    
End Sub
Sub CenterWindow(ByVal hwnd As Long)

    Dim wRect As RECT
    
    Dim x As Integer
    Dim y As Integer

    Dim ret As Long
    
    ret = GetWindowRect(hwnd, wRect)
    
    x = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hwnd, vbNull, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub

Public Function SaveDialog(ByVal hwnd As Long, Filter As String, Title As String, InitDir As String) As String
 
    Dim ofn As OPENFILENAME
    Dim A As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = hwnd
    ofn.hInstance = App.hInstance
    
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    A = GetSaveFileName(ofn)

    If (A) Then
        SaveDialog = Trim$(ofn.lpstrFile)
    Else
        SaveDialog = ""
    End If

End Function

Public Function ShowProperties(Filename As String, OwnerhWnd As Long) As Long
        
    '     'open a file properties property page for specified file if return value
    '     '<=32 an error occurred
    '     'From: Delphi code provided by "Ian Land" (iml@dircon.co.uk)
    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
     
    '     'Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
       
    '     'call the API
    r = ShellExecuteEX(SEI)
 
    '     'return the instance handle as a sign of success
    ShowProperties = SEI.hInstApp
       
End Function

Public Sub ColorSQL(rtb As Control, ByVal sSearch As String, ByVal Color As ColorConstants)

    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim Sql As String
        
    lPos = 1
        
    Sql = UCase$(rtb.Text)
    
    Do While lPos < Len(Sql)
        
        sTmp = Mid(Sql, lPos, Len(Sql))
        
        lWhere = InStr(sTmp, UCase$(sSearch))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            rtb.SelStart = lPos - 2
            rtb.SelLength = Len(sSearch)
            rtb.SelColor = Color
            'rtb.SelBold = True
            rtb.SelLength = 0
            'rtb.SelBold = False
        Else
            Exit Do
        End If
    Loop
    
End Sub
Public Sub ColorReporte(rtb As Control, ByVal sSearch As String)

    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim Sql As String
        
    lPos = 1
        
    Sql = UCase$(rtb.Text)
    
    Do While lPos < Len(Sql)
        
        sTmp = Mid(Sql, lPos, Len(Sql))
        
        lWhere = InStr(sTmp, UCase$(sSearch))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            rtb.SelStart = lPos - 2
            rtb.SelLength = Len(sSearch)
            rtb.SelBold = True
            rtb.SelLength = 0
            rtb.SelBold = False
        Else
            Exit Do
        End If
    Loop
    
End Sub

Public Function LeeIni(ByVal Seccion As String, ByVal Llave As String, ByVal ArchivoIni As String) As String

    Dim lret As Long
    Dim ret As String
    
    Dim Buffer As String
    
    Buffer = String$(255, " ")
    
    lret = GetPrivateProfileString(Seccion, Llave, "", Buffer, Len(Buffer), ArchivoIni)
    
    Buffer = Trim$(Buffer)
    ret = Left$(Buffer, Len(Buffer) - 1)
    
    LeeIni = ret
    
End Function

Public Sub GrabaIni(ByVal ArchivoIni As String, ByVal Seccion As String, ByVal Llave As String, ByVal Valor)

    Dim ret As Long
    
    ret = WritePrivateProfileString(Seccion, Llave, CStr(Valor), ArchivoIni)
    
End Sub


Public Sub Shell_Email()

    On Local Error Resume Next
    ShellExecute Main.hwnd, vbNullString, "mailto:proyectovb@hotmail.com", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub
Public Sub Shell_PaginaWeb()

    On Local Error Resume Next
    ShellExecute Main.hwnd, vbNullString, "http://members.xoom.com/proyectovb/", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub

Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' Make a windows "region" based on a given picture box'
    ' picture. This done by passing on the picture line-
    ' by-line and for each sequence of non-transparent
    ' pixels a region is created that is added to the
    ' complete region. I tried to optimize it so it's
    ' fairly fast, but some more optimizations can
    ' always be done - mainly storing the transparency
    ' data in advance, since what takes the most time is
    ' the GetPixel calls, not Create/CombineRgn
    
    Dim x As Long, y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    Dim hdc As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hdc = picSkin.hdc
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    
    ' The transparent color is always the color of the
    ' top-left pixel in the picture. If you wish to
    ' bypass this constraint, you can set the tansparent
    ' color to be a fixed color (such as pink), or
    ' user-configurable
    TransparentColor = GetPixel(hdc, 0, 0)
    
    For y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            
            If GetPixel(hdc, x, y) = TransparentColor Or x = PicWidth Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function
Public Function VBGetFileSize(ByVal Archivo As String) As Long

    Dim lngHandle As Long
    Dim lret As Long
    Dim ret As Long
    
    lngHandle = CreateFile(Archivo, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    lret = GetFileSize(lngHandle, ret)
    CloseHandle lngHandle
    
    VBGetFileSize = Round((lret / 1024), 2)
    
End Function

Public Sub SetToolBarFlat(tlbTemp As Toolbar)
        
    Dim lngStyle As Long
    Dim lngResult As Long
    Dim lngHWND As Long

    lngHWND = FindWindowEx(tlbTemp.hwnd, 0&, "ToolbarWindow32", vbNullString)
    lngStyle = SendMessage(lngHWND, TB_GETSTYLE, &O0, &O0)
    lngStyle = lngStyle Or TBSTYLE_FLAT
    lngResult = SendMessage(lngHWND, TB_SETSTYLE, 0, lngStyle)
    tlbTemp.Refresh
    
End Sub
Public Sub Hourglass(hwnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hwnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If
    DoEvents
    
End Sub
Public Function VBOpenFile(ByVal Archivo As String) As Boolean

    Dim ret As Boolean
    Dim lret As Long
    Dim of As OFSTRUCT
    
    ret = False
    
    lret = OpenFile(Archivo, of, OF_EXIST)
    
    If of.nErrCode = 0 Then ret = True
    
    VBOpenFile = ret
    
End Function

Public Function VBGetTempFileName() As String

    Dim ret As String
    
    ret = String$(260, 0)
    
    GetTempFileName gsTempPath, "ANAPTMP", 0, ret
    
    ret = Left$(ret, InStr(1, ret, Chr$(0)) - 1)
    
    SetFileAttributes ret, FILE_ATTRIBUTE_TEMPORARY
    
    VBGetTempFileName = ret
        
End Function

Public Function VBArchivoSinPath(ByVal ArchivoConPath As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(ArchivoConPath) To 1 Step -1
        If Mid$(ArchivoConPath, k, 1) = "\" Then
            ret = Mid$(ArchivoConPath, k + 1)
            Exit For
        End If
    Next k
    
    VBArchivoSinPath = ret
    
End Function
Public Function VBGetTempPath() As String

    Dim ret As String
    
    ret = String(100, Chr$(0))
    
    GetTempPath 100, ret
    
    ret = Left$(ret, InStr(ret, Chr$(0)) - 1)
    
    VBGetTempPath = ret
    
End Function


