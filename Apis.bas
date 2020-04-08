Attribute VB_Name = "MApis"
Option Explicit

'Structures Needed For Registry Prototypes
Public Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

' Point struct for ClientToScreen
Private Type PointAPI
    X As Long
    Y As Long
End Type

Public Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function IsDebuggerPresent Lib "kernel32" () As Long
Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hWndLock&)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd&, _
                                              ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, _
                                                                          ByVal nIndex&)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, _
                                                        ByVal nIndex&, ByVal dwNewLong&)

Private Declare Function ClientToScreen& Lib "user32" (ByVal hWnd&, lpPoint As PointAPI)
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
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
    Flags As Long
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

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&

Public Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

Public Function ConfigurarPath(hWnd As Long) As String

    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    Dim ret As String
    
    With udtBI
        'Set the owner window
        .hWndOwner = hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    If ret <> "\" Then
        ret = sPath
    Else
        ret = ""
    End If

    ConfigurarPath = ret
    
End Function

Public Function GetScreenPoint(X As Long, Y As Long, bReturn As Boolean)
  ' this function calls ClientToScreen to convert the passed client point to
  ' a screen point and returns the x or y point depending on the value of bReturn
  
  Dim pt As PointAPI
  
  ' plug the point into the point struct
  pt.X = X
  pt.Y = Y
  
  ' call for the conversion
  Call ClientToScreen(Main.hWnd, pt)
  
  ' return the desired value
  If bReturn Then
    GetScreenPoint = pt.X
  Else
    GetScreenPoint = pt.Y
  End If
  
End Function

Public Sub FlatLviewHeader(lvw As Control)

    Dim lS As Long
    Dim lHwnd As Long

    ' Set the Buttons mode of the ListView's header control:
    lHwnd = SendMessageByLong(lvw.hWnd, LVM_GETHEADER, 0, 0)
    
    If (lHwnd <> 0) Then
        lS = GetWindowLong(lHwnd, GWL_STYLE)
        lS = lS And Not HDS_BUTTONS
        SetWindowLong lHwnd, GWL_STYLE, lS
    End If

End Sub


Public Sub ShowProgress(Mode As Boolean)

    Dim rc As RECT

    Main.staBar.Panels(3).Visible = Mode
    
    If Mode Then
        With Main.pgbStatus
            .Left = Main.staBar.Panels(3).Left
            .Top = Main.staBar.Top + 2
            .Width = Main.staBar.Panels(3).Width
            .Height = Main.staBar.Height - 2
            .Visible = True
            .Max = 100
            .Value = 1
            .ZOrder 0
        End With
    Else
        Main.pgbStatus.Visible = False
    End If
    
End Sub
Sub CenterWindow(ByVal hWnd As Long)

    Dim wRect As RECT
    
    Dim X As Integer
    Dim Y As Integer

    Dim ret As Long
    
    ret = GetWindowRect(hWnd, wRect)
    
    X = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    Y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hWnd, vbNull, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub

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
        .hWnd = OwnerhWnd
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

Public Sub ColorReporte(rtb As Control, ByVal sSearch As String, Optional bUnderline As Boolean = False, Optional ByVal bItalic As Boolean = False)

    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim Sql As String
        
    lPos = 1
        
    Sql = rtb.text
    
    Do While lPos < Len(Sql)
        
        sTmp = Mid(Sql, lPos, Len(Sql))
        
        lWhere = InStr(sTmp, sSearch)
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            
            rtb.SelStart = lPos - 2
            rtb.SelLength = Len(sSearch)
            
            'If Not rtb.SelBold Then
                rtb.SelBold = True
                rtb.SelUnderline = bUnderline
                rtb.SelItalic = bItalic
            'End If
            rtb.SelLength = 0
            rtb.SelBold = False
            rtb.SelUnderline = False
            rtb.SelItalic = False
        Else
            Exit Do
        End If
    Loop
    
End Sub

'limpia todos los nodos de un treeview
Public Sub ClearTreeView(ByVal tvHwnd As Long)

    Dim lNodeHandle As Long
    
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        DoEvents
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
        If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
        Else
            Exit Do
        End If
    Loop

    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Function LeeIni(ByVal Seccion As String, ByVal LLave As String, ByVal ArchivoIni As String) As String

    Dim lRet As Long
    Dim ret As String
    
    Dim Buffer As String
    
    Buffer = String$(255, " ")
    
    lRet = GetPrivateProfileString(Seccion, LLave, "", Buffer, Len(Buffer), ArchivoIni)
    
    Buffer = Trim$(Buffer)
    ret = Left$(Buffer, Len(Buffer) - 1)
    
    LeeIni = ret
    
End Function

Public Sub GrabaIni(ByVal ArchivoIni As String, ByVal Seccion As String, ByVal LLave As String, ByVal Valor)

    Dim ret As Long
    
    ret = WritePrivateProfileString(Seccion, LLave, CStr(Valor), ArchivoIni)
    
End Sub


Public Sub Shell_Email()

    On Local Error Resume Next
    ShellExecute Main.hWnd, vbNullString, "mailto:lnunez@vbsoftware.cl", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub
Public Sub Shell_PaginaWeb()

    On Local Error Resume Next
    ShellExecute Main.hWnd, vbNullString, "http://www.vbsoftware.cl/", vbNullString, "C:\", SW_SHOWMAXIMIZED
    Err = 0
    
End Sub

Public Sub Hourglass(hWnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hWnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If
    DoEvents
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' This function will return true if we are running in the IDE (development) mode else it returns false.
'
' Great for enableling error interception code, eg:
'   If Not InDevelopmentMode Then On Error GoTo ErrorHandler
'
Public Function InDevelopmentMode() As Boolean
   InDevelopmentMode = Not CBool(GetModuleHandle(App.ExeName))
End Function
