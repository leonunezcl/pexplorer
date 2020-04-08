VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFuncFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function MakeTempFile() As String
   Dim sBuffer As String, sPath As String
   Dim nCut As Integer
   Dim dl As Long

   sBuffer = Space$(MAX_PATH)
   dl = GetTempPath(MAX_PATH, sBuffer)
   If dl Then
      sPath = Trim$(Mid$(sBuffer, 1, dl))
   Else
      sPath = App.Path
   End If

   sBuffer = Space$(MAX_PATH)
   dl = GetTempFileName(App.Path, "pn_", 0, sBuffer)
   nCut = InStr(1, sBuffer, Chr(0))
   If nCut Then sBuffer = Trim$(Mid$(sBuffer, 1, nCut - 1))

   If MyFuncFiles.FileExist(sBuffer) Then Kill sBuffer
   MakeTempFile = sBuffer
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



Public Function AppPathFile(sFileName As String) As String
   Dim sFullName As String
   sFullName = App.Path
   If Right$(sFullName, 1) <> "\" Then sFullName = sFullName & "\"
   AppPathFile = sFullName & sFileName
End Function
Function AttachPath(sFileName As String, sPath As String) As String
   If Len(Trim(ExtractPath(sFileName))) = 0 Then
      AttachPath = FixPath(sPath) & sFileName
   Else
      If InStr(1, sFileName, "\") = 0 Then
        AttachPath = sFileName
      Else
        AttachPath = FixPath(sPath) & sFileName
      End If
   End If
End Function


Function FixPath(ByVal sPath As String) As String
   If Len(Trim(sPath)) = 0 Then
      FixPath = ""
   ElseIf Right$(sPath, 1) <> "\" Then
      FixPath = sPath & "\"
   Else
      FixPath = sPath
   End If
End Function



Public Function PathArchivo(ByVal Archivo As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(Archivo) To 1 Step -1
        If Mid$(Archivo, k, 1) = "\" Then
            ret = Mid$(Archivo, 1, k)
            Exit For
        End If
    Next k
    
    PathArchivo = ret
    
End Function


Function ExtractFileName(sFileIn As String) As String
   Dim i As Integer
   For i = Len(sFileIn) To 1 Step -1
      If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
   Next
   ExtractFileName = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
End Function

Function ExtractPath(sPathIn As String) As String
   Dim i As Integer
   For i = Len(sPathIn) To 1 Step -1
      If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
   Next
   ExtractPath = Left$(sPathIn, i)
End Function

' Checks wether file exist (handles wildcards too)
Public Function FileExist(ByVal sFile As String) As Boolean

   If Len(Trim(sFile)) = 0 Then
      ' Nothing given
      FileExist = False
      Exit Function
   ElseIf Right(sFile, 1) = "\" Or Right(sFile, 1) = ":" Then
      ' Just a part of a path or drive... (not complete)
      FileExist = False
      Exit Function
   ElseIf Dir(sFile) = "" Then
      ' Not there...
      FileExist = False
      Exit Function
   End If

   ' After all that torture, it must exist...
   FileExist = True
   Exit Function
ExistErrorHandler:
   FileExist = False
End Function

Public Function ExtractFileExt(sFileName As String) As String
   Dim i As Integer
   For i = Len(sFileName) To 1 Step -1
      If InStr(".", Mid$(sFileName, i, 1)) Then Exit For
   Next
   ExtractFileExt = Right$(sFileName, Len(sFileName) - i)
End Function
Public Function VBGetFileSize(ByVal Archivo As String) As Double

    Dim lngHandle As Long
    Dim lRet As Double
    Dim ret As Long
    Dim of As OFSTRUCT
    
    lngHandle = OpenFile(Archivo, of, 0&)
    lRet = GetFileSize(lngHandle, ret)
    CloseHandle lngHandle
    
    VBGetFileSize = Round((lRet / 1024), 1)
    
End Function

'obtener la fecha de creacion del archivo
Public Function VBGetFileTime(ByVal Archivo As String) As String

    Dim ret As String
    Dim lngHandle As Long
    Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME
    Dim Fecha As String
    Dim Hora As String
    
    Dim of As OFSTRUCT
    
    lngHandle = OpenFile(Archivo, of, 0&)
    
    GetFileTime lngHandle, Ft1, Ft1, Ft2
    
    FileTimeToLocalFileTime Ft2, Ft1
    
    FileTimeToSystemTime Ft1, SysTime
    
    CloseHandle lngHandle
    
    Fecha = Format(Trim(Str$(SysTime.wDay)), "00") & "/" & Format(Trim$(Str$(SysTime.wMonth)), "00") + "/" + LTrim(Str$(SysTime.wYear))
    Hora = Format(Trim(Str$(SysTime.wHour)), "00") & ":" & Format(Trim$(Str$(SysTime.wMinute)), "00") + ":" + LTrim(Str$(SysTime.wSecond))
    
    VBGetFileTime = Fecha & " " & Hora
    
End Function



Public Function VBOpenFile(ByVal Archivo As String) As Boolean

    On Local Error Resume Next
    
    Dim ret As Boolean
    Dim lRet As Long
    Dim of As OFSTRUCT
    Dim nFreeFile As Integer
    
    ret = False
    nFreeFile = FreeFile
    
    Open Archivo For Input Shared As #nFreeFile
    
    If Err = 0 Then
        ret = True
    Else
        ret = False
    End If
    
    Close #nFreeFile
    
    VBOpenFile = ret
    
End Function

