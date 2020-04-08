Attribute VB_Name = "mZip"
Option Explicit

' ======================================================================================
' Name:     mzip
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     1 January 2000
'
' Requires: Info-ZIP's Zip32.DLL v2.32, renamed to vbzip10.dll
'           cUnzip.cls
'
' Copyright � 2000 Steve McMahon for vbAccelerator
' --------------------------------------------------------------------------------------
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
' --------------------------------------------------------------------------------------
'
' Part of the implementation of cUnzip.cls, a class which gives a
' simple interface to Info-ZIP's excellent, free zipping library
' (Zip32.DLL).
'
' This sample uses decompression code by the Info-ZIP group.  The
' original Info-Zip sources are freely available from their website
' at
'     http://www.cdrcom.com/pubs/infozip/
'
' Please ensure you visit the site and read their free source licensing
' information and requirements before using their code in your own
' application.
'
' ======================================================================================


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

' argv
Private Type ZIPnames
    s(0 To 1023) As String
End Type

' Callback large "string" (sic)
Private Type CBChar
    ch(0 To 4096) As Byte
End Type

' Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

' Store the callback functions
Private Type ZIPUSERFUNCTIONS
    lPtrPrint As Long          ' Pointer to application's print routine
    lptrPassword As Long       ' Pointer to application's password routine.
    lptrComment As Long
    lptrService As Long        ' callback function designed to be used for allowing the
                               ' app to process Windows messages, or cancelling the operation
                               ' as well as giving option of progress.  If this function returns
                               ' non-zero, it will terminate what it is doing.  It provides the app
                               ' with the name of the archive member it has just processed, as well
                               ' as the original size.
End Type

Public Type ZPOPT
  Date           As String ' US Date (8 Bytes Long) "12/31/98"?
  szRootDir      As String ' Root Directory Pathname (Up To 256 Bytes Long)
  szTempDir      As String ' Temp Directory Pathname (Up To 256 Bytes Long)
  fTemp          As Long   ' 1 If Temp dir Wanted, Else 0
  fSuffix        As Long   ' Include Suffixes (Not Yet Implemented!)
  fEncrypt       As Long   ' 1 If Encryption Wanted, Else 0
  fSystem        As Long   ' 1 To Include System/Hidden Files, Else 0
  fVolume        As Long   ' 1 If Storing Volume Label, Else 0
  fExtra         As Long   ' 1 If Excluding Extra Attributes, Else 0
  fNoDirEntries  As Long   ' 1 If Ignoring Directory Entries, Else 0
  fExcludeDate   As Long   ' 1 If Excluding Files Earlier Than Specified Date, Else 0
  fIncludeDate   As Long   ' 1 If Including Files Earlier Than Specified Date, Else 0
  fVerbose       As Long   ' 1 If Full Messages Wanted, Else 0
  fQuiet         As Long   ' 1 If Minimum Messages Wanted, Else 0
  fCRLF_LF       As Long   ' 1 If Translate CR/LF To LF, Else 0
  fLF_CRLF       As Long   ' 1 If Translate LF To CR/LF, Else 0
  fJunkDir       As Long   ' 1 If Junking Directory Names, Else 0
  fGrow          As Long   ' 1 If Allow Appending To Zip File, Else 0
  fForce         As Long   ' 1 If Making Entries Using DOS File Names, Else 0
  fMove          As Long   ' 1 If Deleting Files Added Or Updated, Else 0
  fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
  fUpdate        As Long   ' 1 If Updating Zip File-Overwrite Only If Newer, Else 0
  fFreshen       As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
  fJunkSFX       As Long   ' 1 If Junking SFX Prefix, Else 0
  fLatestTime    As Long   ' 1 If Setting Zip File Time To Time Of Latest File In Archive, Else 0
  fComment       As Long   ' 1 If Putting Comment In Zip File, Else 0
  fOffsets       As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
  fPrivilege     As Long   ' 1 If Not Saving Privileges, Else 0
  fEncryption    As Long   ' Read Only Property!!!
  fRecurse       As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories, Else 0
  fRepair        As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
  flevel         As Byte   ' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type

'This assumes zip32.dll is in your \windows\system directory!
Private Declare Function ZpInit Lib "vbzip10.dll" (ByRef tUserFn As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "vbzip10.dll" (ByRef tOpts As ZPOPT) As Long ' Set Zip options
Private Declare Function ZpGetOptions Lib "vbzip10.dll" () As ZPOPT ' used to check encryption flag only
Private Declare Function ZpArchive Lib "vbzip10.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action

' Object for callbacks:
Private m_cZip As cZip
Private m_bCancel As Boolean

Private Function plAddressOf(ByVal lPtr As Long) As Long
   ' VB Bug workaround fn
   plAddressOf = lPtr
End Function

Public Function VBZip( _
      cZipObject As cZip, _
      tZPOPT As ZPOPT, _
      sFileSpecs() As String, _
      iFileCount As Long _
   ) As Long
Dim tUser As ZIPUSERFUNCTIONS
Dim lR As Long
Dim i As Long
Dim sZipFile As String
Dim tZipName As ZIPnames

   m_bCancel = False
   Set m_cZip = cZipObject

   If Not Len(Trim$(m_cZip.BasePath)) = 0 Then
      ChDir m_cZip.BasePath
   End If

   ' Set address of callback functions
   tUser.lPtrPrint = plAddressOf(AddressOf ZipPrintCallback)
   tUser.lptrPassword = plAddressOf(AddressOf ZipPasswordCallback)
   tUser.lptrComment = plAddressOf(AddressOf ZipCommentCallback)
   tUser.lptrService = plAddressOf(AddressOf ZipServiceCallback)  ' not coded yet :-)
   lR = ZpInit(tUser)

   ' Set options
   lR = ZpSetOptions(tZPOPT)
   
   ' Go for it!
   For i = 1 To iFileCount
      tZipName.s(i - 1) = sFileSpecs(i)
   Next i
   tZipName.s(i) = vbNullChar
   
   sZipFile = cZipObject.ZipFile
   lR = ZpArchive(iFileCount, sZipFile, tZipName)
   
   VBZip = lR

End Function

Private Function ZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
Dim iPos As Long
Dim sInfo As String
Dim bCancel As Boolean
    
'-- Always Put This In Callback Routines!
On Error Resume Next
    
   ' Check we've got a message:
   If x > 1 And x < 32000 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), mname, x
      ' Convert to VB string:
      sInfo = StrConv(b, vbUnicode)
      iPos = InStr(sInfo, vbNullChar)
      If iPos > 0 Then
         sInfo = Left$(sInfo, iPos - 1)
      End If
      m_cZip.Service sInfo, bCancel
      If bCancel Then
         ZipServiceCallback = 1
      Else
         ZipServiceCallback = 0
      End If
   End If
End Function

Private Function ZipPrintCallback( _
      ByRef fname As CBChar, _
      ByVal x As Long _
   ) As Long
Dim iPos As Long
Dim sFile As String
   On Error Resume Next
   
   ' Check we've got a message:
   If x > 1 And x < 32000 Then
      ' If so, then get the readable portion of it:
      ReDim b(0 To x) As Byte
      CopyMemory b(0), fname, x
      ' Convert to VB string:
      sFile = StrConv(b, vbUnicode)
      If iPos > 0 Then
         sFile = Left$(sFile, iPos - 1)
      End If
      
      ' Fix up backslashes:
      ReplaceSection sFile, "/", "\"
      
      ' Tell the caller about it
      m_cZip.ProgressReport sFile
   End If
   ZipPrintCallback = 0
End Function

Private Function ZipCommentCallback( _
      ByRef s1 As CBChar _
   ) As CBChar
   ' always put this in callback routines!
   On Error Resume Next
   ' not supported always return \0
   s1.ch(0) = vbNullString
   ZipCommentCallback = s1
End Function

Private Function ZipPasswordCallback( _
      ByRef pwd As CBCh, _
      ByVal x As Long, _
      ByRef s2 As CBCh, _
      ByRef Name As CBCh _
   ) As Long

Dim bCancel As Boolean
Dim sPassword As String
Dim b() As Byte
Dim lSize As Long

On Error Resume Next

   ' The default:
   ZipPasswordCallback = 1
    
   If m_bCancel Then
      Exit Function
   End If
   
   ' Ask for password:
   m_cZip.PasswordRequest sPassword, bCancel
      
   sPassword = Trim$(sPassword)
   
   ' Cancel out if no useful password:
   If bCancel Or Len(sPassword) = 0 Then
      m_bCancel = True
      Exit Function
   End If
   
   ' Put password into return parameter:
   lSize = Len(sPassword)
   If lSize > 254 Then
      lSize = 254
   End If
   b = StrConv(sPassword, vbFromUnicode)
   CopyMemory pwd.ch(0), b(0), lSize
   
   ' Ask UnZip to process it:
   ZipPasswordCallback = 0
       
End Function

Private Function ReplaceSection(ByRef sString As String, ByVal sToReplace As String, ByVal sReplaceWith As String) As Long
Dim iPos As Long
Dim iLastPos As Long
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, sString, "/")
      If (iPos > 1) Then
         Mid$(sString, iPos, 1) = "\"
         iLastPos = iPos + 1
      End If
   Loop While Not (iPos = 0)
   ReplaceSection = iLastPos

End Function
