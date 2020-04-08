Attribute VB_Name = "mVarios"
Option Explicit

' VB5 -> msvbvm50.dll
Private Declare Function VarPtrArray& Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any)
Private Declare Sub RtlMoveMemory Lib "kernel32" (dst As Any, src As Any, ByVal nBytes&)
Private Declare Sub RtlZeroMemory Lib "kernel32" (dst As Any, ByVal nBytes&)
Private Type SAFEARRAY1D
    cDims           As Integer
    fFeatures       As Integer
    cbElements      As Long
    cLocks          As Long
    pvData          As Long
    cElements       As Long
    lLbound         As Long
End Type

Private Const ARR_CHUNK& = 1024

Private Const SND_SYNC As Long = &H0
Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_LOOP As Long = &H8
Private Const SND_NOSTOP As Long = &H10
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Play a wave file - keep this? (yes, for the time being)
Public Const WAVE_ACCESSED As Integer = 0
Public Const WAVE_ANALYSE As Integer = 1
Public Const WAVE_ERROR As Integer = 2
Public Const WAVE_EXIT As Integer = 3
Public Const WAVE_OK As Integer = 4
Public Const WAVE_READY As Integer = 5
Public Const WAVE_SORRY As Integer = 6
Public Const WAVE_STANDBY As Integer = 7
Public Const WAVE_STARTUP As Integer = 8
Public Const WAVE_THANKYOU As Integer = 9

Public gbMatchCase As Integer
Public gbWholeWord As Integer
Public gsFindText As String
Public gbLastPos As Integer
Public glbPathPBackup As String

Private gsBlackKeywords As String
Public gsBlackKeywords2 As String
Private gsBlueKeyWords As String

Public gsInforme As String
Public gsLastPath As String

'opciones de analisis
Private Type eOptAnalisis
    Value As Integer
End Type

Public Ana_Archivo() As eOptAnalisis
Public Ana_General() As eOptAnalisis
Public Ana_Variables() As eOptAnalisis
Public Ana_Rutinas() As eOptAnalisis
Public Ana_Opciones() As eOptAnalisis

'opciones de configurar para los archivos
Private Type eAnaArchivos
    Nomenclatura As String
    Clase As String
End Type
Public glbAnaArchivos() As eAnaArchivos

'opciones de configurar para los controles
Private Type eAnaControles
    Nomenclatura As String
    Clase As String
End Type
Public glbAnaControles() As eAnaControles

'tipos de variables
Private Type eAnaTipoVariables
    Nomenclatura As String
    TipoVar As String
End Type
Public glbAnaTipoVariables() As eAnaTipoVariables

'tipos de datos
Private Type eAnaAmbitoDatos
    Ambito As String
    Nomenclatura As String
End Type
Public glbAmbitoDatos() As eAnaAmbitoDatos

Public Enum enumTipoObjeto
    DAO = 1
    ADO = 2
    OTR = 3
End Enum

'tipos de objetos
Private Type eAnaObjetos
    TipoObj As enumTipoObjeto
    Nombre As String
End Type
Public glbArrObj() As eAnaObjetos

Public glbLinXArch As Long
Public glbLarVar As Long
Public glbLinXRuti As Long
Public glbMaxNumParam As Long

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
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Sub sort(xarray() As String)

    Dim Front, Back, i As Integer
    Dim Temp As String
    Dim Arrsize As Integer
    Dim Excom As Integer
    Dim Exswap As Integer
    
    Arrsize = UBound(xarray)
    
    For Front = 0 To Arrsize
        For Back = Front To Arrsize
            Excom = Excom + 1
            
            If xarray(Front) > xarray(Back) Then
                Temp = xarray(Front)
                xarray(Front) = xarray(Back)
                xarray(Back) = Temp
                Exswap = Exswap + 1
            End If
        Next Back
    Next Front
    
End Sub

Public Function isequal(ByVal sDum1 As String, ByVal sDum2 As String)

    Dim fret As Boolean
    
    If LenB(sDum1) = LenB(sDum2) Then
        fret = (InStrB(1, sDum1, sDum2, vbBinaryCompare) <> 0)
    End If

    isequal = fret
    
End Function


Public Function BinarySearch(xarray() As String, ByVal Search As String) As Boolean
  
    Dim hit As Boolean, fTop As Long
    Dim fBott As Long, fTmp As Long
    Dim texto As String
    
    hit = False
    fTop = UBound(xarray)
    fBott = 0
    fTmp = (fTop + fBott) / 2
  
    Do While (Not hit) And (fTop >= fBott)
        texto = xarray(fTmp)
        
        'If isequal(asciiText(texto), asciiText(Search)) Then
        If isequal(texto, Search) Then
            hit = True
        'ElseIf asciiText(Search) < asciiText(texto) Then
        ElseIf Search < texto Then
            fTop = fTmp - 1
        Else
            fBott = fTmp + 1
        End If
        
        fTmp = (fTop + fBott) / 2
        
        If hit Then Exit Do
    Loop
    
    BinarySearch = hit
    
End Function

Private Function asciiText(text As String)
  Dim i As Integer, s As String, c As String
  
  s = ""
  For i = 1 To Len(text)
    c = Mid$(text, i, 1)
    Select Case c
      Case "á": c = "a"
      Case "é": c = "e"
      Case "í": c = "i"
      Case "ó": c = "o"
      Case "ú": c = "u"
      Case "ñ": c = "n"
      Case "ü": c = "u"
      Case "â": c = "a"
      Case "ê": c = "e"
      Case "î": c = "i"
      Case "ô": c = "o"
      Case "û": c = "u"
      Case "Á": c = "A"
      Case "É": c = "E"
      Case "Í": c = "I"
      Case "Ó": c = "O"
      Case "Ú": c = "U"
      Case "Ñ": c = "N"
      Case "Ü": c = "U"
      Case "À": c = "A"
      Case "È": c = "E"
      Case "Ì": c = "I"
      Case "Ò": c = "O"
      Case "Ù": c = "U"
    End Select
    s = s + c
  Next i
  asciiText = s
End Function

Public Sub CargaOpcionesVarias()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim Valor As Variant
    
    glbPathPBackup = LeeIni("pbackup", "path", C_INI)
    
    'cargar valores de analisis de archivo
    ReDim Ana_Opciones(C_ANA_OPCIONES)
    
    For k = 1 To UBound(Ana_Opciones)
        Valor = LeeIni("ana_opciones", CStr(k), C_INI)
        If Valor <> "" Then
            Ana_Opciones(k).Value = CInt(Valor)
        Else
            Ana_Opciones(k).Value = 1
        End If
    Next k
    
    'cargar opciones de objetos
    ReDim glbArrObj(C_ANA_OBJETOS)
    
    'agregar objetos de ado
    glbArrObj(1).Nombre = "Connection": glbArrObj(1).TipoObj = ADO
    glbArrObj(2).Nombre = "Recordset": glbArrObj(2).TipoObj = ADO
    glbArrObj(3).Nombre = "Command": glbArrObj(3).TipoObj = ADO
    glbArrObj(4).Nombre = "Parameter": glbArrObj(4).TipoObj = ADO
    glbArrObj(5).Nombre = "Parameters": glbArrObj(5).TipoObj = ADO
    glbArrObj(6).Nombre = "Field": glbArrObj(6).TipoObj = ADO
    glbArrObj(7).Nombre = "Fields": glbArrObj(7).TipoObj = ADO
    glbArrObj(8).Nombre = "Record": glbArrObj(8).TipoObj = ADO
    glbArrObj(9).Nombre = "Stream": glbArrObj(9).TipoObj = ADO
    glbArrObj(10).Nombre = "Property": glbArrObj(10).TipoObj = ADO
    glbArrObj(11).Nombre = "Properties": glbArrObj(11).TipoObj = ADO
    glbArrObj(12).Nombre = "Error": glbArrObj(12).TipoObj = ADO
    glbArrObj(13).Nombre = "Errors": glbArrObj(13).TipoObj = ADO
    
    'objetos de dao
    glbArrObj(14).Nombre = "Container": glbArrObj(14).TipoObj = DAO
    glbArrObj(15).Nombre = "Containers": glbArrObj(15).TipoObj = DAO
    glbArrObj(16).Nombre = "DBEngine": glbArrObj(16).TipoObj = DAO
    glbArrObj(17).Nombre = "Database": glbArrObj(17).TipoObj = DAO
    glbArrObj(18).Nombre = "Databases": glbArrObj(18).TipoObj = DAO
    glbArrObj(19).Nombre = "Document": glbArrObj(19).TipoObj = DAO
    glbArrObj(20).Nombre = "Documents": glbArrObj(20).TipoObj = DAO
    glbArrObj(21).Nombre = "Group": glbArrObj(21).TipoObj = DAO
    glbArrObj(22).Nombre = "Groups": glbArrObj(22).TipoObj = DAO
    glbArrObj(23).Nombre = "Index": glbArrObj(23).TipoObj = DAO
    glbArrObj(24).Nombre = "Indexes": glbArrObj(24).TipoObj = DAO
    glbArrObj(25).Nombre = "QueryDef": glbArrObj(25).TipoObj = DAO
    glbArrObj(26).Nombre = "QueryDefs": glbArrObj(26).TipoObj = DAO
    glbArrObj(27).Nombre = "Relation": glbArrObj(27).TipoObj = DAO
    glbArrObj(28).Nombre = "Relations": glbArrObj(28).TipoObj = DAO
    glbArrObj(29).Nombre = "TableDef": glbArrObj(29).TipoObj = DAO
    glbArrObj(30).Nombre = "TableDefs": glbArrObj(30).TipoObj = DAO
    glbArrObj(31).Nombre = "User": glbArrObj(31).TipoObj = DAO
    glbArrObj(32).Nombre = "Users": glbArrObj(32).TipoObj = DAO
    glbArrObj(33).Nombre = "Workspace": glbArrObj(33).TipoObj = DAO
    glbArrObj(34).Nombre = "Workspaces": glbArrObj(34).TipoObj = DAO
    
    'otros objetos
    glbArrObj(35).Nombre = "Button": glbArrObj(35).TipoObj = OTR
    glbArrObj(36).Nombre = "Buttons": glbArrObj(36).TipoObj = OTR
    glbArrObj(37).Nombre = "ButtonMenu": glbArrObj(37).TipoObj = OTR
    glbArrObj(38).Nombre = "ButtonMenus": glbArrObj(38).TipoObj = OTR
    glbArrObj(39).Nombre = "ColumnHeader": glbArrObj(39).TipoObj = OTR
    glbArrObj(40).Nombre = "ColumnHeaders": glbArrObj(40).TipoObj = OTR
    glbArrObj(41).Nombre = "ComboItem": glbArrObj(41).TipoObj = OTR
    glbArrObj(42).Nombre = "ComboItems": glbArrObj(42).TipoObj = OTR
    glbArrObj(43).Nombre = "Control": glbArrObj(43).TipoObj = OTR
    glbArrObj(44).Nombre = "Controls": glbArrObj(44).TipoObj = OTR
    glbArrObj(45).Nombre = "Collection": glbArrObj(45).TipoObj = OTR
    glbArrObj(46).Nombre = "DataObject": glbArrObj(46).TipoObj = OTR
    glbArrObj(47).Nombre = "ImageCombo": glbArrObj(47).TipoObj = OTR
    glbArrObj(48).Nombre = "ImageList": glbArrObj(48).TipoObj = OTR
    glbArrObj(49).Nombre = "ListImage": glbArrObj(49).TipoObj = OTR
    glbArrObj(50).Nombre = "ListItem": glbArrObj(50).TipoObj = OTR
    glbArrObj(51).Nombre = "ListView": glbArrObj(51).TipoObj = OTR
    glbArrObj(52).Nombre = "Node": glbArrObj(52).TipoObj = OTR
    glbArrObj(53).Nombre = "Nodes": glbArrObj(53).TipoObj = OTR
    glbArrObj(54).Nombre = "Panel": glbArrObj(54).TipoObj = OTR
    glbArrObj(55).Nombre = "Panels": glbArrObj(55).TipoObj = OTR
    glbArrObj(56).Nombre = "ProgressBar": glbArrObj(56).TipoObj = OTR
    glbArrObj(57).Nombre = "Slider": glbArrObj(57).TipoObj = OTR
    glbArrObj(58).Nombre = "StatusBar": glbArrObj(58).TipoObj = OTR
    glbArrObj(59).Nombre = "Tab": glbArrObj(59).TipoObj = OTR
    glbArrObj(60).Nombre = "Tabs": glbArrObj(60).TipoObj = OTR
    glbArrObj(61).Nombre = "TabStrip": glbArrObj(61).TipoObj = OTR
    glbArrObj(62).Nombre = "Toolbar": glbArrObj(62).TipoObj = OTR
    glbArrObj(63).Nombre = "TreeView": glbArrObj(63).TipoObj = OTR
        
    glbArrObj(64).Nombre = "Recordset": glbArrObj(64).TipoObj = DAO
    glbArrObj(65).Nombre = "Command": glbArrObj(65).TipoObj = DAO
    glbArrObj(66).Nombre = "Parameter": glbArrObj(66).TipoObj = DAO
    glbArrObj(67).Nombre = "Parameters": glbArrObj(67).TipoObj = DAO
    glbArrObj(68).Nombre = "Field": glbArrObj(68).TipoObj = DAO
    glbArrObj(69).Nombre = "Fields": glbArrObj(69).TipoObj = DAO
    glbArrObj(70).Nombre = "Property": glbArrObj(70).TipoObj = DAO
    glbArrObj(71).Nombre = "Properties": glbArrObj(71).TipoObj = DAO
    glbArrObj(72).Nombre = "Error": glbArrObj(72).TipoObj = DAO
    glbArrObj(73).Nombre = "Errors": glbArrObj(73).TipoObj = DAO
    glbArrObj(74).Nombre = "Object": glbArrObj(74).TipoObj = OTR
    
    Err = 0
    
End Sub

'subraya todo aquello que fue analizado y que se detecto no usado
Public Sub ColorizeAnalisisVB(RTF As RichTextBox)

    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    sBuffer = RTF.text
    sTmpWord = ""
    With RTF
        For nI = 1 To Len(sBuffer)
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z", "_", 1 To 9
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI, 2) = vbCrLf Then
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords2, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelStrikeThru = True
                            .SelBold = True
                            .SelColor = RGB(255, 0, 0)
                            .SelText = Mid$(gsBlackKeywords2, nWordPos + 1, Len(sTmpWord))
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
    End With

End Sub

Sub MakeSound(nEvent As Integer, Optional bWait, Optional bForce)
   If IsMissing(bForce) Then bForce = True '(GetIniString(sIniFile, "Options", "WaveSounds", "1") = "1")
   If bForce Then

      If IsMissing(bWait) Then bWait = False
      If Ana_Opciones(2).Value = 0 Then Exit Sub
        
      Select Case nEvent
      Case WAVE_ACCESSED
         PlayWave MyFuncFiles.AppPathFile("Accessed.wav"), bWait
      Case WAVE_ANALYSE
         PlayWave MyFuncFiles.AppPathFile("Analyse.wav"), bWait
      Case WAVE_ERROR
         PlayWave MyFuncFiles.AppPathFile("Error.wav"), bWait
      Case WAVE_EXIT
         PlayWave MyFuncFiles.AppPathFile("Exit.wav"), bWait
      Case WAVE_OK
         PlayWave MyFuncFiles.AppPathFile("Ok.wav"), bWait
      Case WAVE_READY
         PlayWave MyFuncFiles.AppPathFile("Ready.wav"), bWait
      Case WAVE_SORRY
         PlayWave MyFuncFiles.AppPathFile("Sorry.wav"), bWait
      Case WAVE_STANDBY
         PlayWave MyFuncFiles.AppPathFile("StandBy.wav"), bWait
      Case WAVE_STARTUP
         PlayWave MyFuncFiles.AppPathFile("Startup.wav"), bWait
      Case WAVE_THANKYOU
         PlayWave MyFuncFiles.AppPathFile("ThankYou.wav"), bWait
      End Select
   End If
   
End Sub



Sub PlayWave(sSoundFile As String, Optional bWait)
   If MyFuncFiles.FileExist(sSoundFile) Then
      Dim dl As Long, wFlags As Long
      If IsMissing(bWait) Then bWait = False
      If bWait Then
         wFlags = SND_SYNC Or SND_NODEFAULT
      Else
         wFlags = SND_ASYNC Or SND_NODEFAULT
      End If
      dl = sndPlaySound(sSoundFile, wFlags)
   End If
End Sub


'genera un archivo .html
Public Function GuardarArchivoHtml(ByVal Archivo As String, ByVal Titulo As String) As Boolean

    On Local Error GoTo ErrorGuardarArchivoHtml
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    
    ret = True
    
    nFreeFile = FreeFile
    
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head><title>" & Titulo & "</title></head>"
        Print #nFreeFile, "<body>"
        Print #nFreeFile, gsHtml
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirGuardarArchivoHtml
    
ErrorGuardarArchivoHtml:
    ret = False
    SendMail ("GuardarArchivoHtml : " & Err & " " & Error$)
    Resume SalirGuardarArchivoHtml
    
SalirGuardarArchivoHtml:
    GuardarArchivoHtml = ret
    Err = 0
    
End Function
Public Function RTF2HTML(strRTF As String, Optional strOptions As String, Optional strHeader As String, Optional strFooter As String) As String
    'Version 2.9

    'The current version of this function is available at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip

    'More information can be found at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html

    'Converts Rich Text encoded text to HTML format
    'if you find some text that this function doesn't
    'convert properly please email the text to
    'bradyh@bitstream.net

    'Options:
    '+H              add an HTML header and footer
    '+G              add a generator Metatag
    '+T="MyTitle"    add a title (only works if +H is used)
    Dim strHTML As String
    Dim l As Long
    Dim lTmp As Long
    Dim lTmp2 As Long
    Dim lTmp3 As Long
    Dim lRTFLen As Long
    Dim lBOS As Long                 'beginning of section
    Dim lEOS As Long                 'end of section
    Dim strTmp As String
    Dim strTmp2 As String
    Dim strEOS As String             'string to be added to end of section
    Dim strBOS As String             'string to be added to beginning of section
    Dim strEOP As String             'string to be added to end of paragraph
    Dim strBOL As String             'string to be added to the begining of each new line
    Dim strEOL As String             'string to be added to the end of each new line
    Dim strEOLL As String            'string to be added to the end of previous line
    Dim strCurFont As String         'current font code eg: "f3"
    Dim strCurFontSize As String     'current font size eg: "fs20"
    Dim strCurColor As String        'current font color eg: "cf2"
    Dim strFontFace As String        'Font face for current font
    Dim strFontColor As String       'Font color for current font
    Dim lFontSize As Integer         'Font size for current font
    Const gHellFrozenOver = False    'always false
    Dim gSkip As Boolean             'skip to next word/command
    Dim strCodes As String           'codes for ascii to HTML char conversion
    Dim strCurLine As String         'temp storage for text for current line before being added to strHTML
    Dim strColorTable() As String    'table of colors
    Dim lColors As Long              '# of colors
    Dim strFontTable() As String     'table of fonts
    Dim lFonts As Long               '# of fonts
    Dim strFontCodes As String       'list of font code modifiers
    Dim gSeekingText As Boolean      'True if we have to hit text before inserting a </FONT>
    Dim gText As Boolean             'true if there is text (as opposed to a control code) in strTmp
    Dim strAlign As String           '"center" or "right"
    Dim gAlign As Boolean            'if current text is aligned
    Dim strGen As String             'Temp store for Generator Meta Tag if requested
    Dim strTitle As String           'Temp store for Title if requested

    'setup HTML codes
    strCodes = "&nbsp;  {00}&copy;  {a9}&acute; {b4}&laquo; {ab}&raquo; {bb}&iexcl; {a1}&iquest;{bf}&Agrave;{c0}&agrave;{e0}&Aacute;{c1}"
    strCodes = strCodes & "&aacute;{e1}&Acirc; {c2}&acirc; {e2}&Atilde;{c3}&atilde;{e3}&Auml;  {c4}&auml;  {e4}&Aring; {c5}&aring; {e5}&AElig; {c6}"
    strCodes = strCodes & "&aelig; {e6}&Ccedil;{c7}&ccedil;{e7}&ETH;   {d0}&eth;   {f0}&Egrave;{c8}&egrave;{e8}&Eacute;{c9}&eacute;{e9}&Ecirc; {ca}"
    strCodes = strCodes & "&ecirc; {ea}&Euml;  {cb}&euml;  {eb}&Igrave;{cc}&igrave;{ec}&Iacute;{cd}&iacute;{ed}&Icirc; {ce}&icirc; {ee}&Iuml;  {cf}"
    strCodes = strCodes & "&iuml;  {ef}&Ntilde;{d1}&ntilde;{f1}&Ograve;{d2}&ograve;{f2}&Oacute;{d3}&oacute;{f3}&Ocirc; {d4}&ocirc; {f4}&Otilde;{d5}"
    strCodes = strCodes & "&otilde;{f5}&Ouml;  {d6}&ouml;  {f6}&Oslash;{d8}&oslash;{f8}&Ugrave;{d9}&ugrave;{f9}&Uacute;{da}&uacute;{fa}&Ucirc; {db}"
    strCodes = strCodes & "&ucirc; {fb}&Uuml;  {dc}&uuml;  {fc}&Yacute;{dd}&yacute;{fd}&yuml;  {ff}&THORN; {de}&thorn; {fe}&szlig; {df}&sect;  {a7}"
    strCodes = strCodes & "&para;  {b6}&micro; {b5}&brvbar;{a6}&plusmn;{b1}&middot;{b7}&uml;   {a8}&cedil; {b8}&ordf;  {aa}&ordm;  {ba}&not;   {ac}"
    strCodes = strCodes & "&shy;   {ad}&macr;  {af}&deg;   {b0}&sup1;  {b9}&sup2;  {b2}&sup3;  {b3}&frac14;{bc}&frac12;{bd}&frac34;{be}&times; {d7}"
    strCodes = strCodes & "&divide;{f7}&cent;  {a2}&pound; {a3}&curren;{a4}&yen;   {a5}...     {85}"

    'setup color table
    lColors = 0
    ReDim strColorTable(0)
    lBOS = InStr(strRTF, "\colortbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\red")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strColorTable(lColors)
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 4, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 5, 1)), Mid(strRTF, lBOS + 5, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\green")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 6, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 8, 1)), Mid(strRTF, lBOS + 8, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\blue")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 5, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\red")
                lColors = lColors + 1
            Wend
        End If
    End If

    'setup font table
    lFonts = 0
    ReDim strFontTable(0)
    lBOS = InStr(strRTF, "\fonttbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\f0")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strFontTable(lFonts)
                While ((Mid(strRTF, lBOS, 1) <> " ") And (lBOS <= lEOS))
                    lBOS = lBOS + 1
                Wend
                lBOS = lBOS + 1
                strTmp = Mid(strRTF, lBOS, InStr(lBOS, strRTF, ";") - lBOS)
                strFontTable(lFonts) = strFontTable(lFonts) & strTmp
                lBOS = InStr(lBOS, strRTF, "\f" & (lFonts + 1))
                lFonts = lFonts + 1
            Wend
        End If
    End If

    strHTML = ""
    lRTFLen = Len(strRTF)
    'seek first line with text on it
    lBOS = InStr(strRTF, vbCrLf & "\deflang")
    If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 2
    lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally

    While Not gHellFrozenOver
        strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
        l = lBOS
        While l <= lEOS
            strTmp = Mid(strRTF, l, 1)
            Select Case strTmp
                Case "{"
                    l = l + 1
                Case "}"
                    strCurLine = strCurLine & strEOS
                    strEOS = ""
                    l = l + 1
                Case "\"    'special code
                    l = l + 1
                    strTmp = Mid(strRTF, l, 1)
                    Select Case strTmp
                        Case "b"
                            If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                                'b = bold
                                strCurLine = strCurLine & "<B>"
                                strEOS = "</B>" & strEOS
                                If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                            ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
                                strTmp = "•"     'bullet
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "c"
                            If ((Mid(strRTF, l, 2) = "cf") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                                'cf = color font
                                lTmp = Val(Mid(strRTF, l + 2, 5))
                                If lTmp <= UBound(strColorTable) Then
                                    strCurColor = "cf" & lTmp
                                    strFontColor = "#" & strColorTable(lTmp)
                                    gSeekingText = True
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            Else
                                gSkip = True
                            End If
                        Case "e"
                            If (Mid(strRTF, l, 7) = "emdash ") Then
                                strTmp = "—"
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "f"
                            If IsNumeric(Mid(strRTF, l + 1, 1)) Then
                                'f# = font
                                'first get font number
                                lTmp = l + 2
                                strTmp2 = Mid(strRTF, l + 1, 1)
                                While IsNumeric(Mid(strRTF, lTmp, 1))
                                    strTmp2 = strTmp2 & Mid(strRTF, lTmp2, 1)
                                    lTmp = lTmp + 1
                                Wend
                                lTmp = Val(strTmp2)
                                strCurFont = "f" & lTmp
                                If ((lTmp <= UBound(strFontTable)) And (strFontTable(lTmp) <> strFontTable(0))) Then
                                    'insert codes if lTmp is a valid font # AND the font is not the default font
                                    strFontFace = strFontTable(lTmp)
                                    gSeekingText = True
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            ElseIf ((Mid(strRTF, l + 1, 1) = "s") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                                'fs# = font size
                                'first get font size
                                lTmp = l + 3
                                strTmp2 = Mid(strRTF, l + 2, 1)
                                While IsNumeric(Mid(strRTF, lTmp, 1))
                                    strTmp2 = strTmp2 & Mid(strRTF, lTmp, 1)
                                    lTmp = lTmp + 1
                                Wend
                                lTmp = Val(strTmp2)
                                strCurFontSize = "fs" & lTmp
                                lFontSize = Int((lTmp / 5) - 2)
                                If lFontSize = 2 Then
                                    strCurFontSize = ""
                                    lFontSize = 0
                                Else
                                    gSeekingText = True
                                    If lFontSize > 8 Then lFontSize = 8
                                    If lFontSize < 1 Then lFontSize = 1
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            Else
                                gSkip = True
                            End If
                        Case "i"
                            If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                                strCurLine = strCurLine & "<I>"
                                strEOS = "</I>" & strEOS
                                If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "l"
                            If (Mid(strRTF, l, 10) = "ldblquote ") Then
                                'left doublequote
                                strTmp = "“"
                                l = l + 9
                                gText = True
                            ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
                                'left quote
                                strTmp = "‘"
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "p"
                            If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
                                If (Len(strFontColor & strFontFace) > 0) Then
                                    If Not gSeekingText Then strCurLine = strCurLine & "</FONT>"
                                    strFontColor = ""
                                    strFontFace = ""
                                End If
                                If gAlign Then
                                    strCurLine = strCurLine & "</TD></TR></TABLE><BR>"
                                    gAlign = False
                                End If
                                strCurLine = strCurLine & strEOS
                                strEOS = ""
                                If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5    'catch next \ but skip a space
                            ElseIf (Mid(strRTF, l, 9) = "pnlvlblt\") Then
                                'bulleted list
                                strEOS = ""
                                strBOS = "<UL>"
                                strBOL = "<LI>"
                                strEOL = "</LI>"
                                strEOP = "</UL>"
                                l = l + 7    'catch next \
                            ElseIf (Mid(strRTF, l, 7) = "pntext\") Then
                                l = InStr(l, strRTF, "}")   'skip to end of braces
                            ElseIf (Mid(strRTF, l, 6) = "pntxtb") Then
                                l = InStr(l, strRTF, "}")   'skip to end of braces
                            ElseIf (Mid(strRTF, l, 10) = "pard\plain") Then
                                strCurLine = strCurLine & strEOS & strEOP
                                strEOS = ""
                                strEOP = ""
                                strBOL = ""
                                strEOL = "<BR>"
                                l = l + 3    'catch next \
                            Else
                                gSkip = True
                            End If
                        Case "q"
                            If ((Mid(strRTF, l, 3) = "qc\") Or (Mid(strRTF, l, 3) = "qc ")) Then
                                'qc = centered
                                strAlign = "center"
                                'move "cursor" position to next rtf code
                                If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                                l = l + 1
                            ElseIf ((Mid(strRTF, l, 3) = "qr\") Or (Mid(strRTF, l, 3) = "qr ")) Then
                                'qr = right justified
                                strAlign = "right"
                                'move "cursor" position to next rtf code
                                If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                                l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "r"
                            If (Mid(strRTF, l, 7) = "rquote ") Then
                                'reverse quote
                                strTmp = "’"
                                l = l + 6
                                gText = True
                            ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
                                'reverse doublequote
                                strTmp = "”"
                                l = l + 9
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "s"
                            'strikethrough
                            If ((Mid(strRTF, l, 7) = "strike\") Or (Mid(strRTF, l, 7) = "strike ")) Then
                                strCurLine = strCurLine & "<STRIKE>"
                                strEOS = "</STRIKE>" & strEOS
                                l = l + 6
                            Else
                                gSkip = True
                            End If
                        Case "t"
                            If (Mid(strRTF, l, 4) = "tab ") Then
                                strTmp = "&#9;"   'tab
                                l = l + 2
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "u"
                            'underline
                            If ((Mid(strRTF, l, 3) = "ul ") Or (Mid(strRTF, l, 3) = "ul\")) Then
                                strCurLine = strCurLine & "<U>"
                                strEOS = "</U>" & strEOS
                                l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "'"
                            'special characters
                            strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
                            lTmp = InStr(strCodes, strTmp2)
                            If lTmp = 0 Then
                                strTmp = Chr("&H" & Mid(strTmp2, 2, 2))
                            Else
                                strTmp = Trim(Mid(strCodes, lTmp - 8, 8))
                            End If
                            l = l + 1
                            gText = True
                        Case "~"
                            strTmp = " "
                            gText = True
                        Case "{", "}", "\"
                            gText = True
                        Case vbLf, vbCr, vbCrLf    'always use vbCrLf
                            strCurLine = strCurLine & vbCrLf
                        Case Else
                            gSkip = True
                    End Select
                    If gSkip = True Then
                        'skip everything up until the next space or "\" or "}"
                        While InStr(" \}", Mid(strRTF, l, 1)) = 0
                            l = l + 1
                        Wend
                        gSkip = False
                        If (Mid(strRTF, l, 1) = "\") Then l = l - 1
                    End If
                    l = l + 1
                Case vbLf, vbCr, vbCrLf
                    l = l + 1
                Case Else
                    gText = True
            End Select
            If gText Then
                If ((Len(strFontColor & strFontFace) > 0) And gSeekingText) Then
                    If Len(strAlign) > 0 Then
                        gAlign = True
                        If strAlign = "center" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""center""><TD>"
                        ElseIf strAlign = "right" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""right""><TD>"
                        End If
                        strAlign = ""
                    End If
                    If Len(strFontFace) > 0 Then
                        strFontCodes = strFontCodes & " FACE=" & strFontFace
                    End If
                    If Len(strFontColor) > 0 Then
                        strFontCodes = strFontCodes & " COLOR=" & strFontColor
                    End If
                    If Len(strCurFontSize) > 0 Then
                        strFontCodes = strFontCodes & " SIZE = " & lFontSize
                    End If
                    strCurLine = strCurLine & "<FONT" & strFontCodes & ">"
                    strFontCodes = ""
                End If
                strCurLine = strCurLine & strTmp
                l = l + 1
                gSeekingText = False
                gText = False
            End If
        Wend

        lBOS = lEOS + 2
        lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
        strHTML = strHTML & strEOLL & strBOS & strBOL & strCurLine & vbCrLf
        strEOLL = strEOL
        If Len(strEOL) = 0 Then strEOL = "<BR>"

        If lEOS = 0 Then GoTo finally
        strBOS = ""
        strCurLine = ""
    Wend

finally:
    strHTML = strHTML & strEOS
    'clear up any hanging fonts
    If (Len(strFontColor & strFontFace) > 0) Then strHTML = strHTML & "</FONT>" & vbCrLf

    'Add Generator Metatag if requested
    If InStr(strOptions, "+G") <> 0 Then
        strGen = "<META NAME=""GENERATOR"" CONTENT=""RTF2HTML by Brady Hegberg"">"
    Else
        strGen = ""
    End If

    'Add Title if requested
    If InStr(strOptions, "+T") <> 0 Then
        lTmp = InStr(strOptions, "+T") + 3
        lTmp2 = InStr(lTmp + 1, strOptions, """")
        strTitle = Mid(strOptions, lTmp, lTmp2 - lTmp)
    Else
        strTitle = ""
    End If

    'add header and footer if requested
    If InStr(strOptions, "+H") <> 0 Then strHTML = strHeader & vbCrLf _
            & strHTML _
            & strFooter
    RTF2HTML = strHTML
End Function

'convertir el archivo .rtf en archivo .html
Public Function RichToHTML(rtbRichTextBox As RichTextLib.RichTextBox, Optional lngStartPosition As Long, Optional lngEndPosition As Long) As String

    Dim blnBold As Boolean, blnUnderline As Boolean, blnStrikeThru As Boolean
    Dim blnItalic As Boolean, strLastFont As String, lngLastFontColor As Long
    Dim strHTML As String, lngColor As Long, lngRed As Long, lngGreen As Long
    Dim lngBlue As Long, lngCurText As Long, strHex As String, intLastAlignment As Integer

    Const AlignLeft = 0, AlignRight = 1, AlignCenter = 2
    
    'check for lngStartPosition ad lngEndPosition
    
    If IsMissing(lngStartPosition&) Then lngStartPosition& = 0
    If IsMissing(lngEndPosition&) Then lngEndPosition& = Len(rtbRichTextBox.text)
    
    lngLastFontColor& = -1 'no color

    rtbRichTextBox.Visible = False
    gsCadena = rtbRichTextBox.text
    
    gsHtml = "<code>"
    DoEvents
   For lngCurText& = lngStartPosition& To lngEndPosition&
       rtbRichTextBox.SelStart = lngCurText&
       rtbRichTextBox.SelLength = 1
   
          If intLastAlignment% <> rtbRichTextBox.SelAlignment Then
             intLastAlignment% = rtbRichTextBox.SelAlignment
              
                Select Case rtbRichTextBox.SelAlignment
                   Case AlignLeft: gsHtml = gsHtml & "<p align=left>"
                   Case AlignRight: gsHtml = gsHtml & "<p align=right>"
                   Case AlignCenter: gsHtml = gsHtml & "<p align=center>"
                End Select
                
          End If
   
          If blnBold <> rtbRichTextBox.SelBold Then
               If rtbRichTextBox.SelBold = True Then
                 gsHtml = gsHtml & "<b>"
               Else
                 gsHtml = gsHtml & "</b>"
               End If
             blnBold = rtbRichTextBox.SelBold
          End If

          If blnUnderline <> rtbRichTextBox.SelUnderline Then
               If rtbRichTextBox.SelUnderline = True Then
                 gsHtml = gsHtml & "<u>"
               Else
                 gsHtml = gsHtml & "</u>"
               End If
             blnUnderline = rtbRichTextBox.SelUnderline
          End If
   

          If blnItalic <> rtbRichTextBox.SelItalic Then
               If rtbRichTextBox.SelItalic = True Then
                 gsHtml = gsHtml & "<i>"
               Else
                 gsHtml = gsHtml & "</i>"
               End If
             blnItalic = rtbRichTextBox.SelItalic
          End If


          If blnStrikeThru <> rtbRichTextBox.SelStrikeThru Then
               If rtbRichTextBox.SelStrikeThru = True Then
                 gsHtml = gsHtml & "<s>"
               Else
                 gsHtml = gsHtml & "</s>"
               End If
             blnStrikeThru = rtbRichTextBox.SelStrikeThru
          End If

         If strLastFont$ <> rtbRichTextBox.SelFontName Then
            strLastFont$ = rtbRichTextBox.SelFontName
            gsHtml = gsHtml + "<font face=""" & strLastFont$ & """>"
         End If

         If lngLastFontColor& <> rtbRichTextBox.SelColor Then
            lngLastFontColor& = rtbRichTextBox.SelColor
            
            ''Get hexidecimal value of color
            strHex$ = Hex(rtbRichTextBox.SelColor)
            strHex$ = String$(6 - Len(strHex$), "0") & strHex$
            strHex$ = Right$(strHex$, 2) & Mid$(strHex$, 3, 2) & Left$(strHex$, 2)
            
            gsHtml = gsHtml + "<font color=#" & strHex$ & ">"
        End If
         
        On Error Resume Next
        
        If Asc(Mid$(gsCadena, lngCurText + 1, 1)) <> 13 Then
            gsHtml = gsHtml + rtbRichTextBox.SelText
        Else
            gsHtml = gsHtml + rtbRichTextBox.SelText & "<br>"
        End If
            
   Next lngCurText&
    gsHtml = gsHtml & "</code>"
    rtbRichTextBox.Visible = True
RichToHTML = gsHtml

End Function

Public Sub ColorizeVB(RTF As RichTextBox)
    ' #VBIDEUtils#************************************************************
    ' * Programmer Name : Waty Thierry
    ' * Web Site : http://www.vbdiamond.com
    ' * E-Mail :
    ' * Date : 30/10/98
    ' * Time : 14:47
    ' * Module Name : Colorize_Module
    ' * Module Filename : Colorize.bas
    ' * Procedure Name : ColorizeVB
    ' * Parameters :
    ' * rtf As RichTextBox
    ' **********************************************************************
    ' * Comments : Colorize in black, blue, green the VB keywords
    ' *
    ' *
    ' **********************************************************************
    
    Dim sBuffer As String
    Dim nI As Long
    Dim nJ As Long
    Dim sTmpWord As String
    Dim nStartPos As Long
    Dim nSelLen As Long
    Dim nWordPos As Long
            
    sBuffer = RTF.text
    sTmpWord = ""
    
    Call ShowProgress(True)
    Main.pgbStatus.Min = 1
    Main.pgbStatus.Max = Len(sBuffer) + 2

    With RTF
        For nI = 1 To Len(sBuffer)
            Main.pgbStatus.Value = nI
            Main.staBar.Panels(4).text = Round(nI * 100 / Main.pgbStatus.Max, 0) & " %"
            DoEvents
            Select Case Mid$(sBuffer, nI, 1)
                Case "A" To "Z", "a" To "z", "_"
                    If sTmpWord = "" Then nStartPos = nI
                    sTmpWord = sTmpWord & Mid$(sBuffer, nI, 1)
                
                Case Chr$(34)
                    nSelLen = 1
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI + 1, 1) = Chr$(34) Then
                            nI = nI + 2
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                
                Case Chr$(39)
                    .SelStart = nI - 1
                    nSelLen = 0
                    For nJ = 1 To 9999999
                        If Mid$(sBuffer, nI, 2) = vbCrLf Then
                            Exit For
                        Else
                            nSelLen = nSelLen + 1
                            nI = nI + 1
                        End If
                    Next
                    .SelLength = nSelLen
                    .SelColor = RGB(0, 127, 0)
                
                Case Else
                    If Not (Len(sTmpWord) = 0) Then
                        .SelStart = nStartPos - 1
                        .SelLength = Len(sTmpWord)
                        nWordPos = InStr(1, gsBlackKeywords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = RGB(0, 0, 0)
                            .SelText = Mid$(gsBlackKeywords, nWordPos + 1, Len(sTmpWord))
                        End If
                        nWordPos = InStr(1, gsBlueKeyWords, "*" & sTmpWord & "*", 1)
                        If nWordPos <> 0 Then
                            .SelColor = QBColor(9) 'RGB(0, 0, 127)
                            .SelText = Mid$(gsBlueKeyWords, nWordPos + 1, Len(sTmpWord))
                        End If
                        If UCase$(sTmpWord) = "REM" Then
                            .SelStart = nI - 4
                            .SelLength = 3
                            For nJ = 1 To 9999999
                                If Mid$(sBuffer, nI, 2) = vbCrLf Then
                                    Exit For
                                Else
                                    .SelLength = .SelLength + 1
                                    nI = nI + 1
                                End If
                            Next
                            .SelColor = RGB(0, 127, 0)
                            .SelText = LCase$(.SelText)
                        End If
                    End If
                    sTmpWord = ""
            End Select
        Next
        .SelStart = 0
    End With

    Call ShowProgress(False)
    
End Sub
Public Sub InitColorize()
' **********************************************************************
' * Comments : Initialize the VB keywords
' *
' *
' **********************************************************************

    gsBlackKeywords = "*Abs*Add*AddItem*AppActivate*Array*Asc*Atn*Beep*Begin*BeginProperty*ChDir*ChDrive*Choose*Chr*Clear*Collection*Command*Cos*CreateObject*CurDir*DateAdd*DateDiff*DatePart*DateSerial*DateValue*Day*DDB*DeleteSetting*Dir*DoEvents*EndProperty*Environ*EOF*Err*Exp*FileAttr*FileCopy*FileDateTime*FileLen*Fix*Format*FV*GetAllSettings*GetAttr*GetObject*GetSetting*Hex*Hide*Hour*InputBox*InStr*Int*Int*IPmt*IRR*IsArray*IsDate*IsEmpty*IsError*IsMissing*IsNull*IsNumeric*IsObject*Item*Kill*LCase*Left*Len*Load*Loc*LOF*Log*LTrim*Me*Mid*Minute*MIRR*MkDir*Month*Now*NPer*NPV*Oct*Pmt*PPmt*PV*QBColor*Raise*Randomize*Rate*Remove*RemoveItem*Reset*RGB*Right*RmDir*Rnd*RTrim*SaveSetting*Second*SendKeys*SetAttr*Sgn*Shell*Sin*Sin*SLN*Space*Sqr*Str*StrComp*StrConv*Switch*SYD*Tan*Text*Time*Time*Timer*TimeSerial*TimeValue*Trim*TypeName*UCase*Unload*Val*VarType*WeekDay*Width*Year*"
    gsBlueKeyWords = "*#Const*#Else*#ElseIf*#End If*#If*Alias*Alias*And*As*Base*Binary*Boolean*Byte*ByVal*Call*Case*CBool*CByte*CCur*CDate*CDbl*CDec*CInt*CLng*Close*Compare*Const*CSng*CStr*Currency*CVar*CVErr*Decimal*Declare*DefBool*DefByte*DefCur*DefDate*DefDbl*DefDec*DefInt*DefLng*DefObj*DefSng*DefStr*DefVar*Dim*Do*Double*Each*Else*ElseIf*End*Enum*Eqv*Erase*Error*Exit*Explicit*False*For*Function*Get*Global*GoSub*GoTo*If*Imp*In*Input*Input*Integer*Is*LBound*Let*Lib*Like*Line*Lock*Long*Loop*LSet*Name*New*Next*Not*Object*On*Open*Option*Or*Output*Print*Private*Property*Public*Put*Random*Read*ReDim*Resume*Return*RSet*Seek*Select*Set*Single*Spc*Static*String*Stop*Sub*Tab*Then*Then*True*Type*UBound*Unlock*Variant*Wend*While*With*Xor*Nothing*To*Friend*"

End Sub
'carga la nomenclatura para el ambito donde se declaran las variables
Public Sub CargaNomenclaturaAmbitoDatos()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim Valor As Variant
        
    ReDim glbAmbitoDatos(0)
    
    Valor = LeeIni("analisis_ambito", "numero", C_INI)
        
    'hay tipos de variables registrados ?
    If Valor = "" Then
        Call ConfiguraAmbitoXDefecto
    End If
    
    Valor = LeeIni("analisis_ambito", "numero", C_INI)
    
    ReDim glbAmbitoDatos(CInt(Valor))
    
    For k = 1 To UBound(glbAmbitoDatos)
        Valor = LeeIni("analisis_ambito", "ambito" & k, C_INI)
        
        glbAmbitoDatos(k).Nomenclatura = Left$(Valor, InStr(1, Valor, ",") - 1)
        glbAmbitoDatos(k).Ambito = Mid$(Valor, InStr(1, Valor, ",") + 1)
    Next k
    
    Err = 0
    
End Sub

'carga la nomenclatura de controles
Public Sub CargaNomenclaturaControles()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim Valor As Variant
        
    ReDim glbAnaControles(0)
    
    '***
    'lineas x archivo
    Valor = LeeIni("analisis_controles", "numero", C_INI)
        
    'hay controles registrados
    If Valor = "" Then
        Call ConfiguraControlesXDefecto
    End If
    
    Valor = LeeIni("analisis_controles", "numero", C_INI)
    
    ReDim glbAnaControles(CInt(Valor))
    
    For k = 1 To UBound(glbAnaControles)
        Valor = LeeIni("analisis_controles", "ctl" & k, C_INI)
        
        glbAnaControles(k).Nomenclatura = Left$(Valor, InStr(1, Valor, ",") - 1)
        glbAnaControles(k).Clase = Mid$(Valor, InStr(1, Valor, ",") + 1)
    Next k
    
    Err = 0
    
End Sub
'carga la configuracion de la nomenclatura de archivos
Public Sub CargaNomenclaturaDeArchivos()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim Valor As Variant
        
    ReDim glbAnaArchivos(0)
    
    '***
    'lineas x archivo
    Valor = LeeIni("analisis_archivos", "numero", C_INI)
        
    'hay controles registrados
    If Valor = "" Then
        Call ConfiguraArchivosXDefecto
    End If
    
    Valor = LeeIni("analisis_archivos", "numero", C_INI)
    
    ReDim glbAnaArchivos(CInt(Valor))
    
    For k = 1 To UBound(glbAnaArchivos)
        Valor = LeeIni("analisis_archivos", "arch" & k, C_INI)
        
        glbAnaArchivos(k).Nomenclatura = Left$(Valor, InStr(1, Valor, ",") - 1)
        glbAnaArchivos(k).Clase = Mid$(Valor, InStr(1, Valor, ",") + 1)
    Next k

    Err = 0
    
End Sub

'carga la nomenclatura del tipo de variables
Public Sub CargaNomenclaturaTipoVariables()

    On Local Error Resume Next
    
    Dim k As Integer
    Dim Valor As Variant
        
    ReDim glbAnaTipoVariables(0)
    
    Valor = LeeIni("analisis_tipo_variables", "numero", C_INI)
        
    'hay tipos de variables registrados ?
    If Valor = "" Then
        Call ConfiguraTipoVariablesXDefecto
    End If
    
    Valor = LeeIni("analisis_tipo_variables", "numero", C_INI)
    
    ReDim glbAnaTipoVariables(CInt(Valor))
    
    For k = 1 To UBound(glbAnaTipoVariables)
        Valor = LeeIni("analisis_tipo_variables", "tivar" & k, C_INI)
        
        If Len(Valor) > 0 Then
            glbAnaTipoVariables(k).Nomenclatura = Left$(Valor, InStr(1, Valor, ",") - 1)
            glbAnaTipoVariables(k).TipoVar = Mid$(Valor, InStr(1, Valor, ",") + 1)
        End If
    Next k
    
    Err = 0
    
End Sub

'carga las opciones de analisis desde archivo .ini
Public Sub CargaOpcionesDeAnalisis()

    On Local Error Resume Next
    
    Dim Valor As Variant
    Dim k As Integer
    
    '***
    'lineas x archivo
    Valor = LeeIni("analisis", "lineas_x_archivo", C_INI)
    
    If Valor = "" Then Valor = 500
    
    glbLinXArch = Valor
    
    'largo minimo de variables
    Valor = LeeIni("analisis", "largo_variable", C_INI)
    
    If Valor = "" Then Valor = 3
    
    glbLarVar = Valor
    
    'lineas x rutina
    Valor = LeeIni("analisis", "lineas_x_rutina", C_INI)
    
    If Valor = "" Then Valor = 40
    
    glbLinXRuti = Valor
    
    'maximo numero de parametros
    Valor = LeeIni("analisis", "max_parametros_x_rutina", C_INI)
    
    If Valor = "" Then Valor = 5
    
    glbMaxNumParam = Valor
    
    '***
    
    'cargar valores de analisis de archivo
    ReDim Ana_Archivo(C_ANA_ARCHIVOS)
    
    For k = 1 To UBound(Ana_Archivo)
        Valor = LeeIni("ana_archivo", CStr(k), C_INI)
        If Valor <> "" Then
            Ana_Archivo(k).Value = CInt(Valor)
        Else
            Ana_Archivo(k).Value = 1
        End If
    Next k
    
    'cargar valores de analisis general
    ReDim Ana_General(C_ANA_GENERAl)
    
    For k = 1 To UBound(Ana_General)
        Valor = LeeIni("ana_general", CStr(k), C_INI)
        If Valor <> "" Then
            Ana_General(k).Value = CInt(Valor)
        Else
            Ana_General(k).Value = 1
        End If
    Next k
    
    'cargar valores de analisis de variables
    ReDim Ana_Variables(C_ANA_VARIABLES)
    
    For k = 1 To UBound(Ana_Variables)
        Valor = LeeIni("ana_variables", CStr(k), C_INI)
        If Valor <> "" Then
            Ana_Variables(k).Value = CInt(Valor)
        Else
            Ana_Variables(k).Value = 1
        End If
    Next k
    
    'cargar valores de analisis de rutinas
    ReDim Ana_Rutinas(C_ANA_RUTINAS)
    
    For k = 1 To UBound(Ana_Rutinas)
        Valor = LeeIni("ana_rutinas", CStr(k), C_INI)
        If Valor <> "" Then
            Ana_Rutinas(k).Value = CInt(Valor)
        Else
            Ana_Rutinas(k).Value = 1
        End If
    Next k
    
    Err = 0
    
End Sub
'configura ambito x defecto
Private Sub ConfiguraAmbitoXDefecto()

    Call GrabaIni(C_INI, "analisis_ambito", "numero", "11")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito1", "r,ByRef")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito2", "v,ByVal")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito3", "em,Miembro Enumeración")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito4", "f,Friend")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito5", "glb,Global")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito6", "loc,Local")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito7", "mpri,Privada al módulo")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito8", "mpub,Pública al módulo")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito9", "sl,Statica local")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito10", "sta,Statica al módulo")
    Call GrabaIni(C_INI, "analisis_ambito", "ambito11", "et,Elemento tipo")
    
End Sub

'configura la nomenclatura de archivos
Private Sub ConfiguraArchivosXDefecto()

    Call GrabaIni(C_INI, "analisis_archivos", "numero", "9")
    Call GrabaIni(C_INI, "analisis_archivos", "arch1", "frm,Form")
    Call GrabaIni(C_INI, "analisis_archivos", "arch2", "frm,MDIChild")
    Call GrabaIni(C_INI, "analisis_archivos", "arch3", "frm,MDIForm")
    Call GrabaIni(C_INI, "analisis_archivos", "arch4", "bas,Module")
    Call GrabaIni(C_INI, "analisis_archivos", "arch5", "cls,Class")
    Call GrabaIni(C_INI, "analisis_archivos", "arch6", "ctl,UserControl")
    Call GrabaIni(C_INI, "analisis_archivos", "arch7", "pag,PropertyPage")
    Call GrabaIni(C_INI, "analisis_archivos", "arch8", "res,ResFile32")
    Call GrabaIni(C_INI, "analisis_archivos", "arch9", "rel,RelatedDoc")
    
End Sub

'configura un set de controles x defecto
Private Sub ConfiguraControlesXDefecto()

    Call GrabaIni(C_INI, "analisis_controles", "numero", "81")
    Call GrabaIni(C_INI, "analisis_controles", "ctl1", "frm,Form")
    Call GrabaIni(C_INI, "analisis_controles", "ctl2", "mdi,MDIForm")
    Call GrabaIni(C_INI, "analisis_controles", "ctl3", "pic,PictureBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl4", "lbl,Label")
    Call GrabaIni(C_INI, "analisis_controles", "ctl5", "txt,TextBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl6", "fra,Frame")
    Call GrabaIni(C_INI, "analisis_controles", "ctl7", "cmd,CommandButton")
    Call GrabaIni(C_INI, "analisis_controles", "ctl8", "chk,CheckBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl9", "opt,OptionButton")
    Call GrabaIni(C_INI, "analisis_controles", "ctl10", "cbo,ComboBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl11", "lis,ListBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl12", "hsb,HScrollBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl13", "vsb,VScrollBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl14", "tmr,Timer")
    Call GrabaIni(C_INI, "analisis_controles", "ctl15", "drv,DriveListBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl16", "dir,DirListBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl17", "fil,FileListBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl18", "shp,Shape")
    Call GrabaIni(C_INI, "analisis_controles", "ctl19", "lin,Line")
    Call GrabaIni(C_INI, "analisis_controles", "ctl20", "img,Image")
    Call GrabaIni(C_INI, "analisis_controles", "ctl21", "dat,Data")
    Call GrabaIni(C_INI, "analisis_controles", "ctl22", "ole,OLE")
    Call GrabaIni(C_INI, "analisis_controles", "ctl23", "lvw,ListView")
    Call GrabaIni(C_INI, "analisis_controles", "ctl24", "tbr,Toolbar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl25", "stb,StatusBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl26", "pgb,ProgressBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl27", "tvw,TreeView")
    Call GrabaIni(C_INI, "analisis_controles", "ctl28", "iml,ImageList")
    Call GrabaIni(C_INI, "analisis_controles", "ctl29", "sld,Slider")
    Call GrabaIni(C_INI, "analisis_controles", "ctl30", "rtb,RichTextBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl31", "tab,TabStrip")
    Call GrabaIni(C_INI, "analisis_controles", "ctl32", "icbo,ImageCombo")
    Call GrabaIni(C_INI, "analisis_controles", "ctl33", "sschk,SSCheck")
    Call GrabaIni(C_INI, "analisis_controles", "ctl34", "ssfra,SSFrame")
    Call GrabaIni(C_INI, "analisis_controles", "ctl35", "sscmd,SSCommand")
    Call GrabaIni(C_INI, "analisis_controles", "ctl36", "sspan,SSPanel")
    Call GrabaIni(C_INI, "analisis_controles", "ctl37", "ssopt,SSOption")
    Call GrabaIni(C_INI, "analisis_controles", "ctl38", "ssrib,SSRibbon")
    Call GrabaIni(C_INI, "analisis_controles", "ctl39", "sstab,SSTab")
    Call GrabaIni(C_INI, "analisis_controles", "ctl40", "cdlg,CommonDialog")
    Call GrabaIni(C_INI, "analisis_controles", "ctl41", "dlis,DBList")
    Call GrabaIni(C_INI, "analisis_controles", "ctl42", "dcbo,DBCombo")
    Call GrabaIni(C_INI, "analisis_controles", "ctl43", "ani,Animation")
    Call GrabaIni(C_INI, "analisis_controles", "ctl44", "upd,UpDown")
    Call GrabaIni(C_INI, "analisis_controles", "ctl45", "fgri,MSFlexGrid")
    Call GrabaIni(C_INI, "analisis_controles", "ctl46", "wbro,WebBrowser")
    Call GrabaIni(C_INI, "analisis_controles", "ctl47", "scri,Scriptlet")
    Call GrabaIni(C_INI, "analisis_controles", "ctl48", "miav,MSIAV")
    Call GrabaIni(C_INI, "analisis_controles", "ctl49", "micdr,MSICDROM")
    Call GrabaIni(C_INI, "analisis_controles", "ctl50", "miole,MSIOLEReg")
    Call GrabaIni(C_INI, "analisis_controles", "ctl51", "mipri,MSIPrint")
    Call GrabaIni(C_INI, "analisis_controles", "ctl52", "pctrl,PathControl")
    Call GrabaIni(C_INI, "analisis_controles", "ctl53", "sgrph,StructuredGraphicsControl")
    Call GrabaIni(C_INI, "analisis_controles", "ctl54", "sprit,SpriteControl")
    Call GrabaIni(C_INI, "analisis_controles", "ctl55", "secctl,SequencerControl")
    Call GrabaIni(C_INI, "analisis_controles", "ctl56", "nsf,NSFile")
    Call GrabaIni(C_INI, "analisis_controles", "ctl57", "mie,Msie")
    Call GrabaIni(C_INI, "analisis_controles", "ctl58", "adodat,Adodc")
    Call GrabaIni(C_INI, "analisis_controles", "ctl59", "dgrid,DataGrid")
    Call GrabaIni(C_INI, "analisis_controles", "ctl60", "dalis,DataList")
    Call GrabaIni(C_INI, "analisis_controles", "ctl61", "dacbo,DataCombo")
    Call GrabaIni(C_INI, "analisis_controles", "ctl62", "mth,MonthView")
    Call GrabaIni(C_INI, "analisis_controles", "ctl63", "dtp,DTPicker")
    Call GrabaIni(C_INI, "analisis_controles", "ctl64", "fsbar,FlatScrollBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl65", "mhgri,MSHFlexGrid")
    Call GrabaIni(C_INI, "analisis_controles", "ctl66", "mcht,MSChart")
    Call GrabaIni(C_INI, "analisis_controles", "ctl67", "drpt,DataRepeater")
    Call GrabaIni(C_INI, "analisis_controles", "ctl68", "ine,Inet")
    Call GrabaIni(C_INI, "analisis_controles", "ctl69", "wsk,Winsock")
    Call GrabaIni(C_INI, "analisis_controles", "ctl70", "mapse,MAPISession")
    Call GrabaIni(C_INI, "analisis_controles", "ctl71", "mapme,MAPIMessages")
    Call GrabaIni(C_INI, "analisis_controles", "ctl72", "mmctl,MMControl")
    Call GrabaIni(C_INI, "analisis_controles", "ctl73", "pclp,PictureClip")
    Call GrabaIni(C_INI, "analisis_controles", "ctl74", "mcom,MSComm")
    Call GrabaIni(C_INI, "analisis_controles", "ctl75", "msk,MaskEdBox")
    Call GrabaIni(C_INI, "analisis_controles", "ctl76", "cbar,CoolBar")
    Call GrabaIni(C_INI, "analisis_controles", "ctl77", "swiz,SubWizard")
    Call GrabaIni(C_INI, "analisis_controles", "ctl78", "sinf,SysInfo")
    Call GrabaIni(C_INI, "analisis_controles", "ctl79", "aplug,ActiveXPlugin")
    Call GrabaIni(C_INI, "analisis_controles", "ctl80", "npla,NSPlay")
    Call GrabaIni(C_INI, "analisis_controles", "ctl81", "amov,ActiveMovie")
        
End Sub
'configura el tipo de variables x defecto
Private Sub ConfiguraTipoVariablesXDefecto()

    Call GrabaIni(C_INI, "analisis_tipo_variables", "numero", "28")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar1", "byt,Byte")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar2", "cl,Collection")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar3", "ctl,Control")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar4", "Cur,Currency")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar5", "db,Database")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar6", "dte,Date")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar7", "dec,Decimal")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar8", "dbl,Double")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar9", "ds,Dynaset")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar10", "err,Error")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar11", "fld,Field")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar12", "idx,Index")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar13", "int,Integer")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar14", "lng,Long")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar15", "obj,Object")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar16", "qdf,QueryDef")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar17", "rs,Recordset")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar18", "rpt,Report")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar19", "sng,Single")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar20", "ss,Snapshot")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar21", "str,String")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar22", "tbl,Table")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar23", "tdf,TableDef")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar24", "var,Variant")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar25", "ws,Workspace")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar26", "cn,Connection")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar27", "cmd,Command")
    Call GrabaIni(C_INI, "analisis_tipo_variables", "tivar28", "par,Parameter")
        
End Sub

Public Sub FontStuff(ByVal Titulo As String, picDraw As PictureBox, Optional ByVal Angulo As Integer = 90)
    
    On Error GoTo GetOut
    Dim f As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
    Dim FONTSIZE As Integer
    FONTSIZE = 10 'Val(txtSize.Text)
    
    f.lfEscapement = 10 * Angulo 'Val(txtDegree.Text) 'rotation angle, in tenths
    FontName = "Tahoma" + Chr$(0) 'null terminated
    f.lfFaceName = FontName
    f.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(f)
    hPrevFont = SelectObject(picDraw.hDC, hFont)
    
    picDraw.CurrentX = 3
    'picDraw.CurrentY = 310
    
    If Angulo > 0 Then
        picDraw.CurrentY = picDraw.Height - 10
    End If
    
    picDraw.Print Titulo
        
    '  Clean up, restore original font
    hFont = SelectObject(picDraw.hDC, hPrevFont)
    DeleteObject hFont
    
    Exit Sub
GetOut:
    Exit Sub

End Sub

Public Function ContarTipoDependencias(ByVal Tipo As eTipoDepencia) As Integer

    Dim k As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For k = 1 To UBound(Proyecto.aDepencias)
        If Proyecto.aDepencias(k).Tipo = Tipo Then
            ret = ret + 1
        End If
    Next k
    
    ContarTipoDependencias = ret
    
End Function

Public Function ContarTipoRutinas(ByVal Indice As Integer, ByVal Tipo As eTipoRutinas) As Integer

    Dim r As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For r = 1 To UBound(Proyecto.aArchivos(Indice).aRutinas)
        If Proyecto.aArchivos(Indice).aRutinas(r).Tipo = Tipo Then
            ret = ret + 1
        End If
    Next r
    
    ContarTipoRutinas = ret
    
End Function
Public Function ContarTiposDeArchivos(ByVal Tipo As eTipoArchivo) As Integer

    Dim k As Integer
    
    Dim ret As Integer
    
    ret = 0
    
    For k = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(k).TipoDeArchivo = Tipo Then
            ret = ret + 1
        End If
    Next k
    
    ContarTiposDeArchivos = ret
    
End Function

'muestra las propiedades del archivo seleccionado en el proyecto.
Public Sub PropiedadesArchivo()

    Dim k As Integer
    Dim Archivo As String
    
    If Main.lvwFiles.SelectedItem Is Nothing Then
        MsgBox "Debe seleccionar un archivo.", vbCritical
        Exit Sub
    End If
            
    Archivo = LCase$(Main.lvwFiles.SelectedItem.text)
    
    'no es componente es un archivo del proyecto
    For k = 1 To UBound(Proyecto.aDepencias)
        If LCase$(MyFuncFiles.VBArchivoSinPath(Proyecto.aDepencias(k).ContainingFile)) = Archivo Then
            Call ShowProperties(Proyecto.aDepencias(k).ContainingFile, Main.hwnd)
            Exit Sub
        End If
    Next k
        
    For k = 1 To UBound(Proyecto.aArchivos)
        If LCase$(MyFuncFiles.VBArchivoSinPath(Proyecto.aArchivos(k).PathFisico)) = Archivo Then
            Call ShowProperties(Proyecto.aArchivos(k).PathFisico, Main.hwnd)
        End If
    Next k
        
End Sub

Public Sub SendMail(ByVal MsgError As String)

    Dim Msg As String
    
    Msg = "Desea reportar este error : " & vbNewLine & vbNewLine
    Msg = Msg & MsgError
    
    If Confirma(Msg) = vbYes Then
        Call Shell_Email
    End If
    
End Sub

Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Sub Copiar(ByVal hwnd As Long)

    Dim ret As Long
    
    ret = SendMessage(hwnd, WM_COPY, 0, 0)
    
End Sub

Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function



Public Sub CargaRutinas(ByVal frm As Form, ByVal Tipo As eTipoRutinas)

    Dim k As Integer
    Dim itmx As ListItem
    Dim j As Integer
    Dim r As Integer
    
    Call Hourglass(frm.hwnd, True)
    
    j = 1
    For k = 1 To UBound(Proyecto.aArchivos)
'        MsgBox Proyecto.aArchivos(k).Nombre
        If Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 1, 1
                    Set itmx = frm.lview.ListItems(j)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_BAS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 2, 2
                    Set itmx = frm.lview.ListItems(j)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_OCX Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 4, 4
                    Set itmx = frm.lview.ListItems(j)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        ElseIf Proyecto.aArchivos(k).TipoDeArchivo = TIPO_ARCHIVO_CLS Then
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).Tipo = Tipo Then
                    frm.lview.ListItems.Add , , Proyecto.aArchivos(k).Nombre, 3, 3
                    Set itmx = frm.lview.ListItems(j)
                    itmx.SubItems(1) = Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    itmx.SubItems(2) = Proyecto.aArchivos(k).aRutinas(r).Nombre
                    itmx.SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas
                    j = j + 1
                End If
            Next r
        End If
    Next k
    
    Call Hourglass(frm.hwnd, False)
    
    Set itmx = Nothing
    
End Sub

'busca una
Public Function MyInstr(ByVal Search As String, ByVal What As String, _
                        Optional ByVal Todos As Boolean = True) As Boolean
            
    Dim StringArray() As String
    Dim SearchLen As Integer
    Dim k As Integer
    Dim p As Integer
    Dim c As Integer
    Dim Buffer As String
    Dim ret As Boolean
    Dim Chars As String
    Dim Ntokens As Long
    Dim lSearch As String
    Dim p1 As Integer
    Dim p2 As Integer
    
    ret = False
    p = 1
    c = 0
    Buffer = Search

    'verificar si hay alguna coincidencia en el codigo
    If Len(Search) = 0 Or InStr(Search, What) = 0 Then                  'viene en blanco
        MyInstr = False
        Exit Function
    End If
    
    'verificar que no se analizen cadenas
    lSearch = Search
    If InStr(1, lSearch, Chr$(34)) > 0 Then
        Do
            p1 = InStr(1, lSearch, Chr$(34))
            
            If p1 > 0 Then
                'buscar la otra posicion
                p2 = InStr(p1 + 1, lSearch, Chr$(34))
                If p2 > 0 Then
                    lSearch = Left$(lSearch, p1 - 1) & Mid$(lSearch, p2 + 1)
                Else
                    Search = lSearch
                    Exit Do
                End If
            Else
                Search = lSearch
                Exit Do
            End If
        Loop
    End If
    
    Select Case Right$(What, 1)
        Case "!", "@", "#", "$", "%", "&"
            What = Left$(What, Len(What) - 1)
    End Select
    
    If Len(Search) = 0 Or InStr(Search, What) = 0 Then                  'viene en blanco
        MyInstr = False
        Exit Function
    End If

    Ntokens = Tokenize04(Search, StringArray(), "+-*/.,&@#%[]{};!^:$()=\<> ", True)
        
    'validar que no existan caracteres basic
    Select Case Right$(What, 1)
        Case "!", "@", "#", "$", "%", "&"
            What = Left$(What, Len(What) - 1)
    End Select
    
    'ahora ciclar x todas las cadenas encontradas
    For k = 0 To UBound(StringArray())
        If LCase$(StringArray(k)) = LCase$(What) Then
            ret = True
            Exit For
        End If
    Next k
    
    MyInstr = ret
    
End Function


Public Function Tokenize04&(Expression$, ResultTokens$(), Delimiters$, Optional IncludeEmpty As Boolean)

' Tokenize02 by Donald, donald@xbeat.net
' modified by G.Beckmann, G.Beckmann@NikoCity.de
        
    Dim cExp&, ubExpr&
    Dim cDel&, ubDelim&
    Dim aExpr%(), aDelim%()
    Dim sa1 As SAFEARRAY1D, sa2 As SAFEARRAY1D
    Dim cTokens&, iPos&
 
    ubExpr = Len(Expression)
    ubDelim = Len(Delimiters)
    
    sa1.cbElements = 2
    sa1.cElements = ubExpr
    sa1.cDims = 1
    sa1.pvData = StrPtr(Expression)
    RtlMoveMemory ByVal VarPtrArray(aExpr), VarPtr(sa1), 4
    
    sa2.cbElements = 2
    sa2.cElements = ubDelim
    sa2.cDims = 1
    sa2.pvData = StrPtr(Delimiters)
    RtlMoveMemory ByVal VarPtrArray(aDelim), VarPtr(sa2), 4
  
    If IncludeEmpty Then
        ReDim Preserve ResultTokens(ubExpr)
    Else
        ReDim Preserve ResultTokens(ubExpr \ 2)
    End If
    
    ubDelim = ubDelim - 1
    For cExp = 0 To ubExpr - 1
'        DoEvents
        For cDel = 0 To ubDelim
'            DoEvents
            If aExpr(cExp) = aDelim(cDel) Then
                If cExp > iPos Then
                    ResultTokens(cTokens) = Mid$(Expression, iPos + 1, cExp - iPos)
                    cTokens = cTokens + 1
                ElseIf IncludeEmpty Then
                    ResultTokens(cTokens) = vbNullString
                    cTokens = cTokens + 1
                End If
                iPos = cExp + 1
                Exit For
            End If
        Next cDel
    Next cExp
  
    '/ remainder
    If (cExp > iPos) Or IncludeEmpty Then
        ResultTokens(cTokens) = Mid$(Expression, iPos + 1)
        cTokens = cTokens + 1
    End If
  
    '/ erase or shrink
    If cTokens = 0 Then
        Erase ResultTokens()
    Else
        ReDim Preserve ResultTokens(cTokens - 1)
    End If
  
    '/ return ubound
    Tokenize04 = cTokens - 1
    
    '/ tidy up
    RtlZeroMemory ByVal VarPtrArray(aExpr), 4
    RtlZeroMemory ByVal VarPtrArray(aDelim), 4
    
End Function



