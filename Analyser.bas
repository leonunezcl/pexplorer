Attribute VB_Name = "Analyser"
Option Explicit
' File analyser.

' Procedure type
Public Const PT_DECLARE As Integer = 0
Public Const PT_PROPERTY As Integer = 1
Public Const PT_SUB As Integer = 2
Public Const PT_FUNCTION As Integer = 3
Public Const PT_API As Integer = 4

' Procedure scopy
Public Const PS_NONE As Integer = 0
Public Const PS_PUBLIC As Integer = 1
Public Const PS_PRIVATE As Integer = 2

' Module type
Public Const MT_FORM As Integer = 0
Public Const MT_MODULE As Integer = 1
Public Const MT_CLASS As Integer = 2
Public Const MT_CONTROL As Integer = 3
Public Const MT_PROPERTY As Integer = 4
Public Const MT_DOCUMENT As Integer = 5
Public Const MT_DESIGNER As Integer = 6

' .Selected values:
' vbUnchecked = 0   Unchecked (default).
' vbChecked   = 1   Checked.
' vbGrayed    = 2   Grayed.

Public Type ProcedureState    ' Declaration and procedures
   Name As String             ' Short name for display
   Syntax As String           ' Full line of code of the procedure definition (Not for declaration)
   IndexName As String        ' Short name for index listing
   Type As Integer            ' See PT_* constants
   Scope As Integer           ' See PS_* constants
   Static As Boolean          ' Is it defined STATIC?
   Lines As Integer           ' Number of code lines
   code() As String           ' The code... (just text - one based array)
   Selected As Integer        ' 0, 1 or 2 (see CheckBox.Value)
   ListIndex As Integer       ' Index pointer in outline for this item
   Predef As Boolean          ' declaracion de ambito
End Type
Public Type CONTROLSTATE
   Name As String             ' Name given to identify the control
   Type As String             ' Type of control, eg label, textbox
   Library As String          ' OCX or DLL etc from where it comes from, eg VB, MSComDlg
   Elements As Integer        ' Number of elements (count) in collection, mostly 1, never zero
End Type
Public Type ModuleState
   PathFile As String         ' Filename with path
   file As String             ' Filename without path
   Name As String             ' Form/Module/Class name (eg frmMain, modSupport)
   Type As Integer            ' File type (Form, Module or Class) (see MT_* constants)
   Selected As Integer        ' 0, 1 or 2 (see CheckBox.Value)
   ListIndex As Integer       ' Index pointer in outline for this item
   BinaryFile As String       ' El archivo con el recurso binario
   IconData As String         ' The icon picture data (use LoadIcon() to to obtain picture data for display)
   CtrlElements As Long       ' Total number of elements of all controls
   CtrlCount As Integer       ' Number of controls in form (zero for modules and classes)
   CtrlSelect As Integer      ' 0, 1 or 2 (see CheckBox.Value)
   CtrlLIndex As Integer      ' Index pointer in outline for this item
   Control() As CONTROLSTATE  ' The controls... (one based array)
   ProcCount As Integer       ' Count declaration as a procedure too
   Proc() As ProcedureState   ' The procedures... (one based array)
   SelCount As Integer        ' Selected children count
   ChildCount As Integer      ' Total number of children (includes procedures, control section and declaration section)
   Exposed As Boolean
   Implements As String
End Type

Public Mdl() As ModuleState   ' The information holder for above types (one based array)
Public MdCount As Integer     ' Number of Mdl() elements (makes it easy to create more of them)
Public MdSelected As Integer  ' Number of selected elements

Public PrCount As Integer     ' Total number of procedures (and declarations)
Public PrSelected As Integer  ' Number of selected procedures

'Public CtCount As Integer     ' Total number of control groups (forms)

' -------------------------------------------------------------

Private Type PrcExtractState
   Name As String             ' Short name for display
   IndexName As String        ' Short name for index listing
   Type As Integer            ' See PT_* constants
   Scope As Integer           ' See PS_* constants
   Static As Boolean          ' Is it defined STATIC?
   Predef As Boolean
End Type

' -------------------------------------------------------------

' Project file (VBP) information
Public Type FileState
   file As String
   Name As String
End Type
Public Type ProjectState        ' Information for Project Information page
   Type As String
   Loaded As Boolean             ' True if project file was partially or wholly loaded/analysed.
   Bit32 As Boolean              ' True if 32 bit (Win 95 / NT)specific information is retrieved
   Bit16 As Boolean              ' True if 16 bit (Win 3.x) specific information is retrieved
   StartupForm As String
   StartupFile As String
   FormCount As Integer
   Form() As FileState
   ModuleCount As Integer
   Module() As FileState
   ClassCount As Integer
   Class() As FileState
   ControlCount As Integer          ' VB5 related
   UControl() As FileState          ' VB5 related
   PropertyCount As Integer         ' VB5 related
   PropertyPg() As FileState        ' VB5 related
   DocumentCount As Integer         ' VB5 related
   UDocument() As FileState         ' VB5 related
   RelatedCount As Integer          ' VB5 related
   RelatedDoc() As FileState        ' VB5 related
   DesignerCount As Integer
   UDesigner() As FileState
   ReferenceCount As Integer
   Reference() As FileState
   ObjectCount As Integer
   Object() As FileState         ' '.Object().Name' not in use
   IconForm As String
   IconPoint As Integer          ' Array Pointer into Mdl()
   HelpFile As String
   HelpContextID As String
   Title As String
   ExeName32 As String
   ExeName16 As String
   Path32 As String
   Path16 As String
   Command32 As String
   Command16 As String
   Name As String
   StartMode As String           ' 0 - Standalone,  1-OLE Server
   Description As String
   OLEServer32 As String         ' 'CompatibleExe32=""'
   OLEServer16 As String         ' 'CompatibleExe=""'
   CompileArg As String          ' 'CondComp=""'
   MajorVersion As Integer
   MinorVersion As Integer
   RevisionVersion As Integer
   AutoVersion As Boolean
   Comments As String
   CompanyName As String
   FileDescription As String
   Copyright As String
   TradeMarks As String
   ProductName As String
   Resource32 As String
   Resource16 As String
End Type
Public glbProjectState As ProjectState
' --------------------------------------------------------------
' This is it. Give'm the filename (with optional file type [see MT_* constants])
' and it returns element number if all successfull.
' If things go wrong, you get '-1' back.
'
Function AnalyseFile(sFile As String, Optional nType) As Integer

    Const DECLARE_OFF As Integer = 0
    Const DECLARE_WATCH As Integer = 1
    Const DECLARE_SEPERATOR As Integer = 2

    Dim i As Integer, n As Integer, nHandle As Integer, nBuffer As Integer, nPrPoint As Integer, nLimit As Integer
    'Dim nFileSize As Long
    Dim Id As Integer
    Dim bFileOpen As Boolean, bCodeSection As Boolean, bFound As Boolean, bBuffer As Boolean
    Dim sString As String, sUpper As String, sBuffer() As String
    Dim ProcInfo As PrcExtractState
    Dim CtrlInfo As CONTROLSTATE
    Dim StartProc As Boolean
    Dim LineaPaso As String
    Dim sBinExt As String
    
    bFileOpen = False

    If Not InDevelopmentMode Then On Error GoTo AF_ErrorHandler

    If Not MyFuncFiles.FileExist(sFile) Then
        AnalyseFile = -1
        Exit Function                 ' No file, no analyse - bye.
    End If

    bCodeSection = False
    bBuffer = True                   ' Start buffering code
    nBuffer = 0                      ' Buffer line (for easy redimension the sBuffer() array)
    Erase sBuffer                    ' Code buffer

    'Open the file !!!
    nHandle = FreeFile
    Open sFile For Input Access Read Shared As #nHandle
        bFileOpen = True
       
        MdCount = MdCount + 1
        ReDim Preserve Mdl(MdCount)

        AnalyseFile = MdCount

        Mdl(MdCount).PathFile = sFile
        Mdl(MdCount).file = MyFuncFiles.ExtractFileName(sFile)
        Mdl(MdCount).Name = ""                       ' Get that from 'Attribute VB_Name = "frm..."'
        Mdl(MdCount).Selected = vbUnchecked
        Mdl(MdCount).ListIndex = -1

        Mdl(MdCount).IconData = ""
        Mdl(MdCount).BinaryFile = ""
        
        Mdl(MdCount).CtrlElements = 0
        Mdl(MdCount).CtrlCount = 0
        Mdl(MdCount).CtrlSelect = vbUnchecked
        Mdl(MdCount).CtrlLIndex = -1
        
        Mdl(MdCount).ProcCount = 0
        
        Mdl(MdCount).SelCount = 0
        Mdl(MdCount).ChildCount = 0
        
        If IsMissing(nType) Then
            Select Case UCase$(MyFuncFiles.ExtractFileExt(sFile))
                Case "FRM"
                    Mdl(MdCount).Type = MT_FORM
                Case "BAS"
                    Mdl(MdCount).Type = MT_MODULE
                Case "CLS"
                    Mdl(MdCount).Type = MT_CLASS
                Case "CTL"
                    Mdl(MdCount).Type = MT_CONTROL
                Case "PAG"
                    Mdl(MdCount).Type = MT_PROPERTY
                Case "DOB"
                    Mdl(MdCount).Type = MT_DOCUMENT
                Case "DSR"
                    Mdl(MdCount).Type = MT_DESIGNER
                Case Else
                    Mdl(MdCount).Type = -1
            End Select
        Else
            Mdl(MdCount).Type = CInt(nType)
        End If
        
        nLimit = 1
        
        Do While Not EOF(nHandle)  ' Loop until end of file.
            Line Input #nHandle, sString
            sUpper = UCase$(Trim$(sString))
                                    
            If MatchString(sUpper, "ATTRIBUTE ") Then
                ' Internal section almost over, ready for the code section
                If MatchString(sUpper, "ATTRIBUTE VB_NAME") Then
                    ' Obtain the assigned name of the form/module/class
                    'Attribute VB_Name = "Form1"
                    n = InStr(sString, "=")
                    If n > 0 Then
                        sString = Trim$(Mid$(sString, n + 1))
                        Mdl(MdCount).Name = StripQuotes(sString)
                    End If
                ElseIf MatchString(sUpper, "ATTRIBUTE VB_EXPOSED = TRUE") Then
                    Mdl(MdCount).Exposed = True
                End If

                bCodeSection = True
                GoTo EndOfFileLoop
            End If

            If bCodeSection Then
                If IsProcedure(sUpper) Then
                    nLimit = 0
                    If nBuffer > 0 Then
                        'agregar lo leido para la seccion general
                        If Mdl(MdCount).ProcCount = 0 Or Not StartProc Then
                            Mdl(MdCount).ProcCount = 1
                            ReDim Preserve Mdl(MdCount).Proc(1 To Mdl(MdCount).ProcCount)
                            PrCount = PrCount + 1

                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Name = "(Declarations)"
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Syntax = ""
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).IndexName = "(Declarations)"
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Type = PT_DECLARE
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Scope = PS_NONE
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Static = False
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = 0
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Selected = vbUnchecked
                            Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).ListIndex = -1

                            nPrPoint = Mdl(MdCount).ProcCount
                            StartProc = True
                            
                            ' Add the code...
                            For i = 1 To UBound(sBuffer)
                               Mdl(MdCount).Proc(nPrPoint).Lines = Mdl(MdCount).Proc(nPrPoint).Lines + 1
                               ReDim Preserve Mdl(MdCount).Proc(nPrPoint).code(1 To Mdl(MdCount).Proc(nPrPoint).Lines)
                               Mdl(MdCount).Proc(nPrPoint).code(Mdl(MdCount).Proc(nPrPoint).Lines) = sBuffer(i)
                            Next
                            
                            nBuffer = 0
                            bBuffer = True
                        End If
                    End If
                ElseIf MatchString(sUpper, "END SUB") Or _
                    MatchString(sUpper, "END FUNCTION") Or _
                    MatchString(sUpper, "END PROPERTY") Then
                    
                    'agregar el fin de procedimiento
                    nBuffer = nBuffer + 1
                    ReDim Preserve sBuffer(nBuffer)
                    sBuffer(nBuffer) = sString
                                        
                    'agregar los procedimientos
                    Mdl(MdCount).ProcCount = Mdl(MdCount).ProcCount + 1
                    ReDim Preserve Mdl(MdCount).Proc(1 To Mdl(MdCount).ProcCount)
                    PrCount = PrCount + 1

                    'extraer el nombre del procedimiento
                    LineaPaso = ""
                    For i = LBound(sBuffer) To UBound(sBuffer)
                        'verificar el tipo de linea para procesar
                        sString = sBuffer(i)
                        If Len(sString) > 0 And Left$(sString, 1) <> "'" Then
                            If Right$(sString, 1) = "_" Then
                                LineaPaso = LineaPaso & Trim$(Left$(sString, Len(sString) - 1))
                            Else
                                LineaPaso = LineaPaso & Trim$(sString)
                                sString = LineaPaso
                                Exit For
                            End If
                        End If
                    Next i
                        
                    'extraer informacion del procedimiento
                    ProcInfo = ExtractProcedure(sString)
    
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Name = ProcInfo.Name
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Syntax = Trim$(sString)
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).IndexName = ProcInfo.IndexName
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Type = ProcInfo.Type
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Scope = ProcInfo.Scope
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Static = ProcInfo.Static
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = 0
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Selected = vbUnchecked
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).ListIndex = -1
                    Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Predef = ProcInfo.Predef
                    
                    ' Empty out the buffer. Add some code already
                    For i = 1 To UBound(sBuffer)
                        Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines + 1
                        ReDim Preserve Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(1 To Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines)
                        Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines) = sBuffer(i)
                    Next
                    
                    nBuffer = 0
                    nLimit = 1
                    bBuffer = True
                    
                    GoTo EndOfFileLoop
                End If

                If bBuffer Then
                    ' Do not add spaces in top of declaration section
                    If MdCount = 0 And nBuffer = 0 Then GoTo EndOfFileLoop
                                
                    ' Buffer the code...
                    If Not StartProc Or nLimit = 0 Then
                        nBuffer = nBuffer + 1
                        If nBuffer = 1 Then
                            ReDim sBuffer(1)
                        Else
                            ReDim Preserve sBuffer(nBuffer)
                        End If
                        sBuffer(nBuffer) = RTrim$(sString)
                    Else
                        If nLimit = 1 Then
                            If Len(sString) > 0 Then
                                nBuffer = nBuffer + 1
                                If nBuffer = 1 Then
                                    ReDim sBuffer(1)
                                Else
                                    ReDim Preserve sBuffer(nBuffer)
                                End If
                                sBuffer(nBuffer) = RTrim$(sString)
                            End If
                        End If
                    End If
                Else
                    ' Add code to procedure
                    'Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines + 1
                    'ReDim Preserve Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(1 To Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines)
                    'Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines) = RTrim$(sString)
                End If

            GoTo EndOfFileLoop
        End If
        ' Internal section -------------------------------------------------------------------------------------------------------------------

      If MatchString(sUpper, "BEGIN ") Then
         ' Form Control
         CtrlInfo = ExtractControl(sString)

         If Mdl(MdCount).CtrlCount > 0 Then
            ' Find control name
            bFound = False
            For n = 1 To Mdl(MdCount).CtrlCount
               If Mdl(MdCount).Control(n).Name = CtrlInfo.Name And _
                  Mdl(MdCount).Control(n).Type = CtrlInfo.Type And _
                  Mdl(MdCount).Control(n).Library = CtrlInfo.Library Then

                  Mdl(MdCount).Control(n).Elements = Mdl(MdCount).Control(n).Elements + 1
                  Mdl(MdCount).CtrlElements = Mdl(MdCount).CtrlElements + 1
                  bFound = True
                  Exit For
               End If
            Next
            If bFound Then GoTo EndOfFileLoop
         End If

         Mdl(MdCount).CtrlCount = Mdl(MdCount).CtrlCount + 1
         ReDim Preserve Mdl(MdCount).Control(1 To Mdl(MdCount).CtrlCount)
         Mdl(MdCount).Control(Mdl(MdCount).CtrlCount).Name = CtrlInfo.Name
         Mdl(MdCount).Control(Mdl(MdCount).CtrlCount).Type = CtrlInfo.Type
         Mdl(MdCount).Control(Mdl(MdCount).CtrlCount).Library = CtrlInfo.Library
         Mdl(MdCount).Control(Mdl(MdCount).CtrlCount).Elements = 1
         Mdl(MdCount).CtrlElements = Mdl(MdCount).CtrlElements + 1

      ElseIf MatchString(sUpper, "ICON ") Then
         ' Form Icon

         If EmptyString(Mdl(MdCount).IconData) Then
            '  Icon = "FormFile.frx":0000
            '       ^               ^
            n = InStr(sString, "=")
            If n > 0 Then
               sUpper = Trim$(Mid$(sString, n + 1))
               sString = Trim$(Mid$(sString, n + 1))
               ExtractIcon sString, MdCount
               n = InStr(sUpper, ":")
               sUpper = MyFuncFiles.AttachPath(StripQuotes(Left(sUpper, n - 1)), MyFuncFiles.ExtractPath(Mdl(MdCount).PathFile))
               Mdl(MdCount).BinaryFile = sUpper
            End If
         End If
      ElseIf MatchString(sUpper, "OLEOBJECTBLOB ") Or MatchString(sUpper, "TOOLBOXBITMAP ") Then
        If EmptyString(Mdl(MdCount).IconData) Then
            n = InStr(sString, "=")
            If n > 0 Then
               sUpper = Trim$(Mid$(sString, n + 1))
               sString = Trim$(Mid$(sString, n + 1))
               ExtractIcon sString, MdCount
               n = InStr(sUpper, ":")
               sUpper = MyFuncFiles.AttachPath(StripQuotes(Left(sUpper, n - 1)), MyFuncFiles.ExtractPath(Mdl(MdCount).PathFile))
               Mdl(MdCount).BinaryFile = sUpper
            End If
        End If
      End If

EndOfFileLoop:    ' -----------------------------------------------------------------------------------------------------------------------
   Loop

   If nBuffer > 0 Then
      ' Some code in the buffer... Save it
      If Not InDevelopmentMode Then On Error Resume Next

      If Mdl(MdCount).ProcCount = 0 Then
         ' No procedures defined, must be code from the declaration section
         Mdl(MdCount).ProcCount = 1
         ReDim Preserve Mdl(MdCount).Proc(1 To Mdl(MdCount).ProcCount)
         PrCount = PrCount + 1
            
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Name = "(Declarations)"
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Syntax = ""
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).IndexName = "(Declarations)"
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Type = PT_DECLARE
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Scope = PS_NONE
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Static = False
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = 0
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Selected = vbUnchecked
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).ListIndex = -1
      End If

' BUG IN SYSTEM !!!!!
'
' When a form is loaded and later on a class or modules with only declarations a bug will
' appear. (two code lines from here) - Subscript outta range !!!
' Somehow a dummy entry (the one above) is not created !!!!

      ' Add the code...
      For i = 1 To nBuffer
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines = Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines + 1
         ReDim Preserve Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(1 To Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines)
         Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).code(Mdl(MdCount).Proc(Mdl(MdCount).ProcCount).Lines) = sBuffer(i)
      Next
      If Not InDevelopmentMode Then On Error GoTo AF_ErrorHandler
   
   End If

   Close #nHandle

   Mdl(MdCount).ChildCount = Mdl(MdCount).ProcCount
   If Mdl(MdCount).CtrlCount > 0 Then
      Mdl(MdCount).ChildCount = Mdl(MdCount).ChildCount + 1
   End If

   Exit Function

AF_ErrorHandler:
   MsgBox "Problem analysing program file. Information is partly restricted"

   On Error Resume Next

   If MdCount > 0 Then
      If MdCount = UBound(Mdl) Then
         Mdl(MdCount).ChildCount = Mdl(MdCount).ProcCount
         If Mdl(MdCount).CtrlCount > 0 Then
            Mdl(MdCount).ChildCount = Mdl(MdCount).ChildCount + 1
         End If
      End If
   End If

   If bFileOpen Then Close #nHandle

End Function

' Syntax in file: Begin VB.Menu mnuExit             --> make into: mnuExit, Menu, (VB)
'                 Begin ComctlLib.ImageList Images  -->            Images, ImageList, (ComctlLib)
'                                                                  [Name], [Type], [(Library)]
'
Private Function ExtractControl(ByVal sString As String) As CONTROLSTATE
   Dim nMark As Integer
   Dim sText As String, sName As String
   sText = Trim$(sString)

   nMark = InStr(sText, " ")
   If nMark = 0 Then
      ExtractControl.Name = ""
      ExtractControl.Type = ""
      ExtractControl.Library = ""
      Exit Function
   End If

   sText = Trim$(Mid$(sText, nMark + 1))

   nMark = InStr(sText, " ")
   If nMark = 0 Then
      nMark = InStr(sText, ".")
      If nMark = 0 Then
         ExtractControl.Name = sText
         ExtractControl.Type = "[Unknow object]"
         ExtractControl.Library = ""
      Else
         ExtractControl.Name = "[Unnamed control]"
         ExtractControl.Type = Mid$(sText, nMark + 1)
         ExtractControl.Library = Left(sText, nMark - 1)
      End If
      Exit Function
   End If

   sName = Trim$(Mid$(sText, nMark + 1))
   sText = Trim$(Left$(sText, nMark - 1))

   nMark = InStr(sText, ".")
   If nMark = 0 Then
      ExtractControl.Name = sName
      ExtractControl.Type = sText
      ExtractControl.Library = ""
   Else
      ExtractControl.Name = sName
      ExtractControl.Type = Mid$(sText, nMark + 1)
      ExtractControl.Library = Left(sText, nMark - 1)
   End If

End Function

' Return True if current line is a procedure
Function IsProcedure(sUpper As String) As Boolean
'   Dim sUpper As String
   Dim bValid As Boolean

'   sUpper = UCase$(Trim(sString))
   bValid = False

   If MatchString(sUpper, "PRIVATE ") Then                  ' Speed up scan by minimising "If" statements
      If MatchString(sUpper, "PRIVATE SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE PROPERTY ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE STATIC SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE STATIC FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE STATIC PROPERTY ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PRIVATE DECLARE SUB ") Then
         'bValid = True
      ElseIf MatchString(sUpper, "PRIVATE DECLARE FUNCTION ") Then
         'bValid = True
      End If

   ElseIf MatchString(sUpper, "PUBLIC ") Then
      If MatchString(sUpper, "PUBLIC SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC PROPERTY ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC STATIC SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC STATIC FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC STATIC PROPERTY ") Then
         bValid = True
      ElseIf MatchString(sUpper, "PUBLIC DECLARE SUB ") Then
         'bValid = True
      ElseIf MatchString(sUpper, "PUBLIC DECLARE FUNCTION ") Then
         'bValid = True
      End If

   ElseIf MatchString(sUpper, "STATIC ") Then
      If MatchString(sUpper, "STATIC SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "STATIC FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "STATIC PROPERTY ") Then
         bValid = True
      End If
   ElseIf MatchString(sUpper, "FRIEND ") Then
      If MatchString(sUpper, "FRIEND SUB ") Then
         bValid = True
      ElseIf MatchString(sUpper, "FRIEND FUNCTION ") Then
         bValid = True
      ElseIf MatchString(sUpper, "FRIEND PROPERTY ") Then
         bValid = True
      End If
   ElseIf MatchString(sUpper, "SUB ") Then
      bValid = True
   ElseIf MatchString(sUpper, "FUNCTION ") Then
      bValid = True
   ElseIf MatchString(sUpper, "PROPERTY ") Then
      bValid = True
   ElseIf MatchString(sUpper, "DECLARE SUB ") Then
       'bValid = True
   ElseIf MatchString(sUpper, "DECLARE FUNCTION ") Then
       'bValid = True
   End If

   IsProcedure = bValid

End Function

Private Function ExtractProcedure(ByVal sString As String) As PrcExtractState
   Dim nMark As Integer, nType As Integer, nScope As Integer
   Dim sName As String, sIndexName As String, sUpper As String, _
       sPrefix As String, sSuffix As String
   Dim bStatic As Boolean
   Dim bPredef As Boolean
   
   sName = "-Unknow procedure declaration-"
   sIndexName = "-Unknow procedure-"
   nType = -1
   nScope = -1
   bStatic = False
   
   sString = Trim$(sString)
   sUpper = UCase$(sString)
                         
   If MatchString(sUpper, "PRIVATE ") Then                  ' Speed up scan by minimising "If" statements
      If MatchString(sUpper, "PRIVATE SUB ") Then
         nType = PT_SUB
         nScope = PS_PRIVATE
         bStatic = False
         sPrefix = Left$(sString, 12)
         nMark = 13
      ElseIf MatchString(sUpper, "PRIVATE FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PRIVATE
         bStatic = False
         sPrefix = "Private Function "
         nMark = 18
      ElseIf MatchString(sUpper, "PRIVATE PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PRIVATE
         bStatic = False
         sPrefix = Left$(sString, 21)
         nMark = 22
      ElseIf MatchString(sUpper, "PRIVATE STATIC SUB ") Then
         nType = PT_SUB
         nScope = PS_PRIVATE
         bStatic = True
         sPrefix = "Private Static Sub "
         nMark = 20
      ElseIf MatchString(sUpper, "PRIVATE STATIC FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PRIVATE
         bStatic = True
         sPrefix = "Private Static Function "
         nMark = 25
      ElseIf MatchString(sUpper, "PRIVATE STATIC PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PRIVATE
         bStatic = True
         sPrefix = Left$(sString, 28)
         nMark = 29
      ElseIf MatchString(sUpper, "PRIVATE DECLARE SUB ") Then
         nType = PT_API
         nScope = PS_PRIVATE
         bStatic = True
         sPrefix = "Private Declare Sub "
         nMark = 21
      ElseIf MatchString(sUpper, "PRIVATE DECLARE FUNCTION ") Then
         nType = PT_API
         nScope = PS_PRIVATE
         bStatic = True
         sPrefix = "Private Declare Function "
         nMark = 26
      End If

   ElseIf MatchString(sUpper, "PUBLIC ") Then
      If MatchString(sUpper, "PUBLIC SUB ") Then
         nType = PT_SUB
         nScope = PS_PUBLIC
         bStatic = False
         sPrefix = "Public Sub "
         nMark = 12
      ElseIf MatchString(sUpper, "PUBLIC FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PUBLIC
         bStatic = False
         sPrefix = Left$(sString, 16)
         nMark = 17
      ElseIf MatchString(sUpper, "PUBLIC PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PUBLIC
         bStatic = False
         sPrefix = Left$(sString, 20)
         nMark = 21
      ElseIf MatchString(sUpper, "PUBLIC STATIC SUB ") Then
         nType = PT_SUB
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = "Public Static Sub "
         nMark = 19
      ElseIf MatchString(sUpper, "PUBLIC STATIC FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = "Public Static Function "
         nMark = 24
      ElseIf MatchString(sUpper, "PUBLIC STATIC PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 27)
         nMark = 28
      ElseIf MatchString(sUpper, "PUBLIC DECLARE SUB ") Then
         nType = PT_API
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = "Public Declare Sub "
         nMark = 20
      ElseIf MatchString(sUpper, "PUBLIC DECLARE FUNCTION ") Then
         nType = PT_API
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = "Public Declare Function "
         nMark = 25
      End If

   ElseIf MatchString(sUpper, "STATIC ") Then
      If MatchString(sUpper, "STATIC SUB ") Then
         nType = PT_SUB
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 11)
         nMark = 12
      ElseIf MatchString(sUpper, "STATIC FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 16)
         nMark = 17
      ElseIf MatchString(sUpper, "STATIC PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 20)
         nMark = 21
      End If
   ElseIf MatchString(sUpper, "FRIEND ") Then
      If MatchString(sUpper, "FRIEND SUB ") Then
         nType = PT_SUB
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 11)
         nMark = 12
      ElseIf MatchString(sUpper, "FRIEND FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 16)
         nMark = 17
      ElseIf MatchString(sUpper, "FRIEND PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 20)
         nMark = 21
      ElseIf MatchString(sUpper, "FRIEND STATIC PROPERTY ") Then
         nType = PT_PROPERTY
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 27)
         nMark = 28
      ElseIf MatchString(sUpper, "FRIEND STATIC FUNCTION ") Then
         nType = PT_FUNCTION
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 23)
         nMark = 24
      ElseIf MatchString(sUpper, "FRIEND STATIC SUB ") Then
         nType = PT_SUB
         nScope = PS_PUBLIC
         bStatic = True
         sPrefix = Left$(sString, 18)
         nMark = 19
      End If
   ElseIf MatchString(sUpper, "SUB ") Then
      nType = PT_SUB
      nScope = PS_PUBLIC
      bStatic = True
      sPrefix = "Sub "
      nMark = 5
      bPredef = True
   ElseIf MatchString(sUpper, "FUNCTION ") Then
      nType = PT_FUNCTION
      nScope = PS_PUBLIC
      bStatic = False
      sPrefix = "Function "
      nMark = 10
      bPredef = True
   ElseIf MatchString(sUpper, "PROPERTY ") Then
      nType = PT_PROPERTY
      nScope = PS_PUBLIC
      bStatic = False
      sPrefix = Left$(sString, 13)
      nMark = 14
      bPredef = True
   ElseIf MatchString(sUpper, "DECLARE SUB ") Then
      nType = PT_API
      nScope = PS_PUBLIC
      bStatic = True
      sPrefix = "Declare Sub "
      nMark = 13
      bPredef = True
   ElseIf MatchString(sUpper, "DECLARE FUNCTION ") Then
      nType = PT_API
      nScope = PS_PUBLIC
      bStatic = True
      sPrefix = "Declare Function "
      nMark = 18
      bPredef = True
   End If
   
   If nMark > 0 Then
      If nType <> PT_API Then
          sString = Trim$(Mid$(sString, nMark))
    
          ' Chop of the parameters
          nMark = InStr(sString, "(")
          If nMark > 0 Then
             If Mid(sString, nMark + 1) = ")" Then
                sSuffix = "()"
             Else
                sSuffix = "(...)"
             End If
             sString = Trim$(Left$(sString, nMark - 1))
          Else
             sSuffix = ""
          End If
    
          sName = sPrefix & sString & sSuffix
      Else
          sString = Trim$(Mid$(sString, nMark))
          nMark = InStr(sString, " ")
          If nMark > 0 Then
             sString = Trim$(Left$(sString, nMark - 1))
          End If
      End If
      
      sIndexName = sString
   End If

   ExtractProcedure.Name = sName
   ExtractProcedure.IndexName = sIndexName
   ExtractProcedure.Type = nType
   ExtractProcedure.Scope = nScope
   ExtractProcedure.Static = bStatic
   ExtractProcedure.Predef = bPredef
End Function


Function ProcType(nMIndex As Integer, nPIndex As Integer) As String
   Dim sString As String

   Select Case Mdl(nMIndex).Proc(nPIndex).Type
   Case PT_DECLARE
      sString = "Declarations"
   Case PT_PROPERTY
      sString = "Property"
   Case PT_SUB
      sString = "Sub"
   Case PT_FUNCTION
      sString = "Function"
   Case Else
      sString = ""
   End Select

   Select Case Mdl(nMIndex).Proc(nPIndex).Scope
   Case PS_PUBLIC
      sString = IIf(EmptyString(sString), "", sString & ", ") & "Public"
   Case PS_PRIVATE
      sString = IIf(EmptyString(sString), "", sString & ", ") & "Private"
   End Select

   If Mdl(nMIndex).Proc(nPIndex).Static Then
      sString = IIf(EmptyString(sString), "", sString & ", ") & "Static"
   End If

   ProcType = sString

End Function

'  Icon = "FormFile.frx":0000
'       ^               ^
'         |-----------------| = Parameter
'
' There are 2 ways that graphics is stored. This 12 bytes header and a 28 bytes.
'
Private Sub ExtractIcon(sString As String, nIndex As Integer)
   Dim n As Integer, nHandle As Integer
   Dim nOffset As Long, nFileSize As Long, nSize As Long
   Dim sFile As String, sData As String, sBytes As String
   Dim bFileOpen As Boolean

   bFileOpen = False

   On Error GoTo EI_ErrorHandler

   n = InStr(sString, ":")
   If n < 1 Then Exit Sub
        
   sFile = MyFuncFiles.AttachPath(StripQuotes(Left(sString, n - 1)), MyFuncFiles.ExtractPath(Mdl(nIndex).PathFile))
   sString = "&H" & Trim$(Mid$(sString, n + 1))
   nOffset = Val(sString) + 1

   If Not MyFuncFiles.FileExist(sFile) Then Exit Sub

   nHandle = FreeFile
   Open sFile For Binary Access Read Shared As #nHandle
   bFileOpen = True
   nFileSize = LOF(nHandle)

   If (nOffset + 12) > nFileSize Then GoTo EI_ErrorHandler

   ' Get the header...
   Seek #nHandle, nOffset
   sData = Mid$(Input(12, #nHandle), 9, 4)

   ' Byte 9 to 12 (long) contains data size
   sBytes = "&H" & Right("00" & Hex(Asc(Mid$(sData, 4, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 3, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 2, 1))), 2) & _
                   Right("00" & Hex(Asc(Mid$(sData, 1, 1))), 2)
   nSize = Val(sBytes)

   If (nOffset + 11 + nSize) > nFileSize Then GoTo EI_ErrorHandler

   ' Get the data (position: nOffset + 13 - Already in position)
   Mdl(nIndex).IconData = Input(nSize, #nHandle)

   ' That's it, the icon data is obtained
   Close #nHandle
   bFileOpen = False
   Exit Sub

EI_ErrorHandler:
   If bFileOpen Then Close #nHandle
End Sub

' Loads the icon into frmMain.picImage holder.
' Parameter: the Mdl() element number to load
' Returns: True if successfull.
'
Function LoadIcon(nIndex As Integer) As Boolean
   Dim sTempFile As String
   Dim nHandle As Integer
   Dim bFileOpen As Boolean

   bFileOpen = False

   On Error GoTo LI_ErrorHandler

   If nIndex = -1 Then
      ' use main form icon
      Main.picImage.Picture = Main.Icon

   Else
      If EmptyString(Mdl(nIndex).IconData) Then GoTo LI_ErrorHandler

      sTempFile = MyFuncFiles.MakeTempFile
      If EmptyString(sTempFile) Then GoTo LI_ErrorHandler
      If MyFuncFiles.FileExist(sTempFile) Then Kill sTempFile

      ' Save image data to temp file, then load into PictureBox. Delete file when finished
      nHandle = FreeFile
      Open sTempFile For Binary Access Write Lock Write As #nHandle
      bFileOpen = True
      Put #nHandle, 1, Mdl(nIndex).IconData
      Close nHandle
      bFileOpen = False

      Main.picImage.Picture = LoadPicture(sTempFile)
   
      On Error Resume Next
      Kill sTempFile
   End If

   LoadIcon = True
   Exit Function

LI_ErrorHandler:
   If bFileOpen Then Close #nHandle
   LoadIcon = False
End Function

Function StripQuotes(ByVal sString As String) As String
   If Asc(Left(sString, 1)) = 34 And Asc(Right(sString, 1)) = 34 Then
      StripQuotes = Mid$(sString, 2, Len(sString) - 2)
   Else
      StripQuotes = sString
   End If
End Function

' Gimme a VBP file and if all goes ok I will return the extracted information
' Analyse the files prior the VBP !!
'
Public Function AnalyseVBP(sVBPFile As String) As ProjectState
   Dim i As Integer, n As Integer, nHandle As Integer
   Dim nFileSize As Long
   Dim bFileOpen As Boolean, bFirstFile As Boolean
   Dim sString As String, sKey As String, sValue As String, sFile As String, sName As String, sPath As String
   'Dim Pj As ProjectState

   AnalyseVBP.Loaded = False

   If Not InDevelopmentMode Then
      ' Intercept error only in run-time mode (in development mode gimme a VB error box so I debug it)
      On Error GoTo ProjectScanError
   End If

   ' Only VBP files can be analysed
   If UCase$(MyFuncFiles.ExtractFileExt(sVBPFile)) <> "VBP" Then GoTo ProjectScanAbort

   ' Ofcause the file must exist...
   If Not MyFuncFiles.FileExist(sVBPFile) Then GoTo ProjectScanAbort

   sPath = MyFuncFiles.FixPath(MyFuncFiles.ExtractPath(sVBPFile))
   
   nHandle = FreeFile
   Open sVBPFile For Input Access Read Shared As #nHandle
   bFileOpen = True

   nFileSize = LOF(nHandle)
   bFirstFile = True

   ' Initialise some values... (not really required, but less chance of errors)
   AnalyseVBP.Type = ""
   AnalyseVBP.Bit32 = False
   AnalyseVBP.Bit16 = False

   AnalyseVBP.StartupForm = ""
   AnalyseVBP.StartupFile = ""
   AnalyseVBP.FormCount = 0
   AnalyseVBP.ModuleCount = 0
   AnalyseVBP.ClassCount = 0
   AnalyseVBP.ControlCount = 0      ' VB 5 related
   AnalyseVBP.PropertyCount = 0     ' VB 5 related
   AnalyseVBP.DocumentCount = 0     ' VB 5 related
   AnalyseVBP.RelatedCount = 0      ' VB 5 related
   AnalyseVBP.ReferenceCount = 0
   AnalyseVBP.ObjectCount = 0
   AnalyseVBP.IconForm = ""
   AnalyseVBP.IconPoint = -1
   AnalyseVBP.HelpFile = ""
   AnalyseVBP.HelpContextID = ""
   AnalyseVBP.Title = ""
   AnalyseVBP.ExeName32 = ""
   AnalyseVBP.ExeName16 = ""
   AnalyseVBP.Path32 = ""
   AnalyseVBP.Path16 = ""
   AnalyseVBP.Command32 = ""
   AnalyseVBP.Command16 = ""
   AnalyseVBP.Name = ""
   AnalyseVBP.StartMode = ""
   AnalyseVBP.Description = ""
   AnalyseVBP.OLEServer32 = ""
   AnalyseVBP.OLEServer16 = ""
   AnalyseVBP.CompileArg = ""
   AnalyseVBP.MajorVersion = 0
   AnalyseVBP.MinorVersion = 0
   AnalyseVBP.RevisionVersion = 0
   AnalyseVBP.AutoVersion = False
   AnalyseVBP.Comments = ""
   AnalyseVBP.CompanyName = ""
   AnalyseVBP.FileDescription = ""
   AnalyseVBP.Copyright = ""
   AnalyseVBP.TradeMarks = ""
   AnalyseVBP.ProductName = ""
   AnalyseVBP.Resource32 = ""
   AnalyseVBP.Resource16 = ""

   AnalyseVBP.Loaded = True

   Do While Not EOF(nHandle)  ' Loop until end of file.
      Line Input #nHandle, sString
      
      ' The project file line exist of '[Key] = [Value]'
      ' Use the '=' (equal sign) to separate the key and the value.
      n = InStr(sString, "=")
      If n > 0 Then
         sKey = UCase$(Trim$(Left$(sString, n - 1)))
         sValue = Trim$(Mid$(sString, n + 1))
      Else
         GoTo ProjectScanLoop
      End If

      ' Find out what I got and what to do with it...
      Select Case sKey
      Case "TYPE"
        If Not EmptyString(sValue) Then
            AnalyseVBP.Type = sValue
        End If
      Case "STARTUP"
        n = InStr(sValue, "=")
        AnalyseVBP.StartupForm = StripQuotes(Trim$(Mid$(sValue, n + 1)))
      Case "FORM"   'FORMULARIOS
         If Not EmptyString(sValue) Then
            AnalyseVBP.FormCount = AnalyseVBP.FormCount + 1
            ReDim Preserve AnalyseVBP.Form(1 To AnalyseVBP.FormCount)
            AnalyseVBP.Form(AnalyseVBP.FormCount).file = sValue
            ' Use Mdl() to obtain the name
            If MdCount > 0 Then
               sValue = UCase$(MyFuncFiles.ExtractFileName(sValue))
               For i = 1 To MdCount
                  If UCase$(Mdl(i).file) = sValue Then
                     AnalyseVBP.Form(AnalyseVBP.FormCount).Name = Mdl(i).Name
                     Exit For
                  End If
               Next
            End If
            'If bFirstFile Then
            '   AnalyseVBP.StartupForm = AnalyseVBP.Form(AnalyseVBP.FormCount).Name
            '   AnalyseVBP.StartupFile = AnalyseVBP.Form(AnalyseVBP.FormCount).file
            '   bFirstFile = False
            'End If
         End If
      
      Case "MODULE" 'MODULOS .BAS
         n = InStr(sValue, ";")
         If n > 0 Then
            AnalyseVBP.ModuleCount = AnalyseVBP.ModuleCount + 1
            ReDim Preserve AnalyseVBP.Module(1 To AnalyseVBP.ModuleCount)
            AnalyseVBP.Module(AnalyseVBP.ModuleCount).file = Trim$(Mid$(sValue, n + 1))
            AnalyseVBP.Module(AnalyseVBP.ModuleCount).Name = Trim$(Left(sValue, n - 1))
            'If bFirstFile Then
            '   AnalyseVBP.StartupForm = "Sub Main() en " & AnalyseVBP.Module(AnalyseVBP.ModuleCount).Name
            '   AnalyseVBP.StartupFile = AnalyseVBP.Module(AnalyseVBP.ModuleCount).file
            '   bFirstFile = False
            'End If
         End If
      
      Case "CLASS"  'CLASES
         n = InStr(sValue, ";")
         If n > 0 Then
            AnalyseVBP.ClassCount = AnalyseVBP.ClassCount + 1
            ReDim Preserve AnalyseVBP.Class(1 To AnalyseVBP.ClassCount)
            AnalyseVBP.Class(AnalyseVBP.ClassCount).file = Trim$(Mid$(sValue, n + 1))
            AnalyseVBP.Class(AnalyseVBP.ClassCount).Name = Trim$(Left(sValue, n - 1))
         End If
      
      Case "USERCONTROL"    'CONTROLES DE USUARIO
         If Not EmptyString(sValue) Then
            AnalyseVBP.ControlCount = AnalyseVBP.ControlCount + 1
            ReDim Preserve AnalyseVBP.UControl(1 To AnalyseVBP.ControlCount)
            AnalyseVBP.UControl(AnalyseVBP.ControlCount).file = sValue
            ' Use Mdl() to obtain the name
            If MdCount > 0 Then
               sValue = UCase$(MyFuncFiles.ExtractFileName(sValue))
               For i = 1 To MdCount
                  If UCase$(Mdl(i).file) = sValue Then
                     AnalyseVBP.UControl(AnalyseVBP.ControlCount).Name = Mdl(i).Name
                     Exit For
                  End If
               Next
            End If
'            If bFirstFile Then
'               AnalyseVBP.StartupForm = AnalyseVBP.UControl(AnalyseVBP.ControlCount).Name
'               AnalyseVBP.StartupFile = AnalyseVBP.UControl(AnalyseVBP.ControlCount).File
'               bFirstFile = False
'            End If
         End If

      Case "PROPERTYPAGE"   'PAGINAS DE PROPIEDADES
         If Not EmptyString(sValue) Then
            AnalyseVBP.PropertyCount = AnalyseVBP.PropertyCount + 1
            ReDim Preserve AnalyseVBP.PropertyPg(1 To AnalyseVBP.PropertyCount)
            AnalyseVBP.PropertyPg(AnalyseVBP.PropertyCount).file = sValue
            ' Use Mdl() to obtain the name
            If MdCount > 0 Then
               sValue = UCase$(MyFuncFiles.ExtractFileName(sValue))
               For i = 1 To MdCount
                  If UCase$(Mdl(i).file) = sValue Then
                     AnalyseVBP.PropertyPg(AnalyseVBP.PropertyCount).Name = Mdl(i).Name
                     Exit For
                  End If
               Next
            End If
         End If

      Case "USERDOCUMENT"   'DOCUMENTOS DE USUARIO
         If Not EmptyString(sValue) Then
            AnalyseVBP.DocumentCount = AnalyseVBP.DocumentCount + 1
            ReDim Preserve AnalyseVBP.UDocument(1 To AnalyseVBP.DocumentCount)
            AnalyseVBP.UDocument(AnalyseVBP.DocumentCount).file = sValue
            ' Use Mdl() to obtain the name
            If MdCount > 0 Then
               sValue = UCase$(MyFuncFiles.ExtractFileName(sValue))
               For i = 1 To MdCount
                  If UCase$(Mdl(i).file) = sValue Then
                     AnalyseVBP.UDocument(AnalyseVBP.DocumentCount).Name = Mdl(i).Name
                     Exit For
                  End If
               Next
            End If
         End If

      Case "RELATEDDOC" 'DOCUMENTOS RELACIONADOS
         If Not EmptyString(sValue) Then
            AnalyseVBP.RelatedCount = AnalyseVBP.RelatedCount + 1
            ReDim Preserve AnalyseVBP.RelatedDoc(1 To AnalyseVBP.RelatedCount)
            AnalyseVBP.RelatedDoc(AnalyseVBP.RelatedCount).file = sValue
            AnalyseVBP.RelatedDoc(AnalyseVBP.RelatedCount).Name = ""
         End If
      Case "DESIGNER"   'DISEÑADORES
         If Not EmptyString(sValue) Then
            AnalyseVBP.DesignerCount = AnalyseVBP.DesignerCount + 1
            ReDim Preserve AnalyseVBP.UDesigner(1 To AnalyseVBP.DesignerCount)
            AnalyseVBP.UDesigner(AnalyseVBP.DesignerCount).file = sValue
            ' Use Mdl() to obtain the name
            If MdCount > 0 Then
               sValue = UCase$(MyFuncFiles.ExtractFileName(sValue))
               For i = 1 To MdCount
                  If UCase$(Mdl(i).file) = sValue Then
                     AnalyseVBP.UDesigner(AnalyseVBP.DesignerCount).Name = Mdl(i).Name
                     Exit For
                  End If
               Next
            End If
         End If
      Case "REFERENCE"
        Call CargaReferencias(sString)
        
         i = 0
         Do While True
            n = InStr(sValue, "#")
            If n = 0 Then Exit Do
            i = i + 1
            sValue = Mid$(sValue, n + 1)
            If i = 3 Then
               n = InStr(sValue, "#")
               If n = 0 Then
                  sFile = Trim(sValue)
                  sName = ""
               Else
                  sFile = Trim$(Left$(sValue, n - 1))
                  sName = Trim$(Mid$(sValue, n + 1))
               End If
               Exit Do
            End If
         Loop
         If Not EmptyString(sFile) Then
            AnalyseVBP.ReferenceCount = AnalyseVBP.ReferenceCount + 1
            ReDim Preserve AnalyseVBP.Reference(1 To AnalyseVBP.ReferenceCount)
            AnalyseVBP.Reference(AnalyseVBP.ReferenceCount).file = sFile
            AnalyseVBP.Reference(AnalyseVBP.ReferenceCount).Name = sName
         End If
      
      Case "OBJECT"
        Call CargaComponentes(sString)
        
         n = InStr(sValue, ";")
         If n > 0 Then
            AnalyseVBP.ObjectCount = AnalyseVBP.ObjectCount + 1
            ReDim Preserve AnalyseVBP.Object(1 To AnalyseVBP.ObjectCount)
            AnalyseVBP.Object(AnalyseVBP.ObjectCount).file = Trim$(Mid$(sValue, n + 1))
            AnalyseVBP.Object(AnalyseVBP.ObjectCount).Name = ""
         End If

      Case "ICONFORM"
         AnalyseVBP.IconForm = StripQuotes(sValue)
         AnalyseVBP.IconPoint = -1
         If MdCount > 0 Then
            For i = 1 To MdCount
               If Mdl(i).Name = AnalyseVBP.IconForm Then
                  AnalyseVBP.IconPoint = i
                  Exit For
               End If
            Next
         End If
      Case "HELPFILE"
         AnalyseVBP.HelpFile = StripQuotes(sValue)
      Case "TITLE"
         AnalyseVBP.Title = StripQuotes(sValue)
      Case "EXENAME32"
         AnalyseVBP.ExeName32 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.ExeName32) Then AnalyseVBP.Bit32 = True
      Case "EXENAME"
         AnalyseVBP.ExeName16 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.ExeName16) Then AnalyseVBP.Bit16 = True
      Case "PATH32"
         AnalyseVBP.Path32 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Path32) Then AnalyseVBP.Bit32 = True
      Case "PATH"
         AnalyseVBP.Path16 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Path16) Then AnalyseVBP.Bit16 = True
      Case "COMMAND32"
         AnalyseVBP.Command32 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Command32) Then AnalyseVBP.Bit32 = True
      Case "COMMAND"
         AnalyseVBP.Command16 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Command16) Then AnalyseVBP.Bit16 = True
      Case "NAME"
         AnalyseVBP.Name = StripQuotes(sValue)
      Case "HELPCONTEXTID"
         AnalyseVBP.HelpContextID = StripQuotes(sValue)
      Case "STARTMODE"
         Select Case Val(sValue)
         Case 0
            AnalyseVBP.StartMode = "Standalone"
         Case 0
            AnalyseVBP.StartMode = "OLE Server"
         End Select
      Case "DESCRIPTION"
         AnalyseVBP.Description = StripQuotes(sValue)
      Case "COMPATIBLEEXE32"
         AnalyseVBP.OLEServer32 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.OLEServer32) Then AnalyseVBP.Bit32 = True
      Case "COMPATIBLEEXE"
         AnalyseVBP.OLEServer16 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.OLEServer16) Then AnalyseVBP.Bit16 = True
      Case "CONDCOMP"
         AnalyseVBP.CompileArg = StripQuotes(sValue)

      Case "RESFILE32"
         AnalyseVBP.Resource32 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Resource32) Then AnalyseVBP.Bit32 = True
      Case "RESFILE16"
         AnalyseVBP.Resource16 = StripQuotes(sValue)
         If Not EmptyString(AnalyseVBP.Resource16) Then AnalyseVBP.Bit16 = True
      
      Case "MAJORVER"
         AnalyseVBP.MajorVersion = Val(sValue)
      Case "MINORVER"
         AnalyseVBP.MinorVersion = Val(sValue)
      Case "REVISIONVER"
         AnalyseVBP.RevisionVersion = Val(sValue)
      Case "AUTOINCREMENTVAR"
         AnalyseVBP.AutoVersion = (sValue = "1")

      Case "VERSIONCOMMENTS"
         AnalyseVBP.Comments = StripQuotes(sValue)
      Case "VERSIONCOMPANYNAME"
         AnalyseVBP.CompanyName = StripQuotes(sValue)
      Case "VERSIONFILEDESCRIPTION"
         AnalyseVBP.FileDescription = StripQuotes(sValue)
      Case "VERSIONLEGALCOPYRIGHT"
         AnalyseVBP.Copyright = StripQuotes(sValue)
      Case "VERSIONLEGALTRADEMARKS"
         AnalyseVBP.TradeMarks = StripQuotes(sValue)
      Case "VERSIONPRODUCTNAME"
         AnalyseVBP.ProductName = StripQuotes(sValue)
'      Case ""
      End Select

ProjectScanLoop:

   Loop
   
   On Error Resume Next
   Close #nHandle
   
   Exit Function

ProjectScanError:
   If bFileOpen Then Close #nHandle
   MsgBox "Problema mientras se busca archivo de proyecto." & vbCrLf & "Error #" & Err.Number & ": " & Err.Description, vbCritical

ProjectScanAbort:
   On Error Resume Next
End Function
