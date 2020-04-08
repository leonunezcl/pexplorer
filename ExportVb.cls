VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportVb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eExportType
    HTMLFile
    TextFile
    WordDoc
End Enum

Private msBlueWords(97)  As String
Private colBlueWords As Collection
Private msTitle        As String
Private msBody         As String
Private msDelimiter    As String
Private msHTML         As String

Public Sub ExportCode(Component As VBComponent, ExportTo As eExportType)

    Select Case ExportTo
        Case HTMLFile
            LoadBlueWords
            ConvertToHtml Component
        Case TextFile
            ConvertToText
        Case WordDoc
            ConvertToDoc
    End Select
End Sub

Private Function ConvertToHtml(Component As VBIDE.VBComponent) As String
'    Dim msArrLines() As String
    Dim x         As Long                'Loop Counter
'    Dim lPosTic      As Long             'Position of Tick (') in line
'    Dim lPosQuote    As Long           'Position on Quote (") in line
'    Dim lTicLine     As Long            'Top line of continuous ticks before function
'    Dim lLocBracket  As Long         'Location of Bracket '('in line
'    Dim lLocSpace    As Long           'Location of Space ''in line
'    Dim sTrim        As String               'Trim$ of line
'    Dim sLines       As String              'Joined string of array msArrLines()
'    Dim sStart       As String              'First half of a quote/tick line
'    Dim sEnd         As String                'Send half of a quote/tick line
'    Dim bContinued   As Boolean          'Commented and continued line (has a '_')
'    Dim bPrevLineTic As Boolean        'Previous was commented line
'    Dim sHeader      As String             'Header String that contains links to fx's
'    Dim sFunction    As String           'Name of functions
    Dim sCurrentLine As String
    Dim sHTML        As String
    Dim lPlace       As String
    Dim colFunctions As Collection  'Collection of functions
    Dim sPart1 As String
    Dim sPart2 As String

    msTitle = "<HTML>" & vbCrLf & _
           "<META HTTP-EQUIV=" & Chr$(34) & "Content-Type" & Chr$(34) & " " & _
           "CONTENT=" & Chr$(34) & "text/html;charset=windows-1252" & Chr$(34) & ">" & vbCrLf & _
           "<TITLE>" & Component.Name & "</TITLE>" & vbCrLf & _
           "<A NAME='Top'></A>" & vbCrLf & _
           "<P><B><FONT SIZE=4 FACE='Times New Roman' COLOR='#000080'>" & _
           Component.Name & _
           "</FONT>" & _
           "<FONT SIZE=3 FACE='Times New Roman' COLOR='#666666'>" & _
           "&nbsp;&nbsp;&nbsp;&nbsp;</FONT></B>" & _
           "<BR>" & vbCrLf & _
           "<hr><BR>" & vbCrLf & "<BODY bgcolor='#FFFFFF'><Font Size = -1 Face = 'Courier New'>" & vbCrLf
    
    'Nothing passed
    If Component.CodeModule.CountOfLines = 0 Then Exit Function
    ReDim msArrLines(Component.CodeModule.CountOfLines - 1)
  
    sHTML = msTitle
For x = 1 To Component.CodeModule.CountOfLines
    sCurrentLine = Component.CodeModule.Lines(x, 1)
    If Len(sCurrentLine) = 0 Then
        sHTML = sHTML & "<br>"
    Else
        sCurrentLine = Replace$(sCurrentLine, "&", "&amp;")
        sCurrentLine = Replace$(sCurrentLine, "<", "&lt;")
        sCurrentLine = Replace$(sCurrentLine, ">", "&gt;")
        
        'Remove spaces at the end of any line
        sCurrentLine = RTrim$(sCurrentLine)
        
        If (Trim$(sCurrentLine) Like "[']*") Then
            sHTML = sHTML & "<Font Color='#007700'>" & Replace$(sCurrentLine, Chr$(32), "&nbsp;") & "</Font><br>" & vbCrLf
        Else
            If InStr(1, sCurrentLine, "'", vbBinaryCompare) Then
                lPlace = InStr(1, sCurrentLine, "'", vbBinaryCompare)
                sPart1 = Left$(sCurrentLine, lPlace - 1)
                sPart1 = BlueWords(sPart1)
                sPart2 = Right$(sCurrentLine, Len(sCurrentLine) - lPlace + 1)
                'sHTML = sHTML & Replace$(sPart1, Chr$(32), "&nbsp;")
                sHTML = sHTML & "<Font Color='#007700'>" & Replace$(sPart2, Chr$(32), "&nbsp;") & "</Font><br>" & vbCrLf
            Else
                sHTML = sHTML & BlueWords(sCurrentLine) & "<br>" & vbCrLf
                'sHTML = sHTML & Replace$(sCurrentLine, Chr$(32), "&nbsp;") & "<br>" & vbCrLf
            End If
        End If
    End If
Next
sHTML = sHTML & vbCrLf & "</BODY></HTML>"

    Dim iFile As Integer
    iFile = FreeFile
    Open "c:\test.htm" For Output As iFile
        Print #iFile, sHTML
    Close #iFile
    
End Function

Private Function ConvertToText()

End Function

Private Function ConvertToDoc()

End Function

Private Function BlueWords(ByVal sString As String) As String
    Dim lngI As Long
    Dim sBlue As String

    ' Add beggining and trailing spaces
    sString = "&nbsp;" & Replace$(sString, Chr$(32), "&nbsp;") & "&nbsp;"

    ' Find Blue work
    For lngI = 1 To colBlueWords.Count

        ' Get current blue
        sBlue = Trim$(colBlueWords(lngI))

        ' Can have values before
        If (sBlue Like "[#]*") Then

            ' Remove #
            sBlue = Mid$(sBlue, 2)

            ' Check the string is found
            If (sString Like "*[!0-z]" & sBlue & "[&]nbsp[;]*" Or sString Like "*[&]nbsp[;]" & sBlue & "[&]nbsp[;]*") Then

                ' Replace$ all with blue words
                sString = Replace$(sString, sBlue & "&nbsp;", "<Font Color='#000080'>" & sBlue & "</Font>&nbsp;")
            End If


        ' Can have values after
        ElseIf (sBlue Like "*[#]") Then

            ' Remove #
            sBlue = Mid$(sBlue, 1, (Len(sBlue) - 1))

            ' Check the string is found
            If (sString Like "*[&]nbsp[;]" & sBlue & "[!0-z]*" Or sString Like "*[&]nbsp[;]" & sBlue & "[&]nbsp[;]*") Then

                ' Replace$ all with blue words
                sString = Replace$(sString, "&nbsp;" & sBlue, "&nbsp;<Font Color='#000080'>" & sBlue & "</Font>")
            End If

        ' Stand alone (spaces)
        Else
            ' Replace$ all with blue words
            sString = Replace$(sString, "&nbsp;" & sBlue & "&nbsp;", "&nbsp;<Font Color='#000080'>" & sBlue & "</Font>&nbsp;")

        End If
    Next lngI

    ' Remove 1st & last space...
    sString = Mid$(sString, Len("&nbsp;") + 1, Len(sString) - (2& * Len("&nbsp;")))

    ' Return replaced...
    BlueWords = sString

End Function

Public Sub OutputHtmlToFile(Optional ByVal sFileName As String = "C:\TEMP\VbToHtml.html")
On Error Resume Next
    
    Dim iFile As Integer
    
    ' First things first
    'If (Len(msHTML) = 0) Then ConvertToHtml
    
    'Write events to event log
    iFile = FreeFile
    Open sFileName For Output As iFile
        Print #iFile, msHTML
    Close #iFile
End Sub

Private Sub LoadBlueWords()
    Set colBlueWords = New Collection
    colBlueWords.Add "And", "And"
    colBlueWords.Add "As", "As"
    colBlueWords.Add "Access", "Access"
    colBlueWords.Add "Between", "Between"
    colBlueWords.Add "Binary", "Binary"
    colBlueWords.Add "Boolean#", "Boolean#"
    colBlueWords.Add "By", "By"
    colBlueWords.Add "#ByRef", "#ByRef"
    colBlueWords.Add "Byte#", "Byte#"
    colBlueWords.Add "#ByVal", "#ByVal"
    colBlueWords.Add "Call", "Call"
    colBlueWords.Add "Case", "Case"
    colBlueWords.Add "Close", "Close"
    colBlueWords.Add "Compare", "Compare"
    colBlueWords.Add "Const", "Const"
    colBlueWords.Add "Control#", "Control#"
    colBlueWords.Add "Create", "Create"
    colBlueWords.Add "Currency#", "Currency#"
    colBlueWords.Add "Date#", "Date#"
    colBlueWords.Add "Debug.Print", "Debug.Print"
    colBlueWords.Add "Decimal#", "Decimal#"
    colBlueWords.Add "Declare", "Declare"
    colBlueWords.Add "DefBool", "DefBool"
    colBlueWords.Add "DefDate", "DefDate"
    colBlueWords.Add "DefDbl", "DefDbl"
    colBlueWords.Add "DefInt", "DefInt"
    colBlueWords.Add "DefLng", "DefLng"
    colBlueWords.Add "DefObj", "DefObj"
    colBlueWords.Add "DefStr", "DefStr"
    colBlueWords.Add "DefVar", "DefVar"
    colBlueWords.Add "Dim", "Dim"
    colBlueWords.Add "Do", "Do"
    colBlueWords.Add "Double#", "Double#"
    colBlueWords.Add "Each", "Each"
    colBlueWords.Add "Else", "Else"
    colBlueWords.Add "ElseIf", "ElseIf"
    colBlueWords.Add "End", "End"
    colBlueWords.Add "Erase", "Erase"
    colBlueWords.Add "Error", "Error"
    colBlueWords.Add "Execute", "Execute"
    colBlueWords.Add "Exists", "Exists"
    colBlueWords.Add "Exit", "Exit"
    colBlueWords.Add "Explicit", "Explicit"
    colBlueWords.Add "False", "False"
    colBlueWords.Add "For", "For"
    colBlueWords.Add "Form#", "Form#"
    colBlueWords.Add "From", "From"
    colBlueWords.Add "Function", "Function"
    colBlueWords.Add "Get", "Get"
    colBlueWords.Add "GoTo", "GoTo"
    colBlueWords.Add "If", "If"
    colBlueWords.Add "In", "In"
    colBlueWords.Add "Input", "Input"
    colBlueWords.Add "Integer#", "Integer#"
    colBlueWords.Add "LBound#", "LBound#"
    colBlueWords.Add "Let", "Let"
    colBlueWords.Add "Line", "Line"
    colBlueWords.Add "Like", "Like"
    colBlueWords.Add "Long#", "Long#"
    colBlueWords.Add "Loop", "Loop"
    colBlueWords.Add "New", "New"
    colBlueWords.Add "Next", "Next"
    colBlueWords.Add "#Not", "#Not"
    colBlueWords.Add "Null", "Null"
    colBlueWords.Add "On", "On"
    colBlueWords.Add "Or", "Or"
    colBlueWords.Add "Open", "Open"
    colBlueWords.Add "Option", "Option"
    colBlueWords.Add "#Optional", "#Optional"
    colBlueWords.Add "Output", "Output"
    colBlueWords.Add "Preserve", "Preserve"
    colBlueWords.Add "Primary", "Primary"
    colBlueWords.Add "Print", "Print"
    colBlueWords.Add "Private", "Private"
    colBlueWords.Add "Procedure", "Procedure"
    colBlueWords.Add "Property", "Property"
    colBlueWords.Add "Public", "Public"
    colBlueWords.Add "Read", "Read"
    colBlueWords.Add "ReDim", "ReDim"
    colBlueWords.Add "Recordset#", "Recordset#"
    colBlueWords.Add "Reset", "Reset"
    colBlueWords.Add "Resume", "Resume"
    colBlueWords.Add "Select", "Select"
    colBlueWords.Add "Set", "Set"
    colBlueWords.Add "Single#", "Single#"
    colBlueWords.Add "String#", "String#"
    colBlueWords.Add "Sub", "Sub"
    colBlueWords.Add "To", "To"
    colBlueWords.Add "TableDef#", "TableDef#"
    colBlueWords.Add "Then", "Then"
    colBlueWords.Add "True", "True"
    colBlueWords.Add "Until", "Until"
    colBlueWords.Add "UBound#", "UBound#"
    colBlueWords.Add "Variant#", "Variant#"
    colBlueWords.Add "Wend", "Wend"
    colBlueWords.Add "While", "While"
    colBlueWords.Add "With", "With"
    colBlueWords.Add "WithEvents", "WithEvents"
End Sub




