Attribute VB_Name = "modDynHighlight"
Option Explicit

Public Enum hColEnum
  hRed = 0
  hGreen = 1
  hBlue = 2
End Enum

Public Enum hSectionEnum
  hKeyWords = 0
  hFuncSubs = 1
  hComments = 2
  hStringDelim = 3
  hText = 4
  hOperators = 5
  hScope = 6
End Enum

Public Type hCommentType
  Comment As String
  SingleLine As Boolean
End Type

Public KeyWords() As String
Public KeyWordsCount As Integer
Public KeyWordsCol As OLE_COLOR

Public Comments() As hCommentType
Public CommentsCount As Integer
Public CommentsMultiStart As String
Public CommentsMultiEnd As String
Public CommentsCol As OLE_COLOR

Public FuncSubs() As String
Public FuncSubsCount As Integer
Public FuncSubsCol As OLE_COLOR

Public Operators() As String
Public OperatorsCount As Integer
Public OperatorsCol As OLE_COLOR

Public Scope() As String
Public ScopeCount As Integer
Public ScopeCol As OLE_COLOR

Public StringDelim As String
Public StringDelimCol As OLE_COLOR

Public BlockDelim As String

Public TextCol As OLE_COLOR

'Sets what is used to separate keywords/funcs/subs/etc.
Private Delim As String
'Sets the text back to normal after the color's been
'altered...
Private RTFNormal As String
'Creates a delimited list based on the arrays...
Private KeyWordLst As String, FuncSubLst As String
Private CommentLst As String, OperatorLst As String
Private ScopeLst As String

'Sets the switches for the different programming
'sections...
Private KeyWordSwitch As String
Private FuncSubSwitch As String
Private CommentSwitch As String
Private StringDelimSwitch As String
Private TextSwitch As String
Private OperatorSwitch As String
Private ScopeSwitch As String


'Loads a file...
Public Function LoadFile(sFileName As String) As String
Dim f As Byte
Dim sFile As String

'Closes all open files...
Close
'Gets an open file...
f = FreeFile
'Opens the file and puts the content into the
'sFile variable...
Open sFileName For Input As #f
  sFile = Input(LOF(f), f)
Close #f

'Returns the contents of the file...
LoadFile = sFile
End Function

'This loads a syntax file into memory...
'At the end the following are modified:
'Comments() = Comment array
'Keywords() = Keywords array
'FuncSubs() = Function/Sub array
'CommentsMultiStart = Comment w/ the syntax for
'  for starting a multi-line comment like in C/C++
'CommentsMultiEnd = Sets the end for a multi-line comment
'
'All variables w/ the ending "Count" contain the total
'# of items for the associated array so we know
'how many to go through in later code...
'
'Any variables ending w/ "Col" contain the colors for
'that area...the colors aren't defined in the syntax
'file -- you have to do that on your own...
Public Function LoadSyntaxFile(sFileName As String) As Boolean

    Dim f As Byte
    Dim isMult As Boolean, isSingle As Boolean
    Dim cLn As String, cCat As String
    Dim comLft As String, comRgt As String
    Dim comMultStart As String, comMultEnd As String
    
    'Clears up all mem taken up by the arrays...
    'We can do this b/c we're loading all the info.
    'into the array...
    'We also do this in case the syntax file must be
    'loaded a couple times during program execution...
    'If the syntax is modified at runtime, then this will
    'also help!!  (c;
    Erase KeyWords, Comments, FuncSubs, Operators
    'Free up some memory...
    KeyWordLst = ""
    CommentLst = ""
    FuncSubLst = ""
    OperatorLst = ""
    
    'Closes all open files...
    Close
    'Gets an open file...
    f = FreeFile
    'Opens the file and puts the content into the
    'sFile variable...
    On Error GoTo ErrExit
    Open sFileName For Input As #f
      Do While Not EOF(f)
        Line Input #f, cLn
          
          'Make it lower case for comparison purposes
          'later on...
          cLn = LCase$(cLn)
          
          'Sets the category if there's [] somewhere in the
          'current line...
          If Left$(cLn, 1) = "[" And Right$(cLn, 1) = "]" Then cCat = cLn
          
          'Sets info. based on the category in the syntax
          'file...
          If (LCase$(cLn) <> LCase$(cCat)) And (Left$(cLn, 1) <> ";") And (cLn <> "") Then
            Select Case LCase$(cCat)
                Case "[keywords]"
                    'Increments the keyword counter by 1...
                    KeyWordsCount = KeyWordsCount + 1
                    'ReArranges the array to reflect the new
                    'keyword added to the list...
                    ReDim Preserve KeyWords(KeyWordsCount) As String
                    'Sets the added item to the current line...
                    KeyWords(KeyWordsCount) = Trim$(cLn)
               
                Case "[operators]"
                    'Increments the operator counter by 1...
                    OperatorsCount = OperatorsCount + 1
                    'ReArranges the array to reflect the new
                    'operator added to the list...
                    ReDim Preserve Operators(OperatorsCount) As String
                    'Sets the added line to the array...
                    Operators(OperatorsCount) = Trim$(cLn)
            
                Case "[scope]"
                    'Increments the operator counter by 1...
                    ScopeCount = ScopeCount + 1
                    'ReArranges the array to reflect the new
                    'operator added to the list...
                    ReDim Preserve Scope(ScopeCount) As String
                    'Sets the added line to the array...
                    Scope(ScopeCount) = Trim$(cLn)
               
                Case "[func subs]"
                    'Increment func sub counter...
                    FuncSubsCount = FuncSubsCount + 1
                    'ReArrange func sub array...
                    ReDim Preserve FuncSubs(FuncSubsCount) As String
                    'Put in the info...
                    FuncSubs(FuncSubsCount) = Trim$(cLn)
               
                  Case "[comments]"
                     'This gets the text to the Left$/Right$ of a
                     'string so we can see if it's MultiLine or
                     'SingeLine
                     comRgt = GetTokenTxt(cLn, "=", False)
                     comLft = GetTokenTxt(cLn, "=", True)
                     
                     'Sets the comment info. based on if it's
                     'MultiLine or SingleLine or both or neither
                     Select Case LCase$(comLft)
                        Case "multiline"
                         isMult = True
                        Case "singleline"
                         isSingle = True
                        
                        'Sets the MultiLine variables
                        'if they've been defined in the syntax
                        'file...
                        Case "multilinestart"
                          CommentsMultiStart = comRgt
                        Case "multilineend"
                          CommentsMultiEnd = comRgt
                      End Select
                     
                     If isSingle = True Then
                       'Increments the comment counter by 1...
                       CommentsCount = CommentsCount + 1
                       'ReArrange the comment array...
                       ReDim Preserve Comments(CommentsCount) As hCommentType
                       'Adds the current line and comment info.
                       'to the array...
                       Comments(CommentsCount).Comment = comRgt
                       Comments(CommentsCount).SingleLine = isSingle
                     End If
                     
                     'Make sure they're not set on the next
                     'cycle through....
                     isMult = False
                     isSingle = False
                  Case "[string delim]"
                     If Trim$(cLn) <> "" Then StringDelim = cLn
                     
                  Case "[block]"
                     If Trim$(cLn) <> "" Then BlockDelim = Trim$(cLn)
                  
                End Select
          End If
      Loop
    Close #f
    
    'Returns if the file was loaded properly or not...
    LoadSyntaxFile = True
    Exit Function

'There was a problem, so return false and exit...
ErrExit:
  'There was an error so return false so we know that
  'it wasn't loaded correctly....
  LoadSyntaxFile = False
  Exit Function
End Function
Public Function GetTokenTxt(s As String, t As String, isBefore As Boolean) As String
Dim c As Integer
Dim cRet As String

c = InStr(1, s, t)
If c <> 0 Then 'found token
   If isBefore = True Then
        'Return the text before the token
        GetTokenTxt = Mid(s, 1, c - 1)
      Else
        'Return the text after the token
        GetTokenTxt = Mid(s, c + 1)
   End If
End If
End Function

Public Function InsertHeader(sRTF As String) As String
Dim curr As String
Dim after As String, bef As String
Dim tblPos As Long, secTblPos As Long

'{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Tahoma;}{\f3\froman\fprq2 Times New Roman;}}
'{\colortbl\red0\green0\blue0;\red0\green0\blue160;\red0\green128\blue0;}
'\deflang1033\pard\plain\f3\fs20
'\par \plain\f3\fs20\cf1 blue\plain\f3\fs20\cf2 green\plain\f3\fs20
'\par }
'\cf1 = blue
'\cf2 = green

'Insert color RTF info.
tblPos = InStr(1, sRTF, "{\colortbl")
If tblPos <> 0 Then
   secTblPos = InStr(tblPos, sRTF, "}")
   curr = Mid(sRTF, tblPos, secTblPos - tblPos + 1)
   bef = Mid(sRTF, 1, tblPos - 1)
   after = Mid(sRTF, secTblPos + 1)
   
   sRTF = bef & GetTotalColorTbl & after
End If

'\cf1 = KeyWords
'\cf2 = FuncSubs
'\cf3 = Comment
'\cf4 = StringDelim
'\cf5 = Text
InsertHeader = sRTF
End Function

'Returns the entire color table header info. for RTF...
Public Function GetTotalColorTbl() As String
Dim retRTF As String
retRTF = "{\colortbl\red0\green0\blue0;" & _
         GetColorTblInfo(hKeyWords) & GetColorTblInfo(hFuncSubs) & _
         GetColorTblInfo(hComments) & GetColorTblInfo(hStringDelim) & _
         GetColorTblInfo(hText) & GetColorTblInfo(hOperators) & _
         GetColorTblInfo(hScope) & _
         "}"

'Sets the color switch info. for each section...
KeyWordSwitch = "\b\cf1"
FuncSubSwitch = "\cf2"
CommentSwitch = "\cf3"
StringDelimSwitch = "\cf4"
TextSwitch = "\cf5"
OperatorSwitch = "\cf6"
ScopeSwitch = "\cf7"

GetTotalColorTbl = retRTF
End Function

'Returns the colortable RTF switch info. for each
'specific item...
Public Function GetColorTblInfo(section As hSectionEnum) As String
Dim retRTF As String
Dim cGreen As Integer, cRed As Integer, cBlue As Integer
Dim cCol As Single

'Sets the variable based on the section desired...
Select Case section
  Case 0 'KeyWords
    cCol = KeyWordsCol
  Case 1 'FuncSubs
    cCol = FuncSubsCol
  Case 2 'Comments
    cCol = CommentsCol
  Case 3 'StringDelim
    cCol = StringDelimCol
  Case 4 'Text
    cCol = TextCol
  Case 5 'Operators
    cCol = OperatorsCol
  Case 6 'Scope
    cCol = ScopeCol
End Select

cRed = GetRGBValue(cCol, hRed)
cGreen = GetRGBValue(cCol, hGreen)
cBlue = GetRGBValue(cCol, hBlue)

retRTF = Trim("\red" & cRed & "\green" & cGreen & "\blue" & cBlue & ";")

GetColorTblInfo = retRTF
End Function

'Returns an RGB value for a color specified as a
'# returned from a common dialog or &HFF, etc.
Public Function GetRGBValue(Col As Single, CurrType As hColEnum) As Integer
Dim color As Single
Dim Red As Byte
Dim Green As Byte
Dim Blue As Byte
Dim ret As Integer

On Error Resume Next
color = Col
Red = color Mod 256
Green = (color And &HFF00FF00) / 256
Blue = Int(color / 65536)

On Error Resume Next
Select Case CurrType
  Case 0 'Red
    ret = Red
  Case 1 'Green
    ret = Green
  Case 2 'Blue
    ret = Blue
End Select

GetRGBValue = ret
End Function

'Loads all the arrays into a delimited list all at once
'so you don't have to keep calling all 3 at once in
'other functions or subs...
Public Function LoadTotalArrayLst() As Boolean
  'The KeyWordLst, FuncSubLst, etc. variables are all
  'private variables for the module...
  On Error GoTo ErrExit
  KeyWordLst = LoadArrayStrLst(hKeyWords)
  FuncSubLst = LoadArrayStrLst(hFuncSubs)
  CommentLst = LoadArrayStrLst(hComments)
  OperatorLst = LoadArrayStrLst(hOperators)
  ScopeLst = LoadArrayStrLst(hScope)
  
  LoadTotalArrayLst = True
  Exit Function
  
ErrExit:
    LoadTotalArrayLst = False
    Exit Function
End Function
Public Function LoadArrayStrLst(section As hSectionEnum) As String
Dim retStr As String, i As Integer

'Selects which area you want and returns a list of
'keywords from the area like:
'
'And/True/False/Or/As/Integer/etc.
'
'This is good so we can use a simple InStr to locate
'the item in the list...speeds things up...
'The Delim variable is a private module-level variable
'that contains the separator...in the example above
'it's the "/"
  Select Case section
    Case 0 'KeyWords
      For i = 1 To KeyWordsCount
        retStr = retStr & KeyWords(i) & Delim
      Next i
    Case 1 'FuncSubs
      For i = 1 To FuncSubsCount
        retStr = retStr & FuncSubs(i) & Delim
      Next i
    Case 2 'Comments
      For i = 1 To CommentsCount
        retStr = retStr & Comments(i).Comment & Delim
      Next i
    Case 5 'Operators
      For i = 1 To OperatorsCount
        retStr = retStr & Operators(i) & Delim
      Next i
    Case 6 'Scope
      For i = 1 To ScopeCount
        retStr = retStr & Scope(i) & Delim
      Next i
  End Select
  
  LoadArrayStrLst = Delim & retStr
End Function

Private Function GetNumTokensInStr(s As String, Optional Token As String = ",") As Long
Dim i As Long
Dim TokenCount As Long

     TokenCount = 0
     'Count the # of ,'s are in the list...
     For i = 1 To Len(s)
         i = InStr(i, s, Token)
         'If there's no more, then exit!!
         If i = 0 Then
            Exit For
         End If
         
         'Add 1 to list...
         TokenCount = TokenCount + 1
     Next i
     
GetNumTokensInStr = TokenCount
End Function

Public Function ClearRTFTabs(cWrd As String) As String
Const RTFTAB = "\tab"
Dim i As Long
Dim cAft As String, cBef As String

  Do
    i = InStr(1, cWrd, RTFTAB)
    If i <> 0 Then
       cBef = Mid(1, 1, i)
       cAft = Mid(cWrd, i + RTFTAB)
       
       cWrd = cBef & cAft
    End If
  Loop While i <> 0
  
  ClearRTFTabs = cWrd
End Function
'\cf1 = KeyWords
'\cf2 = FuncSubs
'\cf3 = CommentCol
'\cf4 = StringDelim
'\cf5 = TextCol

Public Function QuickCheck(cChr As String, totalChrLen As Long, isNewLine As Boolean) As String
Static isComment As Boolean, isOperator As Boolean
Static Op As String, isString As Boolean

Dim isBlock As Boolean
Dim cLnTrm As String, i As Long, cOld As Long
Dim cBef As String, cCurrLine As String
Dim cBlock As String
Dim cIsFoundKeyWord As Long, cIsFoundFuncSub As Long, cIsFoundComment As Long
Dim cIsFoundOperator As Long, cIsFoundString As Long
Dim cIsFoundScope As Long

'This checks to see if the user put in a comment that's
'to the side instead of on a line by itself...
'If they did, then we need to ignore any words
'after they're commented
If isComment = True And isNewLine = False Then
   cChr = CommentSwitch & cChr & RTFNormal & TextSwitch
   QuickCheck = cChr
   Exit Function
End If

'If it's a comment that's been set previously and it's
'gone on to a new line, then we need to make sure it's
'set back to normal so highlighting can continue
'normally...
If isComment = True And isNewLine = True Then
   isComment = False
End If


''We need to see if it's the string delimiter and if
''it is, then we don't want any keywords that might
''be in the string to be colored incorrectly...
'If Left(Trim(cChr), Len(StringDelim)) = StringDelim Then
'   cChr = StringDelimSwitch & cChr '& Mid(cChr, 1, Len(cChr) - Len(StringDelim))
'   isString = True
'End If

'If it's a block statement like a ; in JavaScript, then
'we need to get the word by itself and add in the
'block later on after it's all been evaluated...
If Right(cChr, Len(BlockDelim)) = BlockDelim Then
   cChr = Mid(cChr, 1, Len(cChr) - Len(BlockDelim))
   isBlock = True
End If

 'See if it's a word, not just a space...
  If Trim(cChr) <> "" Then
     
     'Checks to see if the word passed is in the list
     cIsFoundKeyWord = InStr(1, KeyWordLst, Delim & LCase(Trim(cChr)) & Delim)
     cIsFoundFuncSub = InStr(1, FuncSubLst, Delim & LCase(Trim(cChr)) & Delim)
     cIsFoundComment = InStr(1, CommentLst, Delim & Left(LCase(Trim(cChr)), 1) & Delim)
     cIsFoundOperator = InStr(1, OperatorLst, Delim & Trim(cChr) & Delim)
     'cIsFoundScope = InStr(1, ScopeLst, Delim & Trim(cChr) & Delim)
     cIsFoundString = InStr(1, Trim(cChr), StringDelim)
     
     If Mid(Trim(cChr), 1, 2) <> StringDelim & StringDelim Then
     If InStr(cIsFoundString + 1, Trim(cChr), StringDelim) = 0 Then
     If isString = True And cIsFoundString = 0 Then 'InStr(cIsFoundString + 1, Trim(cChr), StringDelim) = 0 Then
              QuickCheck = cChr
              Exit Function
            ElseIf isString = True And cIsFoundString <> 0 Then
              isString = False
            ElseIf isString = False And cIsFoundString <> 0 Then
              isString = True
            ElseIf isString = False And cIsFoundString = 0 Then
              isString = False
     End If
     End If
     End If
     
     
     
     
     
     '     isString = True
     '     QuickCheck = cChr
     '     Exit Function
     '   ElseIf isString = True Then
     '     If InStr(1, Trim(cChr), StringDelim) <> 0 Then
     '        isString = False
     '        QuickCheck = cChr
     '        Exit Function
     '     End If
     'End If
     
     'If cIsFoundScope <> 0 Then
     '   cChr = ScopeSwitch & cChr & RTFNormal '& TextSwitch
     'End If
     
     If cIsFoundOperator <> 0 Then
        cChr = OperatorSwitch & cChr & RTFNormal & TextSwitch
     End If
     '
     ''Found an operator on the end!
     'If cIsFoundOperator <> 0 And isOperator = False Then
     '   Dim noOp As String
     '   Dim nDelim As Long
     '   Dim cKeyWrd As String
     '
     '   nDelim = InStr(cIsFoundOperator, OperatorLst, Delim) - cIsFoundOperator + 1
     '   Op = Mid(OperatorLst, cIsFoundOperator + 1, nDelim)
     '   cKeyWrd = Mid(cChr, 1, Len(cChr) - Len(Op))
     '
     '   cChr = QuickCheck(cKeyWrd, totalChrLen, isNewLine)
     '   isOperator = True
     '   'MsgBox (cKeyWrd & vbCrLf & nDelim & vbCrLf & Op & vbCrLf & cChr & vbCrLf & cIsFoundKeyWord & vbCrLf & cIsFoundOperator)
     'End If
     
     
     'Found a keyword!!
     If cIsFoundKeyWord <> 0 Then
        cChr = KeyWordSwitch & cChr & RTFNormal & TextSwitch
     End If
     
     'Found a function/sub!!
     If cIsFoundFuncSub <> 0 Then
        cChr = FuncSubSwitch & cChr & RTFNormal & TextSwitch
     End If
     
     'Found a comment!!
     If cIsFoundComment <> 0 And cIsFoundKeyWord = 0 And _
        cIsFoundFuncSub = 0 Then
        cChr = CommentSwitch & cChr & RTFNormal & TextSwitch
        isComment = True
     End If
     
     'Found none of them so it's normal text...
     If (cIsFoundKeyWord = 0) And (cIsFoundFuncSub = 0) And _
        (cIsFoundComment = 0) Then
        cChr = RTFNormal & TextSwitch & cChr & RTFNormal '& TextSwitch
     End If
     
     If isBlock = True Then
        cChr = cChr & BlockDelim
        
        isBlock = False
     End If
     'If isOperator = True Then
     '   cChr = cChr & Op
     '
     '   isOperator = False
     '   Op = ""
     'End If
  End If
  
'Make sure the variables aren't used in the next
'round w/ values already set...we need to assume
'a keyword/comment/etc. hasn't been found...
cIsFoundKeyWord = 0
cIsFoundComment = 0
cIsFoundFuncSub = 0

QuickCheck = cChr
End Function

'Goes through and inserts the RTF info. line-by-line...
Public Function EvaluateAndHighlightLine(cLine As String) As String
Static isMultiLineComment As Boolean, isInStringDelim As Boolean
Dim cLnTrm As String, i As Long, cOld As Long
Dim cChr As String, cBef As String, cCurrLine As String
Dim cIsFoundKeyWord As Long, cIsFoundFuncSub As Long, cIsFoundComment As Long
Dim totalChrLen As Long
Dim isNewLine As Boolean

cLnTrm = Trim(cLine)

'If it's a comment line, then we need to set it apart
'as a comment...
If Left(cLnTrm, Len(CommentsMultiStart)) = CommentsMultiStart Then
   isMultiLineComment = True
   cLine = RTFNormal & CommentSwitch & cLine
End If

'If it's at the end of a multi-line comment, then we
'need to stop it from highlighting...
If Right(cLnTrm, Len(CommentsMultiEnd)) = CommentsMultiEnd Then
   isMultiLineComment = False
   cLine = CommentSwitch & cLine & TextSwitch & RTFNormal
End If

'If it's part of a multi-line comment then exit
'b/c there's no sense evaluating a commented line...
If isMultiLineComment = True Then
   EvaluateAndHighlightLine = cLine
   Exit Function
End If

'Checks to see if line is a single-line comment...if it
'is, then it will color it green and exit the function...
For i = 1 To CommentsCount
  If Left(LCase(cLnTrm), Len(Comments(i).Comment)) = Comments(i).Comment Then
     cLine = CommentSwitch & cLine & RTFNormal
     EvaluateAndHighlightLine = cLine
     Exit Function
  End If
Next i

'We now know that the line is not any kind of comment
'If it were, then it would have already been colored
'and we would have exited the function...
cOld = 1
'cCurrLine = cLine
For i = 1 To Len(cLine)
  'Finds the first instance of a space...
  i = InStr(i, cLine, Chr(32))
  'Sets what the current word or function, etc. is...
  If i <> 0 Then
       cChr = Mid(cLine, cOld, i - cOld + 1)
       
     Else 'It's the rest of the line if there's no
          'more spaces
       cChr = Mid(cLine, cOld)
       cChr = QuickCheck(cChr, totalChrLen, isNewLine)
       cCurrLine = cCurrLine & cChr
       Exit For
  End If

     'Get the text before the current word...
     'We have to check this b/c if it's at the start
     'of a new line, then cOld is at 1 and it would
     'return the first character of the string...and
     'that does us no good...
     If cOld > 1 Then
          cBef = Mid(cCurrLine, 1, totalChrLen)
            isNewLine = False
          Else
            isNewLine = True
     End If
     
     'Goes and checks to see if it's in the lists...
     'If it is, then the appropriate colors are added
     cChr = QuickCheck(cChr, totalChrLen, isNewLine)
     
     'Adds the current line to the variable...
     cCurrLine = cCurrLine & cChr
     
     'Keeps track of the length of the added RTF
     'code so we know what text to extract for cBef
     totalChrLen = totalChrLen + Len(cChr)
  
  'Keeps the old index for the next loop so we know
  'where the beginning for the next word is...
  cOld = i
Next i

'cIsFoundKeyWord = InStr(1, KeyWordLst, Delim & LCase(Trim(cChr)) & Delim)
'cIsFoundFuncSub = InStr(1, FuncSubLst, LCase(Trim(cChr)))
'cIsFoundComment = InStr(1, CommentLst, LCase(Trim(cChr)))


'cCurrLine = cCurrLine & cChr
'MsgBox (cCurrLine)
EvaluateAndHighlightLine = cCurrLine
End Function

Public Function HighlightStrings(sRTF As String) As String
Dim i As Long, nex As Long
Dim bef As String, aft As String
Dim strEval As String

  For i = 1 To Len(sRTF)
    i = InStr(i, sRTF, StringDelim)
    If i <> 0 Then
         nex = InStr(i + 1, sRTF, StringDelim)
           If nex <> 0 Then
              bef = Mid(sRTF, 1, i - 1)
              aft = Mid(sRTF, nex + 1)
              
              strEval = Mid(sRTF, i, nex - i + 1)
              sRTF = bef & StringDelimSwitch & _
                     strEval & _
                     aft
              i = nex + Len(StringDelimSwitch) + Len(strEval) + 1
              Else
              Exit For
           End If
       Else
         Exit For
    End If
  Next i
  
HighlightStrings = sRTF
End Function


'Highlights the code based on the info. loaded...
Public Function Highlight(sRTF As String) As String
Dim cNumLines As Long, i As Long, cOld As Long
Dim cLn As String, cTotal As String

  'Sets how the array list will be separated...
  Delim = "/"
  'Puts the RTF back to normal after it's been
  'altered...
  RTFNormal = "\plain\f3\fs16"
  'Puts in the color info so now we can reference
  'the colors using RTF switches such as \cf1, \cf2,
  'etc.
  sRTF = InsertHeader(sRTF)
  'The correct header references have now been added...
  
  'Loads the delimited list into the list variables
  'If it didn't load correctly, then display a MsgBox
  'letting the user know something didn't turn out
  'how it should have...
  If LoadTotalArrayLst = False Then MsgBox ("Error highlighting file"), vbExclamation
  
  'Gets the total # of lines in the RTF...
  cNumLines = GetNumTokensInStr(sRTF, vbCrLf)
  
  cOld = 1
  For i = 1 To Len(sRTF)
    'Finds the position of the next line...
    i = InStr(i, sRTF, vbCrLf)
    'You need to check to see if a vbCrLf does exist...
    'if not, then you need to increment i by 2 b/c
    'vbCrLf is 2 characters long...
    If i = 0 Then
         Exit For
       Else
         i = i + 2
    End If
    'Gets the current line...
    On Error Resume Next
    cLn = Mid(sRTF, cOld, i - cOld)
    
    If Right(cLn, 6) = "\par" & vbCrLf Then
       cLn = Mid(cLn, 1, Len(cLn) - 6)
       'cLn now contains the line of code we're evaluating...
       cLn = EvaluateAndHighlightLine(cLn)
    
       'Put the info. back into it...
       cLn = cLn & "\par" & vbCrLf
    End If
    
    'Adds the line to the total RTF...
    cTotal = cTotal & cLn
    'Keeps track of the older line...
    cOld = i
  Next i
  
  'cTotal = HighlightStrings(cTotal) & "}"
  Highlight = cTotal & "}"
End Function

'Frees up some memory after highlighting is complete...
Public Sub HighlightCleanUp()
'Take out all the arrays...
Erase KeyWords, Comments, FuncSubs, Operators, Scope

    'Free up some memory...
    KeyWordLst = ""
    CommentLst = ""
    FuncSubLst = ""
    OperatorLst = ""
    ScopeLst = ""
    
    Delim = ""
    StringDelim = ""
    BlockDelim = ""
    
    CommentsMultiStart = ""
    CommentsMultiEnd = ""
    
    KeyWordsCount = 0
    FuncSubsCount = 0
    CommentsCount = 0
    OperatorsCount = 0
    ScopeCount = 0
    
End Sub
