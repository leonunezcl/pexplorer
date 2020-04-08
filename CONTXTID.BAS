Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <axButton>
'=====================================================================
Global Const Hlp_Contents = 10    'Main Help Window
Global Const Hlp_Revisions = 30    'Main Help Window
Global Const Hlp_License = 40    'Main Help Window
Global Const Hlp_Tech_Support = 50    'Main Help Window
Global Const Hlp_Properties = 60    'Second Help Window
Global Const Hlp_Methods = 70    'Second Help Window
Global Const Hlp_Events = 80    'Second Help Window
Global Const Hlp_Standard = 90    'Main Help Window
Global Const Hlp_ColorDarkShadow = 100    'Second Help Window
Global Const Hlp_ColorHighlight = 110    'Second Help Window
Global Const Hlp_ColorLightShadow = 120    'Second Help Window
Global Const Hlp_DownPicture = 130    'Second Help Window
Global Const Hlp_DropDown = 140    'Second Help Window
Global Const Hlp_FlatPicture = 150    'Second Help Window
Global Const Hlp_MaskColor = 160    'Second Help Window
Global Const Hlp_Picture = 170    'Second Help Window
Global Const Hlp_PictureAlign = 180    'Second Help Window
Global Const Hlp_ShowFlatGrey = 190    'Second Help Window
Global Const Hlp_Style = 200    'Second Help Window
Global Const Hlp_ShowAbout = 210    'Second Help Window
Global Const Hlp_DropDownClick = 230    'Second Help Window
Global Const Hlp_ButtonGroup = 240    'Second Help Window
Global Const Hlp_ButtonGroupDefault = 250    'Second Help Window
Global Const Hlp_DisabledPicture = 260    'Second Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window


Type HELPWININFO
  wStructSize As Integer
  X As Integer
  Y As Integer
  dX As Integer
  dY As Integer
  wMax As Integer
  rgChMember As String * 2
End Type
    Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
    Declare Function WinHelpByInfo Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As HELPWININFO) As Integer
    Declare Function WinHelpByStr Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$) As Integer
    Declare Function WinHelpByNum Lib "User" Alias "Winhelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&) As Integer

    Dim m_hWndMainWindow as Integer ' hWnd to tell WINHELP the helpfile owner
Dim MainWindowInfo as HELPWININFO
Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path),1) = "\" then
        App.HelpFile = App.Path + "axButton.HLP"
    else
        App.HelpFile = App.Path + "\axButton.HLP"
    end if
MainWindowInfo.wStructSize = 14
    MainWindowInfo.X=256
    MainWindowInfo.Y=256
    MainWindowInfo.dX=512
    MainWindowInfo.dY=512
    MainWindowInfo.rgChMember=Chr$(0)+Chr$(0)
End Sub
Sub QuitHelp ()
    Dim Result as Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each 
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd02", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd03", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowGlossary ()
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(64000))

End Sub
Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, Clng(ContextID))

End Sub
Sub DoHelpMacro (ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal(Macrostring))

End Sub
Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, Clng(0))

End Sub
Sub ShowContentsTab ()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_FINDER, Clng(0))

End Sub
Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, Clng(0))

End Sub

Sub SearchHelp ()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal "" )

End Sub

Sub SearchHelpKeyWord (Argument as String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Sub HelpWindowSize (x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result as Variant
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dx = wx
    MainWindowInfo.dy = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub
