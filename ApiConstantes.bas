Attribute VB_Name = "MmApiCon"
Option Explicit

Public Const RGN_OR = 2

'treeview
Public Const TV_FIRST  As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_DELETEITEM  As Long = (TV_FIRST + 1)
Public Const TVGN_ROOT As Long = &H0
Public Const WM_SETREDRAW As Long = &HB

'ListView
Public Const LVM_FIRST = &H1000                   '// ListView messages
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const GWL_STYLE = (-16)
Public Const HDS_BUTTONS = &H2

'Common Dialog
Public Const OFN_READONLY = &H1
Public Const WM_COPY = &H301
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

'CURSOR
Public Const IDC_WAIT = 32514&
Public Const IDC_ARROW = 32512&
Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

Public Const GWL_WNDPROC = (-4)
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4

Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const MOVEFILE_REPLACE_EXISTING = &H1
Public Const FILE_BEGIN = 0
Public Const CREATE_NEW = 1
Public Const GENERIC_READ = &H80000000
Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Public Const SB_GETRECT = (WM_USER + 10)

Public Const OF_EXIST = &H4000
Public Const OFS_MAXPATHNAME = 256
