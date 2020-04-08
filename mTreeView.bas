Attribute VB_Name = "mTreeView"
Option Explicit

Public Const COLOR_WINDOW As Long = 5
Public Const COLOR_WINDOWTEXT As Long = 8

Public Const WM_SETREDRAW As Long = &HB
Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003
Public Const GWL_STYLE = (-16)

Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_DELETEITEM  As Long = (TV_FIRST + 1)
Public Const TVGN_ROOT As Long = &H0

Public Const TVIF_STATE As Long = &H8

'treeview styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Public Const TVIS_BOLD  As Long = &H10

Public Const TV_FIRST As Long = &H1100
'Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

'Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TV_ITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type

Public Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd&, _
                                              ByVal wMsg&, ByVal wParam&, lParam As Any) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd&, _
                                                                          ByVal nIndex&)
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, _
                                                        ByVal nIndex&, ByVal dwNewLong&)

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
Private Function GetTVBackColour(ByVal hwndTV As Long) As Long

   Dim clrref As Long
         
  'try for the treeview backcolor
   clrref = SendMessage(hwndTV, TVM_GETBKCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'a treeview background is COLOR_WINDOW.
   If clrref = -1 Then
      clrref = GetSysColor(COLOR_WINDOW)  ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVBackColour = clrref
   
End Function


Private Function GetTVForeColour(ByVal hwndTV As Long) As Long

   Dim clrref As Long
         
  'try for the treeview text colour
   clrref = SendMessage(hwndTV, TVM_GETTEXTCOLOR, 0, ByVal 0)
   
  'if clrref = -1, then the color is a system color.
  'In theory, system colors need to be Or'd with &HFFFFFF
  'to retrieve the actual RGB value, but not Or'ing
  'seems to work for me. The default system colour for
  'treeview text is COLOR_WINDOWTEXT.
   If clrref = -1 Then
      clrref = GetSysColor(COLOR_WINDOWTEXT) ' Or &HFFFFFF
   End If
   
  'one way or another, pass it back
   GetTVForeColour = clrref
   
End Function


Private Sub SetTVBackColour(ByVal hwndTV As Long, clrref As Long)
   
   Dim style As Long
   
  'Change the background
   Call SendMessage(hwndTV, TVM_SETBKCOLOR, 0, ByVal clrref)
   
  'reset the treeview style so the
  'tree lines appear properly
   style = GetWindowLong(hwndTV, GWL_STYLE)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
  
End Sub

Public Sub SetTVForeColour(ByVal hwndTV As Long, clrref As Long)

   Dim style As Long
      
  'Change the background
   Call SendMessage(hwndTV, TVM_SETTEXTCOLOR, 0, ByVal clrref)
   
  'reset the treeview style so the
  'tree lines appear properly
   style = GetWindowLong(hwndTV, GWL_STYLE)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
   
End Sub
Public Sub TVBold(ByVal hwndTV As Long, ByVal Tipo As Integer)

    Dim TVI As TV_ITEM
    Dim hitemTV As Long
       
  'get the handle to the treeview item.
  'If the item is selected, use TVGN_CARET.
  'To highlight the first item in the root, use TVGN_ROOT
  'To hilight the first visible, use TVGN_FIRSTVISIBLE
  'To hilight the selected item, use TVGN_CARET
  
    If Tipo = 1 Then
        hitemTV = SendMessage(hwndTV, TVM_GETNEXTITEM, TVGN_ROOT, ByVal 0&)
    Else
        hitemTV = SendMessage(hwndTV, TVM_GETNEXTITEM, TVGN_CARET, ByVal 0&)
    End If
    
  'if a valid handle get and set the
  'item's state attributes
   If hitemTV > 0 Then
   
      With TVI
         .hItem = hitemTV
         .mask = TVIF_STATE
         .stateMask = TVIS_BOLD
          Call SendMessage(hwndTV, TVM_GETITEM, 0&, TVI)
         
         'flip the bold mask state
         .state = TVIS_BOLD
      End With
      
      Call SendMessage(hwndTV, TVM_SETITEM, 0&, TVI)
 
   End If
   
End Sub


