VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////
'
' 2/9/1999
' Program written by Maurizio Fassina
' e-mail: maufass@tin.it
'
'///////////////////////////////////////////////
' TSplitter is a class that supports operations for manipulating
' (move and resizing) controls contained in a form and splitters.
' Splitters can be moved (thereby resizing the controls) by mouse input.

Option Explicit

'//////////////// Types /////////////////
Public Enum SplitType
  spVertical = 1
  spOrizontal
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'//////////////// Costants ////////////////

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_MIDDLE = &H800    ' Riempie la parte centrale.
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_LEFT = &H1
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

'///////////////////// DLLs //////////////////////////
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'///////////////////// Variables //////////////////
Private Const SplitterHeight As Integer = 80
Private bMouseDrag As Boolean
Private bMouseIn As Boolean
Private oldXY As Single
Private TypeOfSplitter As SplitType
Private Percent As Single
Private recSplitter As TRect  'dimensioni dello splitter
Private recWindow As TRect   'finestra da dividere
Private pForm As Form

'//////////////////// Methods ////////////////////
Private Sub Class_Initialize()
    Percent = 0
    bMouseDrag = False
    bMouseIn = False
    oldXY = 0
    Set recWindow = New TRect
    Set recSplitter = New TRect
End Sub
Public Sub Init(ByVal frm As Form, _
                ByVal rcWindow As TRect, _
                ByVal Perc As Single, _
                Optional ByVal sptype As SplitType = spVertical)
    TypeOfSplitter = sptype
    Percent = Perc
    Set pForm = frm
    If frm.ClipControls = True Then
      Err.Raise vbObjectError, "TSplitter.TSplitter", frm.Name + ".ClipControl should be False."
    End If
    If sptype = spVertical Then
      Screen.MouseIcon = LoadResPicture(2, vbResCursor)
    Else
      Screen.MouseIcon = LoadResPicture(1, vbResCursor)
    End If
    Resize rcWindow
End Sub
Public Sub Resize(ByVal rcWindow As TRect)
    recWindow.Copy rcWindow
    
    If TypeOfSplitter = spVertical Then
      recSplitter.SetRectWH Int((rcWindow.Right * Percent) - (SplitterHeight / 2)), _
                            rcWindow.Top, _
                            SplitterHeight, _
                            rcWindow.Height
    Else
      recSplitter.SetRectWH rcWindow.Left, _
                            Int((rcWindow.Bottom * Percent) - (SplitterHeight / 2)), _
                            rcWindow.Right, _
                            SplitterHeight
    End If

End Sub
Public Sub MouseDown()
    If bMouseIn Then bMouseDrag = True
End Sub

Public Sub MouseMove(X As Single, Y As Single)
  Dim bInside As Boolean
  
  'If Button <> vbLeftButton Then Exit Sub
  With recSplitter
    bInside = (X > .Left) And (X < .Right) And (Y > .Top) And (Y < .Bottom)
  End With
  'first entry in splitter area
  If bInside And bMouseIn = False Then
      bMouseIn = True
      SetCapture pForm.hwnd
      Screen.MousePointer = 99
   'mouse exit from splitter area
   ElseIf (Not bInside) And bMouseIn Then
      bMouseIn = False
      If Not bMouseDrag Then
        Screen.MousePointer = vbDefault
        ReleaseCapture
      End If
    End If
    If bMouseDrag Then
        If oldXY > 0 Then DrawLine oldXY
        If TypeOfSplitter = spVertical Then
          DrawLine X
          oldXY = X
        Else
          DrawLine Y
          oldXY = Y
        End If
    End If
End Sub

Public Function MouseUp(X As Single, Y As Single) As Boolean
  MouseUp = False
  If bMouseDrag Then
    bMouseDrag = False
    DrawLine oldXY
    oldXY = 0
    If Not bMouseIn Then
      Screen.MousePointer = vbDefault
      ReleaseCapture
    End If
    If TypeOfSplitter = spVertical Then
      Percent = X / recWindow.Width
      MoveSplitRect Int(X), recSplitter.Top
    Else
      Percent = Y / recWindow.Height
      MoveSplitRect recSplitter.Left, Int(Y)
    End If
    MouseUp = True  ' the Form should resize  the inner rectangles
   End If
End Function


Public Sub ResizeFrame(what As Integer, frm As Control)
 Dim rec As TRect
 Set rec = New TRect
 With recWindow
  If what = 1 And TypeOfSplitter = spVertical Then
    rec.SetRect .Left, .Top, recSplitter.Left, .Bottom
  ElseIf what = 1 And TypeOfSplitter = spOrizontal Then
      rec.SetRect .Left, .Top, .Right, recSplitter.Top
  ElseIf what = 2 And TypeOfSplitter = spVertical Then
      rec.SetRect recSplitter.Right, .Top, .Right, .Bottom
  ElseIf what = 2 And TypeOfSplitter = spOrizontal Then
    rec.SetRect .Left, recSplitter.Bottom, .Right, .Bottom
  End If
  rec.SetControlRect frm
 End With
End Sub
Public Sub DEdge() 'draws the edge
  Dim rcs As RECT
  With recSplitter
    rcs.Left = .Left
    rcs.Bottom = .Bottom
    rcs.Top = .Top
    rcs.Right = .Right
    pForm.DrawWidth = 1
    pForm.DrawMode = vbCopyPen
    DrawEdge pForm.hdc, rcs, EDGE_RAISED, BF_RECT
    pForm.Line (.Right - 2 * Screen.TwipsPerPixelX, .Top)-(.Right - 2 * Screen.TwipsPerPixelX, .Bottom), RGB(127, 127, 127)
    pForm.Line (.Right - Screen.TwipsPerPixelX, .Top)-(.Right - Screen.TwipsPerPixelX, .Bottom), RGB(0, 0, 0)
  End With
End Sub
Private Sub MoveSplitRect(ByVal X As Long, ByVal Y As Long)
    recSplitter.MoveTo X, Y
End Sub

Private Sub DrawLine(xy As Single)
  With pForm
    .DrawStyle = vbSolid
    .DrawWidth = 3
    .DrawMode = vbMergePenNot
    If TypeOfSplitter = spVertical Then
      pForm.Line (Int(xy), recWindow.Top)-(Int(xy), recWindow.Bottom), 0
    Else
      pForm.Line (recWindow.Left, Int(xy))-(recWindow.Right, Int(xy)), 0
    End If
    .DrawMode = vbNop
  End With
End Sub

