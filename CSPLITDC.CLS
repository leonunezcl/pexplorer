VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSplitDDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cSplitDDC
' Filename: cSplitDC.cls
' Author:   SP McMahon
' Date:     07 July 1998
'
' A splitter class using the Desktop window to draw a
' splitter bar, therefore allowing splitting of MDI forms
' as well as standard forms.
' ======================================================================

'// some global declarations
Private bDraw As Boolean
Private rcCurrent As RECT
Private rcNew As RECT
Private rcWindow As RECT

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hDC As Long, ByVal nDrawMode As Long) As Long
Private Const R2_BLACK = 1       '   0
Private Const R2_COPYPEN = 13    '  P
Private Const R2_LAST = 16
Private Const R2_MASKNOTPEN = 3  '  DPna
Private Const R2_MASKPEN = 9     '  DPa
Private Const R2_MASKPENNOT = 5  '  PDna
Private Const R2_MERGENOTPEN = 12        '  DPno
Private Const R2_MERGEPEN = 15   '  DPo
Private Const R2_MERGEPENNOT = 14        '  PDno
Private Const R2_NOP = 11        '  D
Private Const R2_NOT = 6 '  Dn
Private Const R2_NOTCOPYPEN = 4  '  PN
Private Const R2_NOTMASKPEN = 8  '  DPan
Private Const R2_NOTMERGEPEN = 2 '  DPon
Private Const R2_NOTXORPEN = 10  '  DPxn
Private Const R2_WHITE = 16      '   1
Private Const R2_XORPEN = 7      '  DPx
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15

Public Enum eOrientationConstants
    espVertical = 1
    espHorizontal = 2
End Enum
Private m_hWnd As Long
Private m_eOrientation As eOrientationConstants
Private m_lBorder(1 To 4) As Long
Private m_oSplit As Object
Public Enum ESplitBorderTypes
   espbLeft = 1
   espbTop = 2
   espbRight = 3
   espbBottom = 4
End Enum
Private m_bIsMDI As Boolean
Private m_bSplitting As Boolean

Public Property Get SplitObject() As Object
    Set SplitObject = m_oSplit
End Property
Public Property Let SplitObject(ByRef oThis As Object)
    Set m_oSplit = oThis
    On Error Resume Next
    oThis.BorderStyle = 0
    If (m_eOrientation = espHorizontal) Then
        oThis.MousePointer = vbSizeNS
    Else
        oThis.MousePointer = vbSizeWE
    End If
End Property
Public Property Let Border(ByVal eBorderType As ESplitBorderTypes, ByVal lSize As Long)
   m_lBorder(eBorderType) = lSize
End Property
Public Property Get Border(ByVal eBorderType As ESplitBorderTypes) As Long
   Border = m_lBorder(eBorderType)
End Property
Public Property Get Orientation() As eOrientationConstants
    Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As eOrientationConstants)
    m_eOrientation = eOrientation
    If Not (m_oSplit Is Nothing) Then
        If (m_eOrientation = espHorizontal) Then
            m_oSplit.MousePointer = vbSizeNS
            m_lBorder(espbTop) = 64
            m_lBorder(espbBottom) = 64
            m_lBorder(espbLeft) = 0
            m_lBorder(espbRight) = 0
        Else
            m_oSplit.MousePointer = vbSizeWE
            m_lBorder(espbTop) = 0
            m_lBorder(espbBottom) = 0
            m_lBorder(espbLeft) = 64
            m_lBorder(espbRight) = 64
        End If
    End If
End Property

Public Sub SplitterMouseDown( _
        ByVal hwnd As Long, _
        ByVal X As Long, _
        ByVal Y As Long _
    )
Dim tP As POINTAPI

    m_hWnd = hwnd

    ' Send subsequent mouse messages to the owner window
    SetCapture m_hWnd
    ' Get the window rectangle on the desktop of the owner window:
    GetWindowRect m_hWnd, rcWindow
    ' Clip the cursor so it can't move outside the window:
    ClipCursorRect rcWindow
    
    ' Check if this is an MDI form:
    If (ClassName(m_hWnd) = "ThunderMDIForm") Then
        ' Get the inside portion of the MDI form:
        ' I'm assuming you have a caption,menu and border in your MDI here
        rcWindow.Left = rcWindow.Left + GetSystemMetrics(SM_CXBORDER)
        rcWindow.Right = rcWindow.Right - GetSystemMetrics(SM_CXBORDER)
        rcWindow.Bottom = rcWindow.Bottom - GetSystemMetrics(SM_CYBORDER)
        rcWindow.Top = rcWindow.Top + GetSystemMetrics(SM_CYBORDER) * 3 + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYMENU)
        m_bIsMDI = True
    Else
        ' Get the client rectangle of the window in screen coordinates:
        GetClientRect m_hWnd, rcWindow
        tP.X = rcWindow.Left
        tP.Y = rcWindow.Top
        ClientToScreen m_hWnd, tP
        rcWindow.Left = tP.X
        rcWindow.Top = tP.Y
        tP.X = rcWindow.Right
        tP.Y = rcWindow.Bottom
        ClientToScreen m_hWnd, tP
        rcWindow.Right = tP.X
        rcWindow.Bottom = tP.Y
        m_bIsMDI = False
    End If
    bDraw = True  '// start actual drawing from next move message
    
    rcCurrent.Left = 0: rcCurrent.Top = 0: rcCurrent.Right = 0: rcCurrent.Bottom = 0
    
    X = (m_oSplit.Left + X) \ Screen.TwipsPerPixelX
    Y = (m_oSplit.Top + Y) \ Screen.TwipsPerPixelY
    SplitterFormMouseMove X, Y
    
End Sub

Public Sub SplitterFormMouseMove( _
      ByVal X As Long, _
      ByVal Y As Long)
Dim hDC As Long
Dim tP As POINTAPI
Dim hWndClient As Long
    If (bDraw) Then
        '// Draw two rectangles in the screen DC to cause splitting:
        
        ' First get the Desktop DC:
        hDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        ' Set the draw mode to XOR:
        SetROP2 hDC, R2_NOTXORPEN
    
        '// Draw over and erase the old rectangle
        ' (if this is the first time, all the coords will be 0 and nothing will get drawn):
        Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
                
        ' It is simpler to use the mouse cursor position than try to translate
        ' X,Y to screen coordinates!
        GetCursorPos tP
        
        ' Determine where to draw the splitter:
        If (m_eOrientation = espHorizontal) Then
            rcNew.Left = rcWindow.Left + m_lBorder(espbLeft)
            rcNew.Right = rcWindow.Right - m_lBorder(espbRight)
            If (tP.Y >= rcWindow.Top + m_lBorder(espbTop)) And (tP.Y < rcWindow.Bottom - m_lBorder(espbBottom)) Then
                rcNew.Top = tP.Y - 2
                rcNew.Bottom = tP.Y + 2
            Else
                If (tP.Y < rcWindow.Top + m_lBorder(espbTop)) Then
                    rcNew.Top = rcWindow.Top + m_lBorder(espbTop) - 2
                    rcNew.Bottom = rcNew.Top + 5
                Else
                    rcNew.Top = rcWindow.Bottom - m_lBorder(espbBottom) - 2
                    rcNew.Bottom = rcNew.Top + 5
                End If
            End If
        Else
            rcNew.Top = rcWindow.Top + m_lBorder(espbTop)
            rcNew.Bottom = rcWindow.Bottom - m_lBorder(espbBottom)
            If (tP.X >= rcWindow.Left + m_lBorder(espbLeft)) And (tP.X <= rcWindow.Right - m_lBorder(espbRight)) Then
                rcNew.Left = tP.X - 2
                rcNew.Right = tP.X + 2
            Else
                If (tP.X < rcWindow.Left + m_lBorder(espbLeft)) Then
                    rcNew.Left = rcWindow.Left + m_lBorder(espbLeft) - 2
                    rcNew.Right = rcNew.Left + 5
                Else
                    rcNew.Left = rcWindow.Right - m_lBorder(espbRight) - 2
                    rcNew.Right = rcNew.Left + 5
                End If
            End If
        End If
        
        '// Draw the new rectangle
        Rectangle hDC, rcNew.Left, rcNew.Top, rcNew.Right, rcNew.Bottom
        
        ' Store this position so we can erase it next time:
        LSet rcCurrent = rcNew
        
        ' Free the reference to the Desktop DC we got (make sure you do this!)
        DeleteDC hDC
    End If
    
End Sub

Public Function SplitterFormMouseUp( _
      ByVal X As Long, _
      ByVal Y As Long _
   ) As Boolean
Dim hDC As Long
Dim tP As POINTAPI
Dim hWndClient As Long

      '// Don't leave orphaned rectangle on desktop; erase last rectangle.
   If (bDraw) Then
       bDraw = False
       
       ' Release mouse capture:
       ReleaseCapture
       ' Release the cursor clipping region (must do this!):
       ClipCursorClear 0&
       
       ' Get the Desktop DC:
       hDC = CreateDCAsNull("DISPLAY", 0, 0, 0)
       ' Set to XOR drawing mode:
       SetROP2 hDC, R2_NOTXORPEN
       ' Erase the last rectangle:
       Rectangle hDC, rcCurrent.Left, rcCurrent.Top, rcCurrent.Right, rcCurrent.Bottom
       ' Clear up the desktop DC:
       DeleteDC hDC
       
       ' Here we ensure the splitter is within bounds before releasing:
       GetCursorPos tP

       If (tP.X < rcWindow.Left + m_lBorder(espbLeft)) Then
           tP.X = rcWindow.Left + m_lBorder(espbLeft)
       End If
       If (tP.X > rcWindow.Right - m_lBorder(espbRight)) Then
           tP.X = rcWindow.Right - m_lBorder(espbRight)
       End If
       If (tP.Y < rcWindow.Top + m_lBorder(espbTop)) Then
           tP.Y = rcWindow.Top + m_lBorder(espbTop)
       End If
       If (tP.Y > rcWindow.Bottom - m_lBorder(espbBottom)) Then
           tP.Y = rcWindow.Bottom - m_lBorder(espbBottom)
       End If
       ScreenToClient m_hWnd, tP
       
       ' Move the splitter to the validated final position:
       If (m_eOrientation = espHorizontal) Then
           m_oSplit.Top = (tP.Y - 2) * Screen.TwipsPerPixelY
       Else
           m_oSplit.Left = (tP.X - 2) * Screen.TwipsPerPixelX
       End If
       
       ' Return true to tell the owner we have completed splitting:
       SplitterFormMouseUp = True
   End If
   
End Function

Private Sub Class_Initialize()
   m_eOrientation = espVertical
   m_lBorder(espbLeft) = 64
   m_lBorder(espbRight) = 64
End Sub
Private Function ClassName(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function

