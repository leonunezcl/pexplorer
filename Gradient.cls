VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Property Storage Variables
Private mlColor1    As Long
Private mlColor2    As Long
Private mfAngle     As Single

'Property Default Constants
Private Const mlDefColor1   As Long = &HFF  'Red
Private Const mlDefColor2   As Long = 0     'Black
Private Const mfDefAngle    As Single = 315 'Upper-left to lower-right

'Misc Constants
Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians

'TypeDefs
Private Type PointAPI   'API Point structure
    X   As Long
    Y   As Long
End Type

Private Type PointSng   'Internal Point structure
    X   As Single
    Y   As Single
End Type

Private Type RectAPI    'API Rect structure
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'API functions and Constants
Private Const PS_SOLID As Long = 0  'Solid Pen Style (Used for CreatePen())
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RectAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function Draw(picObj As Object) As Boolean

'Note: This class uses API functions to draw. If the
'      destination object is in AutoRedraw mode, the
'      Refresh method for that object must be invoked.

'picObj can be a Form or PictureBox.

Dim lRet    As Long
Dim lIdx    As Long
Dim lTime   As Long
Dim uRect   As RectAPI

'    lTime = GetTickCount()
    
    On Error GoTo LocalError
    
    'Translate System Colors
    If mlColor1 < 0 Then
        mlColor1 = TranslateSystemColor(picObj, mlColor1)
    End If
    If mlColor2 < 0 Then
        mlColor2 = TranslateSystemColor(picObj, mlColor2)
    End If
    
    'Stop the window from updating until we're finished
    lRet = LockWindowUpdate(picObj.hWnd)
    lRet = GetClientRect(picObj.hWnd, uRect)
    If lRet <> 0 Then
        If uRect.Right > 1 And uRect.Bottom > 1 Then
            lIdx = DrawGradient(picObj.hDC, uRect.Right, uRect.Bottom)
            Draw = (lIdx > 0)
        End If
    End If
    
    'My P3-500 took 99 millisecs (.099 secs) to create and draw 1600 lines (that's a
    '800px. x 800px. square at a 45 degree angle). Diagonals produce the most lines.
    'At this speed I can redraw an entire 800px. screen over 10 times per second.
    
    'Same size square at 90 degree angle took 50 millisecs (.050 secs) to create and
    'draw 800 lines. At this speed I can redraw an entire 800px. screen 20 times per second.
    
    'Uncomment the 2 lines below and the 1 at the top of this function to test the times on your PC.
'    lTime = GetTickCount() - lTime
'    MsgBox CStr(lIdx / 2) & " lines drawn in " & CStr(lTime) & " milliseconds"
        
NormalExit:
    'Allow the window to update
    lRet = LockWindowUpdate(0)
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation
    Resume NormalExit

End Function

Private Function TranslateSystemColor(picObj As Object, ByVal lSysColor As Long) As Long

'Translates a System Color to a standard color

Dim lSavePixel  As Long

    'Store the pixel color of picObj at 0,0
    lSavePixel = picObj.Point(0, 0)
    
    'Put a pixel on picObj at 0,0 in the system color
    picObj.PSet (0, 0), lSysColor
    
    'Retrieve the actual color
    TranslateSystemColor = GetPixel(picObj.hDC, 0&, 0&)
    
    'Put the original color back
    picObj.PSet (0, 0), lSavePixel
    
End Function

Private Function GetColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal NbrOfSteps As Long, aRetColors() As Long) As Long

Dim j           As Long
Dim iRed(1)     As Integer
Dim iGreen(1)   As Integer
Dim iBlue(1)    As Integer
Dim iStepDir    As Integer
Dim fRedStep    As Single
Dim fGreenStep  As Single
Dim fBlueStep   As Single
Dim sRed        As String
Dim sGreen      As String
Dim sBlue       As String
Dim sColor      As String

    sColor = FixLen(Hex$(Color1), "000000")
    iBlue(0) = Val("&H" & Left$(sColor, 2))
    iGreen(0) = Val("&H" & Mid$(sColor, 3, 2))
    iRed(0) = Val("&H" & Right$(sColor, 2))
    sColor = FixLen(Hex$(Color2), "000000")
    iBlue(1) = Val("&H" & Left$(sColor, 2))
    iGreen(1) = Val("&H" & Mid$(sColor, 3, 2))
    iRed(1) = Val("&H" & Right$(sColor, 2))
    fBlueStep = Div(CSng(iBlue(1) - iBlue(0)), CSng(NbrOfSteps))
    fGreenStep = Div(CSng(iGreen(1) - iGreen(0)), CSng(NbrOfSteps))
    fRedStep = Div(CSng(iRed(1) - iRed(0)), CSng(NbrOfSteps))
    ReDim aRetColors(NbrOfSteps - 1)
    For j = 1 To Abs(NbrOfSteps)
        sBlue = FixLen(Hex$(iBlue(0) + CInt(fBlueStep * CSng(j))), "00")
        sGreen = FixLen(Hex$(iGreen(0) + CInt(fGreenStep * CSng(j))), "00")
        sRed = FixLen(Hex$(iRed(0) + CInt(fRedStep * CSng(j))), "00")
        aRetColors(j - 1) = CLng("&H" & sBlue & sGreen & sRed)
    Next j
    
    GetColors = NbrOfSteps
        
End Function

Private Function FixLen(ByVal sText As String, ByVal sMask As String) As String

    If Len(sText) < Len(sMask) Then
        FixLen = Left$(sMask, Len(sMask) - Len(sText)) & sText
    Else
        FixLen = Left$(sText, Len(sMask))
    End If
    
End Function

Private Function DrawGradient(ByVal hDC As Long, ByVal lWidth As Long, ByVal lHeight As Long) As Long

Dim iIncX       As Integer
Dim iIncY       As Integer
Dim lIdx        As Long
Dim lRet        As Long
Dim hPen        As Long
Dim hOldPen     As Long
Dim lPointCnt   As Long
Dim laColors()  As Long
Dim fMovX       As Single
Dim fMovY       As Single
Dim fDist       As Single
Dim fAngle      As Single
Dim fLongSide   As Single
Dim uTmpPt      As PointAPI
Dim uaPts()     As PointAPI
Dim uaTmpPts()  As PointSng
    
    On Error GoTo LocalError
    
    'Start with center of rect
    ReDim uaTmpPts(2)
    uaTmpPts(2).X = Int(lWidth / 2)
    uaTmpPts(2).Y = Int(lHeight / 2)
    'Calc distance to furthest edge as if rect were square
    fLongSide = IIf(lWidth > lHeight, lWidth, lHeight)
    fDist = (Sqr((fLongSide ^ 2) + (fLongSide ^ 2)) + 2) / 2
    'Create points to the left and the right at a 0º angle (horizontal)
    uaTmpPts(0).X = uaTmpPts(2).X - fDist
    uaTmpPts(0).Y = uaTmpPts(2).Y
    uaTmpPts(1).X = uaTmpPts(2).X + fDist
    uaTmpPts(1).Y = uaTmpPts(2).Y
    'Lines will be drawn perpendicular to mfAngle so
    'add 90º and correct for 360º wrap
    fAngle = CDbl(mfAngle + 90) - Int(Int(CDbl(mfAngle + 90) / 360#) * 360#)
    'Rotate second and third points to fAngle
    Call RotatePoint(uaTmpPts(2), uaTmpPts(0), fAngle)
    Call RotatePoint(uaTmpPts(2), uaTmpPts(1), fAngle)
    
    'We now have a line that crosses the center and
    'two sides of the rect at the correct angle.
    
    'Calc the starting quadrant, direction of move and amount of
    'first move (fMovX, fMovY) and each incremental move (iIncX, iIncY).
    Select Case mfAngle
        Case 0 To 90
            'Left Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 90 To 180
            'Right Bottom
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to bottom edge; Draw bottom to top
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, lHeight - uaTmpPts(1).Y, lHeight - uaTmpPts(0).Y)
                iIncX = 0
                iIncY = -1
            End If
        Case 180 To 270
            'Right Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to right edge; Draw right to left
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, lWidth - uaTmpPts(1).X, lWidth - uaTmpPts(0).X)
                fMovY = 0
                iIncX = -1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
        Case Else
            'Left Top
            If Abs(uaTmpPts(0).X - uaTmpPts(1).X) <= Abs(uaTmpPts(0).Y - uaTmpPts(1).Y) Then
                'Move line to left edge; Draw left to right
                fMovX = IIf(uaTmpPts(0).X > uaTmpPts(1).X, -uaTmpPts(0).X, -uaTmpPts(1).X)
                fMovY = 0
                iIncX = 1
                iIncY = 0
            Else
                'Move line to top edge; Draw top to bottom
                fMovX = 0
                fMovY = IIf(uaTmpPts(0).Y > uaTmpPts(1).Y, -uaTmpPts(0).Y, -uaTmpPts(1).Y)
                iIncX = 0
                iIncY = 1
            End If
    End Select
    
    'At this point we could calculate where the lines will cross the rect edges, but
    'this would slow things down. The picObj clipping region will take care of this.
    
    'Start with 1000 points and add more if needed. This increases
    'speed by not re-dimming the array in each loop.
    'My P3-500 takes 7 millisecs (.007 secs) to create 5123 points using
    'this schema (that's a 1280px. x 1280px. square at a 45 degree angle).
    ReDim uaPts(999)
    
    'Set the first two points in the array
    uaPts(0).X = uaTmpPts(0).X + fMovX
    uaPts(0).Y = uaTmpPts(0).Y + fMovY
    uaPts(1).X = uaTmpPts(1).X + fMovX
    uaPts(1).Y = uaTmpPts(1).Y + fMovY
    
    lIdx = 2
    'Create the rest of the points by incrementing(iIncX,iIncY) from the first two.
    'Where we stop depends on the direction of travel.
    'We'll continue until both points in a set pass the stop point.
    If iIncX > 0 Then
        While uaPts(lIdx - 1).X < lWidth Or uaPts(lIdx - 2).X < lWidth
            uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
            uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
            lIdx = lIdx + 1
            If (lIdx Mod 1000) = 0 Then
                ReDim Preserve uaPts(UBound(uaPts) + 1000)
            End If
        Wend
    ElseIf iIncX < 0 Then
        While uaPts(lIdx - 1).X > 0 Or uaPts(lIdx - 2).X > 0
            uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
            uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
            lIdx = lIdx + 1
            If (lIdx Mod 1000) = 0 Then
                ReDim Preserve uaPts(UBound(uaPts) + 1000)
            End If
        Wend
    ElseIf iIncY > 0 Then
        While uaPts(lIdx - 1).Y < lHeight Or uaPts(lIdx - 2).Y < lHeight
            uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
            uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
            lIdx = lIdx + 1
            If (lIdx Mod 1000) = 0 Then
                ReDim Preserve uaPts(UBound(uaPts) + 1000)
            End If
        Wend
    ElseIf iIncY < 0 Then
        While uaPts(lIdx - 1).Y > 0 Or uaPts(lIdx - 2).Y > 0
            uaPts(lIdx).X = uaPts(lIdx - 2).X + iIncX
            uaPts(lIdx).Y = uaPts(lIdx - 2).Y + iIncY
            lIdx = lIdx + 1
            If (lIdx Mod 1000) = 0 Then
                ReDim Preserve uaPts(UBound(uaPts) + 1000)
            End If
        Wend
    End If
    
    'Free excess memory
    ReDim Preserve uaPts(lIdx - 1)
    
    'Create the array of colors blending from mlColor1 to mlColor2
    lRet = GetColors(mlColor1, mlColor2, lIdx / 2, laColors)
    
    'Now draw each line in it's own color
    For lIdx = 0 To UBound(uaPts) - 1 Step 2
        'Move to next point
        lRet = MoveToEx(hDC, uaPts(lIdx).X, uaPts(lIdx).Y, uTmpPt)
        'Create the colored pen and select it into the DC
        hPen = CreatePen(PS_SOLID, 1, laColors(Int(lIdx / 2)))
        hOldPen = SelectObject(hDC, hPen)
        'Draw the line
        lRet = LineTo(hDC, uaPts(lIdx + 1).X, uaPts(lIdx + 1).Y)
        'Get the pen back out of the DC and destroy it
        lRet = SelectObject(hDC, hOldPen)
        lRet = DeleteObject(hPen)
    Next lIdx
    
    DrawGradient = lIdx
    
NormalExit:
    'Free the memory
    Erase laColors
    Erase uaPts
    Erase uaTmpPts
    Exit Function
    
LocalError:
    MsgBox Err.Description, vbExclamation, "GradientRect.cls"
    DrawGradient = 0
    Resume NormalExit
    
End Function

Private Sub RotatePoint(uAxisPt As PointSng, uRotatePt As PointSng, fDegrees As Single)

Dim fDX As Single
Dim fDY As Single
Dim fRadians As Single

    fRadians = fDegrees * RADS
    fDX = uRotatePt.X - uAxisPt.X
    fDY = uRotatePt.Y - uAxisPt.Y
    uRotatePt.X = uAxisPt.X + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
    uRotatePt.Y = uAxisPt.Y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    
End Sub



Private Function Div(ByVal dNumer As Double, ByVal dDenom As Double) As Double
    
    If dDenom <> 0 Then
        Div = dNumer / dDenom
    Else
        Div = 0
    End If

End Function

Public Property Let Color1(ByVal lData As Long)
    
    mlColor1 = lData
    
End Property


Public Property Get Color1() As Long
    Color1 = mlColor1
End Property



Public Property Let Color2(ByVal lData As Long)
    
    mlColor2 = lData

End Property


Public Property Get Color2() As Long
    Color2 = mlColor2
End Property



Public Property Let Angle(ByVal fData As Single)
    
'Angles are counter-clockwise and may be
'any Single value from 0 to 359.999999999.

' 135  90 45
'    \ | /
'180 --o-- 0
'    / | \
' 235 270 315

    'Correct angle to ensure between 0 and 359.999999999
    mfAngle = CDbl(fData) - Int(Int(CDbl(fData) / 360#) * 360#)

End Property


Public Property Get Angle() As Single
    Angle = mfAngle
End Property







Private Sub Class_Initialize()

    mlColor1 = mlDefColor1
    mlColor2 = mlDefColor2
    mfAngle = mfDefAngle
    
End Sub


