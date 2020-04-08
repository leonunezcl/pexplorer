Attribute VB_Name = "mTextBoxLineNumbers"
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Const DT_CALCRECT = &H400
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_SETMARGINS& = &HD3
Private Const EC_LEFTMARGIN& = &H1
Private Const EC_RIGHTMARGIN& = &H2

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20)

Private Const WS_EX_RTLREADING = &H2000&
Private Const WS_EX_LTRREADING = &H0

Public Property Get TextBoxRTL(txtThis As TextBox)
   Dim lStyle As Long
   lStyle = GetWindowLong(txtThis.hWnd, GWL_EXSTYLE)
   If (lStyle And WS_EX_RTLREADING) = WS_EX_RTLREADING Then
      TextBoxRTL = True
   End If
End Property

Public Sub DrawLines(picTo As PictureBox, txtThis As TextBox)
Dim lLine As Long
Dim lCount As Long
Dim lCurrent As Long
Dim hBr As Long
Dim lEnd As Long
Dim lhDC As Long
Dim bComplete As Boolean
Dim tR As RECT, tTR As RECT
Dim oCol As OLE_COLOR
Dim lStart As Long
Dim lEndLine As Long
Dim tPO As POINTAPI
Dim lLineHeight As Long
Dim hPen As Long
Dim hPenOld As Long

   'Debug.Print "DrawLines"
   lhDC = picTo.hdc
   DrawText lhDC, "Hy", 2, tTR, DT_CALCRECT
   lLineHeight = tTR.Bottom - tTR.Top + 1
   
   lCount = LineCount(txtThis.hWnd)
   lCurrent = LineForCharacterIndex(txtThis.hWnd, txtThis.SelStart)
   lStart = txtThis.SelStart
   lEnd = txtThis.SelStart + txtThis.SelLength - 1
   If (lEnd > lStart) Then
      lEndLine = LineForCharacterIndex(txtThis.hWnd, lEnd)
   Else
      lEndLine = lCurrent
   End If
   lLine = FirstVisibleLine(txtThis.hWnd)
   GetClientRect picTo.hWnd, tR
   lEnd = tR.Bottom - tR.Top
      
   hBr = CreateSolidBrush(TranslateColor(picTo.BackColor))
   FillRect lhDC, tR, hBr
   DeleteObject hBr
   tR.Left = 2
   tR.Right = tR.Right - 2
   tR.Top = 0
   tR.Bottom = tR.Top + lLineHeight
   
   SetTextColor lhDC, TranslateColor(vbButtonShadow)
   
   tR.Right = tR.Right - 2
   Do
      ' Ensure correct colour:
      If (lLine = lCurrent) Then
         SetTextColor lhDC, TranslateColor(vbWindowText)
      ElseIf (lLine = lEndLine + 1) Then
         SetTextColor lhDC, TranslateColor(vbButtonShadow)
      End If
      ' Draw the line number:
      DrawText lhDC, CStr(lLine + 1), -1, tR, DT_RIGHT
      
      ' Increment the line:
      lLine = lLine + 1
      ' Increment the position:
      OffsetRect tR, 0, lLineHeight
      If (tR.Bottom > lEnd) Or (lLine + 1 > lCount) Then
         bComplete = True
      End If
   Loop While Not bComplete
   
   ' Draw a line...
   tR.Right = tR.Right + 2
   MoveToEx lhDC, tR.Right, 0, tPO
   hPen = CreatePen(PS_SOLID, 1, TranslateColor(vbButtonShadow))
   hPenOld = SelectObject(lhDC, hPen)
   LineTo lhDC, tR.Right, lEnd
   SelectObject lhDC, hPenOld
   
   DeleteObject hPen
   If picTo.AutoRedraw Then
      picTo.Refresh
   End If
   
End Sub

Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

Private Property Get LineCount(ByVal hWnd As Long)
    LineCount = SendMessageLong(hWnd, EM_GETLINECOUNT, 0&, 0&)
End Property

Private Property Get LineForCharacterIndex(ByVal hWnd As Long, ByVal lIndex As Long) As Long
   LineForCharacterIndex = SendMessageLong(hWnd, EM_LINEFROMCHAR, lIndex, 0)
End Property

Private Property Get FirstVisibleLine(ByVal hWnd As Long) As Long
   FirstVisibleLine = SendMessageLong(hWnd, EM_GETFIRSTVISIBLELINE, 0, 0)
End Property

Public Sub SetMargins(ByVal hWnd As Long, ByVal lLeft As Integer, ByVal lRight As Integer)
Dim lMargins As Long
    lMargins = MakeDWord(lRight, lLeft)
    SendMessageLong hWnd, EM_SETMARGINS, ByVal (EC_LEFTMARGIN Or EC_RIGHTMARGIN), lMargins
End Sub

Private Function MakeDWord(wHi As Integer, wLo As Integer) As Long

    If wHi And &H8000& Then
        MakeDWord = (((wHi And &H7FFF&) * 65536) Or (wLo And &HFFFF&)) Or &H80000000
    Else
        MakeDWord = (wHi * 65536) + wLo
    End If

End Function

