VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe rettangolo"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private lLeft As Long
Private lTop As Long
Private lRight As Long
Private lBottom As Long
Public Enum rPosition
    rLeftTop = 1
    rRightTop
    rLeftBottom
    rRightBottom
End Enum

Private Sub Class_Initialize()
    lLeft = 0
    lTop = 0
    lRight = 0
    lBottom = 0
End Sub

Public Property Get Left() As Variant
    Left = lLeft
End Property

Public Property Get Top() As Variant
    Top = lTop
End Property

Public Static Property Get Right() As Variant
    Right = lRight
End Property
Public Property Let Right(ByVal vNewRight As Variant)
  lRight = vNewRight
End Property

Public Property Get Bottom() As Variant
    Bottom = lBottom
End Property

Public Property Get Height() As Variant
    Height = (lBottom - lTop)
End Property

Public Property Get Width() As Variant
    Width = (lRight - lLeft)
End Property
Public Property Let Width(ByVal vNewWidth As Variant)
  lRight = lLeft + vNewWidth
End Property
Public Sub Copy(ByVal Other As TRect)
  lLeft = Other.Left
  lRight = Other.Right
  lTop = Other.Top
  lBottom = Other.Bottom
End Sub
Public Sub SetRect(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
  lLeft = Left
  lTop = Top
  lRight = Right
  lBottom = Bottom
End Sub

Public Sub SetRectWH(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
  lLeft = Left
  lTop = Top
  lRight = lLeft + Width
  lBottom = lTop + Height
End Sub


Public Sub Inflate(ByVal dx As Long, ByVal dy As Long)
    lLeft = lLeft + dx
    lRight = lRight + dx
    lTop = lTop + dy
    lBottom = lBottom + dy
End Sub


Public Sub GetFormRect(frm As Form)
        SetRect frm.ScaleLeft, frm.ScaleTop, frm.ScaleWidth, frm.ScaleHeight
End Sub

Public Sub GetControlRect(ctrl As Control)
        SetRectWH ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height
End Sub
Public Sub MoveControlTo(ctrl As Control, ByVal X As Long, ByVal Y As Long)
    GetControlRect ctrl
    MoveTo X, Y
    SetControlRect ctrl
End Sub
Public Sub MoveTo(ByVal X As Long, ByVal Y As Long)
    Dim tmpW As Long, tmpH As Long
    tmpW = Width
    tmpH = Height
    lLeft = X
    lTop = Y
    lRight = X + tmpW
    lBottom = Y + tmpH
End Sub

Public Sub SetControlRect(ctrl As Control)
    ctrl.Left = lLeft
    ctrl.Top = lTop
    ctrl.Width = Width
    ctrl.Height = Height
End Sub

Public Sub Reposition(ByVal pos As rPosition, ByVal outerRec As TRect)
    Select Case pos
        Case rLeftTop
          MoveTo 0, 0
        Case rRightTop
          MoveTo (outerRec.Width - Width), 0
        Case rLeftBottom
          MoveTo 0, (outerRec.Height - Height)
        Case rRightBottom
           MoveTo (outerRec.Width - Width), (outerRec.Height - Height)
    End Select
End Sub

Public Sub Normalize()
Dim Temp As Long
    If lLeft > lRight Then
       Temp = lLeft
       lLeft = lRight
       lRight = Temp
    End If
    If lTop > lBottom Then
       Temp = lTop
       lTop = lBottom
       lBottom = Temp
    End If
End Sub


