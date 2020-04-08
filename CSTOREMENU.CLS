VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStoreMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tMI() As tMenuItem
Private m_iMenuCount As Long
Private m_sKey As String
Private m_hMenu As Long

Friend Property Get hMenu() As Long
   hMenu = m_hMenu
End Property
Friend Property Let hMenu(ByVal l As Long)
   m_hMenu = l
End Property
Friend Property Get Key() As String
   Key = m_sKey
End Property
Friend Property Let Key(ByVal sKey As String)
   m_sKey = sKey
End Property

Friend Sub Store(ByRef tMI() As tMenuItem, ByVal iMenuCount As Long)
Dim iItem As Long
   m_iMenuCount = iMenuCount
   If (m_iMenuCount > 0) Then
      ReDim m_tMI(1 To m_iMenuCount) As tMenuItem
      For iItem = 1 To m_iMenuCount
         LSet m_tMI(iItem) = tMI(iItem)
      Next iItem
   Else
      Erase m_tMI
   End If
End Sub
Friend Sub Restore(ByRef tMI() As tMenuItem, ByRef iMenuCount As Long)
Dim iItem As Long
   iMenuCount = m_iMenuCount
   ReDim tMI(1 To m_iMenuCount) As tMenuItem
   For iItem = 1 To m_iMenuCount
      LSet tMI(iItem) = m_tMI(iItem)
   Next iItem

End Sub

