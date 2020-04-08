Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SendMessageLong Lib _
   "user32" Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Public Declare Function SendMessage Lib "user32" _
     Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     lParam As Long) As Long

Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160

Public Const ANSI_FIXED_FONT = 11
Public Const ANSI_VAR_FONT = 12
Public Const SYSTEM_FONT = 13
Public Const DEFAULT_GUI_FONT = 17 'win95/98 only

Public Const SM_CXHSCROLL = 21
Public Const SM_CXHTHUMB = 10
Public Const SM_CXVSCROLL = 2

Type SIZE
  cx As Long
  cy As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Declare Function DrawText Lib "user32" Alias "DrawTextA" _
    (ByVal hDc As Long, _
     ByVal lpStr As String, _
     ByVal nCount As Long, _
     lpRect As RECT, _
     ByVal wFormat As Long) As Long

Public Const DT_CALCRECT = &H400

Declare Function SelectObject Lib "gdi32" _
    (ByVal hDc As Long, ByVal hObject As Long) As Long

Declare Function GetTextExtentPoint32 Lib "gdi32" _
    Alias "GetTextExtentPoint32A" _
   (ByVal hDc As Long, _
    ByVal lpsz As String, _
    ByVal cbString As Long, _
    lpSize As SIZE) As Long

Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


' Pocedemiento que ajusta el ancho del Dropdown de un Combobox deacuerdo al contenido
' Usenlo cuando este seguro de que el contenido va ser mas acho que
' el combobox
' Written by JANO     invtql@principal.unjbg.edu.pe
Public Sub ComboAutoSize(frm As Form, cbo As ComboBox)
   Dim r As Long
   Dim i As Long
   Dim NumOfChars As Long
   Dim LongestComboItem As Long
   Dim avgCharWidth As Long
   Dim NewDropDownWidth As Long
   
  'evaluamos cada entrada del Comoboentries, y usando SendMessageLong
  'con CB_GETLBTEXTLEN determinamos la longitud del elemento
  'en la porcion dropdown del combo

   For i = 0 To cbo.ListCount - 1
      NumOfChars = SendMessageLong(cbo.hwnd, CB_GETLBTEXTLEN, i, 0)
      If NumOfChars > LongestComboItem Then LongestComboItem = NumOfChars
    Next
   
  'calculamos la longitud promedio del los carateres usando
  'la función GetFontDialogUnits. Como el tipo de letra de usado en
  'GetFontDialogUnits es ficticio(es otro), el valor de avgCharWidth es solo
  'una aproximacion basada en la cadena utilizada
   avgCharWidth = GetFontDialogUnits(frm)
   
  'calcule que el tamaño del dropdown necesita ser acommodado
  'a la cadena mas larga. Aqui Yo resto 2 porque encontre que
  'en my sistema, usando el font fictisio en GetFontDialogUnits,
  'el ancho es solo un bit mas amplio.
   NewDropDownWidth = (LongestComboItem - 2) * avgCharWidth
   
  'ahora cambiamos el tamaño de porcion DropDown del combo box
   r = SendMessageLong(cbo.hwnd, CB_SETDROPPEDWIDTH, NewDropDownWidth, 0)
 
End Sub

' Esta function calcula el ancho promedio de las caracteres de acuerdo
' al tipo de letra del formulario que se entregue
' Written by JANO     invtql@principal.unjbg.edu.pe
Public Function GetFontDialogUnits(frm As Form) As Long

   Dim hFont As Long
   Dim hFontOld As Long
   Dim r As Long
   Dim avgWidth As Long
   Dim hDc As Long
   Dim tmp As String
   Dim sz As SIZE
   
  'guardamos el hdc del formulario
   hDc = GetDC(frm.hwnd)
   
  'recibimos los atributos del tipo de letra actual
   hFont = GetStockObject(ANSI_VAR_FONT)
   hFontOld = SelectObject(hDc, hFont&)
   
  'calculamos su longitud , y calculamos el ancho promedio de cada carater
   tmp = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
   r = GetTextExtentPoint32(hDc, tmp, 52, sz)
   avgWidth = (sz.cx \ 52)
   
  'Seleccionamos otra vez el anterior tipo de letra y eliminamos/liberamos el hdc anterior
   r = SelectObject(hDc, hFontOld)
   r = DeleteObject(hFont)
   r = ReleaseDC(frm.hwnd, hDc)
   
  'retornamos el valor obtenido
   GetFontDialogUnits = avgWidth

End Function

