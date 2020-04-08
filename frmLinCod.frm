VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLinCod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lineas de Código"
   ClientHeight    =   4650
   ClientLeft      =   1185
   ClientTop       =   2220
   ClientWidth     =   9525
   Icon            =   "frmLinCod.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   310
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   635
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   6
      Top             =   765
      Width           =   1215
   End
   Begin VB.ComboBox cboOpe 
      Height          =   315
      Left            =   4860
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   15
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   4
      Top             =   330
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4620
      Left            =   0
      ScaleHeight     =   306
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   0
      Width           =   360
   End
   Begin VB.TextBox txtLinCod 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Text            =   "0"
      Top             =   15
      Width           =   750
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   4320
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   7620
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "N°"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parámetros"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Variables"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Digite el número de lineas x procedimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   45
      Width           =   3675
   End
End
Attribute VB_Name = "frmLinCod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Const GWL_STYLE = (-16)
Private Const ES_NUMBER = &H2000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'use this code to make a control named Text1 only accept numeric input
Dim lngHwnd As Long
Dim lngReturn As Long

Private Sub cmd_Click(Index As Integer)

    Dim total As Integer
    Dim k As Integer
    Dim r As Integer
    Dim c As Integer
    Dim Rutina As String
    Dim ObjName As String
    
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    If txtLinCod.text = "0" Then
        MsgBox "Digite el número de lineas.", vbCritical
        Exit Sub
    End If
    
    If cboOpe.ListIndex = -1 Then
        MsgBox "Debe seleccionar operador de búsqueda.", vbCritical
        Exit Sub
    End If
    
    Call Hourglass(hwnd, True)
    Call EnabledControls(Me, False)
    
    lvw.ListItems.Clear
    
    total = txtLinCod.text
    
    cboOpe.AddItem "="
    cboOpe.AddItem ">"
    cboOpe.AddItem "<"
    cboOpe.AddItem ">="
    cboOpe.AddItem "<="
    cboOpe.AddItem "<>"
    
    'ver las coincidencias
    c = 1
    
    If cboOpe.ListIndex = 0 Then    '=
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas = total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    ElseIf cboOpe.ListIndex = 1 Then    '>
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas > total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    ElseIf cboOpe.ListIndex = 2 Then    '<
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas < total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    ElseIf cboOpe.ListIndex = 3 Then    '>=
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas >= total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    ElseIf cboOpe.ListIndex = 4 Then    '<=
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas <= total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    Else
        For k = 1 To UBound(Proyecto.aArchivos)
            ObjName = Proyecto.aArchivos(k).ObjectName
            For r = 1 To UBound(Proyecto.aArchivos(k).aRutinas)
                If Proyecto.aArchivos(k).aRutinas(r).NumeroDeLineas <> total Then
                    Rutina = ObjName & "." & Proyecto.aArchivos(k).aRutinas(r).NombreRutina
                    lvw.ListItems.Add , , c
                    lvw.ListItems(c).SubItems(1) = Rutina
                    lvw.ListItems(c).SubItems(2) = UBound(Proyecto.aArchivos(k).aRutinas(r).Aparams)
                    lvw.ListItems(c).SubItems(3) = Proyecto.aArchivos(k).aRutinas(r).nVariables
                    c = c + 1
                End If
            Next r
        Next k
    End If
    
    Call EnabledControls(Me, True)
    Call Hourglass(hwnd, False)
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hwnd, True)
    Call CenterWindow(hwnd)
    
    lngHwnd = GetWindowLong(txtLinCod.hwnd, GWL_STYLE)
    lngReturn = SetWindowLong(txtLinCod.hwnd, GWL_STYLE, lngHwnd Or ES_NUMBER)

    cboOpe.AddItem "="
    cboOpe.AddItem ">"
    cboOpe.AddItem "<"
    cboOpe.AddItem ">="
    cboOpe.AddItem "<="
    cboOpe.AddItem "<>"
        
    With mGradient
        .Angle = 90
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set mGradient = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmLinCod = Nothing
End Sub


