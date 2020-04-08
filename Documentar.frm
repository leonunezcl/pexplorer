VERSION 5.00
Begin VB.Form frmDocumentar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Documentar Proyecto"
   ClientHeight    =   3015
   ClientLeft      =   5085
   ClientTop       =   4410
   ClientWidth     =   6030
   Icon            =   "Documentar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra 
      Caption         =   "Seleccione itemes a documentar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   405
      TabIndex        =   5
      Top             =   450
      Width           =   4155
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Diccionario de datos"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   22
         Top             =   1245
         Width           =   1830
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Archivos"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   990
         Width           =   945
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "Co&mponentes"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   750
         Width           =   1275
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Referencias"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   495
         Width           =   1410
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Proyecto"
         Height          =   240
         Index           =   0
         Left            =   90
         Picture         =   "Documentar.frx":030A
         TabIndex        =   18
         Top             =   255
         Width           =   1200
      End
      Begin VB.CheckBox chkTodo 
         Caption         =   "T&odo"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   2115
         Width           =   795
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Subs"
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1485
         Width           =   945
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Funciones"
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   15
         Top             =   1725
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Variables"
         Height          =   255
         Index           =   8
         Left            =   2250
         TabIndex        =   14
         Top             =   270
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Constantes"
         Height          =   255
         Index           =   9
         Left            =   2250
         TabIndex        =   13
         Top             =   495
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Enumeraciones"
         Height          =   255
         Index           =   11
         Left            =   2250
         TabIndex        =   12
         Top             =   975
         Width           =   1665
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "Arra&ys"
         Height          =   255
         Index           =   12
         Left            =   2250
         TabIndex        =   11
         Top             =   1200
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "Ap&is"
         Height          =   255
         Index           =   7
         Left            =   90
         TabIndex        =   10
         Top             =   1950
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Tipos"
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   9
         Top             =   735
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "Co&ntroles"
         Height          =   255
         Index           =   13
         Left            =   2250
         TabIndex        =   8
         Top             =   1425
         Width           =   1185
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "&Propiedades"
         Height          =   255
         Index           =   14
         Left            =   2250
         TabIndex        =   7
         Top             =   1650
         Width           =   1425
      End
      Begin VB.CheckBox chkDocu 
         Caption         =   "E&ventos"
         Height          =   255
         Index           =   15
         Left            =   2250
         TabIndex        =   6
         Top             =   1890
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Documentar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   4710
      TabIndex        =   4
      Top             =   555
      Width           =   1215
   End
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
      Height          =   420
      Index           =   1
      Left            =   4710
      TabIndex        =   3
      Top             =   1020
      Width           =   1215
   End
   Begin VB.TextBox txtArchivo 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   75
      Width           =   3345
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   2940
      Left            =   15
      ScaleHeight     =   194
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   0
      Top             =   30
      Width           =   360
   End
   Begin VB.Label lblproyecto 
      AutoSize        =   -1  'True
      Caption         =   "Proyecto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   2
      Top             =   90
      Width           =   765
   End
End
Attribute VB_Name = "frmDocumentar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Path As String

'realiza la documentacion del proyecto
Private Function DocumentarProyecto() As Boolean

    Dim ret As Boolean
    
    ret = False
    
    Path = ConfigurarPath(hwnd)
    
    If Path = "\" Then
        GoTo Salir
    End If
    
    Call Hourglass(hwnd, True)
    
    If Not GenerarIndice(Path) Then 'generar el indice
        GoTo Salir
    End If
    
    If chkDocu(0).Value = 1 Then    'proyecto
        If Not DocumentarAnalisisProyecto(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(1).Value = 1 Then    'referencias
        If Not DocumentarReferencias(Path) Then
            GoTo Salir
        End If
    End If
    
    If chkDocu(2).Value = 1 Then    'componentes
        If Not DocumentarComponentes(Path) Then
            GoTo Salir
        End If
    End If
    
    If chkDocu(3).Value = 1 Then    'archivos
        If Not DocumentarArchivos(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(5).Value = 1 Then    'subs
        If Not DocumentarSubs(Path) Then
            GoTo Salir
        End If
    End If
    
    If chkDocu(6).Value = 1 Then    'funciones
        If Not DocumentarFunciones(Path) Then
            GoTo Salir
        End If
    End If
    
    If chkDocu(7).Value = 1 Then    'apis
        If Not DocumentarApis(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(8).Value = 1 Then    'variables
        If Not DocumentarVariables(Path) Then
            GoTo Salir
        End If
    End If
    
    If chkDocu(9).Value = 1 Then    'constantes
        If Not DocumentarConstantes(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(10).Value = 1 Then    'tipos
        If Not DocumentarTipos(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(11).Value = 1 Then    'enumeraciones
        If Not DocumentarEnumeraciones(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(12).Value = 1 Then    'arreglos
        If Not DocumentarArreglos(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(13).Value = 1 Then    'controles
        If Not DocumentarControles(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(14).Value = 1 Then    'propiedades
        If Not DocumentarPropiedades(Path) Then
            GoTo Salir
        End If
    End If
        
    If chkDocu(15).Value = 1 Then    'eventos
        If Not DocumentarEventos(Path) Then
            GoTo Salir
        End If
    End If
        
    Call Hourglass(hwnd, False)
    
    ret = True
    
Salir:
    DocumentarProyecto = ret
    
End Function

'valida que se selecciono alguna forma de impresion
Private Function Validar() As Boolean

    Dim ret As Boolean
    Dim k As Integer
    
    ret = False
    
    For k = 0 To 15
        If chkDocu(k).Value = 1 Then
            ret = True
            Exit For
        End If
    Next k
            
    Validar = ret
    
End Function

Private Sub chkTodo_Click()

    Dim k As Integer
    
    For k = 0 To 15
        chkDocu(k).Value = chkTodo.Value
    Next k
    
End Sub


Private Sub cmd_Click(Index As Integer)

    Dim Msg As String
    
    If Index = 0 Then
        If Validar() Then
            Msg = "Confirma documentar proyecto."
            If Confirma(Msg) = vbYes Then
                If DocumentarProyecto() Then
                    MsgBox "Proyecto documentado con éxito!", vbInformation
                    On Local Error Resume Next
                    ShellExecute Me.hwnd, vbNullString, Path & "index.html", vbNullString, App.Path & "\", SW_SHOWMAXIMIZED
                    Err = 0
                End If
            End If
        Else
            MsgBox "Debe seleccionar un archivo a imprimir.", vbCritical
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()

    Call Hourglass(hwnd, True)
    
    CenterWindow hwnd
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
            
    txtArchivo.text = MyFuncFiles.ExtractFileName(Proyecto.PathFisico)
    
    Call Hourglass(hwnd, False)
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Set mGradient = Nothing
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set frmImprimir = Nothing
    
End Sub


