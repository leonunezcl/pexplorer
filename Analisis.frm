VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnalisis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis del Proyecto"
   ClientHeight    =   6465
   ClientLeft      =   375
   ClientTop       =   2400
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Analisis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   517
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRutinas 
      Caption         =   "Procedimientos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   4560
      TabIndex        =   21
      Top             =   1305
      Width           =   5340
      Begin VB.TextBox txtNumParam 
         Height          =   285
         Left            =   2865
         TabIndex        =   39
         Top             =   2565
         Width           =   855
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Comprobar máximo de parámetros"
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   38
         Top             =   2565
         Width           =   2835
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Lineas de código"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   32
         Top             =   255
         Width           =   1710
      End
      Begin VB.TextBox txtLinRut 
         Height          =   285
         Left            =   1830
         TabIndex        =   31
         Top             =   225
         Width           =   855
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Declarar tipo de retorno de funciones (MyFuncion())."
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   30
         Top             =   510
         Width           =   4170
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Parámetros pasados sin tipo (Byval MyParametro) (MyParametro)."
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   29
         Top             =   735
         Width           =   5145
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Parámetros pasados con ByVal v/s ByRef (MyParametro As String)."
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   28
         Top             =   960
         Width           =   5205
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Control de Errores (On Error Resume Next/On Error Goto ....)"
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   27
         Top             =   1185
         Width           =   5025
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Funciones/Subs públicas en archivos frm."
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   26
         Top             =   1410
         Width           =   3405
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Comentarios"
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   25
         Top             =   1635
         Width           =   2325
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "No usadas."
         Height          =   240
         Index           =   7
         Left            =   75
         TabIndex        =   24
         Top             =   1860
         Width           =   2070
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Vacias (sin código)"
         Height          =   240
         Index           =   8
         Left            =   75
         TabIndex        =   23
         Top             =   2085
         Width           =   1710
      End
      Begin VB.CheckBox chkRuti 
         Caption         =   "Exit (Exit Sub/Exit Function/Exit For/Exit Do/Exit Property)."
         Height          =   240
         Index           =   9
         Left            =   75
         TabIndex        =   22
         Top             =   2310
         Width           =   4650
      End
   End
   Begin VB.Frame fraVariables 
      Caption         =   "Variables/Constantes/Enumeraciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   3495
      TabIndex        =   9
      Top             =   285
      Width           =   4710
      Begin VB.CheckBox chkVar 
         Caption         =   "Analizar otros objetos (Object, Node, ListItem , ...)"
         Height          =   240
         Index           =   21
         Left            =   75
         TabIndex        =   52
         Top             =   5010
         Width           =   4125
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Analizar objetos de ADO"
         Height          =   240
         Index           =   20
         Left            =   75
         TabIndex        =   51
         Top             =   4785
         Width           =   2100
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Analizar objetos de DAO"
         Height          =   240
         Index           =   19
         Left            =   75
         TabIndex        =   50
         Top             =   4560
         Width           =   2100
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Verificar uso de Tipos"
         Height          =   240
         Index           =   18
         Left            =   75
         TabIndex        =   49
         Top             =   4335
         Width           =   1875
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Verificar uso de Enumeraciones"
         Height          =   240
         Index           =   17
         Left            =   75
         TabIndex        =   48
         Top             =   4110
         Width           =   2610
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar visibilidad"
         Height          =   240
         Index           =   16
         Left            =   75
         TabIndex        =   46
         Top             =   3885
         Width           =   1965
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar nomenclatura"
         Height          =   240
         Index           =   15
         Left            =   75
         TabIndex        =   45
         Top             =   3660
         Width           =   2250
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Buscar Goto/Gosub/Return"
         Height          =   240
         Index           =   14
         Left            =   75
         TabIndex        =   44
         Top             =   3435
         Width           =   2610
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Buscar :"
         Height          =   240
         Index           =   13
         Left            =   75
         TabIndex        =   43
         Top             =   3210
         Width           =   1005
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Uso de IIF"
         Height          =   240
         Index           =   12
         Left            =   75
         TabIndex        =   42
         Top             =   2970
         Width           =   2130
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Buscar Stop/Debug/End"
         Height          =   240
         Index           =   11
         Left            =   75
         TabIndex        =   41
         Top             =   2745
         Width           =   2130
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar uso de funciones String V/S Variant"
         Height          =   240
         Index           =   10
         Left            =   75
         TabIndex        =   40
         Top             =   2520
         Width           =   3870
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Largo mínimo de nombre de variable"
         Height          =   240
         Index           =   0
         Left            =   75
         TabIndex        =   20
         Top             =   270
         Width           =   2970
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Declaraciones al viejo estilo basic (Dim MyVar$, I%,J!,K#)."
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   19
         Top             =   495
         Width           =   4530
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Declaración de tipo (Dim MyVar). Se asume variant"
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   18
         Top             =   720
         Width           =   4035
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Private v/s Dim en sección general. Visual Basic > 4."
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   17
         Top             =   945
         Width           =   4380
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar Variables públicas"
         Height          =   240
         Index           =   4
         Left            =   75
         TabIndex        =   16
         Top             =   1170
         Width           =   2745
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar Variables locales Fun/Sub"
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   15
         Top             =   1620
         Width           =   3090
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Comprobar Variables locales"
         Height          =   240
         Index           =   5
         Left            =   75
         TabIndex        =   14
         Top             =   1395
         Width           =   2745
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Public v/s Global en sección general .bas Visual Basic > 4."
         Height          =   240
         Index           =   7
         Left            =   75
         TabIndex        =   13
         Top             =   1845
         Width           =   4455
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Declarar ámbito de constante. (Private/Public)"
         Height          =   240
         Index           =   8
         Left            =   75
         TabIndex        =   12
         Top             =   2070
         Width           =   3705
      End
      Begin VB.CheckBox chkVar 
         Caption         =   "Variables públicas en formularios."
         Height          =   240
         Index           =   9
         Left            =   75
         TabIndex        =   11
         Top             =   2295
         Width           =   2790
      End
      Begin VB.TextBox txtLarvar 
         Height          =   285
         Left            =   3090
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraArchivo 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5340
      Left            =   630
      TabIndex        =   33
      Top             =   870
      Width           =   5430
      Begin VB.CheckBox chkArch 
         Caption         =   "Comprobar si archivo esta siendo usado"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   975
         Width           =   3375
      End
      Begin VB.CheckBox chkArch 
         Caption         =   "Nomenclatura de archivo."
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   2700
      End
      Begin VB.CheckBox chkArch 
         Caption         =   "Nomenclatura de controles estándar + windows 95/98."
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   4350
      End
      Begin VB.CheckBox chkArch 
         Caption         =   "Lineas de código x archivo."
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   735
         Width           =   2325
      End
      Begin VB.TextBox txtLinArch 
         Height          =   285
         Left            =   2490
         TabIndex        =   34
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   810
      TabIndex        =   6
      Top             =   1995
      Width           =   4695
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Option Explicit"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1545
      End
      Begin VB.CheckBox chkGeneral 
         Caption         =   "Comentarios"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   525
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione opciones de análisis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   405
      TabIndex        =   0
      Top             =   15
      Width           =   5835
      Begin MSComctlLib.TabStrip tabAna 
         Height          =   5760
         Left            =   135
         TabIndex        =   5
         Top             =   510
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   10160
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Archivo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "General"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Variables"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Procedimientos"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkTodo 
         Caption         =   "Analizar todo"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   423
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   0
      Width           =   360
   End
   Begin VB.CommandButton cmdAccion 
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
      Height          =   405
      Index           =   1
      Left            =   6315
      TabIndex        =   3
      Top             =   555
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccion 
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
      Height          =   405
      Index           =   0
      Left            =   6300
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Analizo As Boolean
Private mGradient As New clsGradient




Private Sub chkTodo_Click()

    Dim ret As Integer
    Dim k As Integer
    
    ret = chkTodo.Value
    
    'archivo
    For k = 1 To UBound(Ana_Archivo) - 1
        chkArch(k - 1).Value = ret
    Next k
    
    'general
    For k = 1 To UBound(Ana_General)
        chkGeneral(k - 1).Value = ret
    Next k
    
    'variables
    For k = 1 To UBound(Ana_Variables)
        chkVar(k - 1).Value = ret
    Next k
    
    'rutinas
    For k = 1 To UBound(Ana_Rutinas)
        chkRuti(k - 1).Value = ret
    Next k
    
End Sub

Private Sub cmdAccion_Click(Index As Integer)

    If Index = 0 Then
        Call GrabarOpcionesAnalisis
        Analizo = True
    End If
    
    Unload Me
        
End Sub

'grabar las opciones de analisis a archivo .ini
Private Sub GrabarOpcionesAnalisis()

    On Local Error Resume Next
    
    Dim k As Integer
    
    Call GrabaIni(C_INI, "analisis", "lineas_x_archivo", txtLinArch.text)
    glbLinXArch = txtLinArch.text
    
    Call GrabaIni(C_INI, "analisis", "largo_variable", txtLarvar.text)
    glbLarVar = txtLarvar.text
    
    Call GrabaIni(C_INI, "analisis", "lineas_x_rutina", txtLinRut.text)
    glbLinXRuti = txtLinRut.text
    
    'grabar opciones de archivo
    For k = 1 To UBound(Ana_Archivo)
        Call GrabaIni(C_INI, "ana_archivo", CStr(k), chkArch(k - 1).Value)
        Ana_Archivo(k).Value = chkArch(k - 1).Value
    Next k
    
    'opciones de general
    For k = 1 To UBound(Ana_General)
        Call GrabaIni(C_INI, "ana_general", CStr(k), chkGeneral(k - 1).Value)
        Ana_General(k).Value = chkGeneral(k - 1).Value
    Next k
    
    'opciones de variables
    For k = 1 To UBound(Ana_Variables)
        Call GrabaIni(C_INI, "ana_variables", CStr(k), chkVar(k - 1).Value)
        Ana_Variables(k).Value = chkVar(k - 1).Value
    Next k
    
    'opciones de rutinas
    For k = 1 To UBound(Ana_Rutinas)
        Call GrabaIni(C_INI, "ana_rutinas", CStr(k), chkRuti(k - 1).Value)
        Ana_Rutinas(k).Value = chkRuti(k - 1).Value
    Next k
    
    Err = 0
    
End Sub

Private Sub Form_Load()

    On Local Error Resume Next
    
    Dim k As Integer
        
    CenterWindow hwnd
    
    Analizo = False
    
    txtLinArch.text = glbLinXArch
    txtLarvar.text = glbLarVar
    txtLinRut.text = glbLinXRuti
    txtNumParam.text = glbMaxNumParam
    
    'cargar archivo
    For k = 1 To UBound(Ana_Archivo)
        chkArch(k - 1).Value = Ana_Archivo(k).Value
    Next k
    
    'cargar general
    For k = 1 To UBound(Ana_General)
        chkGeneral(k - 1).Value = Ana_General(k).Value
    Next k
    
    'cargar variables
    For k = 1 To UBound(Ana_Variables)
        chkVar(k - 1).Value = Ana_Variables(k).Value
    Next k
    
    'cargar rutinas
    For k = 1 To UBound(Ana_Rutinas)
        chkRuti(k - 1).Value = Ana_Rutinas(k).Value
    Next k
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    'setear frames
    fraGeneral.Left = fraArchivo.Left
    fraGeneral.Top = fraArchivo.Top
    fraGeneral.Height = fraArchivo.Height
    fraGeneral.Width = fraArchivo.Width
    
    fraVariables.Left = fraArchivo.Left
    fraVariables.Top = fraArchivo.Top
    fraVariables.Height = fraArchivo.Height
    fraVariables.Width = fraArchivo.Width
    
    fraRutinas.Left = fraArchivo.Left
    fraRutinas.Top = fraArchivo.Top
    fraRutinas.Height = fraArchivo.Height
    fraRutinas.Width = fraArchivo.Width
    
    fraArchivo.ZOrder 0
    
    Err = 0
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmAnalisis = Nothing
    
End Sub


Private Sub tabAna_Click()

    If tabAna.SelectedItem.Index = 1 Then
        fraArchivo.ZOrder 0
    ElseIf tabAna.SelectedItem.Index = 2 Then
        fraGeneral.ZOrder 0
    ElseIf tabAna.SelectedItem.Index = 3 Then
        fraVariables.ZOrder 0
    Else
        fraRutinas.ZOrder 0
    End If
    
End Sub


